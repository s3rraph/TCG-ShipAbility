import os
import io
import shutil
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import subprocess
import platform
import keyring

import requests
from PIL import Image
from PyPDF2 import PdfMerger
from easypost import EasyPostClient

APP_TITLE = "EasyPost Label PDF Builder"
TARGET_WIDTH_IN = 4.0
TARGET_HEIGHT_IN = 6.0
TARGET_DPI = 300
CACHE_DIR = "label_cache"

KEYRING_SERVICE = "TCG ShipAbility"
KEYRING_ACCOUNT = "easypost"

def get_saved_api_key():
    val = keyring.get_password(KEYRING_SERVICE, KEYRING_ACCOUNT)
    return (val or "").strip()

def set_saved_api_key(value: str):
    keyring.set_password(KEYRING_SERVICE, KEYRING_ACCOUNT, value.strip())

def delete_saved_api_key():
    try:
        keyring.delete_password(KEYRING_SERVICE, KEYRING_ACCOUNT)
    except Exception:
        pass

def ensure_api_key_or_prompt(root):
    key = get_saved_api_key()
    if key:
        return key

    dlg = tk.Toplevel(root)
    dlg.title("Enter EasyPost API Key")
    dlg.transient(root)
    dlg.grab_set()

    # center on screen
    dlg.update_idletasks()
    w = dlg.winfo_width(); h = dlg.winfo_height()
    sw = dlg.winfo_screenwidth(); sh = dlg.winfo_screenheight()
    dlg.geometry(f"+{(sw // 2 - w // 2)}+{(sh // 2 - h // 2)}")

    ttk.Label(dlg, text="EasyPost API Key:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    v_key = tk.StringVar(value="")
    ent = ttk.Entry(dlg, textvariable=v_key, width=60, show="*")
    ent.grid(row=0, column=1, padx=10, pady=10)
    ent.focus_set()

    v_show = tk.BooleanVar(value=False)
    def toggle_show():
        ent.configure(show="" if v_show.get() else "*")
    ttk.Checkbutton(dlg, text="Show", variable=v_show, command=toggle_show)\
        .grid(row=1, column=1, sticky="w", padx=10)

    result = {"ok": False}
    def on_ok():
        val = v_key.get().strip()
        if not val:
            messagebox.showerror("Missing", "Please enter an API key.")
            return
        set_saved_api_key(val)
        result["ok"] = True
        dlg.destroy()
    def on_cancel():
        dlg.destroy()

    bf = ttk.Frame(dlg)
    bf.grid(row=2, column=0, columnspan=2, pady=10)
    ttk.Button(bf, text="Save", command=on_ok).grid(row=0, column=0, padx=5)
    ttk.Button(bf, text="Cancel", command=on_cancel).grid(row=0, column=1, padx=5)
    dlg.wait_window()

    return get_saved_api_key() if result["ok"] else ""

# ---------------- UTILS ----------------
def parse_ids(text): return [line.strip() for line in text.splitlines() if line.strip()]

def open_pdf(path):
    """Open the generated PDF with the default system viewer."""
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.Popen(["open", path])
        else:  # Linux / other
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showwarning("Open PDF", f"Could not open file automatically:\n{e}")

def download_bytes(url, session=None):
    s = session or requests.Session()
    r = s.get(url, timeout=60); r.raise_for_status()
    return r.content

def ensure_fit(img, max_w, max_h):
    w, h = img.size
    if w <= max_w and h <= max_h: return img
    scale = min(max_w / float(w), max_h / float(h))
    new_size = (max(1, int(w * scale)), max(1, int(h * scale)))
    return img.resize(new_size, Image.LANCZOS)

def png_to_rotated_padded_pil(png_bytes, target_w_in, target_h_in, dpi):
    with Image.open(io.BytesIO(png_bytes)) as img:
        rotated = img.rotate(90, expand=True).convert("RGB")
        target_w_px = int(target_w_in * dpi); target_h_px = int(target_h_in * dpi)
        rotated = ensure_fit(rotated, target_w_px, target_h_px)
        canvas = Image.new("RGB", (target_w_px, target_h_px), "white")
        paste_x = 0
        paste_y = (target_h_px - rotated.height) // 2
        canvas.paste(rotated, (paste_x, paste_y))
        return canvas

def pil_to_pdf_bytes(pil_img, dpi):
    buf = io.BytesIO()
    pil_img.save(buf, format="PDF", resolution=dpi)
    return buf.getvalue()

def merge_pdfs_to_file(pdf_blobs, out_path):
    merger = PdfMerger()
    for blob in pdf_blobs: merger.append(io.BytesIO(blob))
    with open(out_path, "wb") as f: merger.write(f)
    merger.close()

# ---------------- CACHE ----------------
def cache_png_path(shipment_id):
    safe = shipment_id.replace("/", "_")
    return os.path.join(CACHE_DIR, f"{safe}.png")

def ensure_cache_dir():
    if not os.path.isdir(CACHE_DIR): os.makedirs(CACHE_DIR, exist_ok=True)

def try_load_cached_png(shipment_id):
    p = cache_png_path(shipment_id)
    if os.path.isfile(p):
        with open(p, "rb") as f: data = f.read()
        if data[:8] == b"\x89PNG\r\n\x1a\n": return data
    return None

def save_cached_png(shipment_id, content_bytes):
    ensure_cache_dir()
    with open(cache_png_path(shipment_id), "wb") as f: f.write(content_bytes)

# ---------------- CORE ----------------
def extract_label_urls(shipment):
    pl = shipment.get("postage_label")
    if not pl: return (None, None)
    return (pl.get("label_pdf_url"), pl.get("label_url"))

def _fallback_generate_label_rest(api_key: str, shipment_id: str, file_format: str, sess: requests.Session) -> bytes:
    url = f"https://api.easypost.com/v2/shipments/{shipment_id}/label"
    auth = (api_key, "")
    r = sess.get(url, params={"file_format": file_format}, auth=auth, timeout=60)
    if r.status_code >= 400:
        r = sess.post(url, json={"file_format": file_format}, auth=auth, timeout=60)
    r.raise_for_status()
    shp = r.json()
    _pdf_url, png_url = extract_label_urls(shp)
    if not png_url:
        raise RuntimeError(f"Fallback label API didn't return a PNG URL for {shipment_id}.")
    return download_bytes(png_url, sess)

def fetch_or_cache_png_for_shipment(client: EasyPostClient, api_key: str, sid: str, status_cb, sess: requests.Session):
    cached = try_load_cached_png(sid)
    if cached:
        status_cb(f"Using cached PNG for {sid}")
        return cached

    status_cb(f"Generating PNG label for {sid}")
    shp = client.shipment.retrieve(sid)

    png_url = None
    try_methods = []
    if hasattr(client.shipment, "generate_label"):
        try_methods.append(("generate_label", lambda: client.shipment.generate_label(shp["id"], file_format="PNG")))
    if hasattr(client.shipment, "label"):
        try_methods.append(("label", lambda: client.shipment.label(shp["id"], file_format="PNG")))
    if hasattr(client.shipment, "convert_label"):
        try_methods.append(("convert_label", lambda: client.shipment.convert_label(shp["id"], file_format="PNG")))

    for name, func in try_methods:
        try:
            updated = func()
            _, png_url = extract_label_urls(updated)
            if png_url:
                break
        except Exception:
            pass

    if not png_url:
        status_cb("SDK method not available—using REST fallback")
        blob = _fallback_generate_label_rest(api_key, shp["id"], "PNG", sess)
    else:
        blob = download_bytes(png_url, sess)

    if blob[:8] != b"\x89PNG\r\n\x1a\n":
        raise RuntimeError(f"{sid}: expected PNG content, got something else")
    save_cached_png(sid, blob)
    return blob

def build_pdf_from_shipments_multipage(client, api_key, shipment_ids, out_path, status_cb):
    sess = requests.Session(); ensure_cache_dir()
    pdf_blobs = []

    for idx, sid in enumerate(shipment_ids, start=1):
        status_cb(f"[{idx}/{len(shipment_ids)}] {sid} -> page")

        try:
            shp = client.shipment.retrieve(sid)
            parcel = shp.get("parcel", {}) or {}
            predefined = parcel.get("predefined_package", "")
            is_letter = str(predefined).lower() in ("letter", "flat", "envelope")
        except Exception:
            is_letter = True

        png_blob = fetch_or_cache_png_for_shipment(client, api_key, sid, status_cb, sess)

        if is_letter:
            pil_img = png_to_rotated_padded_pil(png_blob, TARGET_WIDTH_IN, TARGET_HEIGHT_IN, TARGET_DPI)
        else:
            pil_img = Image.open(io.BytesIO(png_blob)).convert("RGB")

        pdf_blobs.append(pil_to_pdf_bytes(pil_img, TARGET_DPI))

    status_cb("Merging pages into final PDF …")
    merge_pdfs_to_file(pdf_blobs, out_path)
    status_cb("Done.")

# ---------------- GUI ----------------
class App:
    def __init__(self, root):
        self.root = root; root.title(APP_TITLE)
        frame = ttk.Frame(root, padding=10); frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Paste shipment IDs (one per line):").pack(anchor="w")
        self.txt = tk.Text(frame, width=80, height=18); self.txt.pack(fill="both", expand=True, pady=6)

        btns = ttk.Frame(frame); btns.pack(fill="x", pady=6)
        ttk.Button(btns, text="Build PDF...", command=self.on_build).pack(side="left")
        ttk.Button(btns, text="Clear cache", command=self.on_clear_cache).pack(side="left", padx=8)
        ttk.Button(btns, text="Quit", command=root.destroy).pack(side="right")

        self.status = tk.StringVar(value="Ready.")
        ttk.Label(frame, textvariable=self.status).pack(anchor="w", pady=4)

    def set_status(self, msg): self.status.set(msg); self.root.update_idletasks()

    def on_clear_cache(self):
        if not os.path.isdir(CACHE_DIR):
            messagebox.showinfo("Cache", "Cache is already empty."); return
        if messagebox.askyesno("Clear cache", f"Delete '{CACHE_DIR}' and all cached files?"):
            shutil.rmtree(CACHE_DIR, ignore_errors=True)
            messagebox.showinfo("Cache", "Cache cleared.")

    def on_build(self):
        ids = parse_ids(self.txt.get("1.0", "end"))
        if not ids:
            messagebox.showerror("No IDs", "Please paste at least one shipment ID."); return

        key = ensure_api_key_or_prompt(self.root)
        if not key:
            messagebox.showerror("Missing API Key", "Cannot continue without an API key."); return
        client = EasyPostClient(key)

        default_name = f"labels_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        out_path = filedialog.asksaveasfilename(
            title="Save PDF", defaultextension=".pdf", initialfile=default_name, filetypes=[("PDF", "*.pdf")]
        )
        if not out_path: return

        def work():
            try:
                self.set_status("Starting ...")
                build_pdf_from_shipments_multipage(client, key, ids, out_path, self.set_status)
                self.set_status("Completed.")
                open_pdf(out_path)
            except Exception as ex:
                messagebox.showerror("Error", str(ex))
                self.set_status(f"Failed: {ex}")

        threading.Thread(target=work, daemon=True).start()

def main():
    root = tk.Tk(); root.geometry("980x720"); App(root); root.mainloop()

if __name__ == "__main__":
    main()
