# manapool_fulfillment.py
import json
import threading
import csv
import os
from typing import List, Dict, Any

import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import keyring

DEFAULT_SELLER_BASE = "https://manapool.com/api/v1/seller"

KEYRING_SERVICE = "TCG ShipAbility"
MANAPOOL_KEYRING_ACCOUNT = "manapool_api"
CONFIG_FILENAME = "config.json"


def _get_here_path(name: str) -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)


def load_config() -> Dict[str, Any]:
    cfg_path = _get_here_path(CONFIG_FILENAME)
    if not os.path.isfile(cfg_path):
        return {"manapool": {"email": ""}}
    try:
        with open(cfg_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        if not isinstance(loaded, dict):
            return {"manapool": {"email": ""}}
        if "manapool" not in loaded or not isinstance(loaded["manapool"], dict):
            loaded["manapool"] = {"email": ""}
        loaded["manapool"].setdefault("email", "")
        return loaded
    except Exception:
        return {"manapool": {"email": ""}}


def get_saved_mp_api_key() -> str:
    val = keyring.get_password(KEYRING_SERVICE, MANAPOOL_KEYRING_ACCOUNT)
    return (val or "").strip()

def _headers(email: str, api_key: str) -> Dict[str, str]:
    return {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-ManaPool-Email": email,
        "X-ManaPool-Access-Token": api_key,
    }


def _session_with_retries() -> requests.Session:
    from requests.adapters import HTTPAdapter
    try:
        from urllib3.util.retry import Retry
    except Exception:
        from urllib3.util import Retry  # type: ignore

    retry = Retry(
        total=5,
        backoff_factor=0.25,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=frozenset(["GET", "PUT"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry)
    s = requests.Session()
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s


def _put_fulfillment(
    session: requests.Session,
    base_url: str,
    email: str,
    api_key: str,
    order_id: str,
    payload: Dict[str, Any],
) -> requests.Response:
    url = f"{base_url}/orders/{order_id}/fulfillment"
    r = session.put(url, headers=_headers(email, api_key), data=json.dumps(payload), timeout=45)
    return r


def show_fulfillment_window(
    root: tk.Tk,
    rows: List[Dict[str, Any]],
    mp_email: str,
    mp_api_key: str,
    base_url: str = DEFAULT_SELLER_BASE,
    default_status: str = "shipped",
) -> None:
    if not rows:
        messagebox.showinfo("Manapool Fulfillment", "No orders to update.")
        return

    win = tk.Toplevel(root)
    win.title("Manapool Fulfillment Updates")
    win.transient(root)
    win.grab_set()

    win.update_idletasks()
    w = win.winfo_width()
    h = win.winfo_height()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw // 2) - (w // 2)
    y = (sh // 2) - (h // 2)
    win.geometry(f"+{x}+{y}")

    top = tk.Frame(win)
    top.pack(fill="both", expand=True, padx=10, pady=10)

    tree_frame = tk.Frame(top)
    tree_frame.pack(fill="both", expand=True)

    cols = ("seller_label", "customer", "tracking_number", "carrier", "tracking_url", "status")
    tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
    tree.heading("seller_label", text="Seller Label #")
    tree.heading("customer", text="Customer")
    tree.heading("tracking_number", text="Tracking #")
    tree.heading("carrier", text="Carrier")
    tree.heading("tracking_url", text="Tracking URL")
    tree.heading("status", text="Status")

    tree.column("seller_label", width=110, anchor="center")
    tree.column("customer", width=180, anchor="w")
    tree.column("tracking_number", width=140, anchor="w")
    tree.column("carrier", width=100, anchor="center")
    tree.column("tracking_url", width=220, anchor="w")
    tree.column("status", width=90, anchor="center")

    for idx, r in enumerate(rows):
        seller_label = str(r.get("seller_label_number", "") or r.get("manapool.seller_label_number", ""))
        customer = str(r.get("customer_name", "") or r.get("manapool.customer_name", ""))
        tracking_number = str(r.get("tracking_number", "") or "")
        carrier = str(r.get("tracking_company", "") or "")
        tracking_url = str(r.get("tracking_url", "") or "")
        tree.insert(
            "",
            "end",
            iid=str(idx),
            values=(seller_label, customer, tracking_number, carrier, tracking_url, default_status),
        )

    tree.pack(side="left", fill="both", expand=True)
    vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")

    status_var = tk.StringVar(value="Ready.")
    status_bar = tk.Label(top, textvariable=status_var, anchor="w")
    status_bar.pack(fill="x", pady=(6, 0))

    def export_to_csv():
        if not rows:
            messagebox.showinfo("Export CSV", "No rows to export.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Export ManaPool Fulfillment as CSV",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile="manapool_fulfillment.csv",
        )
        if not file_path:
            return

        all_keys = set()
        for row in rows:
            if isinstance(row, dict):
                all_keys.update(row.keys())

        preferred = [
            "mp_order_id",
            "manapool.order_id",
            "seller_label_number",
            "manapool.seller_label_number",
            "customer_name",
            "manapool.customer_name",
            "tracking_company",
            "tracking_number",
            "tracking_url",
        ]
        fieldnames = [k for k in preferred if k in all_keys] + [
            k for k in sorted(all_keys) if k not in preferred
        ]

        try:
            with open(file_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for row in rows:
                    if isinstance(row, dict):
                        writer.writerow(row)
            messagebox.showinfo(
                "Export CSV",
                f"Exported {len(rows)} rows to:\n{file_path}",
            )
        except Exception as ex:
            messagebox.showerror("Export CSV", f"Failed to export CSV:\n{ex}")


    btns = tk.Frame(top)
    btns.pack(fill="x", pady=(8, 0))

    btn_update = tk.Button(btns, text="Update Manapool Fulfillment")
    btn_update.pack(side="left")

    btn_export = tk.Button(btns, text="Export CSV…", command=export_to_csv)
    btn_export.pack(side="left", padx=(8, 0))

    tk.Button(btns, text="Close", command=win.destroy).pack(side="right")

    def _do_update():
        if not mp_email or not mp_api_key:
            messagebox.showerror("Manapool Credentials", "Manapool email and API key are required.")
            return

        def work():
            win.config(cursor="watch")
            btn_update.config(state=tk.DISABLED)
            btn_export.config(state=tk.DISABLED)
            status_var.set("Updating fulfillment on Manapool...")
            win.update_idletasks()

            ok_count = 0
            err_count = 0
            error_rows: List[str] = []

            with _session_with_retries() as session:
                for idx, r in enumerate(rows):
                    order_id = str(
                        r.get("mp_order_id")
                        or r.get("manapool.order_id")
                        or r.get("order_id", "")
                    ).strip()
                    if not order_id:
                        err_count += 1
                        error_rows.append(f"Row {idx}: missing order id")
                        continue

                    tracking_company = str(r.get("tracking_company", "")).strip()
                    tracking_number = str(r.get("tracking_number", "")).strip()
                    tracking_url = str(r.get("tracking_url", "")).strip()

                    payload = {
                        "status": default_status,
                        "tracking_company": tracking_company or None,
                        "tracking_number": tracking_number or None,
                        "tracking_url": tracking_url or None,
                        "in_transit_at": None,
                        "estimated_delivery_at": None,
                        "delivered_at": None,
                    }

                    try:
                        resp = _put_fulfillment(
                            session=session,
                            base_url=base_url,
                            email=mp_email,
                            api_key=mp_api_key,
                            order_id=order_id,
                            payload=payload,
                        )
                        if 200 <= resp.status_code < 300:
                            ok_count += 1
                        else:
                            err_count += 1
                            try:
                                body = resp.json()
                            except Exception:
                                body = resp.text
                            error_rows.append(
                                f"Row {idx}: HTTP {resp.status_code} {resp.reason} - {body}"
                            )
                    except Exception as ex:
                        err_count += 1
                        error_rows.append(f"Row {idx}: {type(ex).__name__}: {ex}")

                    status_var.set(f"Updated {ok_count} order(s); {err_count} error(s)...")
                    win.update_idletasks()

            win.config(cursor="")
            btn_update.config(state=tk.NORMAL)
            btn_export.config(state=tk.NORMAL)

            if err_count == 0:
                status_var.set(f"Done. Updated {ok_count} order(s).")
                messagebox.showinfo(
                    "Manapool Fulfillment",
                    f"Successfully updated {ok_count} order(s) on Manapool.",
                )
            else:
                msg = f"Updated {ok_count} order(s), {err_count} error(s).\n\n"
                msg += "\n".join(error_rows[:10])
                if len(error_rows) > 10:
                    msg += f"\n... and {len(error_rows) - 10} more."
                status_var.set("Completed with errors.")
                messagebox.showwarning("Manapool Fulfillment (Errors)", msg)

        threading.Thread(target=work, daemon=True).start()

    btn_update.config(command=_do_update)


def _launch_fulfillment_from_csv(
    root: tk.Tk,
    email_var: tk.StringVar,
    api_key_var: tk.StringVar,
    base_url_var: tk.StringVar,
) -> None:
    email = email_var.get().strip()
    api_key = api_key_var.get().strip()
    base_url = (base_url_var.get().strip() or DEFAULT_SELLER_BASE).rstrip("/")

    if not email or not api_key:
        messagebox.showwarning(
            "Manapool Credentials",
            "Please enter your ManaPool seller email and API key before loading a CSV.",
        )
        return

    file_path = filedialog.askopenfilename(
        title="Load ManaPool Fulfillment CSV",
        filetypes=[("CSV Files", "*.csv")],
    )
    if not file_path:
        return

    try:
        with open(file_path, "r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows: List[Dict[str, Any]] = [dict(row) for row in reader]

        if not rows:
            messagebox.showinfo("Load CSV", "The selected CSV has no rows.")
            return

        show_fulfillment_window(
            root=root,
            rows=rows,
            mp_email=email,
            mp_api_key=api_key,
            base_url=base_url,
        )
    except Exception as ex:
        messagebox.showerror("Load CSV", f"Failed to load CSV:\n{ex}")


if __name__ == "__main__":
    cfg = load_config()
    cfg_mp = cfg.get("manapool", {}) if isinstance(cfg, dict) else {}
    default_email = cfg_mp.get("email", "")
    default_api_key = get_saved_mp_api_key()

    root = tk.Tk()
    root.title("ManaPool Fulfillment Helper")

    main_frame = tk.Frame(root)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)

    mp_email_var = tk.StringVar(value=default_email)
    mp_api_key_var = tk.StringVar(value=default_api_key)
    base_url_var = tk.StringVar(value=DEFAULT_SELLER_BASE)

    tk.Label(main_frame, text="ManaPool Email:").grid(row=1, column=0, sticky="e", padx=(0, 6), pady=4)
    email_entry = tk.Entry(main_frame, textvariable=mp_email_var, width=40)
    email_entry.grid(row=1, column=1, columnspan=2, sticky="w", pady=4)

    tk.Label(main_frame, text="API Key:").grid(row=2, column=0, sticky="e", padx=(0, 6), pady=4)
    api_entry = tk.Entry(main_frame, textvariable=mp_api_key_var, width=40, show="*")
    api_entry.grid(row=2, column=1, sticky="w", pady=4)

    show_var = tk.BooleanVar(value=False)

    def _toggle_show():
        api_entry.config(show="" if show_var.get() else "*")

    tk.Checkbutton(main_frame, text="Show", variable=show_var, command=_toggle_show).grid(
        row=2, column=2, sticky="w", pady=4
    )

    tk.Label(main_frame, text="Seller API Base URL:").grid(row=3, column=0, sticky="e", padx=(0, 6), pady=4)
    base_entry = tk.Entry(main_frame, textvariable=base_url_var, width=40)
    base_entry.grid(row=3, column=1, columnspan=2, sticky="w", pady=4)

    btn_frame = tk.Frame(main_frame)
    btn_frame.grid(row=4, column=0, columnspan=3, pady=(12, 0), sticky="w")

    tk.Button(
        btn_frame,
        text="Load CSV and Open Fulfillment",
        command=lambda: _launch_fulfillment_from_csv(root, mp_email_var, mp_api_key_var, base_url_var),
    ).pack(side="left")

    tk.Button(btn_frame, text="Quit", command=root.destroy).pack(side="left", padx=(8, 0))

    main_frame.grid_columnconfigure(1, weight=1)

    email_entry.focus_set()
    root.mainloop()
