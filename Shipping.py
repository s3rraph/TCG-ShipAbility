import json
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from fetch_lables import build_pdf_from_shipments_multipage, open_pdf
from easypost import EasyPostClient
import traceback
import keyring

KEYRING_SERVICE = "TCG ShipAbility"
KEYRING_ACCOUNT = "easypost"
CONFIG_FILENAME = "config.json"
OLD_CONFIG_FILENAME = "shipping_config.json" 

DEFAULT_CONFIG = {
    "defaults": {
        "carrier": "USPS",
        "service": "First",
        "label_format": "PNG",
        "country": "US"
    },
    "from_address": {
        "name": "",
        "company": "",
        "phone": "",
        "email": "",
        "street1": "",
        "street2": "",
        "city": "",
        "state": "",
        "zip": "",
        "country": ""
    },
    # Rules apply to LETTERS ONLY (packages skip rules)
    # Each rule: rows with item_count <= max_items get these fields set.
    "rules": [
        {"max_items": 7,    "weight_oz": 1, "machinable": True,  "predefined_package": "Letter"},
        {"max_items": 14,    "weight_oz": 2, "machinable": True,  "predefined_package": "Letter"},
        {"max_items": 36,    "weight_oz": 3.5, "machinable": False,  "predefined_package": "Letter"},
        {"max_items": 80,    "weight_oz": 6, "machinable": True,  "predefined_package": "Flat"},
        {"max_items": 9999, "weight_oz": 1, "machinable": True,  "predefined_package": "Package"}
    ],
    # Package detection
    "detection": {
        "manapool_shipping_equals_package": [0, 4.99, 9.99]
    }
    
}

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

def get_here_path(name):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)

def get_config_path():
    return get_here_path(CONFIG_FILENAME)

def get_old_config_path():
    return get_here_path(OLD_CONFIG_FILENAME)

def _deep_merge(base: dict, override: dict) -> dict:
    out = dict(base)
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out

def load_config():
    """
    Load config from config.json; if missing, try shipping_config.json for backward compatibility.
    Always merge over DEFAULT_CONFIG, so new keys appear with sensible defaults.
    """
    cfg = DEFAULT_CONFIG.copy()
    cfg_path = get_config_path()
    old_path = get_old_config_path()

    path_to_use = None
    if os.path.isfile(cfg_path):
        path_to_use = cfg_path
    elif os.path.isfile(old_path):
        path_to_use = old_path

    if path_to_use:
        try:
            with open(path_to_use, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            cfg = _deep_merge(DEFAULT_CONFIG, loaded)
        except Exception:
            messagebox.showwarning("Config", "Could not read config file. Using defaults.")
            cfg = DEFAULT_CONFIG.copy()

    # Normalize and sort rules
    rules = cfg.get("rules", [])
    if not isinstance(rules, list) or len(rules) == 0:
        cfg["rules"] = DEFAULT_CONFIG["rules"]
    else:
        try:
            cfg["rules"] = sorted(rules, key=lambda r: int(r.get("max_items", 0)))
        except Exception:
            pass
    return cfg

def save_config(cfg):
    """Persist to config.json (stateful)."""
    try:
        with open(get_config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
        return True
    except Exception as e:
        messagebox.showerror("Config", f"Failed to save config: {e}")
        return False

def _norm(name: str) -> str:
    return (name.lower()
            .replace(" ", "")
            .replace("_", "")
            .replace(".", "")
            .replace("(", "")
            .replace(")", "")
            .replace("-", ""))

def _normalize_map(cols):
    return {_norm(c): c for c in cols}

# For auto-detection
TCG_SIG = {"firstname", "lastname", "address1", "address2", "city", "state", "postalcode", "country", "itemcount"}
MP_SIG  = {"shippingname", "shippingline1", "shippingline2", "shippingcity", "shippingstate", "shippingzip", "shippingcountry", "itemcount"}

def detect_format_from_headers(df):
    nset = set(_norm(c) for c in df.columns)
    def score(sig, anchors):
        return len(nset & sig) * 2 + sum(1 for a in anchors if a in nset)
    s_tcg = score(TCG_SIG, anchors={"firstname", "lastname", "postalcode"})
    s_mp  = score(MP_SIG, anchors={"shippingname", "shippingzip"})
    if s_tcg == 0 and s_mp == 0:
        return None
    if s_tcg == s_mp:
        if "shippingname" in nset and not ({"firstname", "lastname"} & nset):
            return "Manapool"
        if ({"firstname", "lastname"} & nset) and "shippingname" not in nset:
            return "TCGPlayer"
        return None
    return "TCGPlayer" if s_tcg > s_mp else "Manapool"

def _in_num_list(val, candidates, tol=1e-6):
    try:
        x = float(val)
    except Exception:
        return False
    for c in candidates:
        try:
            y = float(c)
        except Exception:
            continue
        if abs(x - y) <= tol:
            return True
    return False

class CSVConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Shipping Export to Batch CSV Converter")

        self.config = load_config()

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)

        # -------- Convert tab --------
        self.convert_frame = tk.Frame(self.notebook)
        self.notebook.add(self.convert_frame, text="Convert")

        top = tk.Frame(self.convert_frame)
        top.pack(fill="x", padx=10, pady=8)
        tk.Label(top, text="Select Format:").grid(row=0, column=0, sticky="w")
        self.format_var = tk.StringVar(value="Auto")
        tk.Radiobutton(top, text="Auto",      variable=self.format_var, value="Auto").grid(row=1, column=0, sticky="w")
        tk.Radiobutton(top, text="TCGPlayer", variable=self.format_var, value="TCGPlayer").grid(row=1, column=1, sticky="w", padx=(10,0))
        tk.Radiobutton(top, text="Manapool",  variable=self.format_var, value="Manapool").grid(row=1, column=2, sticky="w", padx=(10,0))

        btns = tk.Frame(self.convert_frame)
        btns.pack(fill="x", padx=10)
        self.load_button = tk.Button(btns, text="Load Shipping Export CSV", command=self.load_csv)
        self.load_button.pack(side="left")
        self.save_button = tk.Button(btns, text="Save as Batch CSV", command=self.save_csv, state=tk.DISABLED)
        self.save_button.pack(side="left", padx=(8,0))
        self.buy_button = tk.Button(btns, text="Buy Labels & Build PDF…",
                            command=self.buy_labels_and_build_pdf,
                            state=tk.DISABLED)
        self.buy_button.pack(side="left", padx=(8, 0))

        
        # Preview container + status
        self.preview_container = tk.Frame(self.convert_frame)
        self.preview_container.pack(fill="both", expand=True, padx=10, pady=10)
        self.preview_container.grid_rowconfigure(0, weight=1)
        self.preview_container.grid_columnconfigure(0, weight=1)

        self.status_var = tk.StringVar(value="")
        status_bar = tk.Label(self.convert_frame, textvariable=self.status_var, anchor="w")
        status_bar.pack(fill="x", padx=10, pady=(0,8))

        self.tree = None
        self.data = None
        self.preview_cols = []
        self._is_package_mask = None  # set per file load

        # -------- Settings tab --------
        self.settings_frame = tk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="Settings")
        self._build_settings_ui()

    # ---------- Settings UI + AUTOSAVE ----------
    def _build_settings_ui(self):
        outer = tk.Frame(self.settings_frame)
        outer.pack(fill="both", expand=True, padx=10, pady=10)

        # guard to avoid saving during initial UI construction
        self._loading_settings = True

        lf_from = ttk.LabelFrame(outer, text="From Address")
        lf_from.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0,10))
        fa = self.config["from_address"]
        fields = [
            ("Name","name"),("Company","company"),("Phone","phone"),("Email","email"),
            ("Street1","street1"),("Street2","street2"),("City","city"),
            ("State","state"),("Zip","zip"),("Country","country")
        ]
        self.from_vars = {}
        for i,(lab,key) in enumerate(fields):
            ttk.Label(lf_from, text=lab).grid(row=i, column=0, sticky="e", padx=4, pady=3)
            var = tk.StringVar(value=fa.get(key,""))
            ttk.Entry(lf_from, textvariable=var, width=28).grid(row=i, column=1, sticky="w")
            self.from_vars[key] = var

        # Rules (letters-only)
        lf_rules = ttk.LabelFrame(outer, text="Rules (applied to LETTERS only)")
        lf_rules.grid(row=1, column=0, columnspan=2, sticky="nsew")
        lf_rules.grid_columnconfigure(0, weight=1)
        lf_rules.grid_rowconfigure(1, weight=1)

        table_frame = tk.Frame(lf_rules)
        table_frame.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.rules_tree = ttk.Treeview(
            table_frame,
            columns=("max_items","weight_oz","machinable","predefined_package"),
            show="headings"
        )
        self.rules_tree.heading("max_items", text="Max Items (<=)")
        self.rules_tree.heading("weight_oz", text="Weight (oz)")
        self.rules_tree.heading("machinable", text="Machinable")
        self.rules_tree.heading("predefined_package", text="Predefined Pkg")
        self.rules_tree.column("max_items", anchor="e", width=140)
        self.rules_tree.column("weight_oz", anchor="e", width=110)
        self.rules_tree.column("machinable", anchor="center", width=110)
        self.rules_tree.column("predefined_package", anchor="w", width=150)
        self.rules_tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.rules_tree.yview)
        self.rules_tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        self._refresh_rules_table()

        btns = tk.Frame(lf_rules)
        btns.grid(row=1, column=0, sticky="w", padx=6, pady=(0,6))
        ttk.Button(btns, text="Add Rule", command=self._add_rule_dialog).pack(side="left", padx=(0,6))
        ttk.Button(btns, text="Edit Selected", command=self._edit_selected_rule_dialog).pack(side="left", padx=(0,6))
        ttk.Button(btns, text="Delete Selected", command=self._delete_selected_rule).pack(side="left", padx=(0,6))
        ttk.Button(btns, text="Move Up", command=lambda: self._move_rule(-1)).pack(side="left", padx=(0,6))
        ttk.Button(btns, text="Move Down", command=lambda: self._move_rule(1)).pack(side="left", padx=(0,6))        
        
        lf_key = ttk.LabelFrame(outer, text="EasyPost")
        lf_key.grid(row=3, column=0, columnspan=2, sticky="w", pady=(0,10), padx=(0,0))

        ttk.Button(lf_key, text="Set API Key…", command=self._set_api_key_dialog).grid(row=0, column=0, padx=6, pady=6)


        watch_vars = list(self.from_vars.values())
        
        for v in watch_vars:
            v.trace_add("write", self._autosave_settings)

        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        self._loading_settings = False  # done building

    # --- Autosave helpers ---
    def _parse_numbers(self, s):
        out = []
        for part in s.split(","):
            part = part.strip()
            if not part:
                continue
            try:
                val = float(part)
                if abs(val - int(val)) < 1e-9:
                    val = int(val)
                out.append(val)
            except Exception:
                # ignore invalid fragments
                pass
        return out or []

    def _update_config_from_vars(self):        
        self.config["from_address"] = {k: v.get().strip() for k, v in self.from_vars.items()}        

    def _autosave_settings(self, *args):
        if getattr(self, "_loading_settings", False):
            return
        self._update_config_from_vars()
        save_config(self.config)
        
    def _set_api_key_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Set EasyPost API Key")
        dlg.transient(self.root); dlg.grab_set()

        # Center on screen
        dlg.update_idletasks()
        w = dlg.winfo_width(); h = dlg.winfo_height()
        sw = dlg.winfo_screenwidth(); sh = dlg.winfo_screenheight()
        x = (sw // 2) - (w // 2); y = (sh // 2) - (h // 2)
        dlg.geometry(f"+{x}+{y}")

        current = get_saved_api_key()

        tk.Label(dlg, text="API Key").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        vKey = tk.StringVar(value=current)
        e = ttk.Entry(dlg, textvariable=vKey, width=48, show="*")
        e.grid(row=0, column=1, sticky="w", padx=(0,6))

        vShow = tk.BooleanVar(value=False)
        def toggle_show():
            e.configure(show="" if vShow.get() else "*")
        ttk.Checkbutton(dlg, text="Show", variable=vShow, command=toggle_show)\
            .grid(row=1, column=1, sticky="w", padx=(0,6))

        def on_ok():
            val = vKey.get().strip()
            if val:
                set_saved_api_key(val)
                messagebox.showinfo("EasyPost", "API key saved securely to your system keychain.")
            else:
                delete_saved_api_key()
                messagebox.showinfo("EasyPost", "API key cleared from keychain.")
            dlg.destroy()

        ttk.Button(dlg, text="OK", command=on_ok).grid(row=2, column=0, padx=6, pady=10)
        ttk.Button(dlg, text="Cancel", command=dlg.destroy).grid(row=2, column=1, padx=6, pady=10, sticky="e")
        e.focus_set()


    # ---------- Rules table ops (auto-save on changes) ----------
    def _refresh_rules_table(self):
        for i in self.rules_tree.get_children():
            self.rules_tree.delete(i)
        for r in self.config.get("rules", []):
            self.rules_tree.insert("", "end", values=(
                r.get("max_items",""),
                r.get("weight_oz",""),
                "True" if r.get("machinable", False) else "False",
                r.get("predefined_package","")
            ))

    def _add_rule_dialog(self):
        self._rule_dialog("Add Rule")

    def _edit_selected_rule_dialog(self):
        sel = self.rules_tree.selection()
        if not sel:
            messagebox.showinfo("Edit Rule", "Select a rule to edit.")
            return
        idx = self.rules_tree.index(sel[0])
        cur = self.config["rules"][idx]
        self._rule_dialog("Edit Rule", current=cur, index=idx)

    def _rule_dialog(self, title, current=None, index=None):
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.transient(self.root)
        dlg.grab_set()

        v_max = tk.StringVar(value=str(current.get("max_items","")) if current else "")
        v_wt  = tk.StringVar(value=str(current.get("weight_oz","")) if current else "")
        v_mach = tk.BooleanVar(value=bool(current.get("machinable", True)) if current else True)
        v_pkg  = tk.StringVar(value=str(current.get("predefined_package","Letter")) if current else "Letter")

        def row(r, label, widget):
            ttk.Label(dlg, text=label).grid(row=r, column=0, sticky="e", padx=6, pady=6)
            widget.grid(row=r, column=1, sticky="w")

        e1 = ttk.Entry(dlg, textvariable=v_max, width=18)
        e2 = ttk.Entry(dlg, textvariable=v_wt, width=18)
        cb = ttk.Checkbutton(dlg, text="Machinable", variable=v_mach)
        e3 = ttk.Entry(dlg, textvariable=v_pkg, width=18)

        row(0, "Max Items (<=)", e1)
        row(1, "Weight (oz)", e2)
        row(2, "", cb)
        row(3, "Predefined Pkg", e3)

        def on_ok():
            try:
                mi = int(v_max.get())
                wt = float(v_wt.get())
            except ValueError:
                messagebox.showerror("Invalid Input", "Max Items must be an integer and Weight(oz) must be a number.")
                return
            new_rule = {
                "max_items": mi,
                "weight_oz": wt,
                "machinable": bool(v_mach.get()),
                "predefined_package": v_pkg.get().strip()
            }
            if current is None:
                self.config["rules"].append(new_rule)
            else:
                self.config["rules"][index] = new_rule
            self.config["rules"].sort(key=lambda r: int(r["max_items"]))
            self._refresh_rules_table()
            save_config(self.config)  # auto-save
            dlg.destroy()

        ttk.Button(dlg, text="OK", command=on_ok).grid(row=4, column=0, padx=6, pady=10)
        ttk.Button(dlg, text="Cancel", command=dlg.destroy).grid(row=4, column=1, padx=6, pady=10, sticky="e")
        e1.focus_set()

    def _delete_selected_rule(self):
        sel = self.rules_tree.selection()
        if not sel:
            return
        idx = self.rules_tree.index(sel[0])
        del self.config["rules"][idx]
        self._refresh_rules_table()
        save_config(self.config)  # auto-save

    def _move_rule(self, direction):
        sel = self.rules_tree.selection()
        if not sel:
            return
        idx = self.rules_tree.index(sel[0])
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.config["rules"]):
            return
        rules = self.config["rules"]
        rules[idx], rules[new_idx] = rules[new_idx], rules[idx]
        self._refresh_rules_table()
        try:
            self.rules_tree.selection_set(self.rules_tree.get_children()[new_idx])
        except Exception:
            pass
        save_config(self.config)  # auto-save

    # ---------- Convert logic ----------
    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files","*.csv")])
        if not file_path:
            return
        try:
            df = pd.read_csv(file_path, dtype=str)

            # Decide format
            format_type = self.format_var.get()
            if format_type == "Auto":
                guessed = detect_format_from_headers(df)
                if not guessed:
                    messagebox.showerror(
                        "Unrecognized CSV",
                        "Could not determine CSV type from headers.\n\n"
                        "TCGplayer expects: FirstName, LastName, Address1, Address2, City, State, PostalCode, Country, Item Count\n"
                        "Manapool expects: shipping_name, shipping_line1, shipping_line2, shipping_city, shipping_state, shipping_zip, shipping_country, item_count"
                    )
                    return
                format_type = guessed  # silent

            norm_map = _normalize_map(df.columns)
            dflts = self.config["defaults"]
            from_addr = self.config["from_address"]

            if format_type == "TCGPlayer":
                def col(key, fallback):
                    return df[norm_map.get(key, fallback)]
                df["Item Count"] = col("itemcount", "Item Count").astype(int)
                df["PostalCode"] = col("postalcode", "PostalCode").astype(str)
                card_count = df["Item Count"]

                first = col("firstname","FirstName").fillna("")
                last  = col("lastname","LastName").fillna("")
                addr2 = df.get(norm_map.get("address2","Address2"), pd.Series([""]*len(df))).fillna("")

                self.data = pd.DataFrame({
                    "to_address.name": (first + " " + last).str.strip(),
                    "to_address.company": "",
                    "to_address.phone": "",
                    "to_address.email": "",
                    "to_address.street1": col("address1","Address1").fillna(""),
                    "to_address.street2": addr2,
                    "to_address.city": col("city","City").fillna(""),
                    "to_address.state": col("state","State").fillna(""),
                    "to_address.zip": df["PostalCode"].str.strip(),
                    "to_address.country": df.get(norm_map.get("country","Country"), pd.Series([""]*len(df))).fillna(dflts["country"]),
                })

                # ---- TCGplayer package detection: string compare on "Product Weight" ----
                if "Product Weight" in df.columns:
                    is_package = df["Product Weight"].astype(str).str.strip().eq("0.00")
                else:
                    is_package = pd.Series([False] * len(df), index=df.index)

                # ✅ APPLY rules + package handling
                self.apply_rules_and_package_logic(card_count, is_package)

            elif format_type == "Manapool":
                def col(key, fallback):
                    return df[norm_map.get(key, fallback)]
                df["item_count"] = col("itemcount","item_count").astype(int)
                df["shipping_zip"] = col("shippingzip","shipping_zip").astype(str)
                card_count = df["item_count"]
                
                # Sort by seller_label_number if present
                seller_label_col = norm_map.get("sellerlabelnumber", "seller_label_number")
                if seller_label_col in df.columns:
                    try:
                        df["seller_label_number"] = df[seller_label_col].astype(float)
                        df = df.sort_values(by="seller_label_number", ascending=True)
                    except Exception:
                        df = df.sort_values(by=seller_label_col, ascending=True)

                addr2 = df.get(norm_map.get("shippingline2","shipping_line2"), pd.Series([""]*len(df))).fillna("")
                self.data = pd.DataFrame({
                    "to_address.name": col("shippingname","shipping_name").fillna(""),
                    "to_address.company": "",
                    "to_address.phone": "",
                    "to_address.email": "",
                    "to_address.street1": col("shippingline1","shipping_line1").fillna(""),
                    "to_address.street2": addr2,
                    "to_address.city": col("shippingcity","shipping_city").fillna(""),
                    "to_address.state": col("shippingstate","shipping_state").fillna(""),
                    "to_address.zip": df["shipping_zip"].str.strip(),
                    "to_address.country": col("shippingcountry","shipping_country").fillna(dflts["country"]),
                })

                # Package detection by 'shipping' price (handles variants)
                det = self.config.get("detection", {})
                mp_pkg_triggers = det.get("manapool_shipping_equals_package", [0, 4.99, 9.99])
                ship_col = None
                for k in ("shipping", "shippingprice", "shipping_total", "shippingtotal", "shippingamount"):
                    if k in norm_map:
                        ship_col = norm_map[k]
                        break
                if ship_col is not None:
                    is_package = df[ship_col].fillna("0").apply(lambda v: _in_num_list(v, mp_pkg_triggers))
                else:
                    is_package = pd.Series([False] * len(df))

                self.apply_rules_and_package_logic(card_count, is_package)
            else:
                raise ValueError("Unknown format type.")

            # Add static fields (do NOT set 'service' here; we set per-row below)
            self.data = self.data.assign(
                **{
                    "from_address.name": from_addr.get("name", ""),
                    "from_address.company": from_addr.get("company", ""),
                    "from_address.phone": from_addr.get("phone", ""),
                    "from_address.email": from_addr.get("email", ""),
                    "from_address.street1": from_addr.get("street1", ""),
                    "from_address.street2": from_addr.get("street2", ""),
                    "from_address.city": from_addr.get("city", ""),
                    "from_address.state": from_addr.get("state", ""),
                    "from_address.zip": from_addr.get("zip", ""),
                    "from_address.country": from_addr.get("country", dflts["country"]),
                    "carrier": dflts.get("carrier", "USPS"),
                    "options.label_format": dflts.get("label_format", "PNG"),
                }
            )

            # Apply per-row service: GroundAdvantage for packages; default for letters
            self._apply_service_per_row()

            self.display_preview()
            self._update_save_state()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV: {e}")

    def apply_rules_and_package_logic(self, card_count_series: pd.Series, is_package_mask: pd.Series):
        """
        LETTER rows: apply the first matching rule (<= max_items).
          - If rule.predefined_package == "Package" (case-insensitive), convert to PACKAGE:
              * mark as package (mask=True)
              * clear parcel.predefined_package (packages have no predefined package)
              * clear L/W/H/Weight (manual entry required)
              * clear options.machinable (not applicable)
          - Else (a letter rule like "Letter"): set weight/machinable/predefined_package and clear L/W/H.
        PACKAGE rows (from detection or converted by rule) remain with blanks for predefined & dims & weight.
        """
        # Start with detected mask, allow rules to promote letters to packages
        pkg_mask = is_package_mask.copy()

        # Ensure columns exist
        needed = [
            "parcel.length", "parcel.width", "parcel.height",
            "parcel.predefined_package", "parcel.weight",
            "options.machinable"
        ]
        for col in needed:
            if col not in self.data.columns:
                self.data[col] = ""

        # Initialize blanks for everyone; rules will fill for letters
        self.data["parcel.length"] = ""
        self.data["parcel.width"] = ""
        self.data["parcel.height"] = ""
        self.data["parcel.predefined_package"] = ""
        self.data["parcel.weight"] = ""
        self.data["options.machinable"] = ""

        rules = self.config["rules"]
        remaining_letter = ~pkg_mask  # only rows not already packages are considered letters initially

        def _apply_letter_rule(hit_idx, rule):
            # If rule says "Package", convert these hits to packages
            want_pkg = str(rule.get("predefined_package", "")).strip().lower() == "package"
            if want_pkg:
                pkg_mask.loc[hit_idx] = True
                self.data.loc[hit_idx, ["parcel.predefined_package",
                                        "parcel.length", "parcel.width", "parcel.height",
                                        "parcel.weight"]] = ""
                self.data.loc[hit_idx, "options.machinable"] = ""
            else:
                # normal letter rule
                self.data.loc[hit_idx, "parcel.weight"] = str(rule.get("weight_oz", ""))
                self.data.loc[hit_idx, "options.machinable"] = str(bool(rule.get("machinable", False)))
                self.data.loc[hit_idx, "parcel.predefined_package"] = str(rule.get("predefined_package", ""))
                self.data.loc[hit_idx, ["parcel.length", "parcel.width", "parcel.height"]] = ""

        # Apply rules in order to remaining letters
        for r in rules:
            try:
                threshold = int(r.get("max_items", 0))
            except Exception:
                continue
            hit = remaining_letter & (card_count_series <= threshold)
            if not hit.any():
                continue
            _apply_letter_rule(hit, r)
            remaining_letter = remaining_letter & ~hit

        # Fallback: if any letters still unmatched, apply last rule to them
        if remaining_letter.any() and len(rules) > 0:
            _apply_letter_rule(remaining_letter, rules[-1])

        # Ensure all package rows (detected or converted) have blanks that force manual entry
        self.data.loc[pkg_mask, ["parcel.predefined_package", "parcel.length",
                                 "parcel.width", "parcel.height", "parcel.weight"]] = ""

        # Expose final package mask for service assignment + preview highlighting
        self._is_package_mask = pkg_mask

    def _apply_service_per_row(self):
        if "service" not in self.data.columns:
            self.data["service"] = ""

        default_service = self.config["defaults"].get("service", "First")
        package_service = "GroundAdvantage"

        if self._is_package_mask is None:
            self.data["service"] = default_service
            return

        pkg_mask = self._is_package_mask
        letter_mask = ~pkg_mask
        self.data.loc[letter_mask, "service"] = default_service
        self.data.loc[pkg_mask,   "service"] = package_service

    # ---------- Preview (with inline package editor + highlighting) ----------
    def display_preview(self):
        for w in self.preview_container.winfo_children():
            w.destroy()

        # hide from_address.* and to_address.company/phone/email
        skip_cols = {"to_address.company", "to_address.phone", "to_address.email"}
        preview_df = self.data[[c for c in self.data.columns if not c.startswith("from_address.") and c not in skip_cols]]

        # reorder columns for readability
        order = []
        order += [c for c in preview_df.columns if c.startswith("to_address.")]
        order += [c for c in preview_df.columns if c == "parcel.predefined_package"]
        order += [c for c in preview_df.columns if c in ("parcel.length", "parcel.width", "parcel.height", "parcel.weight")]
        order += [c for c in preview_df.columns if c.startswith("parcel.") and c not in ("parcel.length","parcel.width","parcel.height","parcel.weight","parcel.predefined_package")]
        order += [c for c in preview_df.columns if c.startswith("options.")]
        order += [c for c in preview_df.columns if c in ("carrier", "service")]
        order += [c for c in preview_df.columns if c not in order]
        preview_df = preview_df[order]
        self.preview_cols = list(preview_df.columns)

        tree_frame = tk.Frame(self.preview_container)
        tree_frame.grid(row=0, column=0, sticky="nsew")
        self.preview_container.grid_rowconfigure(0, weight=1)
        self.preview_container.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_frame, columns=self.preview_cols, show="headings")
        for col in self.preview_cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140, stretch=True)

        try:
            self.tree.tag_configure("needs_dims", background="#ffe6e6")  # soft red
        except Exception:
            pass

        # insert rows with per-row tag based on completeness
        for idx, row in preview_df.iterrows():
            tags = ()
            if self._row_needs_dims(idx):
                tags = ("needs_dims",)
            self.tree.insert("", "end", iid=str(idx), values=list(row), tags=tags)

        self.tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        # Edit controls and bindings
        controls = tk.Frame(self.preview_container)
        controls.grid(row=1, column=0, sticky="w", padx=0, pady=(8, 0))
        tk.Button(controls, text="Edit Row…", command=self._edit_selected_package).pack(side="left")

        self.tree.bind("<Double-1>", self._on_tree_double_click)

        self._update_save_state()

    # ----- Row completeness / highlighting helpers -----
    def _is_package_row(self, idx: int) -> bool:
        try:
            # primary source of truth is the mask set by apply_rules_and_package_logic
            if self._is_package_mask is not None:
                return bool(self._is_package_mask.loc[idx])
            # fallback if mask missing
            return str(self.data.at[idx, "parcel.predefined_package"]) == ""
        except Exception:
            return False

    def _row_needs_dims(self, idx: int) -> bool:
        """Return True if row is a package AND any of L/W/H/Weight is missing or non-positive."""
        if self._is_package_mask is not None and not bool(self._is_package_mask.loc[idx]):
            return False
        if not self._is_package_row(idx):
            return False
        for col in ("parcel.length", "parcel.width", "parcel.height", "parcel.weight"):
            try:
                val = str(self.data.at[idx, col]).strip()
            except Exception:
                return True
            if val == "":
                return True
            try:
                if float(val) <= 0:
                    return True
            except Exception:
                return True
        return False

    def _packages_missing_count(self) -> int:
        if self.data is None:
            return 0
        missing = 0
        for idx in self.data.index:
            if self._row_needs_dims(idx):
                missing += 1
        return missing

    def _update_row_tag(self, idx: int):
        item_id = str(idx)
        if not self.tree.exists(item_id):
            return
        tags = list(self.tree.item(item_id, "tags"))
        tags = [t for t in tags if t != "needs_dims"]
        if self._row_needs_dims(idx):
            tags.append("needs_dims")
        self.tree.item(item_id, tags=tuple(tags))

    def _update_save_state(self):
        missing = self._packages_missing_count()
        if missing > 0:
            self.status_var.set(f"{missing} package row(s) need L/W/H/Weight before export.")
            self.save_button.config(state=tk.DISABLED)
        else:
            self.save_button.config(state=tk.NORMAL)
        
        if missing > 0 or self.data is None or len(self.data) == 0:
            self.buy_button.config(state=tk.DISABLED)
        else:
            self.buy_button.config(state=tk.NORMAL)

    # ---------- Editing packages inline ----------
    def _on_tree_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        try:
            idx = int(item)
        except Exception:
            return
        self._edit_row_by_index(idx)

    def _edit_selected_package(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Edit Package", "Select a package row to edit.")
            return
        try:
            idx = int(sel[0])
        except Exception:
            messagebox.showerror("Edit Package", "Could not determine selected row.")
            return
        self._edit_row_by_index(idx)

    def _edit_row_by_index(self, idx: int):
        """
        Unified editor:
          - For Packages: edit Length, Width, Height, Weight (oz)
          - For Letters:  edit Weight (oz) and Machinable (checkbox) to override rule output
        """
        is_pkg = self._is_package_row(idx)

        if is_pkg:
            cur_L  = str(self.data.at[idx, "parcel.length"]) if "parcel.length" in self.data.columns else ""
            cur_W  = str(self.data.at[idx, "parcel.width"]) if "parcel.width" in self.data.columns else ""
            cur_H  = str(self.data.at[idx, "parcel.height"]) if "parcel.height" in self.data.columns else ""
            cur_Wt = str(self.data.at[idx, "parcel.weight"]) if "parcel.weight" in self.data.columns else ""

            dlg = tk.Toplevel(self.root)
            dlg.title("Edit Package Dimensions")
            dlg.transient(self.root); dlg.grab_set()
            dlg.update_idletasks()
            w = dlg.winfo_width()
            h = dlg.winfo_height()
            sw = dlg.winfo_screenwidth()
            sh = dlg.winfo_screenheight()
            x = (sw // 2) - (w // 2)
            y = (sh // 2) - (h // 2)
            dlg.geometry(f"+{x}+{y}")

            vL  = tk.StringVar(value=cur_L)
            vW  = tk.StringVar(value=cur_W)
            vH  = tk.StringVar(value=cur_H)
            vWt = tk.StringVar(value=cur_Wt)

            def row(r, label, var, width=10):
                ttk.Label(dlg, text=label).grid(row=r, column=0, sticky="e", padx=6, pady=6)
                e = ttk.Entry(dlg, textvariable=var, width=width)
                e.grid(row=r, column=1, sticky="w"); return e

            eL  = row(0, "Length",      vL)
            eW  = row(1, "Width",       vW)
            eH  = row(2, "Height",      vH)
            eWt = row(3, "Weight (oz)", vWt)

            def _pos(s):
                try:
                    return float(s.strip()) > 0
                except Exception:
                    return False

            def on_ok():
                if not (_pos(vL.get()) and _pos(vW.get()) and _pos(vH.get()) and _pos(vWt.get())):
                    messagebox.showerror("Invalid Input", "Enter positive numbers for L, W, H, and Weight (oz).")
                    return
                self.data.at[idx, "parcel.length"] = vL.get().strip()
                self.data.at[idx, "parcel.width"]  = vW.get().strip()
                self.data.at[idx, "parcel.height"] = vH.get().strip()
                self.data.at[idx, "parcel.weight"] = vWt.get().strip()
                self._refresh_preview_row(idx)
                self._update_row_tag(idx)
                self._update_save_state()
                dlg.destroy()

            ttk.Button(dlg, text="OK", command=on_ok).grid(row=4, column=0, padx=6, pady=10)
            ttk.Button(dlg, text="Cancel", command=dlg.destroy).grid(row=4, column=1, padx=6, pady=10, sticky="e")
            eL.focus_set()
            return

        # LETTER editor: weight + machinable override
        cur_Wt = str(self.data.at[idx, "parcel.weight"]) if "parcel.weight" in self.data.columns else ""
        raw_m  = str(self.data.at[idx, "options.machinable"]) if "options.machinable" in self.data.columns else ""
        cur_M  = (raw_m.strip().lower() in ("true", "t", "1", "yes", "y"))

        dlg = tk.Toplevel(self.root)
        dlg.title("Edit Letter Options")
        dlg.transient(self.root); dlg.grab_set()
        
        dlg.update_idletasks()
        w = dlg.winfo_width()
        h = dlg.winfo_height()
        sw = dlg.winfo_screenwidth()
        sh = dlg.winfo_screenheight()
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        dlg.geometry(f"+{x}+{y}")

        vWt = tk.StringVar(value=cur_Wt)
        vMach = tk.BooleanVar(value=cur_M)

        ttk.Label(dlg, text="Weight (oz)").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        eWt = ttk.Entry(dlg, textvariable=vWt, width=10)
        eWt.grid(row=0, column=1, sticky="w")

        cb = ttk.Checkbutton(dlg, text="Machinable", variable=vMach)
        cb.grid(row=1, column=1, sticky="w", padx=6, pady=6)

        def _pos(s):
            try:
                return float(s.strip()) > 0
            except Exception:
                return False

        def on_ok():
            if not _pos(vWt.get()):
                messagebox.showerror("Invalid Input", "Enter a positive number for Weight (oz).")
                return
            # Persist overrides directly into row: these are consumed by _mk_parcel/_mk_options at buy/export time
            self.data.at[idx, "parcel.weight"] = vWt.get().strip()
            self.data.at[idx, "options.machinable"] = "True" if vMach.get() else "False"
            self._refresh_preview_row(idx)
            self._update_row_tag(idx)
            self._update_save_state()
            dlg.destroy()

        ttk.Button(dlg, text="OK", command=on_ok).grid(row=2, column=0, padx=6, pady=10)
        ttk.Button(dlg, text="Cancel", command=dlg.destroy).grid(row=2, column=1, padx=6, pady=10, sticky="e")
        eWt.focus_set()


    def _refresh_preview_row(self, idx: int):
        if not hasattr(self, "preview_cols"):
            return
        values = []
        for col in self.preview_cols:
            values.append("" if col not in self.data.columns else self.data.at[idx, col])
        item_id = str(idx)
        if self.tree.exists(item_id):
            self.tree.item(item_id, values=values)

    # ---------- Save ----------
    def save_csv(self):
        missing = self._packages_missing_count()
        if missing > 0:
            messagebox.showerror(
                "Cannot Export",
                f"{missing} package row(s) are missing L/W/H/Weight. Please fill them before exporting."
            )
            return

        if self.data is None:
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files","*.csv")])
        if not file_path:
            return
        try:
            self.data.to_csv(file_path, index=False)
            messagebox.showinfo("Success", f"Batch CSV saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save CSV: {e}")            

    def _mk_address(self, row, prefix):
        addr = {
            "name":    str(row.get(prefix + "name", "")).strip(),
            "company": str(row.get(prefix + "company", "")).strip(),
            "phone":   str(row.get(prefix + "phone", "")).strip(),
            "email":   str(row.get(prefix + "email", "")).strip(),
            "street1": str(row.get(prefix + "street1", "")).strip(),
            "street2": str(row.get(prefix + "street2", "")).strip(),
            "city":    str(row.get(prefix + "city", "")).strip(),
            "state":   str(row.get(prefix + "state", "")).strip().upper(),
            "zip":     str(row.get(prefix + "zip", "")).strip(),   # ALWAYS string
            "country": (str(row.get(prefix + "country", "")).strip() or "US").upper(),
        }
        # drop blank fields
        return {k: v for k, v in addr.items() if v not in ("", None)}

    def _f_or_none(self, v):
        s = str(v).strip()
        if s == "":
            return None
        try:
            return float(s)
        except Exception:
            return None

    def _mk_parcel(self, row):
        predef = str(row.get("parcel.predefined_package", "")).strip()
        wt     = self._f_or_none(row.get("parcel.weight", ""))
        if predef:
            # LETTER: predefined_package + weight (oz). No dimensions.
            out = {"predefined_package": predef}
            if wt is not None:
                out["weight"] = wt
            return out
        # PACKAGE: L/W/H + weight, all numeric
        L = self._f_or_none(row.get("parcel.length", ""))
        W = self._f_or_none(row.get("parcel.width", ""))
        H = self._f_or_none(row.get("parcel.height", ""))
        out = {}
        if L is not None: out["length"] = L
        if W is not None: out["width"]  = W
        if H is not None: out["height"] = H
        if wt is not None: out["weight"] = wt
        return out

    def _mk_options(self, row):
        label_format = (str(row.get("options.label_format", "PNG")).strip() or "PNG")
        mach_raw = str(row.get("options.machinable", "")).strip().lower()
        opts = {"label_format": label_format}
        if mach_raw in ("true", "t", "1", "yes", "y"):
            opts["machinable"] = True
        elif mach_raw in ("false", "f", "0", "no", "n"):
            opts["machinable"] = False
        return opts

    def _row_to_shipment_create(self, row):
        to_addr   = self._mk_address(row, "to_address.")
        from_addr = self._mk_address(row, "from_address.")
        parcel    = self._mk_parcel(row)
        options   = self._mk_options(row)

        # Remove empty dicts
        payload = {"to_address": to_addr, "from_address": from_addr, "parcel": parcel}
        if options:
            payload["options"] = options
        return payload

    def _explain_easypost_error(self, ex):
        """
        Print everything useful from an EasyPost error so we can see
        which field(s) were invalid.
        """
        try:
            print("EasyPost error class:", type(ex).__name__)
            # common attributes across versions
            msg = getattr(ex, "message", None) or str(ex)
            code = getattr(ex, "code", None)
            http_status = getattr(ex, "http_status", None)
            errors = getattr(ex, "errors", None) or getattr(ex, "error", None)
            json_body = getattr(ex, "json_body", None) or getattr(ex, "body", None)
            http_body = getattr(ex, "http_body", None)

            print("message:", msg)
            if code is not None: print("code:", code)
            if http_status is not None: print("http_status:", http_status)
            if errors:
                print("errors:")
                try:
                    import json as _json
                    print(_json.dumps(errors, indent=2))
                except Exception:   
                    print(errors)
            if json_body:
                print("json_body:")
                try:
                    import json as _json
                    print(_json.dumps(json_body, indent=2))
                except Exception:
                    print(json_body)
            if http_body and not json_body:
                # sometimes only raw body is present
                print("http_body:", http_body)
        except Exception as _:
            pass
        
        
    def buy_labels_and_build_pdf(self):
        """
        Create & buy EasyPost labels for the current preview rows,
        then fetch/rotate/pad/merge into a single PDF.
        SDK-compatible across singular/plural services + kwargs/positional.
        """
        if self.data is None or len(self.data) == 0:
            messagebox.showerror("No Data", "Load and prepare your orders first.")
            return

        missing = self._packages_missing_count()
        if missing > 0:
            messagebox.showerror(
                "Cannot Buy",
                f"{missing} package row(s) are missing L/W/H/Weight. Please fill them before buying."
            )
            return

        pdf_path = filedialog.asksaveasfilename(
            title="Save Merged Label PDF",
            defaultextension=".pdf",
            initialfile="labels.pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not pdf_path:
            print("User canceled Save As dialog; aborting buy.")
            return

        key = get_saved_api_key()
        if not key:
            messagebox.showerror("Missing API Key", "Cannot continue without an EasyPost API key.")
            print("Missing EasyPost API key; aborting buy.")
            return

        client = EasyPostClient(key)

        # ---- SDK compatibility helpers ----
        svc = getattr(client, "shipments", None) or getattr(client, "shipment", None)
        if svc is None:
            messagebox.showerror("EasyPost SDK", "Could not locate EasyPost shipment service on client.")
            return

        def _create(payload):
            """Create a shipment following official EasyPost SDK signature."""
            # EasyPost Python SDK expects named keyword args, not a dict wrapper.
            print("Calling EasyPost shipment.create(...) with keyword args:")
            print(json.dumps(payload, indent=2, default=str))
            return svc.create(
                from_address=payload["from_address"],
                to_address=payload["to_address"],
                parcel=payload["parcel"],
                options=payload.get("options", None),
            )
        
        def _buy(sid, rate_id):
            """Buy a shipment across SDK variants."""
            try:
                # current SDK: pass a rate object
                return svc.buy(sid, rate={"id": rate_id})
            except TypeError:
                pass
            try:
                # kwargs style
                return svc.buy(shipment_id=sid, rate={"id": rate_id})
            except TypeError:
                pass
            # very old style fallback: sometimes accepted a plain rate_id
            return svc.buy(sid, rate_id)

        # Determine processing order (mirror the preview table)
        tree_ids = self.tree.get_children() if self.tree is not None else []
        order = [int(i) for i in tree_ids] if tree_ids else list(self.data.index)

        bought_ids, errors = [], []

        # Disable UI during operation
        if hasattr(self, "load_button"): self.load_button.config(state=tk.DISABLED)
        if hasattr(self, "save_button"): self.save_button.config(state=tk.DISABLED)
        if hasattr(self, "buy_button"):  self.buy_button.config(state=tk.DISABLED)
        self.root.config(cursor="watch"); self.status_var.set("Creating & buying labels…"); self.root.update_idletasks()

        print("\n=== BEGIN BULK BUY ===")
        print(f"Rows to process: {len(order)}")
        print(f"Output PDF: {pdf_path}")

        try:
            for idx in order:
                print(f"\n=== ROW {idx} ===")
                try:
                    row = dict(self.data.loc[idx])
                except Exception as e_idx:
                    print(f"[Row {idx}] Row read error: {e_idx}")
                    traceback.print_exc()
                    errors.append((idx, f"Row read error: {e_idx}"))
                    continue

                try:
                    print("Row data:")
                    print(json.dumps(row, indent=2, default=str))
                except Exception:
                    print("Row JSON dump failed; printing as str:")
                    print(str(row))

                try:
                    # Build EasyPost create payload from your columns
                    create_payload = self._row_to_shipment_create(row)
                    print("Create payload:")
                    print(json.dumps(create_payload, indent=2, default=str))

                    # Create shipment (SDK-tolerant)
                    shp = _create(create_payload)
                    sid = shp.get("id") if isinstance(shp, dict) else getattr(shp, "id", None)
                    print(f"Shipment created: {sid}")

                    # Rate selection
                    want_carrier = str(row.get("carrier", "")).strip()
                    want_service = str(row.get("service", "")).strip()
                    print(f"Desired rate: carrier={want_carrier or '(none)'} service={want_service or '(none)'}")

                    rates = shp.get("rates", []) if isinstance(shp, dict) else getattr(shp, "rates", []) or []
                    print(f"Rates returned ({len(rates)}):")
                    for r in rates:
                        print(f"  - {r.get('carrier')} / {r.get('service')} : {r.get('rate')}  (id={r.get('id')})")

                    chosen_rate_id = None
                    if want_carrier and want_service:
                        for r in rates:
                            if (str(r.get("carrier","")).upper() == want_carrier.upper()
                                and str(r.get("service","")).upper() == want_service.upper()):
                                chosen_rate_id = r.get("id"); print(f"Matched desired rate id={chosen_rate_id}")
                                break
                    if not chosen_rate_id and rates:
                        try:
                            chosen_rate_id = sorted(rates, key=lambda r: float(r.get("rate", "1e12")))[0]["id"]
                            print(f"Falling back to cheapest rate id={chosen_rate_id}")
                        except Exception as sort_ex:
                            print(f"Rate sort failed: {sort_ex}")
                            traceback.print_exc()
                            chosen_rate_id = rates[0]["id"]; print(f"Falling back to first rate id={chosen_rate_id}")

                    if not chosen_rate_id:
                        raise RuntimeError("No rates available for shipment.")

                    # Buy (SDK-tolerant)
                    print(f"Buying shipment {sid} with rate {chosen_rate_id}…")
                    shp_bought = _buy(sid, chosen_rate_id)
                    bought_sid = (shp_bought.get("id") if isinstance(shp_bought, dict)
                                  else getattr(shp_bought, "id", sid))
                    bought_ids.append(bought_sid)
                    print(f"Purchase successful: {bought_sid}")

                    self.status_var.set(f"Bought {len(bought_ids)}/{len(order)}…")
                    self.root.update_idletasks()

                except Exception as ex_row:
                    err_msg = f"{type(ex_row).__name__}: {ex_row}"
                    print(f"[Row {idx}] ERROR: {err_msg}")
                    # If it's an EasyPost InvalidRequestError (invalid parameters), dump details:
                    try:
                        from easypost.errors.api.invalid_request_error import InvalidRequestError
                        if isinstance(ex_row, InvalidRequestError):
                            self._explain_easypost_error(ex_row)
                    except Exception:
                        pass
                    import traceback as _tb
                    _tb.print_exc()
                    errors.append((idx, err_msg))

            if not bought_ids:
                print("\nNo labels purchased; aborting PDF build.")
                messagebox.showerror("Buy Failed", "No labels were purchased. Check console logs for details.")
                return

            # Compose final PDF using your existing builder (rotate/pad/cache → merge)
            def _status_cb(msg):
                self.status_var.set(msg); print(msg); self.root.update_idletasks()

            print("\n=== FETCH & MERGE LABELS INTO PDF ===")
            print(f"Shipment IDs to fetch: {bought_ids}")
            build_pdf_from_shipments_multipage(client, key, bought_ids, pdf_path, _status_cb)
            open_pdf(pdf_path)
            
            if errors:
                msg = f"Saved PDF to:\n{pdf_path}\n\nSome rows failed:\n"
                for i, (ridx, emsg) in enumerate(errors[:10]):
                    msg += f"  Row {ridx}: {emsg}\n"
                if len(errors) > 10:
                    msg += f"  …and {len(errors)-10} more."
                print("\nCompleted with errors.")
                messagebox.showwarning("Completed with Errors", msg)
            else:
                print("\nSuccess. PDF saved.")
                

        finally:
            # always restore cursor
            self.root.config(cursor="")

            # always re-enable Load and Buy
            try:
                if hasattr(self, "load_button"):
                    self.load_button.config(state=tk.NORMAL)
                if hasattr(self, "buy_button"):
                    self.buy_button.config(state=tk.NORMAL)
            except Exception:
                pass

            # let existing logic decide Save button state based on package completeness
            if hasattr(self, "_update_save_state"):
                self._update_save_state()
            else:
                # fallback: if no updater, at least turn Save back on
                try:
                    if hasattr(self, "save_button"):
                        self.save_button.config(state=tk.NORMAL)
                except Exception:
                    pass





if __name__ == "__main__":
    root = tk.Tk()
    app = CSVConverterApp(root)
    root.mainloop()
