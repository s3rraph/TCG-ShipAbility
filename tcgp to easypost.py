import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os

DEFAULT_CARRIER = 'USPS'
DEFAULT_SERVICE = 'First'
DEFAULT_LABEL_FORMAT = 'PNG'
DEFAULT_COUNTRY = 'US'

# Default From Address
FROM_NAME = 'JEREMY CARRILLO'
FROM_COMPANY = 'TCG PLAYABILITY'
FROM_PHONE = ''
FROM_EMAIL = ''
FROM_STREET1 = '17130 VAN BUREN BLVD PMB 807'
FROM_STREET2 = ''
FROM_CITY = 'RIVERSIDE'
FROM_STATE = 'CA'
FROM_ZIP = '92504-5905'
FROM_COUNTRY = 'US'

class CSVConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Shipping Export to Batch CSV Converter")

        self.frame = tk.Frame(root)
        self.frame.pack(padx=10, pady=10)

        self.format_var = tk.StringVar(value="TCGPlayer")
        tk.Label(self.frame, text="Select Format:").pack()
        tk.Radiobutton(self.frame, text="TCGPlayer", variable=self.format_var, value="TCGPlayer").pack(anchor=tk.W)
        tk.Radiobutton(self.frame, text="Manapool", variable=self.format_var, value="Manapool").pack(anchor=tk.W)

        self.load_button = tk.Button(self.frame, text="Load Shipping Export CSV", command=self.load_csv)
        self.load_button.pack(pady=5)

        self.tree = None
        self.data = None

        self.save_button = tk.Button(self.frame, text="Save as Batch CSV", command=self.save_csv, state=tk.DISABLED)
        self.save_button.pack(pady=5)

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return

        try:
            df = pd.read_csv(file_path, dtype=str)
            format_type = self.format_var.get()

            if format_type == "TCGPlayer":
                df['Item Count'] = df['Item Count'].astype(int)
                df['PostalCode'] = df['PostalCode'].astype(str)
                card_count = df['Item Count']
                self.data = pd.DataFrame({
                    'to_address.name': df['FirstName'].fillna('') + ' ' + df['LastName'].fillna(''),
                    'to_address.company': '',
                    'to_address.phone': '',
                    'to_address.email': '',
                    'to_address.street1': df['Address1'],
                    'to_address.street2': df['Address2'],
                    'to_address.city': df['City'],
                    'to_address.state': df['State'],
                    'to_address.zip': df['PostalCode'].str.strip(),
                    'to_address.country': df['Country'].fillna(DEFAULT_COUNTRY),
                    'parcel.weight': card_count.apply(lambda x: 1 if x <= 8 else 2),
                    'options.machinable': card_count.apply(lambda x: False if x > 15 else True)
                })
            elif format_type == "Manapool":
                df['item_count'] = df['item_count'].astype(int)
                df['shipping_zip'] = df['shipping_zip'].astype(str)
                card_count = df['item_count']
                self.data = pd.DataFrame({
                    'to_address.name': df['shipping_name'],
                    'to_address.company': '',
                    'to_address.phone': '',
                    'to_address.email': '',
                    'to_address.street1': df['shipping_line1'],
                    'to_address.street2': df['shipping_line2'],
                    'to_address.city': df['shipping_city'],
                    'to_address.state': df['shipping_state'],
                    'to_address.zip': df['shipping_zip'].str.strip(),
                    'to_address.country': df['shipping_country'].fillna(DEFAULT_COUNTRY),
                    'parcel.weight': card_count.apply(lambda x: 1 if x <= 8 else 2),
                    'options.machinable': card_count.apply(lambda x: False if x > 15 else True)
                })

            # Add static fields
            self.data = self.data.assign(
                **{
                    'from_address.name': FROM_NAME,
                    'from_address.company': FROM_COMPANY,
                    'from_address.street1': FROM_STREET1,
                    'from_address.street2': FROM_STREET2,
                    'from_address.city': FROM_CITY,
                    'from_address.state': FROM_STATE,
                    'from_address.zip': FROM_ZIP,
                    'from_address.country': FROM_COUNTRY,
                    'parcel.length': '',
                    'parcel.width': '',
                    'parcel.height': '',
                    'parcel.predefined_package': 'Letter',
                    'carrier': DEFAULT_CARRIER,
                    'service': DEFAULT_SERVICE,
                    'options.label_format': DEFAULT_LABEL_FORMAT
                }
            )

            self.display_preview()
            self.save_button.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV: {e}")

    def display_preview(self):
        if self.tree:
            self.tree.destroy()

        self.tree = ttk.Treeview(self.frame, columns=list(self.data.columns), show='headings', height=10)
        for col in self.data.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)

        for _, row in self.data.head(20).iterrows():
            self.tree.insert('', 'end', values=list(row))

        self.tree.pack(pady=10)

    def save_csv(self):
        if self.data is None:
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return

        try:
            self.data.to_csv(file_path, index=False)
            messagebox.showinfo("Success", f"Batch CSV saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save CSV: {e}")

if __name__ == '__main__':
    root = tk.Tk()
    app = CSVConverterApp(root)
    root.mainloop()
