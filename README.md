# 🧾 TCG ShipAbility

**TCG ShipAbility** is an all-in-one desktop utility for TCG sellers that converts exported order CSVs from marketplaces like **TCGplayer**, **Manapool**, and others into fully purchasable **EasyPost shipping labels**.

It streamlines your fulfillment flow from **CSV → label → PDF**, complete with rules for weight, machinability, and package type detection.

---

## 🚀 Key Features

- **CSV Ingestion:** Automatically detects and normalizes exports from multiple TCG marketplaces.  
- **Smart Detection:** Automatically identifies letters vs packages based on item count or shipping cost.  
- **Configurable Rules:** Define per-item weight thresholds, machinable status, and default services.  
- **EasyPost Integration:** Buy and merge EasyPost labels directly into a single printable PDF.  
- **Manual Overrides:** Edit both package and letter rows directly from the preview table.  
- **Batch Processing:** Purchase and generate dozens of labels in one click.  
- **Persistent Settings:** All settings and your EasyPost API Key are saved to `config.json`.  
- **Multi-Platform Ready:** Extensible architecture for adding new marketplaces (Shopify, CardTrader, etc.).

---

## 🧩 Installation

1. **Clone the repo**
   ```bash
   git clone https://github.com/yourusername/tcg-shipability.git
   cd tcg-shipability
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the app**
   ```bash
   python tcgp_to_easypost.py
   ```
   *(The app window will open automatically.)*

---

## ⚙️ Configuration

Settings persist to `config.json` automatically.  
To edit defaults manually, open the app and go to the **Settings** tab.

### Configuration fields

| Section | Description |
|----------|--------------|
| **Defaults** | Default carrier, service, label format, and country. |
| **From Address** | Your sender information (used for every label). |
| **Rules** | Determines letter weight / machinability by item count. |
| **Detection** | Defines which shipping price values indicate a package. |
| **EasyPost API Key** | Stored securely in the config; editable from Settings → “Set API Key …”. |

---

## 🖥️ Usage Flow

1. **Select Format:** Choose *Auto*, *TCGplayer*, or *Manapool*.  
2. **Load CSV:** Import your marketplace export.  
3. **Preview Orders:** Verify address, weight, and service.  
4. **Edit Rows:** Double-click any line to modify dimensions, weight, or machinability.  
5. **Save as Batch CSV:** Exports an EasyPost-ready batch file.  
6. **Buy Labels & Build PDF:** Purchases all labels and merges them into a printable 4×6 PDF.

---

## 🧰 Advanced Notes

- **Letter Rules:** Automatically applied to non-package rows; overridden by manual edits.  
- **Package Rows:** Require L/W/H + Weight before purchase.  
- **Sorting:** Manapool CSVs are automatically sorted by `seller_label_number`.  
- **Cache:** Downloaded labels are cached locally for faster rebuilds.  
- **Error Handling:** Invalid addresses or rates will be logged in the console with detailed EasyPost responses.

---

## 🧾 Example Workflow

1. Export your **Manapool orders** CSV.  
2. Launch **TCG ShipAbility** → Format: *Auto*.  
3. Click **Load Shipping Export CSV** → Preview populates.  
4. Edit any letters / packages as needed.  
5. Click **Buy Labels & Build PDF** → select output path.  
6. Print your merged label PDF and ship.

---

## 🔐 EasyPost Integration

- API key stored in `config.json` under `"easypost_api_key"`.  
- You can set/update it via **Settings → Set API Key …**.  
- Uses official [`easypost`](https://pypi.org/project/easypost/) Python SDK.

---

## 🛠️ Extending for New Marketplaces

To add a new CSV type:
1. Define a new header signature in `detect_format_from_headers()`.  
2. Add a format branch in `load_csv()` to map columns to `self.data`.  
3. Call `apply_rules_and_package_logic()` for rule application.  
4. Preview and label generation will work automatically.

---

## 📦 Tech Stack

- **Python 3.9+**  
- **Tkinter GUI**  
- **Pandas** for CSV parsing  
- **EasyPost SDK** for label creation  
- **ReportLab / Pillow** for PDF composition  

---

## 🧑‍💻 Development Notes

- All persistent UI states are stored in `config.json`.  
- Logs print to stdout for debugging during label purchase.  
- Exception traces are safe to view — no private data dumped by default.  

---

## 🏷️ License

MIT License © 2025 TCG Playability / Jeremy Carrillo
