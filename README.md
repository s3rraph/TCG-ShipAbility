# 🧾 TCG ShipAbility

**TCG ShipAbility** is an all-in-one desktop utility for TCG sellers that converts exported order CSVs from marketplaces like **TCGplayer**, **Manapool** into purchased  **EasyPost shipping labels**.

It streamlines your fulfillment flow from **CSV → label → PDF**, complete with rules for weight, machinability, and package type detection.

---

## 🚀 Key Features

- **CSV Ingestion:** Automatically detects and normalizes exports from multiple TCG marketplaces.  
- **Smart Detection:** Automatically identifies letters vs packages.  
- **Configurable Rules:** Define per-item weight thresholds, machinable status, and default services.  
- **EasyPost Integration:** Buy and merge EasyPost labels directly into a single printable PDF.  
- **Manual Overrides:** Edit both package and letter rows directly from the preview table.  
- **Batch Processing:** Purchase and generate dozens of labels in one click.  
- **Persistent Settings:** All settings are saved to `config.json`.  

---

<img src="https://github.com/user-attachments/assets/b527366b-d41f-42dd-bcda-68ae52c5b54d" height="150">
<img src="https://github.com/user-attachments/assets/40fa02cc-94f7-42e8-a4fc-d69651db13af" height="150">
<img src="https://github.com/user-attachments/assets/5f6fbef5-9d19-430a-b7f2-f4825baf4436" height="150">

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
   python Shipping.py
   ```
   
---

## ⚙️ Configuration

Settings persist to `config.json` automatically.  
They can be edited in the **Settings** tab.

### Configuration fields

| Section | Description |
|----------|--------------|
| **From Address** | Your sender information (used for every label). |
| **Rules** | Determines letter weight / machinability by item count. |
| **EasyPost API Key** | Stored in the config; editable from Settings → “Set API Key …”. |

---

## 🖥️ Usage Flow

1. **Select Format:** Choose *Auto*, *TCGplayer*, or *Manapool*.  
2. **Load CSV:** Import your marketplace export.  
3. **Preview Orders:** Verify address, weight, and service.  
4. **Edit Rows:** Double-click any line to modify dimensions, weight, or machinability.  
5. **Buy Labels & Build PDF:** Purchases all labels and merges them into a printable 4×6 PDF.

---

## 🧰 Advanced Notes

- **Save as Batch CSV:** Exports an EasyPost-ready batch file that you can use to purchase from the web portal instead of through this app.  
- **Letter Rules:** Automatically applied to non-package rows; overridden by manual edits.  
- **Package Rows:** Require L/W/H + Weight before purchase.  
- **Error Handling:** Invalid addresses or rates will be logged in the console with detailed EasyPost responses.

---

## 🧾 Example Workflow

1. Export your **Manapool orders** CSV.  
2. Launch **TCG ShipAbility** → Format: *Auto*.  
3. Click **Load Shipping Export CSV** → Preview populates.  
4. Edit any letters / packages as needed.  
5. Click **Buy Labels & Build PDF** → select output path.  
6. Print your merged label PDF and ship.



## 🏷️ License

MIT License © 2025 TCG Playability / S3rraph / Jeremy Carrillo
