# Plugin Spec: SaveOnPeptides Auto Replace Tool

## 🧩 Overview
This UXP Photoshop plugin helps automate text and image replacement using CSV data for peptide product labels and visual generation.

## 🔀 Tabs
- Auto Text Replace
- Auto Image Replace
- Logs

---

## 🔤 Auto Text Replace Tab

### User Flow
1. User clicks "Choose CSV File" → opens file picker
2. User selects a row-processing mode:
   - All rows
   - Just one row
3. On "Process CSV":
   - Prompts for output folder
   - Creates `PNG/` and `PSD/` subfolders inside it
   - For each row:
     - Matches layers named `text1`, `text2`, etc.
     - Replaces text with value from CSV (handling `|br|` as a line break)
     - Applies font size from `fontsize1`, `fontsize2`
     - Retains all other text styles
     - Saves updated file to output folders

---

## 🖼 Auto Image Replace Tab

### User Flow
1. User clicks "Choose Input Folder" (with PSDs + PNGs)
2. User selects CSV file with columns like `psdname1`, `imgname1`, etc.
3. On "Process CSV":
   - Prompts for output folder
   - Creates `PNG/` and `PSD/` subfolders
   - For each row:
     - Opens matching PSD file
     - Replaces smart object/image layers named `img1`, `img2`, etc.
     - Saves updated PSD and PNG to output

---

## 📝 Logs Tab
- Realtime log window visible during processing
- Writes final log to `log.txt` in the output folder after job is complete

## ⚙️ Output Folder Structure
```
/selected_output_folder/
├── PNG/
├── PSD/
└── log.txt
```

---

## 🔄 Interoperability
- CSV format:
  - Text Replace: `text1`, `fontsize1`, `text2`, ...
  - Image Replace: `psdname1`, `imgname1`, ...
- Supported input formats: `.csv`, `.psd`, `.png`

## 🔌 External Tools
- Communicates with MCP server via WebSocket for:
  - Restarting plugin
  - Screenshots
  - OCR verification
  - Logging results