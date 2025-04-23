
# Plugin Specification: TestCursorPlugin

## 🧠 Purpose
A simple UXP plugin for Adobe Photoshop that loads a panel and allows reading of the active document's layer metadata. This will help test Cursor's ability to read and interact with Photoshop documents.

---

## 🔧 Plugin Metadata

- **Name**: Test Cursor Plugin
- **ID**: com.cursor.testplugin
- **Manifest Version**: 5
- **Host App**: Photoshop (minVersion: 23.0.0)
- **Panel ID**: cursorPanel

---

## 🎨 UI

The panel should include:

- A heading that says: `🧠 Test Cursor Panel`
- A button labeled: `Read PSD Metadata`
- A `<pre>` block to display the JSON result from layer metadata
- Clean, readable styling

---

## 📄 Files

### 1. `manifest.json`

- Set `manifestVersion` to `5`
- Use `entrypoints` with:
  - `"type": "panel"`
  - `"mainPath": "index.html"`
- Assign `panelId` as `cursorPanel`
- Define minimum and preferred sizes for the panel

### 2. `index.html`

- Include a heading, a button (`id="readMetadata"`), and a `<pre>` (`id="output"`)
- Load `index.js` using `<script type="module">`

### 3. `index.js`

- Require `photoshop` API using:  
  ```js
  const { app } = require("photoshop");
  ```
- On button click, read all top-level layer data from the current active document
- Display the data as formatted JSON inside the `<pre>` element
- Include error handling if no document is open

---

## ✅ Functional Requirements

- Show the panel in Photoshop under Plugins > Test Cursor Plugin
- Log a message (`"✅ TestCursorPlugin loaded"`) when loaded
- When the user clicks the `Read PSD Metadata` button:
  - Read all layer names, types, and visibility flags
  - Display results as prettified JSON in the panel
- Handle errors gracefully and log to console

---

## 🧪 Test Scenarios

- [ ] Panel loads and is visible in Photoshop
- [ ] Button click reads layer data from open PSD
- [ ] No crash when no document is open
- [ ] Layer info displays in readable JSON

---

## 📁 Output Format Example

```json
[
  {
    "name": "Background",
    "kind": "pixel",
    "visible": true
  },
  {
    "name": "Logo",
    "kind": "smartObject",
    "visible": false
  }
]
```

---

## ✨ Future Extensions (Optional)

- Export the metadata to a CSV or JSON file
- Add layer thumbnails to the output
- Use MCP to send the layer metadata to a remote AI
