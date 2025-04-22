# Example UXP Plugin Patterns

## Folder Structure
```
/plugin
  ├── manifest.json
  ├── index.html
  ├── main.js
  └── style.css
```

## Panel Logic
- Tab switching with `display: none` on inactive `<div>`
- CSV parsing using `File.read()` + `.split(',')`
- Console logging to in-plugin UI: `document.getElementById("log").textContent += ...`

## Events
```js
document.getElementById("processBtn").addEventListener("click", async () => {
    // do work here
});
```