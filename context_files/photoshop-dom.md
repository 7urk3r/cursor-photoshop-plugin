# Photoshop DOM (JavaScript API)

## Key Objects
- `app.activeDocument` – the currently active PSD file
- `app.activeDocument.layers` – array of all top-level layers
- `layer.textItem` – represents text layers with:
    - `.contents` – the text string
    - `.size` – font size in points
    - `.font`, `.color`, `.justification`, etc.

## Usage Examples
```js
let layer = app.activeDocument.artLayers.getByName("text1");
layer.textItem.contents = "Hello\rWorld";
layer.textItem.size = 24;
```

## Smart Object Replacement
```js
let smartObjLayer = app.activeDocument.artLayers.getByName("img1");
smartObjLayer.smartObject.replaceContents(File("/path/to/image.png"));
```