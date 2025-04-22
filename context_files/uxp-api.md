# UXP Plugin API (Photoshop)

## File Picker
```js
const uxp = require('uxp').storage.localFileSystem;
const file = await uxp.getFileForOpening({ types: ['csv'] });
```

## Folder Picker
```js
const folder = await uxp.getFolder();
```

## Writing Files
```js
const file = await folder.createFile("log.txt", { overwrite: true });
await file.write("Log content here");
```

## Permissions (manifest.json)
```json
"permissions": {
  "filesystem": { "read": true, "write": true },
  "launchProcess": true
}
```

## UI (HTML + CSS)
- Plugin UI is rendered in an iframe-like panel
- JS has limited access: use message-passing if needed