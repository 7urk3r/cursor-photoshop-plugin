console.log("✅ TestCursorPlugin loaded");

const { app } = require("photoshop");
const { batchPlay } = require('photoshop').action;
const fs = require('uxp').storage.localFileSystem;
const formats = require('uxp').storage.formats;

// Define log function
function log(message) {
    const logPanel = document.getElementById('logPanel');
    if (logPanel) {
        const logEntry = document.createElement('div');
        logEntry.textContent = message;
        logPanel.appendChild(logEntry);
        // Auto-scroll to bottom
        logPanel.scrollTop = logPanel.scrollHeight;
    }
    console.log(message);
}

// Enhanced state management with performance tracking and step completion
const createPluginState = () => ({
    status: {
        isProcessing: false,
        lastError: null,
        currentOperation: null,
        isInitialized: false,
        steps: {
            csvLoaded: false,
            outputFolderSelected: false,
            processingStarted: false,
            processingComplete: false
        },
        performance: {
            startTime: null,
            endTime: null,
            csvLoadTime: null,
            folderSetupTime: null,
            processingTime: null,
            totalRows: 0,
            processedRows: 0
        }
    },
    data: {
        csvData: null,
        outputFolder: null,
        lastProcessedRow: null
    }
});

// Separate state per tab with enhanced structure
const textReplaceState = createPluginState();
const imageReplaceState = createPluginState();

// Enhanced error handling
class PluginError extends Error {
    constructor(message, code, details = {}) {
        super(message);
        this.name = 'PluginError';
        this.code = code;
        this.details = details;
        this.timestamp = new Date().toISOString();
    }
}

// Cache common BatchPlay commands for better performance
const batchPlayCommands = {
    selectLayer: (layerName) => ({
        _obj: "select",
        _target: [{ _ref: "layer", _name: layerName }],
        makeVisible: true,
        _isCommand: true,
        _options: { dialogOptions: "dontDisplay" }
    }),
    
    setText: (layerName, text) => ({
        _obj: "set",
        _target: [{ _ref: "textLayer", _name: layerName }],
        to: { _obj: "textLayer", textKey: text },
        _isCommand: true,
        _options: { dialogOptions: "dontDisplay" }
    }),
    
    setFontSize: (layerName, size) => ({
        _obj: "set",
        _target: [{ _ref: "textLayer", _name: layerName }],
        to: { _obj: "characterStyle", size: { _unit: "pointsUnit", _value: parseFloat(size) } },
        _isCommand: true,
        _options: { dialogOptions: "dontDisplay" }
    }),

    getCharacterStyle: (layerName) => ({
        _obj: "get",
        _target: [{ _ref: "property", _property: "characterStyle" }, { _ref: "textLayer", _name: layerName }],
        _isCommand: true,
        _options: { dialogOptions: "dontDisplay" }
    })
};

// Optimized delay function with configurable times
const delays = {
    selection: 200,
    textUpdate: 500,
    fontUpdate: 500,
    verification: 300
};

const wait = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// Enhanced MCP relay write with proper file handling, validation, and type checking
async function writeToMCPRelay(data) {
    let file = null;
    
    if (!data) {
        throw new PluginError('No data provided for MCP relay', 'INVALID_INPUT');
    }

    // Validate formats.utf8 is available
    if (!formats || !formats.utf8) {
        throw new PluginError('UTF-8 format not available', 'FORMAT_ERROR');
    }
    
    try {
        console.log("[DEBUG] Getting temporary folder...");
        const tempFolder = await fs.getTemporaryFolder();
        console.log("[DEBUG] Temp folder path:", tempFolder.nativePath);
        
        // Get or create mcp-relay folder
        console.log("[DEBUG] Setting up MCP relay folder...");
        let mcpRelayFolder;
        try {
            mcpRelayFolder = await tempFolder.getEntry('mcp-relay');
            if (!mcpRelayFolder.isFolder) {
                throw new Error('mcp-relay exists but is not a folder');
            }
            console.log("[DEBUG] Found existing MCP relay folder:", mcpRelayFolder.nativePath);
        } catch (error) {
            mcpRelayFolder = await tempFolder.createEntry('mcp-relay', { type: 'folder' });
            console.log("[DEBUG] Created MCP relay folder:", mcpRelayFolder.nativePath);
        }

        // Get or create logs folder
        console.log("[DEBUG] Setting up logs folder...");
        let logsFolder;
        try {
            logsFolder = await mcpRelayFolder.getEntry('logs');
            if (!logsFolder.isFolder) {
                throw new Error('logs exists but is not a folder');
            }
            console.log("[DEBUG] Found existing logs folder:", logsFolder.nativePath);
        } catch (error) {
            logsFolder = await mcpRelayFolder.createEntry('logs', { type: 'folder' });
            console.log("[DEBUG] Created logs folder:", logsFolder.nativePath);
        }

        // Create unique filename for this write
        const timestamp = new Date().getTime();
        const fileName = `mcp-out-${timestamp}.json`;
        
        // Convert data to string with validation
        let jsonString;
        try {
            jsonString = JSON.stringify(data, null, 2);
            if (!jsonString) {
                throw new Error('JSON stringification resulted in empty string');
            }
        } catch (jsonError) {
            throw new PluginError('Failed to stringify data: ' + jsonError.message, 'JSON_ERROR', { jsonError, data });
        }
        
        // Create and write to file with enhanced error handling
        console.log("[DEBUG] Creating and writing MCP relay file...");
        try {
            // First try to create the file
            file = await logsFolder.createFile(fileName, { overwrite: true });
            console.log("[DEBUG] Created file:", file.nativePath);
            
            // Validate file handle
            if (!file || typeof file.write !== 'function') {
                throw new PluginError(
                    'Invalid file handle: write() not available',
                    'FILE_WRITE_TYPE_ERROR',
                    { fileType: typeof file, methods: Object.keys(file || {}) }
                );
            }

            // Write data with explicit format
            await file.write(jsonString, { format: formats.utf8 });
            console.log("[DEBUG] Data written successfully to:", file.nativePath);
            
            return file.nativePath;
            
        } catch (writeError) {
            // If standard write fails, try alternative approach with temporary file
            console.log("[DEBUG] Standard write failed, attempting fallback...");
            
            try {
                // Create a temporary file first
                const tempFile = await fs.createTemporaryFile('mcp-relay-');
                await tempFile.write(jsonString, { format: formats.utf8 });
                
                // Move to final location
                const finalPath = `${logsFolder.nativePath}/${fileName}`;
                await tempFile.moveTo(finalPath, { overwrite: true });
                
                console.log("[DEBUG] Fallback write successful to:", finalPath);
                return finalPath;
                
            } catch (fallbackError) {
                throw new PluginError(
                    'Both standard and fallback write attempts failed',
                    'FILE_WRITE_ERROR',
                    { 
                        originalError: writeError,
                        fallbackError,
                        attempted: file?.nativePath
                    }
                );
            }
        }
        
    } catch (error) {
        console.error("[DEBUG] Error in writeToMCPRelay:", {
            message: error.message,
            stack: error.stack,
            name: error.name,
            code: error.code
        });
        
        throw new PluginError(
            'Failed to write to MCP relay: ' + error.message,
            'MCP_WRITE_ERROR',
            { originalError: error }
        );
    } finally {
        if (file) {
            try {
                // Log final status
                console.log("[DEBUG] Completed MCP relay operation for file:", file.nativePath);
            } catch (cleanupError) {
                console.error('[DEBUG] Error during cleanup:', cleanupError);
            }
        }
    }
}

// Enhanced layer targeting utility with better type handling and logging
function createLayerTarget(layer, options = {}) {
    if (!layer) {
        throw new PluginError('No layer provided for targeting', 'INVALID_LAYER');
    }

    // Log full layer details for debugging with API v2 properties
    console.log("[DEBUG] Creating target for layer:", {
        name: layer.name,
        id: layer._id,  // API v2 uses _id
        kind: layer.kind,
        type: typeof layer,
        hasTextItem: !!layer.textItem,
        textContents: layer.textItem?.contents,
        bounds: layer.bounds,
        visible: layer.visible
    });

    // Enhanced layer type detection with numeric codes
    const layerKind = layer.kind;
    const isTextLayer = layerKind === 'text' || layerKind === 3;
    const isSmartObject = layerKind === 'smartObject' || layerKind === 5;

    console.log("[DEBUG] Layer type detection:", {
        layerKind,
        isTextLayer,
        isSmartObject,
        rawKind: typeof layerKind === 'number' ? layerKind : 'string'
    });

    // Get the appropriate ID for API v2
    const layerId = layer._id;
    if (!layerId) {
        console.error("[DEBUG] Invalid layer ID:", {
            _id: layer._id,
            name: layer.name,
            type: typeof layerId,
            layer: JSON.stringify(layer, null, 2)
        });
        throw new PluginError(
            'Layer ID not found', 
            'INVALID_LAYER_ID',
            { 
                layerName: layer.name,
                layerKind
            }
        );
    }

    // Default targeting options with enhanced type detection
    const defaultOptions = {
        type: isTextLayer ? 'textLayer' : 
              isSmartObject ? 'smartObjectLayer' : 'layer',
        makeVisible: false,
        _isCommand: true,
        dialogOptions: "dontDisplay"
    };

    // Merge with provided options, but ensure type is correct for the layer
    const finalOptions = { ...defaultOptions, ...options };
    
    // Create the base target descriptor with enhanced validation
    const target = {
        _ref: finalOptions.type,
        _id: layerId
    };

    // For text layers, we need special handling
    if (isTextLayer) {
        // When selecting, use ordinal targeting
        if (finalOptions.select) {
            target._enum = "ordinal";
            target._value = "targetEnum";
        }
        // For text operations, include additional text-specific properties
        else {
            target._ref = "textLayer";
            target.textKey = layer.textItem?.contents || "";
        }
    }

    // Log the final target for debugging
    console.log("[DEBUG] Created layer target:", {
        target,
        options: finalOptions,
        isTextLayer,
        layerKind
    });

    // For selection commands
    if (finalOptions.select) {
        const selectCommand = {
            _obj: "select",
            _target: [target],
            makeVisible: finalOptions.makeVisible,
            _isCommand: finalOptions._isCommand,
            _options: { dialogOptions: "dontDisplay" }
        };

        console.log("[DEBUG] Created selection command:", selectCommand);
        return selectCommand;
    }

    return target;
}

// Optimized layer operations with better error handling
async function executeWithRetry(command, maxRetries = 3) {
    let lastError;
    for (let i = 0; i < maxRetries; i++) {
        try {
            const result = await batchPlay([command], {
                synchronousExecution: true,
                modalBehavior: "none",
                _options: { dialogOptions: "dontDisplay" }
            });
            return result;
        } catch (error) {
            lastError = error;
            log(`[DEBUG] Retry ${i + 1} after error: ${error.message}`);
            if (i < maxRetries - 1) {
                await wait(100 * Math.pow(2, i)); // Exponential backoff
            }
        }
    }
    throw lastError;
}

// Optimized text replacement with verification
async function replaceText(layer, newText) {
    try {
        if (!layer?.name) {
            throw new PluginError('Invalid layer object', 'INVALID_LAYER');
        }

        log(`[DEBUG] Starting text replacement for layer: ${layer.name}`);

        // Select layer first
        await batchPlay(
            [{
                _obj: "select",
                _target: [{ _ref: "layer", _name: layer.name }],
                makeVisible: false
            }],
            { 
                synchronousExecution: true,
                modalBehavior: "execute"
            }
        );

        // Update text with line break handling
        const processedText = newText.replace(/\|br\|/g, '\r');
        await batchPlay(
            [{
                _obj: "set",
                _target: [{ _ref: "textLayer", _name: layer.name }],
                to: { _obj: "textLayer", textKey: processedText }
            }],
            {
                synchronousExecution: true,
                modalBehavior: "execute"
            }
        );

        log(`[DEBUG] Text updated for layer: ${layer.name}`);
        return true;

    } catch (error) {
        log(`[DEBUG] Text replacement error for layer ${layer?.name}: ${error.message}`);
        throw new PluginError(
            `Failed to replace text in layer ${layer?.name}`,
            'TEXT_REPLACE_ERROR',
            { originalError: error, layer: layer?.name }
        );
    }
}

async function verifyFontSize(layer, expectedSize) {
    try {
        log(`[DEBUG] Starting font size verification for layer "${layer.name}" (ID: ${layer._id})`);
        log(`[DEBUG] Expected font size: ${expectedSize}pt`);
        
        // Method 1: Check via DOM API textItem property
        let textItemSize;
        try {
            textItemSize = layer.textItem?.size;
            log(`[DEBUG] TextItem size: ${textItemSize}pt (via DOM API)`);
        } catch (textItemError) {
            log(`[DEBUG] TextItem check failed: ${textItemError.message}`);
        }

        // Method 2: Check via text style property using executeAsModal
        let textStyleSize;
        try {
            const result = await app.executeAsModal(async () => {
                return layer.textItem?.style?.size;
            });
            textStyleSize = result;
            log(`[DEBUG] Text style size: ${textStyleSize}pt (via style property)`);
        } catch (styleError) {
            log(`[DEBUG] Style property check failed: ${styleError.message}`);
        }

        // Collect all valid measurements
        const measurements = [
            { method: 'TextItem', size: textItemSize },
            { method: 'Text Style', size: textStyleSize }
        ].filter(m => m.size !== undefined && m.size !== null);

        if (measurements.length === 0) {
            log(`[DEBUG] ❌ No valid font size measurements obtained`);
            return false;
        }

        // Log all measurements
        log(`[DEBUG] All font size measurements:`, {
            expected: expectedSize,
            measurements: measurements.map(m => `${m.method}: ${m.size}pt`)
        });

        // Check for consistency among measurements with tolerance
        const tolerance = 0.1;
        const sizesMatch = measurements.every(m => 
            Math.abs(parseFloat(m.size) - parseFloat(expectedSize)) < tolerance
        );

        if (!sizesMatch) {
            const mismatchedSizes = measurements
                .filter(m => Math.abs(parseFloat(m.size) - parseFloat(expectedSize)) >= tolerance)
                .map(m => `${m.method}: ${m.size}pt`);
            
            log(`[DEBUG] ❌ Font size mismatch detected:`, {
                expected: `${expectedSize}pt`,
                mismatches: mismatchedSizes
            });
            return false;
        }

        log(`[DEBUG] ✅ Font size verified successfully: ${expectedSize}pt`);
        return true;

    } catch (error) {
        log(`[DEBUG] ❌ Font size verification failed with error: ${error.message}`);
        return false;
    }
}

async function updateFontSize(layer, fontSize) {
    try {
        if (!layer?.name) {
            throw new PluginError('Invalid layer object', 'INVALID_LAYER');
        }

        // Ensure fontSize is a valid number
        const size = parseFloat(fontSize);
        if (isNaN(size)) {
            throw new PluginError('Invalid font size', 'INVALID_FONT_SIZE', { fontSize });
        }

        log(`[DEBUG] Starting font size update for layer: ${layer.name} to size: ${size}`);

        // Get initial font size for comparison
        const initialSize = await verifyFontSize(layer, size);
        log(`[DEBUG] Initial font size verification: ${initialSize ? 'matches' : 'differs from'} target size`);

        let updateSuccess = false;
        const updateMethods = [];

        // Method 1: Try DOM API first
        try {
            layer.textItem.size = size;
            updateMethods.push('DOM API');
            log(`[DEBUG] Font size updated via DOM API`);
            
            // Quick verify
            await wait(delays.fontUpdate);
            if (await verifyFontSize(layer, size)) {
                updateSuccess = true;
                log(`[DEBUG] DOM API update verified successfully`);
            }
        } catch (domError) {
            log(`[DEBUG] DOM API update failed: ${domError.message}`);
        }

        // Method 2: Try batchPlay with textStyle if DOM API failed
        if (!updateSuccess) {
            try {
                const command = {
                    _obj: "set",
                    _target: [
                        { 
                            _ref: "textLayer",
                            _id: layer._id 
                        }
                    ],
                    to: { 
                        _obj: "textStyle",
                        size: {
                            _unit: "pointsUnit",
                            _value: size
                        }
                    },
                    _isCommand: true,
                    _options: { 
                        dialogOptions: "dontDisplay"
                    }
                };

                const result = await batchPlay(
                    [command],
                    {
                        synchronousExecution: true,
                        modalBehavior: "execute"
                    }
                );

                updateMethods.push('batchPlay textStyle');
                log(`[DEBUG] Font size updated via batchPlay textStyle`);

                // Quick verify
                await wait(delays.fontUpdate);
                if (await verifyFontSize(layer, size)) {
                    updateSuccess = true;
                    log(`[DEBUG] BatchPlay textStyle update verified successfully`);
                }
            } catch (batchError) {
                log(`[DEBUG] BatchPlay textStyle update failed: ${batchError.message}`);
            }
        }

        // Method 3: Try alternative batchPlay method if previous methods failed
        if (!updateSuccess) {
            try {
                const altCommand = {
                    _obj: "set",
                    _target: [
                        { 
                            _ref: "property",
                            _property: "textStyle"
                        },
                        { 
                            _ref: "textLayer",
                            _id: layer._id 
                        }
                    ],
                    to: { 
                        _obj: "textStyle",
                        size: {
                            _unit: "pointsUnit",
                            _value: size
                        }
                    },
                    _options: { 
                        dialogOptions: "dontDisplay"
                    }
                };

                await batchPlay([altCommand], {
                    synchronousExecution: true,
                    modalBehavior: "execute"
                });

                updateMethods.push('batchPlay property');
                log(`[DEBUG] Font size updated via batchPlay property`);

                // Final verification
                await wait(delays.fontUpdate);
                if (await verifyFontSize(layer, size)) {
                    updateSuccess = true;
                    log(`[DEBUG] BatchPlay property update verified successfully`);
                }
            } catch (altError) {
                log(`[DEBUG] BatchPlay property update failed: ${altError.message}`);
            }
        }

        // Final status check
        if (!updateSuccess) {
            throw new Error(`Font size change could not be verified after trying multiple methods: ${updateMethods.join(', ')}`);
        }

        log(`[Cursor OK] Font size updated and verified for layer: ${layer.name} to ${size}pt`);
        
        // Write to MCP relay for tracking
        await writeToMCPRelay({
            command: "sendLog",
            message: `Font size updated and verified: ${layer.name} -> ${size}pt`,
            metrics: {
                layer: layer.name,
                targetSize: size,
                verified: true,
                methodsUsed: updateMethods
            },
            timestamp: new Date().toISOString()
        });

        return true;

    } catch (error) {
        log(`[DEBUG] Font size update error for layer ${layer?.name}: ${error.message}`);
        throw new PluginError(
            `Failed to update font size in layer ${layer?.name}`,
            'FONT_SIZE_UPDATE_ERROR',
            { originalError: error, layer: layer?.name, fontSize }
        );
    }
}

// Optimized layer processing without verification
async function processLayer(layer, rowData, layerIndex) {
    const layerStart = Date.now();
    const updates = { text: false, fontSize: false };

    try {
        if (!layer?.name || !rowData || !layerIndex) {
            throw new PluginError(
                'Invalid layer processing parameters',
                'INVALID_LAYER_PARAMS',
                { layer: layer?.name, hasRowData: !!rowData, layerIndex }
            );
        }

        log(`[DEBUG] Processing layer: ${layer.name}`);

        // Update text if needed
        if (rowData[`text${layerIndex}`]) {
            try {
                await replaceText(layer, rowData[`text${layerIndex}`]);
                updates.text = true;
            } catch (textError) {
                log(`[DEBUG] Text update failed for layer ${layer.name}: ${textError.message}`);
                throw textError;
            }
        }

        // Update font size if needed
        if (rowData[`fontsize${layerIndex}`]) {
            try {
                await updateFontSize(layer, rowData[`fontsize${layerIndex}`]);
                updates.fontSize = true;
            } catch (fontError) {
                log(`[DEBUG] Font size update failed for layer ${layer.name}: ${fontError.message}`);
                throw fontError;
            }
        }

        return {
            success: true,
            layer: layer.name,
            updates,
            duration: Date.now() - layerStart
        };

    } catch (error) {
        throw new PluginError(
            `Failed to process layer ${layer?.name}`,
            'LAYER_PROCESS_ERROR',
            { originalError: error, layer: layer?.name, updates }
        );
    }
}

async function replaceImage(layer, imagePath) {
    try {
        const imageFile = await fs.getFileForPath(imagePath);
        if (!imageFile) {
            throw new PluginError('Image file not found', 'FILE_NOT_FOUND');
        }
        
        const result = await batchPlay(
            [
                {
                    _obj: "replaceContents",
                    _target: [{ _ref: "smartObjectLayer", _id: layer._id }],
                    using: { _ref: "file", _path: imagePath }
                }
            ],
            { synchronousExecution: true }
        );
        
        if (!result || result.length === 0) {
            throw new PluginError('Failed to replace image', 'BATCHPLAY_ERROR');
        }
        
        return result;
    } catch (error) {
        throw new PluginError('Error replacing image', 'IMAGE_REPLACE_ERROR', { error });
    }
}

// Enhanced folder setup with verification and conditional creation
async function setupTextOutputFolders(outputFolder) {
    try {
        if (!outputFolder) {
            throw new PluginError('Output folder not selected', 'FOLDER_NOT_SELECTED');
        }

        console.log("[DEBUG] Checking output folders in:", outputFolder.nativePath);
        const folders = {};
        
        // Store the base output folder
        folders.baseFolder = outputFolder;
        
        // Check PNG folder, create if not exists
        try {
            console.log("[DEBUG] Checking Text_PNG folder...");
            try {
                folders.pngFolder = await outputFolder.getEntry('Text_PNG');
                if (!folders.pngFolder.isFolder) {
                    throw new Error('Text_PNG exists but is not a folder');
                }
                console.log("[DEBUG] Found existing Text_PNG folder:", folders.pngFolder.nativePath);
            } catch (notFoundError) {
                // If the folder doesn't exist, create it
                console.log("[DEBUG] Text_PNG folder not found, creating new one...");
                folders.pngFolder = await outputFolder.createEntry('Text_PNG', { type: 'folder' });
                console.log("[DEBUG] Created new Text_PNG folder:", folders.pngFolder.nativePath);
            }
            
            // Verify the folder was created
            if (!folders.pngFolder || !folders.pngFolder.isFolder) {
                throw new PluginError('Failed to create PNG folder', 'FOLDER_CREATE_ERROR');
            }
        } catch (error) {
            console.error("[DEBUG] Error with PNG folder:", error);
            throw new PluginError(
                'Failed to create Text_PNG folder',
                'FOLDER_CREATE_ERROR',
                { path: 'Text_PNG', error }
            );
        }
        
        // Check PSD folder, create if not exists
        try {
            console.log("[DEBUG] Checking Text_PSD folder...");
            try {
                folders.psdFolder = await outputFolder.getEntry('Text_PSD');
                if (!folders.psdFolder.isFolder) {
                    throw new Error('Text_PSD exists but is not a folder');
                }
                console.log("[DEBUG] Found existing Text_PSD folder:", folders.psdFolder.nativePath);
            } catch (notFoundError) {
                // If the folder doesn't exist, create it
                console.log("[DEBUG] Text_PSD folder not found, creating new one...");
                folders.psdFolder = await outputFolder.createEntry('Text_PSD', { type: 'folder' });
                console.log("[DEBUG] Created new Text_PSD folder:", folders.psdFolder.nativePath);
            }
            
            // Verify the folder was created
            if (!folders.psdFolder || !folders.psdFolder.isFolder) {
                throw new PluginError('Failed to create PSD folder', 'FOLDER_CREATE_ERROR');
            }
        } catch (error) {
            console.error("[DEBUG] Error with PSD folder:", error);
            throw new PluginError(
                'Failed to create Text_PSD folder',
                'FOLDER_CREATE_ERROR',
                { path: 'Text_PSD', error }
            );
        }

        // Verify both folders are accessible
        try {
            const pngExists = await folders.pngFolder.isEntry;
            const psdExists = await folders.psdFolder.isEntry;
            
            if (!pngExists || !psdExists) {
                throw new PluginError(
                    'Could not verify folder access', 
                    'FOLDER_ACCESS_ERROR',
                    {
                        pngExists,
                        psdExists,
                        pngPath: folders.pngFolder?.nativePath,
                        psdPath: folders.psdFolder?.nativePath
                    }
                );
            }
        } catch (verifyError) {
            console.error("[DEBUG] Folder verification failed:", verifyError);
            throw new PluginError(
                'Failed to verify folder access', 
                'FOLDER_VERIFY_ERROR',
                { error: verifyError }
            );
        }

        // Ensure native paths are properly formatted for M1 Macs
        folders.pngNativePath = folders.pngFolder.nativePath.replace(/\\/g, '/');
        folders.psdNativePath = folders.psdFolder.nativePath.replace(/\\/g, '/');

        console.log("[DEBUG] Output folders ready:", {
            png: folders.pngNativePath,
            psd: folders.psdNativePath
        });
        
        return folders;
    } catch (error) {
        console.error("[DEBUG] Folder setup failed:", error);
        throw new PluginError(
            'Failed to setup output folders', 
            'FOLDER_SETUP_ERROR', 
            { 
                error,
                outputPath: outputFolder?.nativePath 
            }
        );
    }
}

async function setupImageOutputFolders(outputFolder) {
    try {
        if (!outputFolder) {
            throw new PluginError('Output folder not selected', 'FOLDER_NOT_SELECTED');
        }

        const folders = {};
        
        // Create PNG folder for image operations
        try {
            folders.pngFolder = await outputFolder.getEntry('Image_PNG');
        } catch (error) {
            folders.pngFolder = await outputFolder.createEntry('Image_PNG', { type: 'folder' });
        }
        
        // Create PSD folder for image operations
        try {
            folders.psdFolder = await outputFolder.getEntry('Image_PSD');
        } catch (error) {
            folders.psdFolder = await outputFolder.createEntry('Image_PSD', { type: 'folder' });
        }
        
        return folders;
    } catch (error) {
        throw new PluginError('Failed to setup image output folders', 'IMAGE_FOLDER_SETUP_ERROR', { error });
    }
}

// Performance tracking utilities
function startOperation(state, operation) {
    state.status.currentOperation = operation;
    state.status.performance.startTime = Date.now();
    return state.status.performance.startTime;
}

function endOperation(state, operation) {
    const endTime = Date.now();
    const duration = endTime - state.status.performance.startTime;
    state.status.performance[`${operation}Time`] = duration;
    state.status.performance.endTime = endTime;
    return duration;
}

// Enhanced CSV Loading for Text Replace
async function loadTextCSV(file) {
    const startTime = Date.now();
    try {
        console.log("[DEBUG] Starting CSV load:", {
            file: file.name,
            timestamp: new Date().toISOString()
        });

        const fileContent = await file.read();
        const lines = fileContent.split('\n');
        const headers = lines[0].trim().split(',').map(h => h.trim());

        // Validate required columns
        const textColumns = headers.filter(h => h.startsWith('text'));
        const fontSizeColumns = headers.filter(h => h.startsWith('fontsize'));

        if (textColumns.length === 0) {
            throw new PluginError(
                'No text columns found in CSV',
                'INVALID_CSV_FORMAT',
                { headers, foundColumns: textColumns }
            );
        }

        console.log("[DEBUG] CSV columns found:", {
            textColumns,
            fontSizeColumns,
            totalColumns: headers.length
        });

        // Parse data with validation
        const data = [];
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue; // Skip empty lines

            const values = line.split(',').map(v => v.trim());
            if (values.length !== headers.length) {
                console.warn(`[DEBUG] Skipping malformed line ${i + 1}: expected ${headers.length} columns, got ${values.length}`);
                continue;
            }

            const rowData = {};
            let hasData = false;

            headers.forEach((header, index) => {
                const value = values[index];
                rowData[header] = value;

                // Track if row has any text content
                if (header.startsWith('text') && value) {
                    hasData = true;
                }
                // Validate font size values
                if (header.startsWith('fontsize') && value) {
                    const size = parseFloat(value);
                    if (isNaN(size)) {
                        console.warn(`[DEBUG] Invalid font size in row ${i + 1}, column ${header}: ${value}`);
                        rowData[header] = ''; // Clear invalid font size
                    }
                }
            });

            if (hasData) {
                data.push(rowData);
            }
        }

        if (data.length === 0) {
            throw new PluginError(
                'No valid data rows found in CSV',
                'EMPTY_CSV_DATA',
                { totalLines: lines.length }
            );
        }

        // Update state
        textReplaceState.data.csvData = data;
        textReplaceState.status.steps.csvLoaded = true;
        textReplaceState.status.performance.csvLoadTime = Date.now() - startTime;
        textReplaceState.status.performance.totalRows = data.length;

        console.log("[DEBUG] CSV load completed:", {
            rowsLoaded: data.length,
            duration: Date.now() - startTime,
            timestamp: new Date().toISOString()
        });

        // Write to MCP relay
        await writeToMCPRelay({
            command: "sendLog",
            message: `CSV loaded successfully: ${data.length} rows`,
            metrics: {
                duration: Date.now() - startTime,
                rowCount: data.length,
                textColumns: textColumns.length,
                fontSizeColumns: fontSizeColumns.length
            },
            timestamp: new Date().toISOString()
        });

        log(`[Cursor OK] CSV loaded: ${data.length} rows, ${textColumns.length} text columns`);
        return data;

    } catch (error) {
        console.error("[DEBUG] CSV load failed:", {
            error: error.message,
            code: error.code,
            file: file.name
        });

        textReplaceState.status.steps.csvLoaded = false;
        throw new PluginError(
            'Failed to load CSV file',
            'CSV_LOAD_ERROR',
            { originalError: error, file: file.name }
        );
    }
}

// Enhanced CSV Loading for Image Replace
async function loadImageCSV(file) {
    const startTime = startOperation(imageReplaceState, 'csvLoad');
    try {
        const fileContent = await file.read();
        const lines = fileContent.split('\n');
        const headers = lines[0].split(',');
        const data = lines.slice(1).map(line => {
            const values = line.split(',');
            return headers.reduce((obj, header, index) => {
                obj[header.trim()] = values[index]?.trim() || '';
                return obj;
            }, {});
        });
        
        // Update state with performance metrics
        imageReplaceState.data.csvData = data;
        imageReplaceState.status.steps.csvLoaded = true;
        imageReplaceState.status.performance.totalRows = data.length;
        
        const duration = endOperation(imageReplaceState, 'csvLoad');
        log(`[Cursor OK] Image CSV loaded successfully (${duration}ms, ${data.length} rows)`);
        
        // Write performance data to MCP relay
        await writeToMCPRelay({
            command: "sendMetrics",
            tab: "image",
            operation: "csvLoad",
            metrics: {
                duration,
                rowCount: data.length,
                timestamp: new Date().toISOString()
            }
        });
        
        return data;
    } catch (error) {
        imageReplaceState.status.steps.csvLoaded = false;
        log(`Error loading image CSV: ${error.message}`);
        throw error;
    }
}

// Updated PNG save function with proper API v2 format based on official documentation
async function saveAsPNG(doc, outputPath) {
    try {
        log(`[DEBUG] Starting PNG save to: ${outputPath}`);
        
        // First ensure we have a valid file token
        const fs = require('uxp').storage.localFileSystem;
        const tempFolder = await fs.getTemporaryFolder();
        
        // Parse the output path to get just the filename
        const pathParts = outputPath.split('/');
        const fileName = pathParts[pathParts.length - 1];
        
        // Create the file entry in the output directory
        const outputDir = pathParts.slice(0, -1).join('/');
        log(`[DEBUG] Attempting to access output directory: ${outputDir}`);
        
        let outputFile;
        try {
            // Try to get the directory using the path
            const outputDirEntry = await fs.getEntryWithUrl(`file:${outputDir}`);
            if (!outputDirEntry || !outputDirEntry.isFolder) {
                throw new Error('Invalid output directory');
            }
            
            // Create the file in the output directory
            outputFile = await outputDirEntry.createFile(fileName, { overwrite: true });
            log(`[DEBUG] Created file entry at: ${outputFile.nativePath}`);
        } catch (dirError) {
            log(`[DEBUG] Error accessing output directory: ${dirError.message}`);
            log(`[DEBUG] Attempting fallback to temporary folder...`);
            
            // Fallback to temp folder
            outputFile = await tempFolder.createFile(fileName, { overwrite: true });
            log(`[DEBUG] Created temporary file at: ${outputFile.nativePath}`);
        }
        
        // Create a session token for the file
        const sessionToken = fs.createSessionToken(outputFile);
        log(`[DEBUG] Created session token for file`);
        
        // Use the proper save method from the documentation
        const saveDesc = {
            _obj: "exportDocument",
            documentID: doc._id,
            format: {
                _obj: "PNG",
                PNG8: false,
                transparency: true,
                interlaced: false,
                quality: 100
            },
            _path: sessionToken,
            _options: { 
                dialogOptions: "dontDisplay"
            }
        };

        log("[DEBUG] Executing PNG save command with session token...");
        
        const result = await batchPlay(
            [saveDesc],
            {
                synchronousExecution: true,
                modalBehavior: "none"
            }
        );

        log(`[DEBUG] PNG Save completed successfully: ${outputFile.nativePath}`);
        return result;

    } catch (error) {
        log(`[DEBUG] PNG Save Failed: ${error.message}`);
        
        // Fallback method for PS 2025
        try {
            log(`[DEBUG] Attempting fallback PNG save method...`);
            
            // Get the proper constants
            const saveOptions = {
                _obj: "save",
                as: {
                    _obj: "PNGFormat",
                    PNG8: false,
                    transparency: true
                },
                in: { _path: outputPath },
                copy: true,
                saveStages: {
                    _enum: "saveStagesType",
                    _value: "saveSucceeded"
                },
                _options: { 
                    dialogOptions: "dontDisplay"
                }
            };
            
            const fallbackResult = await batchPlay(
                [saveOptions],
                {
                    synchronousExecution: true,
                    modalBehavior: "none"
                }
            );
            
            log(`[DEBUG] Fallback PNG Save completed successfully: ${outputPath}`);
            return fallbackResult;
            
        } catch (fallbackError) {
            log(`[DEBUG] Fallback PNG Save also failed: ${fallbackError.message}`);
            
            // Try one more approach based on the Adobe forums
            try {
                log(`[DEBUG] Attempting last-resort save method using file token...`);
                
                const fs = require('uxp').storage.localFileSystem;
                const tempFolder = await fs.getTemporaryFolder();
                const fileName = outputPath.split('/').pop();
                const tempFile = await tempFolder.createFile(fileName, { overwrite: true });
                const saveToken = fs.createSessionToken(tempFile);
                
                const lastResortSaveDesc = {
                    _obj: "save",
                    as: {
                        _obj: "PNGFormat",
                        PNG8: false
                    },
                    in: { _path: saveToken },
                    copy: true,
                    _options: { 
                        dialogOptions: "dontDisplay"
                    }
                };
                
                const lastResult = await batchPlay(
                    [lastResortSaveDesc],
                    {
                        synchronousExecution: true,
                        modalBehavior: "none"
                    }
                );
                
                log(`[DEBUG] Last-resort PNG Save completed to temp: ${tempFile.nativePath}`);
                return lastResult;
                
            } catch (lastError) {
                log(`[DEBUG] All PNG save methods failed`);
                throw new PluginError(
                    'Failed to save PNG (all methods failed)',
                    'PNG_SAVE_ERROR',
                    { 
                        originalError: error, 
                        fallbackError, 
                        lastError,
                        outputPath 
                    }
                );
            }
        }
    }
}

// Updated PSD save function with proper API v2 format
async function saveAsPSD(doc, outputPath) {
    try {
        log(`[DEBUG] Starting PSD save to: ${outputPath}`);
        
        // Check if outputPath is already a session token
        let sessionToken = outputPath;
        
        // If it's a file path, create a file entry and get a session token
        if (typeof outputPath === 'string' && outputPath.includes('/')) {
            const fs = require('uxp').storage.localFileSystem;
            const tempFolder = await fs.getTemporaryFolder();
            
            // Parse the output path to get just the filename
            const pathParts = outputPath.split('/');
            const fileName = pathParts[pathParts.length - 1];
            
            // Create the file entry in the output directory
            const outputDir = pathParts.slice(0, -1).join('/');
            log(`[DEBUG] Attempting to access output directory for PSD: ${outputDir}`);
            
            let outputFile;
            try {
                // Try to get the directory using the path
                const outputDirEntry = await fs.getEntryWithUrl(`file:${outputDir}`);
                if (!outputDirEntry || !outputDirEntry.isFolder) {
                    throw new Error('Invalid output directory');
                }
                
                // Create the file in the output directory
                outputFile = await outputDirEntry.createFile(fileName, { overwrite: true });
                log(`[DEBUG] Created PSD file entry at: ${outputFile.nativePath}`);
            } catch (dirError) {
                log(`[DEBUG] Error accessing output directory for PSD: ${dirError.message}`);
                log(`[DEBUG] Attempting fallback to temporary folder...`);
                
                // Fallback to temp folder
                outputFile = await tempFolder.createFile(fileName, { overwrite: true });
                log(`[DEBUG] Created temporary PSD file at: ${outputFile.nativePath}`);
            }
            
            // Create a session token for the file
            sessionToken = fs.createSessionToken(outputFile);
            log(`[DEBUG] Created session token for PSD file`);
        }
        
        // Create the save descriptor with proper format
        const saveDesc = {
            _obj: "save",
            as: {
                _obj: "photoshop35Format",
                alphaChannels: true,
                embedColorProfile: true,
                layers: true,
                maximizeCompatibility: true
            },
            in: { _path: sessionToken },
            copy: true,
            lowerCase: true,
            _options: { 
                dialogOptions: "dontDisplay"
            }
        };

        log("[DEBUG] Executing PSD save command...");
        
        const result = await batchPlay(
            [saveDesc],
            {
                synchronousExecution: true,
                modalBehavior: "none"
            }
        );

        log(`[DEBUG] PSD Save completed successfully`);
        return result;

    } catch (error) {
        log(`[DEBUG] PSD Save Failed: ${error.message}`);
        
        // Try fallback method
        try {
            log(`[DEBUG] Attempting fallback PSD save method...`);
            
            const fs = require('uxp').storage.localFileSystem;
            const tempFolder = await fs.getTemporaryFolder();
            const fileName = typeof outputPath === 'string' ? 
                outputPath.split('/').pop() : 
                `document_${Date.now()}.psd`;
                
            const tempFile = await tempFolder.createFile(fileName, { overwrite: true });
            const saveToken = fs.createSessionToken(tempFile);
            
            const fallbackSaveDesc = {
                _obj: "save",
                as: {
                    _obj: "photoshop35Format",
                    maximizeCompatibility: true
                },
                in: { _path: saveToken },
                copy: true,
                _options: { 
                    dialogOptions: "dontDisplay"
                }
            };
            
            const fallbackResult = await batchPlay(
                [fallbackSaveDesc],
                {
                    synchronousExecution: true,
                    modalBehavior: "none"
                }
            );
            
            log(`[DEBUG] Fallback PSD Save completed to temp: ${tempFile.nativePath}`);
            return fallbackResult;
            
        } catch (fallbackError) {
            log(`[DEBUG] All PSD save methods failed`);
            throw new PluginError(
                'Failed to save PSD (all methods failed)',
                'PSD_SAVE_ERROR',
                { 
                    originalError: error, 
                    fallbackError,
                    outputPath 
                }
            );
        }
    }
}

// Add document initialization function
async function ensureDocumentInitialized() {
    try {
        const doc = app.activeDocument;
        if (!doc) {
            throw new PluginError('No active document found', 'NO_ACTIVE_DOC');
        }

        // Check if document has a valid ID
        if (!doc._id) {
            log('[DEBUG] Document has no ID - performing initial save...');
            
            // Get the document name or use a default
            const docName = doc.name || 'Untitled';
            const timestamp = new Date().getTime();
            const tempName = `${docName.replace('.psd', '')}_temp_${timestamp}.psd`;
            
            // Create temp folder if it doesn't exist
            const tempFolder = await fs.getTemporaryFolder();
            const savePath = `${tempFolder.nativePath}/${tempName}`;
            
            log(`[DEBUG] Attempting to initialize document with temp save to: ${savePath}`);
            
            // Perform initial save to get document ID - try multiple approaches for M1 compatibility
            try {
                // First approach - standard save
                const saveCommand = {
                    _obj: "save",
                    as: {
                        _obj: "photoshop35Format",
                        maximizeCompatibility: true
                    },
                    in: { _path: savePath },
                    copy: true,
                    lowerCase: true,
                    _options: { 
                        dialogOptions: "dontDisplay"
                    }
                };

                await batchPlay(
                    [saveCommand],
                    {
                        synchronousExecution: true,
                        modalBehavior: "none"
                    }
                );
            } catch (saveError) {
                log(`[DEBUG] Standard save failed: ${saveError.message}, trying alternative approach...`);
                
                // Second approach - alternative save format for M1
                try {
                    const altSaveCommand = {
                        _obj: "save",
                        as: {
                            _obj: "photoshop35Format",
                            maximizeCompatibility: true
                        },
                        in: savePath,
                        documentID: app.activeDocument._id,
                        copy: true,
                        _options: { 
                            dialogOptions: "dontDisplay"
                        }
                    };
    
                    await batchPlay(
                        [altSaveCommand],
                        {
                            synchronousExecution: true,
                            modalBehavior: "none"
                        }
                    );
                } catch (altSaveError) {
                    log(`[DEBUG] Alternative save also failed: ${altSaveError.message}`);
                    throw altSaveError;
                }
            }

            log(`[Cursor OK] Document initialized with temporary save: ${tempName}`);
            
            // Force refresh document reference
            const refreshedDoc = app.activeDocument;
            
            // Verify document now has ID
            if (!refreshedDoc || !refreshedDoc._id) {
                throw new PluginError('Failed to initialize document ID', 'DOC_INIT_FAILED');
            }

            return true;
        }

        return true;
    } catch (error) {
        log(`[DEBUG] Document initialization failed: ${error.message}`);
        throw new PluginError(
            'Failed to initialize document',
            'DOC_INIT_ERROR',
            { originalError: error }
        );
    }
}

// Update processTextRow to use document initialization
async function processTextRow(row, index, total, folders) {
    console.log("[DEBUG] Starting row processing:", {
        rowIndex: index,
        totalRows: total,
        timestamp: new Date().toISOString()
    });

    // First ensure document is properly initialized
    try {
        await ensureDocumentInitialized();
    } catch (error) {
        log(`[DEBUG] Failed to initialize document: ${error.message}`);
        throw error;
    }

    const doc = app.activeDocument;
    if (!doc) {
        throw new PluginError('No active document found', 'NO_ACTIVE_DOC');
    }

    // Get document info using the correct API v2 properties
    const docInfo = {
        id: doc._id,
        name: doc.name || 'Untitled',
        path: doc.path || null,
        layerCount: doc.layers?.length || 0
    };

    log(`[DEBUG] Document initialized with ID: ${docInfo.id}`);

    const layers = doc.layers;
    const processingStart = Date.now();
    const layerUpdates = [];
    const errors = [];

    try {
        // First pass: identify text layers with enhanced logging
        const textLayers = layers.filter(layer => {
            const isTextLayer = layer.kind === 'text' || layer.kind === 3;
            const match = layer.name.match(/\d+/);
            const index = match?.[0];
            
            console.log("[DEBUG] Layer analysis:", {
                name: layer.name,
                id: layer._id,  // Use _id for API v2
                kind: layer.kind,
                isTextLayer,
                match,
                index,
                hasTextData: index ? !!row[`text${index}`] : false,
                hasFontData: index ? !!row[`fontsize${index}`] : false
            });
            
            return isTextLayer && index;
        }).map(layer => ({
            layer,
            index: layer.name.match(/\d+/)[0],
            hasData: !!row[`text${layer.name.match(/\d+/)[0]}`] || !!row[`fontsize${layer.name.match(/\d+/)[0]}`]
        }));

        if (textLayers.length === 0) {
            throw new PluginError('No text layers found in document', 'NO_TEXT_LAYERS');
        }

        console.log("[DEBUG] Found text layers:", textLayers.map(({ layer, index, hasData }) => ({
            name: layer.name,
            index,
            hasData,
            textContent: row[`text${index}`],
            fontSize: row[`fontsize${index}`]
        })));

        // Process each layer, continuing even if one fails
        for (const { layer, index: layerIndex } of textLayers) {
            if (row[`text${layerIndex}`] || row[`fontsize${layerIndex}`]) {
                try {
                    const result = await processLayer(layer, row, layerIndex);
                    layerUpdates.push(result);
                } catch (layerError) {
                    console.error("[DEBUG] Layer processing failed but continuing:", {
                        layer: layer.name,
                        error: layerError.message,
                        code: layerError.code
                    });
                    errors.push({
                        layer: layer.name,
                        error: layerError.message,
                        code: layerError.code
                    });
                    // Continue with next layer instead of throwing
                    continue;
                }
            }
        }

        // Only attempt saves if at least one layer was processed successfully
        if (layerUpdates.length > 0) {
            try {
                // Get text content for filename
                const text1Content = row.text1 || 'default';
                const text2Content = row.text2 || '';
                
                // Generate filenames for output
                const baseFileName = `${text1Content}_${text2Content}`.replace(/[^a-zA-Z0-9]/g, '_');
                
                // Use the normalized paths for M1 compatibility
                const pngPath = `${folders.pngNativePath || folders.pngFolder.nativePath}/${baseFileName}.png`;
                const psdPath = `${folders.psdNativePath || folders.psdFolder.nativePath}/${baseFileName}.psd`;

                log("[DEBUG] Starting file saves:", {
                    filename: baseFileName,
                    png: pngPath,
                    psd: psdPath,
                    text1: text1Content,
                    text2: text2Content,
                    timestamp: new Date().toISOString()
                });

                // Create file entries first
                let pngFile, psdFile;
                try {
                    // Create PNG file entry
                    pngFile = await folders.pngFolder.createFile(`${baseFileName}.png`, { overwrite: true });
                    log(`[DEBUG] Created PNG file entry: ${pngFile.nativePath}`);
                    
                    // Create PSD file entry
                    psdFile = await folders.psdFolder.createFile(`${baseFileName}.psd`, { overwrite: true });
                    log(`[DEBUG] Created PSD file entry: ${psdFile.nativePath}`);
                } catch (fileCreateError) {
                    log(`[DEBUG] Error creating file entries: ${fileCreateError.message}`);
                    // Continue with path-based approach as fallback
                }

                // Save PNG - use file entry if available, otherwise use path
                if (pngFile) {
                    const fs = require('uxp').storage.localFileSystem;
                    const pngToken = fs.createSessionToken(pngFile);
                    await saveAsPNG(doc, pngToken);
                    log(`[DEBUG] ✅ PNG saved successfully using file token: ${pngFile.nativePath}`);
                } else {
                    await saveAsPNG(doc, pngPath);
                    log(`[DEBUG] ✅ PNG saved successfully using path: ${pngPath}`);
                }

                // Save PSD - use file entry if available, otherwise use path
                if (psdFile) {
                    const fs = require('uxp').storage.localFileSystem;
                    const psdToken = fs.createSessionToken(psdFile);
                    await saveAsPSD(doc, psdToken);
                    log(`[DEBUG] ✅ PSD saved successfully using file token: ${psdFile.nativePath}`);
                } else {
                    await saveAsPSD(doc, psdPath);
                    log(`[DEBUG] ✅ PSD saved successfully using path: ${psdPath}`);
                }

                // Write success to MCP relay
                await writeToMCPRelay({
                    command: "sendLog",
                    message: `Files saved successfully for ${baseFileName}`,
                    data: {
                        filename: baseFileName,
                        png: pngFile ? pngFile.nativePath : pngPath,
                        psd: psdFile ? psdFile.nativePath : psdPath,
                        timestamp: new Date().toISOString()
                    }
                });

            } catch (saveError) {
                log(`[DEBUG] ❌ Save operation failed: ${saveError.message}`);
                errors.push({
                    type: 'SAVE_ERROR',
                    error: saveError.message,
                    code: saveError.code
                });
            }
        }

        return {
            success: layerUpdates.length > 0,
            processedLayers: layerUpdates.length,
            failedLayers: errors.length,
            errors,
            textLayers: textLayers.length,
            layerUpdates,
            duration: Date.now() - processingStart
        };

    } catch (error) {
        console.error("[DEBUG] Row processing error:", {
            rowIndex: index,
            error: error.message,
            code: error.code,
            duration: Date.now() - processingStart
        });

        throw new PluginError(
            'Row processing error',
            'ROW_PROCESS_ERROR',
            {
                originalError: error,
                rowIndex: index,
                processedLayers: layerUpdates.length,
                failedLayers: errors.length,
                errors,
                duration: Date.now() - processingStart
            }
        );
    }
}

// Update processImageRow to use document initialization
async function processImageRow(row, index, total) {
    if (!row || index === undefined || total === undefined) {
        throw new PluginError(
            'Invalid image row processing parameters',
            'INVALID_ROW_PARAMS',
            { index, total, hasRow: !!row }
        );
    }

    // Ensure document is initialized before processing
    await ensureDocumentInitialized();

    const doc = app.activeDocument;
    if (!doc) {
        throw new PluginError('No active document found', 'NO_DOCUMENT');
    }

    const layers = doc.layers;
    const startTime = Date.now();
    
    try {
        let processedLayers = 0;
        for (const layer of layers) {
            if (layer.kind === 'smartObject') {
                const layerIndex = layer.name.match(/\d+/)?.[0];
                if (layerIndex && row[`imgname${layerIndex}`]) {
                    const imagePath = `${imageReplaceState.data.outputFolder.nativePath}/${row[`imgname${layerIndex}`]}`;
                    await replaceImage(layer, imagePath);
                    processedLayers++;
                }
            }
        }
        
        // Update progress with validation
        if (imageReplaceState.status.performance) {
            imageReplaceState.status.performance.processedRows++;
            imageReplaceState.data.lastProcessedRow = row;
        }
        
        // Update status with progress percentage
        const progress = Math.round((index + 1) / total * 100);
        const statusElement = document.getElementById('imageStatus');
        if (statusElement) {
            statusElement.textContent = `Processing: ${progress}% complete (${processedLayers} layers updated)`;
        }
        
        // Log progress metrics with enhanced data
        const duration = Date.now() - startTime;
        await writeToMCPRelay({
            command: "sendProgress",
            tab: "image",
            progress: {
                current: index + 1,
                total,
                percentage: progress,
                rowDuration: duration,
                processedLayers,
                timestamp: new Date().toISOString()
            }
        });
        
    } catch (error) {
        throw new PluginError(
            `Error processing image row ${index + 1}/${total}`,
            'IMAGE_ROW_PROCESS_ERROR',
            { row, error, index, total }
        );
    }
}

// Update processTextReplacement to pass folders to processTextRow
async function processTextReplacement() {
    const textStatus = document.getElementById('textStatus');
    const processType = document.querySelector('input[name="processType"]:checked')?.value;
    
    if (!processType) {
        log('[Cursor OK] Process type not selected');
        return;
    }
    
    try {
        // Validate state with detailed errors
        if (!textReplaceState.data.csvData || !textReplaceState.data.csvData.length) {
            throw new PluginError('No CSV data available', 'CSV_NOT_LOADED');
        }

        if (!textReplaceState.data.outputFolder) {
            throw new PluginError('Output folder not selected', 'FOLDER_NOT_SELECTED');
        }

        const doc = app.activeDocument;
        if (!doc) {
            throw new PluginError('No active document found', 'NO_DOCUMENT');
        }

        // Initialize document once at the start
        textStatus.textContent = 'Initializing document...';
        try {
            await ensureDocumentInitialized();
            log('[DEBUG] Document initialized successfully');
        } catch (initError) {
            log(`[DEBUG] Document initialization failed: ${initError.message}`);
            throw initError;
        }

        // Update state and UI
        textReplaceState.status.isProcessing = true;
        textReplaceState.status.currentOperation = 'text_replacement';
        textStatus.textContent = 'Setting up output folders...';
        
        // Setup output folders using text-specific function
        let folders;
        try {
            folders = await setupTextOutputFolders(textReplaceState.data.outputFolder);
        } catch (folderError) {
            console.error("[DEBUG] Folder setup failed:", folderError);
            throw new PluginError(
                'Failed to setup output folders',
                'FOLDER_SETUP_FAILED',
                { originalError: folderError }
            );
        }
        
        textStatus.textContent = 'Processing...';
        const startTime = Date.now();
        
        try {
            if (processType === 'current') {
                const currentRow = textReplaceState.data.csvData[0];
                if (!currentRow) {
                    throw new PluginError('No data in first row', 'EMPTY_ROW');
                }
                
                await processTextRow(currentRow, 0, 1, folders);
                textStatus.textContent = 'Current row processed successfully';
                
            } else {
                const total = textReplaceState.data.csvData.length;
                log("[DEBUG] Starting batch processing:", {
                    totalRows: total,
                    timestamp: new Date().toISOString()
                });
                
                for (let i = 0; i < total; i++) {
                    const row = textReplaceState.data.csvData[i];
                    if (!row) {
                        throw new PluginError(`Empty row at index ${i}`, 'EMPTY_ROW');
                    }
                    
                    textStatus.textContent = `Processing row ${i + 1} of ${total}...`;
                    await processTextRow(row, i, total, folders);
                }
                textStatus.textContent = 'All rows processed successfully';
            }

            const duration = Date.now() - startTime;
            
            await writeToMCPRelay({
                command: "sendLog",
                message: `Text replacement completed successfully. Type: ${processType}`,
                metrics: {
                    duration,
                    processType,
                    totalRows: textReplaceState.data.csvData.length,
                    processedRows: textReplaceState.status.performance.processedRows,
                    outputFolders: {
                        png: folders.pngFolder.nativePath,
                        psd: folders.psdFolder.nativePath
                    }
                },
                timestamp: new Date().toISOString()
            });

            log('[Cursor OK] Text replacement completed');
            log(`[Cursor OK] Files saved to:`);
            log(`  PNG folder: ${folders.pngFolder.nativePath}`);
            log(`  PSD folder: ${folders.psdFolder.nativePath}`);
        } catch (error) {
            console.error("[DEBUG] Text replacement error:", error);
            throw error;
        }
    } catch (error) {
        console.error("[DEBUG] Text replacement error:", error);
        throw error;
    }
}

// Update processImageReplacement to use document initialization
async function processImageReplacement() {
    const imageStatus = document.getElementById('imageStatus');
    const processType = document.querySelector('input[name="processTypeImg"]:checked')?.value;
    
    try {
        // Validate state
        if (!imageReplaceState.data.csvData) {
            throw new PluginError('Please load a CSV file first', 'CSV_NOT_LOADED');
        }

        if (!imageReplaceState.data.outputFolder) {
            throw new PluginError('Please select an output folder first', 'FOLDER_NOT_SELECTED');
        }

        const doc = app.activeDocument;
        if (!doc) {
            throw new PluginError('No active document found', 'NO_DOCUMENT');
        }

        // Update state
        imageReplaceState.status.isProcessing = true;
        imageReplaceState.status.currentOperation = 'image_replacement';
        
        // Setup output folders using image-specific function
        imageStatus.textContent = 'Setting up output folders...';
        const folders = await setupImageOutputFolders(imageReplaceState.data.outputFolder);
        
        imageStatus.textContent = 'Processing...';

        if (processType === 'current') {
            const currentRow = imageReplaceState.data.csvData[0];
            await processImageRow(currentRow);
            imageStatus.textContent = 'Current row processed successfully';
        } else {
            for (const row of imageReplaceState.data.csvData) {
                await processImageRow(row);
            }
            imageStatus.textContent = 'All rows processed successfully';
        }

        // Write to MCP relay
        await writeToMCPRelay({
            command: "sendLog",
            message: `Image replacement completed. Type: ${processType}`,
            timestamp: new Date().toISOString()
        });

        log('[Cursor OK] Image replacement completed');
    } catch (error) {
        await handleError(error, 'image');
    } finally {
        imageReplaceState.status.isProcessing = false;
        imageReplaceState.status.currentOperation = null;
    }
}

// Event Listeners for Text Replace
document.getElementById('loadCSV').addEventListener('click', async () => {
    try {
        const file = await fs.getFileForOpening({ types: ['csv'] });
        if (file) {
            await loadTextCSV(file);
            document.getElementById('textStatus').textContent = 'CSV file loaded successfully';
        }
    } catch (error) {
        log(`[Cursor OK] Error loading CSV: ${error.message}`);
        document.getElementById('textStatus').textContent = 'Error loading CSV file';
    }
});

// Event Listeners for Image Replace
document.getElementById('loadInputFolder').addEventListener('click', async () => {
    try {
        imageReplaceState.data.outputFolder = await fs.getFolder();
        if (imageReplaceState.data.outputFolder) {
            log(`[Cursor OK] Input folder selected: ${imageReplaceState.data.outputFolder.nativePath}`);
            document.getElementById('imageStatus').textContent = `Input folder: ${imageReplaceState.data.outputFolder.nativePath}`;
        }
    } catch (error) {
        log(`Error selecting input folder: ${error.message}`);
        document.getElementById('imageStatus').textContent = 'Error selecting input folder';
    }
});

document.getElementById('loadCSVImages').addEventListener('click', async () => {
    try {
        const file = await fs.getFileForOpening({ types: ['csv'] });
        if (file) {
            await loadImageCSV(file);
            document.getElementById('imageStatus').textContent = 'CSV file loaded successfully';
        }
    } catch (error) {
        log(`Error loading CSV: ${error.message}`);
        document.getElementById('imageStatus').textContent = 'Error loading CSV file';
    }
});

// Enhanced folder selection handler
document.getElementById('selectOutputFolder').addEventListener('click', async () => {
    try {
        console.log("[DEBUG] Requesting output folder selection...");
        textReplaceState.data.outputFolder = await fs.getFolder();
        
        if (textReplaceState.data.outputFolder) {
            console.log("[DEBUG] Setting up output folders in:", textReplaceState.data.outputFolder.nativePath);
            
            // Immediately try to setup the folders
            const folders = await setupTextOutputFolders(textReplaceState.data.outputFolder);
            
            textReplaceState.status.steps.outputFolderSelected = true;
            log(`[Cursor OK] Text output folders created:`);
            log(`  PNG: ${folders.pngFolder.nativePath}`);
            log(`  PSD: ${folders.psdFolder.nativePath}`);
            
            document.getElementById('textStatus').textContent = `Output folders ready: ${textReplaceState.data.outputFolder.nativePath}`;
        }
    } catch (error) {
        console.error("[DEBUG] Output folder setup failed:", error);
        textReplaceState.status.steps.outputFolderSelected = false;
        log(`Error setting up output folders: ${error.message}`);
        document.getElementById('textStatus').textContent = 'Error setting up output folders';
        
        // Attempt to show more specific error message
        if (error.code === 'FOLDER_CREATE_ERROR') {
            document.getElementById('textStatus').textContent = 'Failed to create output folders. Please check folder permissions.';
        } else if (error.code === 'FOLDER_VERIFY_ERROR') {
            document.getElementById('textStatus').textContent = 'Failed to verify output folders were created.';
        }
    }
});

document.getElementById('selectOutputFolderImg').addEventListener('click', async () => {
    try {
        imageReplaceState.data.outputFolder = await fs.getFolder();
        if (imageReplaceState.data.outputFolder) {
            imageReplaceState.status.steps.outputFolderSelected = true;
            log(`[Cursor OK] Image output folder selected: ${imageReplaceState.data.outputFolder.nativePath}`);
            document.getElementById('imageStatus').textContent = `Output folder: ${imageReplaceState.data.outputFolder.nativePath}`;
        }
    } catch (error) {
        imageReplaceState.status.steps.outputFolderSelected = false;
        log(`Error selecting output folder: ${error.message}`);
        document.getElementById('imageStatus').textContent = 'Error selecting output folder';
    }
});

// Process button event listeners
document.getElementById('processText').addEventListener('click', processTextReplacement);
document.getElementById('processImages').addEventListener('click', processImageReplacement);

// Restart Plugin button event listener
document.getElementById('restartPlugin').addEventListener('click', async () => {
    const button = document.getElementById('restartPlugin');
    const originalText = button.textContent;
    
    try {
        // Disable button and show loading state
        button.disabled = true;
        button.textContent = 'Sending restart command...';
        
        // Send restart command
        await writeToMCPRelay({
            command: "restartPlugin",
            timestamp: new Date().toISOString()
        });
        
        // Log success
        log("🌀 Restart command sent to MCP");
        
        // Show success state briefly
        button.textContent = '✅ Restart command sent';
        setTimeout(() => {
            button.textContent = originalText;
            button.disabled = false;
        }, 2000);
        
    } catch (error) {
        // Log error
        log(`❌ Error sending restart command: ${error.message}`);
        
        // Show error state
        button.textContent = '❌ Failed to send restart command';
        setTimeout(() => {
            button.textContent = originalText;
            button.disabled = false;
        }, 2000);
        
        // Log detailed error to console
        console.error('Restart command failed:', error);
    }
});

// Simple tab management
function initializeTabs() {
    // Get tab elements
    const textReplaceTab = document.getElementById('textReplaceTab');
    const imageReplaceTab = document.getElementById('imageReplaceTab');
    const logsTab = document.getElementById('logsTab');

    // Get panel elements
    const textReplacePanel = document.getElementById('textReplacePanel');
    const imageReplacePanel = document.getElementById('imageReplacePanel');
    const logsPanel = document.getElementById('logsPanel');

    // Function to switch tabs
    function switchTab(tabId) {
        // Hide all panels
        [textReplacePanel, imageReplacePanel, logsPanel].forEach(panel => {
            if (panel) panel.style.display = 'none';
        });

        // Deactivate all tabs
        [textReplaceTab, imageReplaceTab, logsTab].forEach(tab => {
            if (tab) tab.classList.remove('active');
        });

        // Show selected panel and activate tab
        const selectedPanel = document.getElementById(`${tabId}Panel`);
        const selectedTab = document.getElementById(`${tabId}Tab`);
        
        if (selectedPanel) selectedPanel.style.display = 'block';
        if (selectedTab) selectedTab.classList.add('active');
    }

    // Add click handlers
    if (textReplaceTab) textReplaceTab.addEventListener('click', () => switchTab('textReplace'));
    if (imageReplaceTab) imageReplaceTab.addEventListener('click', () => switchTab('imageReplace'));
    if (logsTab) logsTab.addEventListener('click', () => switchTab('logs'));

    // Set initial tab
    switchTab('textReplace');
}

// Initialize plugin
document.addEventListener('DOMContentLoaded', () => {
    log('[Cursor OK] Plugin initialized');
    
    // Initialize tabs
    initializeTabs();

    // Write to MCP relay
    writeToMCPRelay({
        command: "sendLog",
        message: "Plugin loaded and UI initialized.",
        timestamp: new Date().toISOString()
    }).catch(error => {
        log(`[Cursor OK] MCP relay write failed: ${error.message}`);
    });
});