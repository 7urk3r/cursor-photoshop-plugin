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

// Add a global stop processing flag
let stopProcessingRequested = false;

// Function to check if processing should stop
function shouldStopProcessing() {
    return stopProcessingRequested;
}

// Function to request stopping the processing
function requestStopProcessing() {
    stopProcessingRequested = true;
    log("[DEBUG] Stop processing requested by user");
    
    // Update UI to show stopping
    const statusElement = document.getElementById('textStatus');
    if (statusElement) {
        statusElement.textContent = 'Stopping...';
        statusElement.style.color = 'orange';
    }
    
    // Update stop button
    const stopButton = document.getElementById('stopProcessingBtn');
    if (stopButton) {
        stopButton.disabled = true;
        stopButton.textContent = 'Stopping...';
    }
    
    return true;
}

// Function to reset stop processing flag
function resetStopProcessing() {
    stopProcessingRequested = false;
    
    // Update UI elements
    const stopButton = document.getElementById('stopProcessingBtn');
    if (stopButton) {
        stopButton.disabled = false;
        stopButton.textContent = 'Stop Processing';
        stopButton.style.display = 'none';
    }
    
    return true;
}

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

// Image replacement function removed - plugin now focuses on text only

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
        
        // Verify PNG folder is accessible
        try {
            const pngExists = await folders.pngFolder.isEntry;
            
            if (!pngExists) {
                throw new PluginError(
                    'Could not verify PNG folder access', 
                    'FOLDER_ACCESS_ERROR',
                    {
                        pngExists,
                        pngPath: folders.pngFolder?.nativePath
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

        console.log("[DEBUG] Output folder ready:", {
            png: folders.pngNativePath
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
// Image CSV loading function removed - plugin now focuses on text only

// FIXED: Simple and reliable PNG save function using session tokens
async function saveAsPNG(doc, outputPath) {
    try {
        log(`[DEBUG] Starting PNG save to: ${outputPath}`);
        console.log(`[CURSOR SAVE] Starting PNG save to: ${outputPath}`);
        
        const fs = require('uxp').storage.localFileSystem;
        
        // Parse the path to get directory and filename
        const pathParts = outputPath.split('/');
        const fileName = pathParts.pop();
        const dirPath = pathParts.join('/');
        
        log(`[DEBUG] Parsed path - Directory: ${dirPath}, Filename: ${fileName}`);
        console.log(`[CURSOR SAVE] Parsed path - Directory: ${dirPath}, Filename: ${fileName}`);
        
        // Get the directory and create the file
        const directory = await fs.getFolder(dirPath);
        const outputFile = await directory.createFile(fileName, { overwrite: true });
        
        log(`[DEBUG] Created output file: ${outputFile.nativePath}`);
        console.log(`[CURSOR SAVE] Created output file: ${outputFile.nativePath}`);
        
        // CRITICAL FIX: Create a session token for batchPlay to use
        const sessionToken = fs.createSessionToken(outputFile);
        
        log(`[DEBUG] Created session token for file`);
        console.log(`[CURSOR SAVE] Created session token for file`);
        
        // Use compatible save command with PNG format
        const result = await batchPlay(
            [
                {
                    _obj: "save",
                    as: { 
                        _obj: "PNGFormat",
                        PNG8: false,
                        transparency: true,
                        interlaced: false,
                        compression: 6
                    },
                    in: sessionToken, // Use the session token here
                    documentID: doc._id,
                    copy: true,
                    lowerCase: true,
                    _options: { dialogOptions: "dontDisplay" }
                }
            ],
            { synchronousExecution: true }
        );
        
        log(`[DEBUG] ✅ PNG save completed successfully`);
        console.log(`[CURSOR SAVE] ✅ PNG save completed successfully`);
        
        return { success: true, method: "sessionToken", path: outputFile.nativePath, result };
        
    } catch (error) {
        log(`[DEBUG] ❌ PNG save failed: ${error.message}`);
        console.log(`[CURSOR SAVE] ❌ PNG save failed: ${error.message}`);
        throw new PluginError('PNG save failed', 'PNG_SAVE_ERROR', { error, outputPath });
    }
}

// PSD save function removed - no longer needed

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

// Add a new function to directly save files using the File API
async function directSaveFile(doc, outputPath, fileType) {
    try {
        console.log(`[CURSOR SAVE] DIRECT SAVE: Starting direct save to ${outputPath} as ${fileType}`);
        
        // Ensure we have a valid path
        if (!outputPath || typeof outputPath !== 'string') {
            throw new Error('Invalid output path for direct save');
        }
        
        // Get the file system module
        const fs = require('uxp').storage.localFileSystem;
        
        // Parse the path
        const pathParts = outputPath.split('/');
        const fileName = pathParts.pop();
        const dirPath = pathParts.join('/');
        
        console.log(`[CURSOR SAVE] DIRECT SAVE: Parsed path - Directory: ${dirPath}, Filename: ${fileName}`);
        
        // Get the directory
        const directory = await fs.getFolder(dirPath);
        console.log(`[CURSOR SAVE] DIRECT SAVE: Got directory: ${directory.nativePath}`);
        
        // Create the file entry
        const file = await directory.createFile(fileName, { overwrite: true });
        console.log(`[CURSOR SAVE] DIRECT SAVE: Created file entry: ${file.nativePath}`);
        
        // Create a file token
        const fileToken = fs.createSessionToken(file);
        console.log(`[CURSOR SAVE] DIRECT SAVE: Created file token`);
        
        // Save based on file type
        if (fileType.toLowerCase() === 'png') {
            // Save as PNG
            const exportDesc = {
                _obj: "exportDocument",
                documentID: doc._id,
                format: {
                    _obj: "PNG",
                    PNG8: false,
                    transparency: true,
                    interlaced: false,
                    quality: 100
                },
                in: fileToken,
                _options: { 
                    dialogOptions: "dontDisplay"
                }
            };
            
            console.log(`[CURSOR SAVE] DIRECT SAVE: Executing PNG export`);
            await batchPlay([exportDesc], { synchronousExecution: true });

        } else {
            throw new Error(`Unsupported file type: ${fileType}`);
        }
        
        // Verify the file exists
        try {
            const fileEntry = await directory.getEntry(fileName);
            const fileSize = await fileEntry.size;
            console.log(`[CURSOR SAVE] DIRECT SAVE: ✅ File saved and verified: ${fileEntry.nativePath}, size: ${fileSize} bytes`);
            return {
                success: true,
                method: "direct-save",
                path: fileEntry.nativePath,
                size: fileSize
            };
        } catch (verifyError) {
            console.log(`[CURSOR SAVE] DIRECT SAVE: ❌ File verification failed: ${verifyError.message}`);
            throw verifyError;
        }
    } catch (error) {
        console.log(`[CURSOR SAVE] DIRECT SAVE: ❌ Failed: ${error.message}`);
        console.log(`[CURSOR SAVE] DIRECT SAVE: Error stack: ${error.stack}`);
        throw error;
    }
}

// Update processTextRow to use the new direct save function
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
    let filesSaved = { png: false };

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
                
                // Ensure we have valid folder paths
                if (!folders || !folders.pngFolder) {
                    log(`[DEBUG] ❌ Invalid folder structure for saving: ${JSON.stringify(folders)}`);
                    console.log(`[CURSOR SAVE] ❌ Invalid folder structure for saving: ${JSON.stringify(folders)}`);
                    throw new Error('Invalid folder structure for saving');
                }
                
                // Get native paths with proper formatting
                const pngNativePath = folders.pngNativePath || folders.pngFolder.nativePath;
                
                // Ensure paths are properly formatted for M1 Mac
                const pngPath = `${pngNativePath.replace(/\\/g, '/')}/${baseFileName}.png`;

                log("[DEBUG] Starting file saves for row " + index + ":", {
                    filename: baseFileName,
                    png: pngPath,
                    text1: text1Content,
                    text2: text2Content,
                    timestamp: new Date().toISOString()
                });
                console.log(`[CURSOR SAVE] Starting file saves for row ${index}:`, {
                    filename: baseFileName,
                    png: pngPath
                });

                // Save PNG file
                let pngSaved = false;
                try {
                    console.log(`[CURSOR SAVE] Attempting PNG save...`);
                    const pngSaveResult = await saveAsPNG(doc, pngPath);
                    console.log(`[CURSOR SAVE] ✅ PNG save result:`, pngSaveResult);
                    pngSaved = true;
                    filesSaved.png = true;
                } catch (pngError) {
                    console.log(`[CURSOR SAVE] ❌ PNG save failed: ${pngError.message}`);
                    errors.push({
                        type: 'PNG_SAVE',
                        error: pngError.message,
                        path: pngPath
                    });
                }



                // Verify files were created
                try {
                    log(`[DEBUG] Verifying saved files exist...`);
                    console.log(`[CURSOR SAVE] Verifying saved files exist...`);
                    const fs = require('uxp').storage.localFileSystem;
                    
                    // Verify PNG file
                    if (pngSaved) {
                        try {
                            // Extract directory path and filename
                            const pathParts = pngPath.split('/');
                            const fileName = pathParts.pop();
                            const dirPath = pathParts.join('/');
                            
                            // Get the directory
                            const pngDir = await fs.getFolder(dirPath);
                            const pngExists = await pngDir.getEntry(fileName);
                            
                            if (pngExists) {
                                log(`[DEBUG] ✅ PNG file verified: ${pngExists.nativePath}`);
                                console.log(`[CURSOR SAVE] ✅ PNG file verified: ${pngExists.nativePath}`);
                                log(`[DEBUG] PNG file size: ${await pngExists.size} bytes`);
                                console.log(`[CURSOR SAVE] PNG file size: ${await pngExists.size} bytes`);
                                filesSaved.png = true;
                            } else {
                                log(`[DEBUG] ❌ PNG file not found after save`);
                                console.log(`[CURSOR SAVE] ❌ PNG file not found after save`);
                                filesSaved.png = false;
                            }
                        } catch (pngVerifyError) {
                            log(`[DEBUG] ❌ PNG file verification failed: ${pngVerifyError.message}`);
                            console.log(`[CURSOR SAVE] ❌ PNG file verification failed: ${pngVerifyError.message}`);
                            filesSaved.png = false;
                        }
                    }
                    

                } catch (verifyError) {
                    log(`[DEBUG] File verification error: ${verifyError.message}`);
                    console.log(`[CURSOR SAVE] File verification error: ${verifyError.message}`);
                }

                // Log file save summary
                if (pngSaved) {
                    log(`[DEBUG] ✅ Row ${index} processing complete with PNG file saved:`);
                    console.log(`[CURSOR SAVE] ✅ Row ${index} processing complete with PNG file saved:`);
                    log(`[DEBUG]   - PNG: ${pngPath}`);
                    console.log(`[CURSOR SAVE]   - PNG: ${pngPath}`);
                } else {
                    log(`[DEBUG] ❌ Row ${index} processing complete but NO FILES SAVED`);
                    console.log(`[CURSOR SAVE] ❌ Row ${index} processing complete but NO FILES SAVED`);
                }

                // Write success to MCP relay for external monitoring
                try {
                    await writeToMCPRelay({
                        status: 'success',
                        files: {
                            png: pngSaved ? pngPath : null
                        },
                        timestamp: new Date().toISOString()
                    });
                } catch (mcpError) {
                    log(`[DEBUG] MCP relay write failed: ${mcpError.message}`);
                }
            } catch (saveError) {
                log(`[DEBUG] Critical save error: ${saveError.message}`);
                console.log(`[CURSOR SAVE] Critical save error: ${saveError.message}`);
                errors.push({
                    type: 'save',
                    error: saveError.message
                });
            }
        }

        // Return processing results
        return {
            success: layerUpdates.length > 0,
            processed: layerUpdates.length,
            errors: errors.length > 0 ? errors : null,
            duration: Date.now() - processingStart,
            filesSaved: filesSaved
        };
    } catch (error) {
        log(`[DEBUG] Row processing error: ${error.message}`);
        console.log(`[CURSOR SAVE] Row processing error: ${error.message}`);
        throw error;
    }
}

// Test function for isolated PNG save testing
async function runSingleSaveTest() {
    log("Running single save test...");
    
    try {
        // Get the activeDocument and outputFolder from textReplaceState
        const doc = app.activeDocument;
        const outputFolder = textReplaceState.data.outputFolder;
        
        // Check if both document and folder exist
        if (!doc) {
            log("❌ Error: No active document found");
            return;
        }
        
        if (!outputFolder) {
            log("❌ Error: No output folder selected");
            return;
        }
        
        // Setup output folders to get the PNG folder
        const folders = await setupTextOutputFolders(outputFolder);
        
        // Call the existing saveAsPNG function with a test filename
        const testPath = `${folders.pngFolder.nativePath}/test-output.png`;
        log(`Testing PNG save to: ${testPath}`);
        
        const result = await saveAsPNG(doc, testPath);
        log(`✅ Test save completed: ${JSON.stringify(result)}`);
        
    } catch (error) {
        log(`❌ Test save failed: ${error.message}`);
        console.error("Test save error:", error);
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

// Second image replacement function removed - plugin now focuses on text only

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
        
        // Verify PNG folder is accessible
        try {
            const pngExists = await folders.pngFolder.isEntry;
            
            if (!pngExists) {
                throw new PluginError(
                    'Could not verify PNG folder access', 
                    'FOLDER_ACCESS_ERROR',
                    {
                        pngExists,
                        pngPath: folders.pngFolder?.nativePath
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

        console.log("[DEBUG] Output folder ready:", {
            png: folders.pngNativePath
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

// Image output folder setup removed - plugin now focuses on text only

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
// Second image CSV loading function removed - plugin now focuses on text only

// Updated PNG save function with proper API v2 format
/**
 * CRITICAL FIX: Simplified PNG save function using proper session tokens and executeAsModal
 * Based on the analysis, this addresses the core issues preventing PNG files from being saved
 */
async function saveAsPNG(doc, outputPath) {
    try {
        log(`[DEBUG] Starting PNG save to: ${outputPath}`);
        console.log(`[CURSOR SAVE] Starting PNG save to: ${outputPath}`);
        
        const fs = require('uxp').storage.localFileSystem;
        
        // CRITICAL FIX: Parse the path to get directory and filename
        const pathParts = outputPath.split('/');
        const fileName = pathParts.pop();
        const dirPath = pathParts.join('/');
        
        log(`[DEBUG] Parsed path - Directory: ${dirPath}, Filename: ${fileName}`);
        console.log(`[CURSOR SAVE] Parsed path - Directory: ${dirPath}, Filename: ${fileName}`);
        
        // Get the directory and create the file
        const directory = await fs.getFolder(dirPath);
        const outputFile = await directory.createFile(fileName, { overwrite: true });
        
        log(`[DEBUG] Created output file: ${outputFile.nativePath}`);
        console.log(`[CURSOR SAVE] Created output file: ${outputFile.nativePath}`);
        
        // CRITICAL FIX: Create a session token for batchPlay to use
        const sessionToken = fs.createSessionToken(outputFile);
        
        log(`[DEBUG] Created session token for file`);
        console.log(`[CURSOR SAVE] Created session token for file`);
        
        // CRITICAL FIX: Use compatible save command with PNG format
        const result = await batchPlay(
            [
                {
                    _obj: "save",
                    as: { 
                        _obj: "PNGFormat",
                        PNG8: false,
                        transparency: true,
                        interlaced: false,
                        compression: 6
                    },
                    in: sessionToken, // Use the session token directly
                    documentID: doc._id,
                    copy: true,
                    lowerCase: true,
                    _options: { dialogOptions: "dontDisplay" }
                }
            ],
            { synchronousExecution: true }
        );
        
        log(`[DEBUG] ✅ PNG save completed successfully`);
        console.log(`[CURSOR SAVE] ✅ PNG save completed successfully`);
        
        return { success: true, method: "sessionToken", path: outputFile.nativePath, result };
        
    } catch (error) {
        log(`[DEBUG] ❌ PNG save failed: ${error.message}`);
        console.log(`[CURSOR SAVE] ❌ PNG save failed: ${error.message}`);
        throw new PluginError('PNG save failed', 'PNG_SAVE_ERROR', { error, outputPath });
    }
}

// PSD save function removed - no longer needed

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

// Add a new function to directly save files using the File API
async function directSaveFile(doc, outputPath, fileType) {
    try {
        console.log(`[CURSOR SAVE] DIRECT SAVE: Starting direct save to ${outputPath} as ${fileType}`);
        
        // Ensure we have a valid path
        if (!outputPath || typeof outputPath !== 'string') {
            throw new Error('Invalid output path for direct save');
        }
        
        // Get the file system module
        const fs = require('uxp').storage.localFileSystem;
        
        // Parse the path
        const pathParts = outputPath.split('/');
        const fileName = pathParts.pop();
        const dirPath = pathParts.join('/');
        
        console.log(`[CURSOR SAVE] DIRECT SAVE: Parsed path - Directory: ${dirPath}, Filename: ${fileName}`);
        
        // Get the directory
        const directory = await fs.getFolder(dirPath);
        console.log(`[CURSOR SAVE] DIRECT SAVE: Got directory: ${directory.nativePath}`);
        
        // Create the file entry
        const file = await directory.createFile(fileName, { overwrite: true });
        console.log(`[CURSOR SAVE] DIRECT SAVE: Created file entry: ${file.nativePath}`);
        
        // Create a file token
        const fileToken = fs.createSessionToken(file);
        console.log(`[CURSOR SAVE] DIRECT SAVE: Created file token`);
        
        // Save based on file type
        if (fileType.toLowerCase() === 'png') {
            // Save as PNG
            const exportDesc = {
                _obj: "exportDocument",
                documentID: doc._id,
                format: {
                    _obj: "PNG",
                    PNG8: false,
                    transparency: true,
                    interlaced: false,
                    quality: 100
                },
                in: fileToken,
                _options: { 
                    dialogOptions: "dontDisplay"
                }
            };
            
            console.log(`[CURSOR SAVE] DIRECT SAVE: Executing PNG export`);
            await batchPlay([exportDesc], { synchronousExecution: true });

        } else {
            throw new Error(`Unsupported file type: ${fileType}`);
        }
        
        // Verify the file exists
        try {
            const fileEntry = await directory.getEntry(fileName);
            const fileSize = await fileEntry.size;
            console.log(`[CURSOR SAVE] DIRECT SAVE: ✅ File saved and verified: ${fileEntry.nativePath}, size: ${fileSize} bytes`);
            return {
                success: true,
                method: "direct-save",
                path: fileEntry.nativePath,
                size: fileSize
            };
        } catch (verifyError) {
            console.log(`[CURSOR SAVE] DIRECT SAVE: ❌ File verification failed: ${verifyError.message}`);
            throw verifyError;
        }
    } catch (error) {
        console.log(`[CURSOR SAVE] DIRECT SAVE: ❌ Failed: ${error.message}`);
        console.log(`[CURSOR SAVE] DIRECT SAVE: Error stack: ${error.stack}`);
        throw error;
    }
}

// Update processTextRow to use the new direct save function
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
    let filesSaved = { png: false };

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
                
                // Ensure we have valid folder paths
                if (!folders || !folders.pngFolder) {
                    log(`[DEBUG] ❌ Invalid folder structure for saving: ${JSON.stringify(folders)}`);
                    console.log(`[CURSOR SAVE] ❌ Invalid folder structure for saving: ${JSON.stringify(folders)}`);
                    throw new Error('Invalid folder structure for saving');
                }
                
                // Get native paths with proper formatting
                const pngNativePath = folders.pngNativePath || folders.pngFolder.nativePath;
                
                // Ensure paths are properly formatted for M1 Mac
                const pngPath = `${pngNativePath.replace(/\\/g, '/')}/${baseFileName}.png`;

                log("[DEBUG] Starting file saves for row " + index + ":", {
                    filename: baseFileName,
                    png: pngPath,
                    text1: text1Content,
                    text2: text2Content,
                    timestamp: new Date().toISOString()
                });
                console.log(`[CURSOR SAVE] Starting file saves for row ${index}:`, {
                    filename: baseFileName,
                    png: pngPath
                });

                // Save PNG file
                let pngSaved = false;
                try {
                    console.log(`[CURSOR SAVE] Attempting PNG save...`);
                    const pngSaveResult = await saveAsPNG(doc, pngPath);
                    console.log(`[CURSOR SAVE] ✅ PNG save result:`, pngSaveResult);
                    pngSaved = true;
                    filesSaved.png = true;
                } catch (pngError) {
                    console.log(`[CURSOR SAVE] ❌ PNG save failed: ${pngError.message}`);
                    errors.push({
                        type: 'PNG_SAVE',
                        error: pngError.message,
                        path: pngPath
                    });
                }



                // Verify files were created
                try {
                    log(`[DEBUG] Verifying saved files exist...`);
                    console.log(`[CURSOR SAVE] Verifying saved files exist...`);
                    const fs = require('uxp').storage.localFileSystem;
                    
                    // Verify PNG file
                    if (pngSaved) {
                        try {
                            // Extract directory path and filename
                            const pathParts = pngPath.split('/');
                            const fileName = pathParts.pop();
                            const dirPath = pathParts.join('/');
                            
                            // Get the directory
                            const pngDir = await fs.getFolder(dirPath);
                            const pngExists = await pngDir.getEntry(fileName);
                            
                            if (pngExists) {
                                log(`[DEBUG] ✅ PNG file verified: ${pngExists.nativePath}`);
                                console.log(`[CURSOR SAVE] ✅ PNG file verified: ${pngExists.nativePath}`);
                                log(`[DEBUG] PNG file size: ${await pngExists.size} bytes`);
                                console.log(`[CURSOR SAVE] PNG file size: ${await pngExists.size} bytes`);
                                filesSaved.png = true;
                            } else {
                                log(`[DEBUG] ❌ PNG file not found after save`);
                                console.log(`[CURSOR SAVE] ❌ PNG file not found after save`);
                                filesSaved.png = false;
                            }
                        } catch (pngVerifyError) {
                            log(`[DEBUG] ❌ PNG file verification failed: ${pngVerifyError.message}`);
                            console.log(`[CURSOR SAVE] ❌ PNG file verification failed: ${pngVerifyError.message}`);
                            filesSaved.png = false;
                        }
                    }
                    

                } catch (verifyError) {
                    log(`[DEBUG] File verification error: ${verifyError.message}`);
                    console.log(`[CURSOR SAVE] File verification error: ${verifyError.message}`);
                }

                // Log file save summary
                if (pngSaved) {
                    log(`[DEBUG] ✅ Row ${index} processing complete with PNG file saved:`);
                    console.log(`[CURSOR SAVE] ✅ Row ${index} processing complete with PNG file saved:`);
                    log(`[DEBUG]   - PNG: ${pngPath}`);
                    console.log(`[CURSOR SAVE]   - PNG: ${pngPath}`);
                } else {
                    log(`[DEBUG] ❌ Row ${index} processing complete but NO FILES SAVED`);
                    console.log(`[CURSOR SAVE] ❌ Row ${index} processing complete but NO FILES SAVED`);
                }

                // Write success to MCP relay for external monitoring
                try {
                    await writeToMCPRelay({
                        status: 'success',
                        files: {
                            png: pngSaved ? pngPath : null
                        },
                        timestamp: new Date().toISOString()
                    });
                } catch (mcpError) {
                    log(`[DEBUG] MCP relay write failed: ${mcpError.message}`);
                }
            } catch (saveError) {
                log(`[DEBUG] Critical save error: ${saveError.message}`);
                console.log(`[CURSOR SAVE] Critical save error: ${saveError.message}`);
                errors.push({
                    type: 'save',
                    error: saveError.message
                });
            }
        }

        // Return processing results
        return {
            success: layerUpdates.length > 0,
            processed: layerUpdates.length,
            errors: errors.length > 0 ? errors : null,
            duration: Date.now() - processingStart,
            filesSaved: filesSaved
        };
    } catch (error) {
        log(`[DEBUG] Row processing error: ${error.message}`);
        console.log(`[CURSOR SAVE] Row processing error: ${error.message}`);
        throw error;
    }
}

// Image replacement processing function removed - plugin now focuses on text only

// Event Listeners for Text Replace
document.getElementById('loadCSV').addEventListener('click', async () => {
    try {
        const file = await fs.getFileForOpening({ types: ['csv'] });
        if (file) {
            await loadTextCSV(file);
            
            // CRITICAL FIX: Reset row counter when new CSV is loaded and show current row status
            textReplaceState.status.performance.processedRows = 0;
            const currentRow = (textReplaceState.status.performance.processedRows || 0) + 1;
            const totalRows = textReplaceState.data.csvData?.length || 0;
            
            document.getElementById('textStatus').textContent = `CSV loaded: ${totalRows} rows. Ready to process row ${currentRow}`;
        }
    } catch (error) {
        log(`[Cursor OK] Error loading CSV: ${error.message}`);
        document.getElementById('textStatus').textContent = 'Error loading CSV file';
    }
});

// Image replace event listeners removed - plugin now focuses on text only

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

// Image output folder event listener removed - plugin now focuses on text only

// Process button event listeners
document.getElementById('processText').addEventListener('click', processTextReplacement);

// Test Save button event listener
document.getElementById('testSaveBtn').addEventListener('click', runSingleSaveTest);

// Restart Plugin button event listener
document.getElementById('restartPlugin').addEventListener('click', async () => {
    const button = document.getElementById('restartPlugin');
    const originalText = button.textContent;
    
    try {
        // Disable button and show loading state
        button.disabled = true;
        button.textContent = 'Restarting...';
        
        // CRITICAL FIX: Reset the row counter to start from row 1
        textReplaceState.status.performance.processedRows = 0;
        
        // Update status to show we're back to row 1
        const textStatus = document.getElementById('textStatus');
        if (textStatus && textReplaceState.data.csvData) {
            textStatus.textContent = `Ready to process row 1/${textReplaceState.data.csvData.length}`;
        }
        
        // Send restart command
        await writeToMCPRelay({
            command: "restartPlugin",
            message: "Plugin restarted - row counter reset to 1",
            timestamp: new Date().toISOString()
        });
        
        // Log success
        log("🌀 Plugin restarted - back to row 1");
        
        // Show success state briefly
        button.textContent = '✅ Reset to row 1';
        setTimeout(() => {
            button.textContent = originalText;
            button.disabled = false;
        }, 2000);
        
    } catch (error) {
        // Log error
        log(`❌ Error restarting plugin: ${error.message}`);
        
        // Show error state
        button.textContent = '❌ Restart failed';
        setTimeout(() => {
            button.textContent = originalText;
            button.disabled = false;
        }, 2000);
        
        // Log detailed error to console
        console.error('Restart failed:', error);
    }
});

// Simple tab management
function initializeTabs() {
    // Get tab elements
    const textReplaceTab = document.getElementById('textReplaceTab');
    const logsTab = document.getElementById('logsTab');

    // Get panel elements
    const textReplacePanel = document.getElementById('textReplacePanel');
    const logsPanel = document.getElementById('logsPanel');

    // Function to switch tabs
    function switchTab(tabId) {
        // Hide all panels
        [textReplacePanel, logsPanel].forEach(panel => {
            if (panel) panel.style.display = 'none';
        });

        // Deactivate all tabs
        [textReplaceTab, logsTab].forEach(tab => {
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

    // Add CSS for the stop button
    const style = document.createElement('style');
    style.textContent = `
        .stop-button {
            background-color: #d9534f !important;
            color: white !important;
        }
    `;
    document.head.appendChild(style);
});

// Add event listener for stop button
document.addEventListener('DOMContentLoaded', function() {
    // Add existing event listeners
    
    // Add stop button event listener
    const stopButton = document.getElementById('stopProcessingBtn');
    if (stopButton) {
        stopButton.addEventListener('click', requestStopProcessing);
    }
    
    // Other existing event listeners
});

// Main text replacement processing function
async function processTextReplacement() {
    const textStatus = document.getElementById('textStatus');
    const processType = document.querySelector('input[name="processType"]:checked')?.value;
    
    // Reset the stop flag when starting a new process
    resetStopProcessing();
    
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
        
        // Show stop button
        const stopButton = document.getElementById('stopProcessingBtn');
        if (stopButton) {
            stopButton.style.display = 'block';
        }
        
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
            // Process single row or all rows
            if (processType === 'current') {
                const currentRowIndex = textReplaceState.status.performance.processedRows || 0;
                if (currentRowIndex < textReplaceState.data.csvData.length) {
                    const row = textReplaceState.data.csvData[currentRowIndex];
                    textStatus.textContent = `Processing row ${currentRowIndex + 1}/${textReplaceState.data.csvData.length}...`;
                    const progressElement = document.getElementById('processingProgress');
                    if (progressElement) {
                        progressElement.textContent = `Processing row ${currentRowIndex + 1} of ${textReplaceState.data.csvData.length}`;
                    }
                    await processTextRow(row, currentRowIndex, textReplaceState.data.csvData.length, folders);
                    
                    // Update processed rows counter for "current" mode
                    textReplaceState.status.performance.processedRows = currentRowIndex + 1;
                    
                    textStatus.textContent = `Processed row ${currentRowIndex + 1}/${textReplaceState.data.csvData.length}`;
                } else {
                    textStatus.textContent = 'No more rows to process';
                }
            } else {
                // Process all rows with stop check
                const startIndex = textReplaceState.status.performance.processedRows || 0;
                for (let i = startIndex; i < textReplaceState.data.csvData.length; i++) {
                    // Check if processing should stop
                    if (shouldStopProcessing()) {
                        textStatus.textContent = 'Processing stopped by user';
                        log('[DEBUG] Processing stopped by user request');
                        break;
                    }
                    
                    const row = textReplaceState.data.csvData[i];
                    textStatus.textContent = `Processing row ${i + 1}/${textReplaceState.data.csvData.length}...`;
                    
                    // Update progress element
                    const progressElement = document.getElementById('processingProgress');
                    if (progressElement) {
                        progressElement.textContent = `Processing row ${i + 1} of ${textReplaceState.data.csvData.length}`;
                    }
                    
                    await processTextRow(row, i, textReplaceState.data.csvData.length, folders);
                    textReplaceState.status.performance.processedRows = i + 1;
                }
                
                if (!shouldStopProcessing()) {
                    textStatus.textContent = 'All rows processed';
                }
            }
            
            // Hide stop button
            if (stopButton) {
                stopButton.style.display = 'none';
            }
            
            // Reset stop flag
            resetStopProcessing();
            
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
                        png: folders.pngFolder.nativePath
                    }
                },
                timestamp: new Date().toISOString()
            });

            log('[Cursor OK] Text replacement completed');
            log(`[Cursor OK] Files saved to:`);
            log(`  PNG folder: ${folders.pngNativePath}`);
            
        } catch (error) {
            console.error("[DEBUG] Text replacement error:", error);
            throw error;
        }
    } catch (error) {
        console.error("[DEBUG] Text replacement error:", error);
        textStatus.textContent = `Error: ${error.message}`;
        textStatus.style.color = 'red';
        throw error;
    } finally {
        // Hide stop button
        const stopButton = document.getElementById('stopProcessingBtn');
        if (stopButton) {
            stopButton.style.display = 'none';
        }
        
        // Reset stop flag
        resetStopProcessing();
        
        textReplaceState.status.isProcessing = false;
    }
}

async function processCSV(csvData) {
    // Reset stop processing flag at the start of each CSV processing
    stopProcessingRequested = false;
    
    try {
        log(`[DEBUG] Starting CSV processing with ${csvData.length} rows`);
        
        // Update UI to show processing started
        const statusElement = document.getElementById('textStatus');
        const progressElement = document.getElementById('processingProgress');
        
        if (statusElement) {
            statusElement.textContent = 'Processing...';
            statusElement.style.color = '#4CAF50';
        }
        
        // Show stop button during processing
        const stopButton = document.getElementById('stopProcessingBtn');
        if (stopButton) {
            stopButton.style.display = 'block';
            stopButton.disabled = false;
        }
        
        // Setup output folders
        const folders = await setupTextOutputFolders(textReplaceState.data.outputFolder);
        log(`[DEBUG] Output folders setup complete`);
        
        // Get processing type
        const processType = document.querySelector('input[name="processType"]:checked').value;
        
        // Process rows based on selection
        const results = [];
        let processedCount = 0;
        let errorCount = 0;
        let savedFileCount = 0;
        
        if (processType === 'current') {
            // Process only the current row
            const currentRow = textReplaceState.data.currentRow;
            if (!currentRow) {
                throw new PluginError('No current row selected', 'NO_CURRENT_ROW');
            }
            
            const rowIndex = textReplaceState.data.currentRowIndex;
            
            try {
                // Update UI with current progress
                if (progressElement) {
                    progressElement.textContent = `Processing row ${rowIndex + 1}...`;
                }
                
                log(`[DEBUG] Processing single row ${rowIndex + 1}`);
                const result = await processTextRow(currentRow, rowIndex + 1, 1, folders);
                
                // Track successful saves
                if (result.success) {
                    processedCount++;
                    
                                            // Check if files were saved
                        const filesSaved = result.filesSaved || {};
                        if (filesSaved.png) {
                            savedFileCount += (filesSaved.png ? 1 : 0);
                        }
                }
                
                // Track errors
                if (result.errors && result.errors.length > 0) {
                    errorCount += result.errors.length;
                }
                
                results.push(result);
            } catch (rowError) {
                log(`[DEBUG] Error processing row ${rowIndex + 1}: ${rowError.message}`);
                errorCount++;
                results.push({
                    success: false,
                    error: rowError.message,
                    rowIndex: rowIndex + 1
                });
            }
        } else {
            // Process all rows
            for (let i = 0; i < csvData.length; i++) {
                // Check if processing should stop
                if (shouldStopProcessing()) {
                    log(`[DEBUG] Processing stopped by user after ${processedCount} rows`);
                    break;
                }
                
                const row = csvData[i];
                
                try {
                    // Update UI with current progress
                    if (progressElement) {
                        progressElement.textContent = `Processing row ${i + 1} of ${csvData.length}...`;
                    }
                    
                    log(`[DEBUG] Processing row ${i + 1} of ${csvData.length}`);
                    const result = await processTextRow(row, i + 1, csvData.length, folders);
                    
                    // Track successful saves
                    if (result.success) {
                        processedCount++;
                        
                        // Check if files were saved
                        const filesSaved = result.filesSaved || {};
                        if (filesSaved.png) {
                            savedFileCount += (filesSaved.png ? 1 : 0);
                        }
                    }
                    
                    // Track errors
                    if (result.errors && result.errors.length > 0) {
                        errorCount += result.errors.length;
                    }
                    
                    results.push(result);
                    
                    // Pause briefly to allow UI updates and prevent freezing
                    await new Promise(resolve => setTimeout(resolve, 100));
                    
                } catch (rowError) {
                    log(`[DEBUG] Error processing row ${i + 1}: ${rowError.message}`);
                    errorCount++;
                    results.push({
                        success: false,
                        error: rowError.message,
                        rowIndex: i + 1
                    });
                    
                    // Continue with next row instead of stopping the whole process
                    continue;
                }
            }
        }
        
        // Hide stop button after processing
        if (stopButton) {
            stopButton.style.display = 'none';
        }
        
        // Update UI with final status
        if (statusElement) {
            if (stopProcessingRequested) {
                statusElement.textContent = `Stopped: ${processedCount} rows processed, ${savedFileCount} files saved, ${errorCount} errors`;
                statusElement.style.color = 'orange';
            } else if (errorCount > 0) {
                statusElement.textContent = `Completed with issues: ${processedCount} rows processed, ${savedFileCount} files saved, ${errorCount} errors`;
                statusElement.style.color = 'orange';
            } else {
                statusElement.textContent = `Completed: ${processedCount} rows processed, ${savedFileCount} files saved`;
                statusElement.style.color = '#4CAF50';
            }
        }
        
        // Clear progress
        if (progressElement) {
            progressElement.textContent = '';
        }
        
        // Log completion statistics
        log(`[DEBUG] CSV processing ${stopProcessingRequested ? 'stopped' : 'completed'}:`);
        log(`[DEBUG] - Rows processed: ${processedCount} of ${csvData.length}`);
        log(`[DEBUG] - Files saved: ${savedFileCount}`);
        log(`[DEBUG] - Errors: ${errorCount}`);
        
        // Show output folder paths
        log(`[DEBUG] Files saved to:`);
        log(`[DEBUG] - PNG: ${folders.pngNativePath}`);
        
        // Write final status to MCP relay
        try {
            await writeToMCPRelay({
                status: stopProcessingRequested ? 'stopped' : 'completed',
                statistics: {
                    totalRows: csvData.length,
                    processedRows: processedCount,
                    savedFiles: savedFileCount,
                    errors: errorCount
                },
                timestamp: new Date().toISOString()
            });
        } catch (mcpError) {
            log(`[DEBUG] Failed to write final status to MCP relay: ${mcpError.message}`);
        }
        
        return {
            success: processedCount > 0,
            processed: processedCount,
            total: csvData.length,
            errors: errorCount,
            stopped: stopProcessingRequested,
            savedFiles: savedFileCount,
            results
        };
    } catch (error) {
        log(`[DEBUG] CSV processing failed with critical error: ${error.message}`);
        log(`[DEBUG] Error stack: ${error.stack}`);
        
        // Update UI to show error
        const statusElement = document.getElementById('textStatus');
        if (statusElement) {
            statusElement.textContent = `Error: ${error.message}`;
            statusElement.style.color = 'red';
        }
        
        // Hide stop button
        const stopButton = document.getElementById('stopProcessingBtn');
        if (stopButton) {
            stopButton.style.display = 'none';
        }
        
        throw error;
    }
}