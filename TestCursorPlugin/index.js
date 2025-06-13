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
async function setupTextOutputFolders() {
    log('[DEBUG] Setting up text output folders');
    
    try {
        const fs = require('uxp').storage.localFileSystem;
        
        // Get desktop folder as base location
        const desktop = await fs.getFolder(fs.domains.userDesktop);
        log(`[DEBUG] Got desktop folder: ${desktop.nativePath}`);
        
        // Create main output folder
        let mainFolder;
        try {
            mainFolder = await desktop.getEntry('SaveOnPeptides_Output');
            log(`[DEBUG] Found existing main output folder: ${mainFolder.nativePath}`);
        } catch (e) {
            log(`[DEBUG] Creating main output folder on desktop`);
            mainFolder = await desktop.createFolder('SaveOnPeptides_Output');
            log(`[DEBUG] Created main output folder: ${mainFolder.nativePath}`);
        }
        
        // Create timestamped subfolder for this session
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0];
        const sessionFolderName = `Session_${timestamp}`;
        
        let sessionFolder;
        try {
            sessionFolder = await mainFolder.getEntry(sessionFolderName);
            log(`[DEBUG] Found existing session folder: ${sessionFolder.nativePath}`);
        } catch (e) {
            log(`[DEBUG] Creating new session folder: ${sessionFolderName}`);
            sessionFolder = await mainFolder.createFolder(sessionFolderName);
            log(`[DEBUG] Created session folder: ${sessionFolder.nativePath}`);
        }
        
        // Create PNG output folder
        let pngFolder;
        try {
            pngFolder = await sessionFolder.getEntry('PNG_Output');
            log(`[DEBUG] Found existing PNG output folder: ${pngFolder.nativePath}`);
        } catch (e) {
            log(`[DEBUG] Creating PNG output folder`);
            pngFolder = await sessionFolder.createFolder('PNG_Output');
            log(`[DEBUG] Created PNG output folder: ${pngFolder.nativePath}`);
        }
        
        // Create PSD output folder
        let psdFolder;
        try {
            psdFolder = await sessionFolder.getEntry('PSD_Output');
            log(`[DEBUG] Found existing PSD output folder: ${psdFolder.nativePath}`);
        } catch (e) {
            log(`[DEBUG] Creating PSD output folder`);
            psdFolder = await sessionFolder.createFolder('PSD_Output');
            log(`[DEBUG] Created PSD output folder: ${psdFolder.nativePath}`);
        }
        
        // Create MCP relay folder for external monitoring
        let mcpFolder;
        try {
            mcpFolder = await sessionFolder.getEntry('MCP_Relay');
            log(`[DEBUG] Found existing MCP relay folder: ${mcpFolder.nativePath}`);
        } catch (e) {
            log(`[DEBUG] Creating MCP relay folder`);
            mcpFolder = await sessionFolder.createFolder('MCP_Relay');
            log(`[DEBUG] Created MCP relay folder: ${mcpFolder.nativePath}`);
        }
        
        // Verify folder permissions by writing test files
        try {
            log(`[DEBUG] Verifying folder write permissions with test files`);
            
            // Test PNG folder
            const pngTestFile = await pngFolder.createFile('test.txt', { overwrite: true });
            await pngTestFile.write('test');
            await pngTestFile.delete();
            log(`[DEBUG] ✅ PNG folder write test passed`);
            
            // Test PSD folder
            const psdTestFile = await psdFolder.createFile('test.txt', { overwrite: true });
            await psdTestFile.write('test');
            await psdTestFile.delete();
            log(`[DEBUG] ✅ PSD folder write test passed`);
            
            // Test MCP folder
            const mcpTestFile = await mcpFolder.createFile('test.txt', { overwrite: true });
            await mcpTestFile.write('test');
            await mcpTestFile.delete();
            log(`[DEBUG] ✅ MCP folder write test passed`);
        } catch (testError) {
            log(`[DEBUG] ❌ Folder permission test failed: ${testError.message}`);
            // Continue anyway, we'll handle errors during actual file operations
        }
        
        // Store native paths for easier access
        const pngNativePath = pngFolder.nativePath;
        const psdNativePath = psdFolder.nativePath;
        const mcpNativePath = mcpFolder.nativePath;
        
        // Log the paths in a format suitable for M1 Mac
        log(`[DEBUG] Output folders setup complete:`);
        log(`[DEBUG] - PNG output: ${pngNativePath.replace(/\\/g, '/')}`);
        log(`[DEBUG] - PSD output: ${psdNativePath.replace(/\\/g, '/')}`);
        log(`[DEBUG] - MCP relay: ${mcpNativePath.replace(/\\/g, '/')}`);
        
        return {
            pngFolder,
            psdFolder,
            mcpFolder,
            pngNativePath: pngNativePath.replace(/\\/g, '/'),
            psdNativePath: psdNativePath.replace(/\\/g, '/'),
            mcpNativePath: mcpNativePath.replace(/\\/g, '/')
        };
    } catch (error) {
        log(`[DEBUG] ❌ Failed to setup output folders: ${error.message}`);
        throw new PluginError('Failed to setup output folders', 'FOLDER_SETUP_ERROR', error);
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

// Updated PNG save function with proper API v2 format
async function saveAsPNG(doc, outputPath) {
    try {
        log(`[DEBUG] Starting PNG save to: ${outputPath}`);
        
        // Check if we're dealing with a file token or a path string
        const isToken = typeof outputPath !== 'string';
        
        // For debugging
        log(`[DEBUG] Save PNG using ${isToken ? 'token' : 'path string'}: ${isToken ? 'File Token' : outputPath}`);
        
        // APPROACH 1: Use direct file system API for most reliable method
        try {
            log(`[DEBUG] APPROACH 1: Attempting direct file system API save...`);
            
            const fs = require('uxp').storage.localFileSystem;
            const formats = require('uxp').storage.formats;
            
            // Parse the path to get directory and filename if it's a string path
            let outputFile;
            if (!isToken) {
                try {
                    // Extract directory path and filename
                    const pathParts = outputPath.split('/');
                    const fileName = pathParts.pop();
                    const dirPath = pathParts.join('/');
                    
                    log(`[DEBUG] Parsed path - Directory: ${dirPath}, Filename: ${fileName}`);
                    
                    // Get the directory
                    const directory = await fs.getFolder(dirPath);
                    log(`[DEBUG] Got directory: ${directory.nativePath}`);
                    
                    // Create or get the file
                    outputFile = await directory.createFile(fileName, { overwrite: true });
                    log(`[DEBUG] Created output file: ${outputFile.nativePath}`);
                } catch (parseError) {
                    log(`[DEBUG] Path parsing error: ${parseError.message}`);
                    throw parseError;
                }
            } else {
                outputFile = outputPath;
                log(`[DEBUG] Using provided file token`);
            }
            
            // Save to temporary file first
            log(`[DEBUG] Creating temporary file for PNG save...`);
            const tempFile = await fs.createTemporaryFile("temp-png-");
            log(`[DEBUG] Created temp file: ${tempFile.nativePath}`);
            
            // Create a session token for the temp file
            const tempToken = fs.createSessionToken(tempFile);
            
            // Use exportDocument to save to temp file
            log(`[DEBUG] Exporting document to temp file...`);
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
                in: tempToken,
                _options: { 
                    dialogOptions: "dontDisplay"
                }
            };
            
            await batchPlay([exportDesc], { synchronousExecution: true });
            log(`[DEBUG] Successfully exported to temp file`);
            
            // Read temp file content
            log(`[DEBUG] Reading temp file content...`);
            const tempContent = await tempFile.read({ format: formats.binary });
            log(`[DEBUG] Read ${tempContent.byteLength} bytes from temp file`);
            
            // Write content to final destination
            log(`[DEBUG] Writing content to final destination...`);
            await outputFile.write(tempContent, { format: formats.binary });
            log(`[DEBUG] ✅ Successfully wrote PNG file: ${outputFile.nativePath}`);
            
            // Clean up temp file
            await tempFile.delete();
            log(`[DEBUG] Cleaned up temp file`);
            
            return { success: true, method: "direct-fs", path: outputFile.nativePath };
        } catch (directFsError) {
            log(`[DEBUG] Direct file system approach failed: ${directFsError.message}`);
            log(`[DEBUG] Error details: ${JSON.stringify(directFsError)}`);
            
            // Continue to next approach
        }
        
        // APPROACH 2: Use exportDocument with modern format
        try {
            log(`[DEBUG] APPROACH 2: Attempting exportDocument method...`);
            
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
                in: isToken ? outputPath : { _path: outputPath },
                _options: { 
                    dialogOptions: "dontDisplay"
                }
            };
            
            const exportResult = await batchPlay(
                [exportDesc],
                {
                    synchronousExecution: true,
                    modalBehavior: "none"
                }
            );
            
            log(`[DEBUG] ✅ Export PNG Save completed successfully`);
            return { success: true, method: "exportDocument", result: exportResult };
        } catch (exportError) {
            log(`[DEBUG] Export PNG Save failed: ${exportError.message}`);
            log(`[DEBUG] Error details: ${JSON.stringify(exportError)}`);
            
            // Continue to next approach
        }
        
        // APPROACH 3: Use basic save with minimal options
        try {
            log(`[DEBUG] APPROACH 3: Attempting basic save method...`);
            
            const saveDesc = {
                _obj: "save",
                as: {
                    _obj: "PNGFormat",
                    PNG8: false
                },
                in: isToken ? outputPath : { _path: outputPath },
                copy: true,
                _options: { 
                    dialogOptions: "dontDisplay"
                }
            };
            
            const saveResult = await batchPlay(
                [saveDesc],
                {
                    synchronousExecution: true,
                    modalBehavior: "none"
                }
            );
            
            log(`[DEBUG] ✅ Basic PNG Save completed successfully`);
            return { success: true, method: "basicSave", result: saveResult };
        } catch (saveError) {
            log(`[DEBUG] Basic PNG Save failed: ${saveError.message}`);
            log(`[DEBUG] Error details: ${JSON.stringify(saveError)}`);
            
            // Continue to next approach
        }
        
        // APPROACH 4: Use quickExport
        try {
            log(`[DEBUG] APPROACH 4: Attempting quickExport method...`);
            
            const quickExportDesc = {
                _obj: "quickExport",
                format: {
                    _enum: "exportFormat",
                    _value: "PNG"
                },
                destination: {
                    _enum: "saveStageType",
                    _value: "saveStageType"
                },
                in: isToken ? outputPath : { _path: outputPath },
                _options: { 
                    dialogOptions: "dontDisplay"
                }
            };
            
            const quickExportResult = await batchPlay(
                [quickExportDesc],
                {
                    synchronousExecution: true,
                    modalBehavior: "none"
                }
            );
            
            log(`[DEBUG] ✅ QuickExport PNG Save completed successfully`);
            return { success: true, method: "quickExport", result: quickExportResult };
        } catch (quickExportError) {
            log(`[DEBUG] QuickExport PNG Save failed: ${quickExportError.message}`);
            log(`[DEBUG] Error details: ${JSON.stringify(quickExportError)}`);
            
            // All approaches failed
            throw new PluginError(
                'All PNG save methods failed',
                'PNG_SAVE_ERROR',
                { 
                    directFsError: directFsError?.message,
                    exportError: exportError?.message,
                    saveError: saveError?.message,
                    quickExportError: quickExportError?.message,
                    outputPath 
                }
            );
        }
    } catch (error) {
        log(`[DEBUG] PNG Save Failed with critical error: ${error.message}`);
        log(`[DEBUG] Stack trace: ${error.stack}`);
        throw error;
    }
}

// Updated PSD save function with proper API v2 format
async function saveAsPSD(doc, outputPath) {
    try {
        log(`[DEBUG] Starting PSD save to: ${outputPath}`);
        
        // Check if we're dealing with a file token or a path string
        const isToken = typeof outputPath !== 'string';
        
        // For debugging
        log(`[DEBUG] Save PSD using ${isToken ? 'token' : 'path string'}: ${isToken ? 'File Token' : outputPath}`);
        
        // APPROACH 1: Use direct file system API for most reliable method
        try {
            log(`[DEBUG] APPROACH 1: Attempting direct file system API save for PSD...`);
            
            const fs = require('uxp').storage.localFileSystem;
            const formats = require('uxp').storage.formats;
            
            // Parse the path to get directory and filename if it's a string path
            let outputFile;
            if (!isToken) {
                try {
                    // Extract directory path and filename
                    const pathParts = outputPath.split('/');
                    const fileName = pathParts.pop();
                    const dirPath = pathParts.join('/');
                    
                    log(`[DEBUG] Parsed path - Directory: ${dirPath}, Filename: ${fileName}`);
                    
                    // Get the directory
                    const directory = await fs.getFolder(dirPath);
                    log(`[DEBUG] Got directory: ${directory.nativePath}`);
                    
                    // Create or get the file
                    outputFile = await directory.createFile(fileName, { overwrite: true });
                    log(`[DEBUG] Created output file: ${outputFile.nativePath}`);
                } catch (parseError) {
                    log(`[DEBUG] Path parsing error: ${parseError.message}`);
                    throw parseError;
                }
            } else {
                outputFile = outputPath;
                log(`[DEBUG] Using provided file token`);
            }
            
            // Save to temporary file first
            log(`[DEBUG] Creating temporary file for PSD save...`);
            const tempFile = await fs.createTemporaryFile("temp-psd-");
            log(`[DEBUG] Created temp file: ${tempFile.nativePath}`);
            
            // Create a session token for the temp file
            const tempToken = fs.createSessionToken(tempFile);
            
            // Use save to save to temp file
            log(`[DEBUG] Saving document to temp file...`);
            const saveDesc = {
                _obj: "save",
                as: {
                    _obj: "photoshop35Format",
                    alphaChannels: true,
                    embedColorProfile: true,
                    layers: true,
                    maximizeCompatibility: true
                },
                in: tempToken,
                copy: true,
                _options: { 
                    dialogOptions: "dontDisplay"
                }
            };
            
            await batchPlay([saveDesc], { synchronousExecution: true });
            log(`[DEBUG] Successfully saved to temp file`);
            
            // Read temp file content
            log(`[DEBUG] Reading temp file content...`);
            const tempContent = await tempFile.read({ format: formats.binary });
            log(`[DEBUG] Read ${tempContent.byteLength} bytes from temp file`);
            
            // Write content to final destination
            log(`[DEBUG] Writing content to final destination...`);
            await outputFile.write(tempContent, { format: formats.binary });
            log(`[DEBUG] ✅ Successfully wrote PSD file: ${outputFile.nativePath}`);
            
            // Clean up temp file
            await tempFile.delete();
            log(`[DEBUG] Cleaned up temp file`);
            
            return { success: true, method: "direct-fs", path: outputFile.nativePath };
        } catch (directFsError) {
            log(`[DEBUG] Direct file system approach failed for PSD: ${directFsError.message}`);
            log(`[DEBUG] Error details: ${JSON.stringify(directFsError)}`);
            
            // Continue to next approach
        }
        
        // APPROACH 2: Use standard save
        try {
            log(`[DEBUG] APPROACH 2: Attempting standard PSD save method...`);
            
            const saveDesc = {
                _obj: "save",
                as: {
                    _obj: "photoshop35Format",
                    alphaChannels: true,
                    embedColorProfile: true,
                    layers: true,
                    maximizeCompatibility: true
                },
                in: isToken ? outputPath : { _path: outputPath },
                copy: true,
                _options: { 
                    dialogOptions: "dontDisplay"
                }
            };

            const result = await batchPlay(
                [saveDesc],
                {
                    synchronousExecution: true,
                    modalBehavior: "none"
                }
            );

            log(`[DEBUG] ✅ PSD Save completed successfully`);
            return { success: true, method: "standardSave", result };
        } catch (error) {
            log(`[DEBUG] Standard PSD Save Failed: ${error.message}`);
            log(`[DEBUG] Error details: ${JSON.stringify(error)}`);
            
            // Try basic approach
            try {
                log(`[DEBUG] APPROACH 3: Attempting basic PSD save method...`);
                
                const basicSaveDesc = {
                    _obj: "save",
                    as: {
                        _obj: "photoshop35Format"
                    },
                    in: isToken ? outputPath : { _path: outputPath },
                    copy: true,
                    _options: { 
                        dialogOptions: "dontDisplay"
                    }
                };
                
                const basicResult = await batchPlay(
                    [basicSaveDesc],
                    {
                        synchronousExecution: true,
                        modalBehavior: "none"
                    }
                );
                
                log(`[DEBUG] ✅ Basic PSD Save completed successfully`);
                return { success: true, method: "basicSave", result: basicResult };
            } catch (basicError) {
                log(`[DEBUG] All PSD save methods failed`);
                log(`[DEBUG] Error details: ${JSON.stringify(basicError)}`);
                throw new PluginError(
                    'Failed to save PSD (all methods failed)',
                    'PSD_SAVE_ERROR',
                    { 
                        directFsError: directFsError?.message,
                        standardError: error.message, 
                        basicError: basicError.message,
                        outputPath 
                    }
                );
            }
        }
    } catch (error) {
        log(`[DEBUG] PSD Save Failed with critical error: ${error.message}`);
        log(`[DEBUG] Stack trace: ${error.stack}`);
        throw error;
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
    let filesSaved = { png: false, psd: false };

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
                if (!folders || !folders.pngFolder || !folders.psdFolder) {
                    log(`[DEBUG] ❌ Invalid folder structure for saving: ${JSON.stringify(folders)}`);
                    throw new Error('Invalid folder structure for saving');
                }
                
                // Get native paths with proper formatting
                const pngNativePath = folders.pngNativePath || folders.pngFolder.nativePath;
                const psdNativePath = folders.psdNativePath || folders.psdFolder.nativePath;
                
                // Ensure paths are properly formatted for M1 Mac
                const pngPath = `${pngNativePath.replace(/\\/g, '/')}/${baseFileName}.png`;
                const psdPath = `${psdNativePath.replace(/\\/g, '/')}/${baseFileName}.psd`;

                log("[DEBUG] Starting file saves for row " + index + ":", {
                    filename: baseFileName,
                    png: pngPath,
                    psd: psdPath,
                    text1: text1Content,
                    text2: text2Content,
                    timestamp: new Date().toISOString()
                });

                // Save PNG file with explicit error handling
                let pngSaveResult = null;
                let pngSaved = false;
                try {
                    log(`[DEBUG] Saving PNG file to: ${pngPath}`);
                    pngSaveResult = await saveAsPNG(doc, pngPath);
                    log(`[DEBUG] ✅ PNG save result: ${JSON.stringify(pngSaveResult)}`);
                    pngSaved = true;
                    filesSaved.png = true;
                } catch (pngError) {
                    log(`[DEBUG] ❌ PNG save failed with error: ${pngError.message}`);
                    log(`[DEBUG] Error details: ${JSON.stringify(pngError)}`);
                    
                    // Add to errors but continue with PSD save
                    errors.push({
                        type: 'PNG_SAVE',
                        error: pngError.message,
                        path: pngPath
                    });
                }

                // Save PSD file with explicit error handling
                let psdSaveResult = null;
                let psdSaved = false;
                try {
                    log(`[DEBUG] Saving PSD file to: ${psdPath}`);
                    psdSaveResult = await saveAsPSD(doc, psdPath);
                    log(`[DEBUG] ✅ PSD save result: ${JSON.stringify(psdSaveResult)}`);
                    psdSaved = true;
                    filesSaved.psd = true;
                } catch (psdError) {
                    log(`[DEBUG] ❌ PSD save failed with error: ${psdError.message}`);
                    log(`[DEBUG] Error details: ${JSON.stringify(psdError)}`);
                    
                    // Add to errors but continue
                    errors.push({
                        type: 'PSD_SAVE',
                        error: psdError.message,
                        path: psdPath
                    });
                }

                // Verify files were created
                try {
                    log(`[DEBUG] Verifying saved files exist...`);
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
                                log(`[DEBUG] PNG file size: ${await pngExists.size} bytes`);
                                filesSaved.png = true;
                            } else {
                                log(`[DEBUG] ❌ PNG file not found after save`);
                                filesSaved.png = false;
                            }
                        } catch (pngVerifyError) {
                            log(`[DEBUG] ❌ PNG file verification failed: ${pngVerifyError.message}`);
                            filesSaved.png = false;
                        }
                    }
                    
                    // Verify PSD file
                    if (psdSaved) {
                        try {
                            // Extract directory path and filename
                            const pathParts = psdPath.split('/');
                            const fileName = pathParts.pop();
                            const dirPath = pathParts.join('/');
                            
                            // Get the directory
                            const psdDir = await fs.getFolder(dirPath);
                            const psdExists = await psdDir.getEntry(fileName);
                            
                            if (psdExists) {
                                log(`[DEBUG] ✅ PSD file verified: ${psdExists.nativePath}`);
                                log(`[DEBUG] PSD file size: ${await psdExists.size} bytes`);
                                filesSaved.psd = true;
                            } else {
                                log(`[DEBUG] ❌ PSD file not found after save`);
                                filesSaved.psd = false;
                            }
                        } catch (psdVerifyError) {
                            log(`[DEBUG] ❌ PSD file verification failed: ${psdVerifyError.message}`);
                            filesSaved.psd = false;
                        }
                    }
                } catch (verifyError) {
                    log(`[DEBUG] File verification error: ${verifyError.message}`);
                }

                // Log file save summary
                if (pngSaved || psdSaved) {
                    log(`[DEBUG] ✅ Row ${index} processing complete with files saved:`);
                    if (pngSaved) log(`[DEBUG]   - PNG: ${pngPath}`);
                    if (psdSaved) log(`[DEBUG]   - PSD: ${psdPath}`);
                } else {
                    log(`[DEBUG] ❌ Row ${index} processing complete but NO FILES SAVED`);
                }

                // Write success to MCP relay for external monitoring
                try {
                    await writeToMCPRelay({
                        status: 'success',
                        files: {
                            png: pngSaved ? pngPath : null,
                            psd: psdSaved ? psdPath : null
                        },
                        timestamp: new Date().toISOString()
                    });
                } catch (mcpError) {
                    log(`[DEBUG] MCP relay write failed: ${mcpError.message}`);
                }
            } catch (saveError) {
                log(`[DEBUG] Critical save error: ${saveError.message}`);
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
        throw error;
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

// Add a global processing control variable
let isProcessingStopped = false;

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
    
    // Disable stop button
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

// Update the processTextReplacement function to check for stop requests
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
            folders = await setupTextOutputFolders();
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
            // Update UI to show stop button
            const processButton = document.getElementById('processText');
            if (processButton) {
                processButton.textContent = 'Stop Processing';
                processButton.classList.add('stop-button');
                processButton.disabled = true; // Disable the main button while stop button is active
            }
            
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
            
            // Reset the button
            if (processButton) {
                processButton.textContent = 'Process CSV';
                processButton.classList.remove('stop-button');
                processButton.disabled = false;
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
                        png: folders.pngFolder.nativePath,
                        psd: folders.psdFolder.nativePath
                    }
                },
                timestamp: new Date().toISOString()
            });

            log('[Cursor OK] Text replacement completed');
            log(`[Cursor OK] Files saved to:`);
            log(`  PNG folder: ${folders.pngNativePath}`);
            log(`  PSD folder: ${folders.psdNativePath}`);
            
        } catch (error) {
            console.error("[DEBUG] Text replacement error:", error);
            throw error;
        }
    } catch (error) {
        console.error("[DEBUG] Text replacement error:", error);
        throw error;
    } finally {
        // Always reset the button state
        const processButton = document.getElementById('processText');
        if (processButton) {
            processButton.textContent = 'Process CSV';
            processButton.classList.remove('stop-button');
            processButton.disabled = false;
        }
        
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

// Update processImageReplacement to use document initialization
async function processImageReplacement() {
    const imageStatus = document.getElementById('imageStatus');
    const processType = document.querySelector('input[name="processTypeImg"]:checked')?.value;
    
    // Reset the stop flag when starting a new process
    isProcessingStopped = false;
    
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

        // Update UI to show stop button
        const processButton = document.getElementById('processImages');
        if (processButton) {
            processButton.textContent = 'Stop Processing';
            processButton.classList.add('stop-button');
            // Change the button to a stop button
            processButton.onclick = () => {
                isProcessingStopped = true;
                imageStatus.textContent = 'Stopping processing...';
                log('[DEBUG] Processing stop requested by user');
                processButton.disabled = true;
            };
        }
        
        // Process single row or all rows
        if (processType === 'current') {
            const currentRowIndex = imageReplaceState.status.performance.processedRows || 0;
            if (currentRowIndex < imageReplaceState.data.csvData.length) {
                const row = imageReplaceState.data.csvData[currentRowIndex];
                imageStatus.textContent = `Processing row ${currentRowIndex + 1}/${imageReplaceState.data.csvData.length}...`;
                await processImageRow(row, currentRowIndex, imageReplaceState.data.csvData.length);
                imageStatus.textContent = `Processed row ${currentRowIndex + 1}/${imageReplaceState.data.csvData.length}`;
            } else {
                imageStatus.textContent = 'No more rows to process';
            }
        } else {
            // Process all rows with stop check
            const startIndex = imageReplaceState.status.performance.processedRows || 0;
            for (let i = startIndex; i < imageReplaceState.data.csvData.length; i++) {
                // Check if processing should stop
                if (isProcessingStopped) {
                    imageStatus.textContent = 'Processing stopped by user';
                    log('[DEBUG] Processing stopped by user request');
                    break;
                }
                
                const row = imageReplaceState.data.csvData[i];
                imageStatus.textContent = `Processing row ${i + 1}/${imageReplaceState.data.csvData.length}...`;
                await processImageRow(row, i, imageReplaceState.data.csvData.length);
                imageReplaceState.status.performance.processedRows = i + 1;
            }
            
            if (!isProcessingStopped) {
                imageStatus.textContent = 'All rows processed';
            }
        }
        
        // Reset the button
        if (processButton) {
            processButton.textContent = 'Process CSV';
            processButton.classList.remove('stop-button');
            processButton.disabled = false;
            // Restore original click handler
            processButton.onclick = processImageReplacement;
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
        // Always reset the button state
        const processButton = document.getElementById('processImages');
        if (processButton) {
            processButton.textContent = 'Process CSV';
            processButton.classList.remove('stop-button');
            processButton.disabled = false;
            // Restore original click handler
            processButton.onclick = processImageReplacement;
        }
        
        imageReplaceState.status.isProcessing = false;
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
            const folders = await setupTextOutputFolders();
            
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

async function processCSV(csvData) {
    // Reset stop processing flag at the start of each CSV processing
    stopProcessingRequested = false;
    
    try {
        log(`[DEBUG] Starting CSV processing with ${csvData.length} rows`);
        
        // Update UI to show processing started
        const statusElement = document.getElementById('textStatus');
        if (statusElement) {
            statusElement.textContent = 'Processing...';
            statusElement.style.color = '#4CAF50';
        }
        
        // Show stop button during processing
        const stopButton = document.getElementById('stopProcessingBtn');
        if (stopButton) {
            stopButton.style.display = 'block';
            stopButton.disabled = false;
            stopButton.textContent = 'Stop Processing';
        }
        
        // Setup output folders
        const folders = await setupTextOutputFolders();
        log(`[DEBUG] Output folders setup complete`);
        
        // Process each row
        const results = [];
        let processedCount = 0;
        let errorCount = 0;
        let savedFileCount = 0;
        
        for (let i = 0; i < csvData.length; i++) {
            // Check if processing should stop
            if (shouldStopProcessing()) {
                log(`[DEBUG] Processing stopped by user after ${processedCount} rows`);
                break;
            }
            
            const row = csvData[i];
            
            try {
                // Update UI with current progress
                if (statusElement) {
                    statusElement.textContent = `Processing row ${i + 1} of ${csvData.length}...`;
                }
                
                log(`[DEBUG] Processing row ${i + 1} of ${csvData.length}`);
                const result = await processTextRow(row, i + 1, csvData.length, folders);
                
                // Track successful saves
                if (result.success) {
                    processedCount++;
                    
                    // Check if files were saved
                    const filesSaved = result.filesSaved || {};
                    if (filesSaved.png || filesSaved.psd) {
                        savedFileCount += (filesSaved.png ? 1 : 0) + (filesSaved.psd ? 1 : 0);
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
        
        // Log completion statistics
        log(`[DEBUG] CSV processing ${stopProcessingRequested ? 'stopped' : 'completed'}:`);
        log(`[DEBUG] - Rows processed: ${processedCount} of ${csvData.length}`);
        log(`[DEBUG] - Files saved: ${savedFileCount}`);
        log(`[DEBUG] - Errors: ${errorCount}`);
        
        // Show output folder paths
        log(`[DEBUG] Files saved to:`);
        log(`[DEBUG] - PNG: ${folders.pngNativePath}`);
        log(`[DEBUG] - PSD: ${folders.psdNativePath}`);
        
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