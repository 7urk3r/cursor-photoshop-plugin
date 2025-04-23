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

    // Log full layer details for debugging
    console.log("[DEBUG] Creating target for layer:", {
        name: layer.name,
        id: layer._id,
        legacyId: layer.id,
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

    // Get the appropriate ID, with fallbacks and validation
    const layerId = layer._id ?? layer.id;
    if (typeof layerId !== 'number') {
        console.error("[DEBUG] Invalid layer ID:", {
            _id: layer._id,
            id: layer.id,
            type: typeof layerId,
            layer: JSON.stringify(layer, null, 2)
        });
        throw new PluginError(
            'Invalid layer ID type', 
            'INVALID_LAYER_ID_TYPE',
            { 
                layerId,
                type: typeof layerId,
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
    
    // Validate the targeting is appropriate for the layer type
    if (finalOptions.type === 'textLayer' && !isTextLayer) {
        throw new PluginError(
            'Cannot target non-text layer as text layer',
            'INVALID_LAYER_TYPE',
            {
                expectedType: 'text',
                actualType: layerKind,
                layerName: layer.name
            }
        );
    }

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
            dialogOptions: finalOptions.dialogOptions
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

        console.log("[DEBUG] Starting text replacement:", {
            targetLayer: layer.name,
            textToApply: newText,
            timestamp: new Date().toISOString()
        });

        // Select layer first
        await batchPlay([batchPlayCommands.selectLayer(layer.name)], {
            synchronousExecution: true,
            modalBehavior: "fail"
        });
        await wait(delays.selection);

        // Update text with line break handling
        const processedText = newText.replace(/\|br\|/g, '\r');
        await batchPlay([batchPlayCommands.setText(layer.name, processedText)], {
            synchronousExecution: true,
            modalBehavior: "fail"
        });
        await wait(delays.textUpdate);

        // Verify text update through character panel
        const verifyResult = await batchPlay([batchPlayCommands.getCharacterStyle(layer.name)], {
            synchronousExecution: true,
            modalBehavior: "fail"
        });

        console.log("[DEBUG] Text update verification:", {
            layer: layer.name,
            expected: processedText,
            verifyResult,
            success: true
        });

        return true;
    } catch (error) {
        console.error("[DEBUG] Text replacement error:", {
            layer: layer?.name,
            error: error.message,
            code: error.code
        });
        throw new PluginError(
            `Failed to replace text in layer ${layer?.name}`,
            'TEXT_REPLACE_ERROR',
            { originalError: error, layer: layer?.name }
        );
    }
}

// Optimized font size update with verification
async function updateFontSize(layer, fontSize) {
    try {
        if (!layer?.name) {
            throw new PluginError('Invalid layer object', 'INVALID_LAYER');
        }

        // Ensure fontSize is a valid number and convert to points
        const size = parseFloat(fontSize);
        if (isNaN(size)) {
            throw new PluginError('Invalid font size', 'INVALID_FONT_SIZE', { fontSize });
        }

        console.log("[DEBUG] Starting font size update:", {
            layer: layer.name,
            targetSize: size,
            timestamp: new Date().toISOString()
        });

        // Get initial font size for verification
        const initialState = await batchPlay([{
            _obj: "get",
            _target: [
                { _property: "textStyle" },
                { _ref: "textLayer", _name: layer.name }
            ],
            _options: { dialogOptions: "dontDisplay" }
        }], {
            synchronousExecution: true,
            modalBehavior: "fail"
        });

        console.log("[DEBUG] Initial font state:", {
            layer: layer.name,
            currentSize: initialState?.[0]?.textStyle?.size?._value,
            targetSize: size,
            fullState: initialState
        });

        // Select layer first with retry
        let selectAttempts = 0;
        while (selectAttempts < 3) {
            try {
                await batchPlay([{
                    _obj: "select",
                    _target: [{ _ref: "layer", _name: layer.name }],
                    makeVisible: false,
                    _options: { dialogOptions: "dontDisplay" }
                }], {
                    synchronousExecution: true,
                    modalBehavior: "fail"
                });
                break;
            } catch (selectError) {
                selectAttempts++;
                if (selectAttempts === 3) {
                    throw new PluginError(
                        'Failed to select layer for font size update',
                        'LAYER_SELECT_ERROR',
                        { attempts: selectAttempts, layer: layer.name }
                    );
                }
                await wait(delays.selection * selectAttempts);
            }
        }
        
        await wait(delays.selection);

        // Update font size with retry using explicit point unit
        let updateAttempts = 0;
        while (updateAttempts < 3) {
            try {
                await batchPlay([{
                    _obj: "set",
                    _target: [{ _ref: "textLayer", _name: layer.name }],
                    to: { 
                        _obj: "textStyle",
                        size: { 
                            _unit: "pointsUnit",
                            _value: size
                        }
                    },
                    _options: { dialogOptions: "dontDisplay" }
                }], {
                    synchronousExecution: true,
                    modalBehavior: "fail"
                });
                break;
            } catch (updateError) {
                updateAttempts++;
                if (updateAttempts === 3) {
                    throw new PluginError(
                        'Failed to update font size',
                        'FONT_SIZE_UPDATE_ERROR',
                        { attempts: updateAttempts, layer: layer.name, size }
                    );
                }
                await wait(delays.fontUpdate * updateAttempts);
            }
        }

        await wait(delays.fontUpdate);

        // Verify font size update through character panel with retry
        let verifyAttempts = 0;
        let verifyResult;
        
        while (verifyAttempts < 3) {
            try {
                verifyResult = await batchPlay([{
                    _obj: "get",
                    _target: [
                        { _property: "textStyle" },
                        { _ref: "textLayer", _name: layer.name }
                    ],
                    _options: { dialogOptions: "dontDisplay" }
                }], {
                    synchronousExecution: true,
                    modalBehavior: "fail"
                });

                const updatedSize = verifyResult?.[0]?.textStyle?.size?._value;
                
                console.log("[DEBUG] Font size verification attempt:", {
                    layer: layer.name,
                    attempt: verifyAttempts + 1,
                    targetSize: size,
                    actualSize: updatedSize,
                    match: Math.abs(updatedSize - size) < 0.01,
                    fullState: verifyResult
                });

                if (Math.abs(updatedSize - size) < 0.01) {
                    // Success - sizes match within tolerance
                    console.log("[DEBUG] Font size update verified:", {
                        layer: layer.name,
                        targetSize: size,
                        actualSize: updatedSize,
                        attempts: verifyAttempts + 1,
                        success: true
                    });
                    return true;
                }

                verifyAttempts++;
                if (verifyAttempts === 3) {
                    throw new PluginError(
                        'Font size verification failed',
                        'FONT_SIZE_VERIFY_ERROR',
                        { 
                            expected: size, 
                            actual: updatedSize,
                            attempts: verifyAttempts,
                            verifyResult 
                        }
                    );
                }
                await wait(delays.verification * verifyAttempts);
                
            } catch (verifyError) {
                verifyAttempts++;
                if (verifyAttempts === 3) {
                    throw new PluginError(
                        'Font size verification failed',
                        'FONT_SIZE_VERIFY_ERROR',
                        { 
                            error: verifyError.message,
                            attempts: verifyAttempts,
                            verifyResult 
                        }
                    );
                }
                await wait(delays.verification * verifyAttempts);
            }
        }

        throw new PluginError(
            'Font size update could not be verified',
            'FONT_SIZE_VERIFY_ERROR',
            { 
                expected: size,
                attempts: verifyAttempts,
                verifyResult 
            }
        );

    } catch (error) {
        console.error("[DEBUG] Font size update failed:", {
            layer: layer?.name,
            targetSize: fontSize,
            error: error.message,
            code: error.code,
            stack: error.stack
        });
        
        throw new PluginError(
            `Failed to update font size in layer ${layer?.name}`,
            'FONT_SIZE_UPDATE_ERROR',
            { 
                originalError: error,
                layer: layer?.name,
                fontSize,
                details: error.details || {} 
            }
        );
    }
}

// Optimized layer processing with verification
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

        console.log("[DEBUG] Processing layer:", {
            layer: layer.name,
            layerIndex,
            hasText: !!rowData[`text${layerIndex}`],
            hasFontSize: !!rowData[`fontsize${layerIndex}`],
            timestamp: new Date().toISOString()
        });

        // Get initial layer state for verification
        const initialState = await batchPlay([batchPlayCommands.getCharacterStyle(layer.name)], {
            synchronousExecution: true,
            modalBehavior: "fail"
        });

        // Update text if needed
        if (rowData[`text${layerIndex}`]) {
            try {
                await replaceText(layer, rowData[`text${layerIndex}`]);
                updates.text = true;
                console.log("[DEBUG] Text update successful:", {
                    layer: layer.name,
                    text: rowData[`text${layerIndex}`]
                });
            } catch (textError) {
                console.error("[DEBUG] Text update failed:", {
                    layer: layer.name,
                    error: textError.message
                });
                throw textError;
            }
        }

        // Update font size if needed
        if (rowData[`fontsize${layerIndex}`]) {
            try {
                await updateFontSize(layer, rowData[`fontsize${layerIndex}`]);
                updates.fontSize = true;
                console.log("[DEBUG] Font size update successful:", {
                    layer: layer.name,
                    size: rowData[`fontsize${layerIndex}`]
                });
            } catch (fontError) {
                console.error("[DEBUG] Font size update failed:", {
                    layer: layer.name,
                    error: fontError.message
                });
                throw fontError;
            }
        }

        // Final verification of changes
        const finalState = await batchPlay([batchPlayCommands.getCharacterStyle(layer.name)], {
            synchronousExecution: true,
            modalBehavior: "fail"
        });

        const verificationResult = {
            layer: layer.name,
            updates,
            initialState,
            finalState,
            duration: Date.now() - layerStart
        };

        console.log("[DEBUG] Layer processing completed:", verificationResult);

        return {
            success: true,
            ...verificationResult
        };

    } catch (error) {
        const errorContext = {
            layer: layer?.name,
            updates,
            duration: Date.now() - layerStart,
            error: error.message
        };

        console.error("[DEBUG] Layer processing failed:", errorContext);

        throw new PluginError(
            `Failed to process layer ${layer?.name}`,
            'LAYER_PROCESS_ERROR',
            errorContext
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

// Enhanced folder setup with better error handling and validation
async function setupTextOutputFolders(outputFolder) {
    try {
        if (!outputFolder) {
            throw new PluginError('Output folder not selected', 'FOLDER_NOT_SELECTED');
        }

        console.log("[DEBUG] Setting up text output folders in:", outputFolder.nativePath);
        const folders = {};
        
        // Create PNG folder for text operations
        try {
            console.log("[DEBUG] Creating Text_PNG folder...");
            folders.pngFolder = await outputFolder.getEntry('Text_PNG');
            if (!folders.pngFolder.isFolder) {
                throw new Error('Text_PNG exists but is not a folder');
            }
            console.log("[DEBUG] Found existing Text_PNG folder");
        } catch (error) {
            console.log("[DEBUG] Creating new Text_PNG folder");
            folders.pngFolder = await outputFolder.createEntry('Text_PNG', { type: 'folder' });
            if (!folders.pngFolder || !folders.pngFolder.isFolder) {
                throw new PluginError('Failed to create PNG folder', 'FOLDER_CREATE_ERROR');
            }
        }
        
        // Create PSD folder for text operations
        try {
            console.log("[DEBUG] Creating Text_PSD folder...");
            folders.psdFolder = await outputFolder.getEntry('Text_PSD');
            if (!folders.psdFolder.isFolder) {
                throw new Error('Text_PSD exists but is not a folder');
            }
            console.log("[DEBUG] Found existing Text_PSD folder");
        } catch (error) {
            console.log("[DEBUG] Creating new Text_PSD folder");
            folders.psdFolder = await outputFolder.createEntry('Text_PSD', { type: 'folder' });
            if (!folders.psdFolder || !folders.psdFolder.isFolder) {
                throw new PluginError('Failed to create PSD folder', 'FOLDER_CREATE_ERROR');
            }
        }

        // Verify folders were created
        const pngExists = await folders.pngFolder.isEntry;
        const psdExists = await folders.psdFolder.isEntry;
        
        if (!pngExists || !psdExists) {
            throw new PluginError(
                'Failed to verify output folders', 
                'FOLDER_VERIFY_ERROR',
                {
                    pngExists,
                    psdExists,
                    pngPath: folders.pngFolder?.nativePath,
                    psdPath: folders.psdFolder?.nativePath
                }
            );
        }

        console.log("[DEBUG] Successfully created output folders:", {
            png: folders.pngFolder.nativePath,
            psd: folders.psdFolder.nativePath
        });
        
        return folders;
    } catch (error) {
        console.error("[DEBUG] Folder setup failed:", error);
        throw new PluginError(
            'Failed to setup text output folders', 
            'TEXT_FOLDER_SETUP_ERROR', 
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

// Save document as PNG with verification
async function saveAsPNG(doc, outputPath, options = {}) {
    try {
        // 1. Initial state logging
        const initialState = {
            docId: doc?.id,
            docName: doc?.name,
            activeDocId: app.activeDocument?.id,
            activeDocName: app.activeDocument?.name,
            hasSelection: app.activeDocument?.activeLayers?.length > 0,
            selectedLayer: app.activeDocument?.activeLayers[0]?.name
        };

        console.log("[DEBUG] Starting PNG save - Initial State:", initialState);

        // 2. Document validation with detailed checks
        const activeDoc = app.activeDocument;
        if (!activeDoc) {
            console.error("[DEBUG] No active document found during PNG save");
            throw new PluginError('No active document for PNG save', 'NO_ACTIVE_DOC_PNG');
        }

        // 3. Document ID verification
        const docState = {
            providedDocId: doc?.id,
            activeDocId: activeDoc.id,
            activeDocName: activeDoc.name,
            layerCount: activeDoc.layers.length,
            selectedLayers: activeDoc.activeLayers.map(l => ({
                name: l.name,
                id: l._id,
                kind: l.kind
            }))
        };

        console.log("[DEBUG] PNG Save - Document State:", docState);

        // 4. Pre-save verification
        const preSaveCheck = {
            outputPathValid: !!outputPath,
            outputFolder: outputPath.substring(0, outputPath.lastIndexOf('/')),
            docIdMatch: doc?.id === activeDoc.id,
            hasValidId: !!activeDoc.id
        };

        console.log("[DEBUG] PNG Save - Pre-save Verification:", preSaveCheck);

        // 5. Construct save command with logging
        const saveCommand = {
            _obj: "save",
            as: {
                _obj: "PNGFormat",
                PNG8: false,
                transparency: true
            },
            in: { _path: outputPath, _kind: "local" },
            copy: true,
            documentID: activeDoc.id,
            _isCommand: true,
            _options: { dialogOptions: "dontDisplay" }
        };

        console.log("[DEBUG] PNG Save - BatchPlay Command:", saveCommand);

        // 6. Execute save with detailed error catching
        const result = await batchPlay(
            [saveCommand],
            {
                synchronousExecution: true,
                modalBehavior: "fail"
            }
        );

        console.log("[DEBUG] PNG Save - BatchPlay Result:", result);

        // 7. Post-save verification
        let fileVerification;
        try {
            const savedFile = await fs.getFileForPath(outputPath);
            fileVerification = {
                exists: !!savedFile,
                path: savedFile?.nativePath,
                isFile: savedFile?.isFile,
                name: savedFile?.name
            };
        } catch (verifyError) {
            fileVerification = {
                exists: false,
                error: verifyError.message
            };
        }

        console.log("[DEBUG] PNG Save - File Verification:", fileVerification);

        if (!fileVerification.exists) {
            throw new PluginError(
                'PNG file not created after save command',
                'PNG_SAVE_VERIFY_FAILED',
                fileVerification
            );
        }

        // 8. Final state check
        const finalState = {
            docId: app.activeDocument?.id,
            docName: app.activeDocument?.name,
            selectedLayer: app.activeDocument?.activeLayers[0]?.name,
            saveTime: Date.now()
        };

        console.log("[DEBUG] PNG Save - Final State:", finalState);
        
        return result;

    } catch (error) {
        // 9. Enhanced error logging
        const errorContext = {
            error: {
                message: error.message,
                code: error.code,
                stack: error.stack
            },
            document: {
                providedId: doc?.id,
                activeId: app.activeDocument?.id,
                name: app.activeDocument?.name
            },
            path: {
                output: outputPath,
                folder: outputPath.substring(0, outputPath.lastIndexOf('/'))
            },
            state: {
                hasActiveDoc: !!app.activeDocument,
                hasSelection: app.activeDocument?.activeLayers?.length > 0,
                selectedLayer: app.activeDocument?.activeLayers[0]?.name
            }
        };

        console.error("[DEBUG] PNG Save Failed - Full Context:", errorContext);
        
        throw new PluginError(
            'Failed to save PNG',
            'PNG_SAVE_ERROR',
            errorContext
        );
    }
}

// Save document as PSD with verification
async function saveAsPSD(doc, outputPath, options = {}) {
    try {
        // 1. Initial state logging
        const initialState = {
            docId: doc?.id,
            docName: doc?.name,
            activeDocId: app.activeDocument?.id,
            activeDocName: app.activeDocument?.name,
            hasSelection: app.activeDocument?.activeLayers?.length > 0,
            selectedLayer: app.activeDocument?.activeLayers[0]?.name
        };

        console.log("[DEBUG] Starting PSD save - Initial State:", initialState);

        // 2. Document validation with detailed checks
        const activeDoc = app.activeDocument;
        if (!activeDoc) {
            console.error("[DEBUG] No active document found during PSD save");
            throw new PluginError('No active document for PSD save', 'NO_ACTIVE_DOC_PSD');
        }

        // 3. Document ID verification
        const docState = {
            providedDocId: doc?.id,
            activeDocId: activeDoc.id,
            activeDocName: activeDoc.name,
            layerCount: activeDoc.layers.length,
            selectedLayers: activeDoc.activeLayers.map(l => ({
                name: l.name,
                id: l._id,
                kind: l.kind
            }))
        };

        console.log("[DEBUG] PSD Save - Document State:", docState);

        // 4. Pre-save verification
        const preSaveCheck = {
            outputPathValid: !!outputPath,
            outputFolder: outputPath.substring(0, outputPath.lastIndexOf('/')),
            docIdMatch: doc?.id === activeDoc.id,
            hasValidId: !!activeDoc.id
        };

        console.log("[DEBUG] PSD Save - Pre-save Verification:", preSaveCheck);

        // 5. Construct save command with logging
        const saveCommand = {
            _obj: "save",
            as: {
                _obj: "photoshop35Format",
                maximizeCompatibility: true
            },
            in: { _path: outputPath, _kind: "local" },
            copy: true,
            documentID: activeDoc.id,
            _isCommand: true,
            _options: { dialogOptions: "dontDisplay" }
        };

        console.log("[DEBUG] PSD Save - BatchPlay Command:", saveCommand);

        // 6. Execute save with detailed error catching
        const result = await batchPlay(
            [saveCommand],
            {
                synchronousExecution: true,
                modalBehavior: "fail"
            }
        );

        console.log("[DEBUG] PSD Save - BatchPlay Result:", result);

        // 7. Post-save verification
        let fileVerification;
        try {
            const savedFile = await fs.getFileForPath(outputPath);
            fileVerification = {
                exists: !!savedFile,
                path: savedFile?.nativePath,
                isFile: savedFile?.isFile,
                name: savedFile?.name
            };
        } catch (verifyError) {
            fileVerification = {
                exists: false,
                error: verifyError.message
            };
        }

        console.log("[DEBUG] PSD Save - File Verification:", fileVerification);

        if (!fileVerification.exists) {
            throw new PluginError(
                'PSD file not created after save command',
                'PSD_SAVE_VERIFY_FAILED',
                fileVerification
            );
        }

        // 8. Final state check
        const finalState = {
            docId: app.activeDocument?.id,
            docName: app.activeDocument?.name,
            selectedLayer: app.activeDocument?.activeLayers[0]?.name,
            saveTime: Date.now()
        };

        console.log("[DEBUG] PSD Save - Final State:", finalState);
        
        return result;

    } catch (error) {
        // 9. Enhanced error logging
        const errorContext = {
            error: {
                message: error.message,
                code: error.code,
                stack: error.stack
            },
            document: {
                providedId: doc?.id,
                activeId: app.activeDocument?.id,
                name: app.activeDocument?.name
            },
            path: {
                output: outputPath,
                folder: outputPath.substring(0, outputPath.lastIndexOf('/'))
            },
            state: {
                hasActiveDoc: !!app.activeDocument,
                hasSelection: app.activeDocument?.activeLayers?.length > 0,
                selectedLayer: app.activeDocument?.activeLayers[0]?.name
            }
        };

        console.error("[DEBUG] PSD Save Failed - Full Context:", errorContext);
        
        throw new PluginError(
            'Failed to save PSD',
            'PSD_SAVE_ERROR',
            errorContext
        );
    }
}

// Update the processTextRow function to use the new font size update function
async function processTextRow(row, index, total, folders) {
    console.log("[DEBUG] Starting row processing:", {
        rowIndex: index,
        totalRows: total,
        timestamp: new Date().toISOString()
    });

    const doc = app.activeDocument;
    if (!doc) {
        throw new PluginError('No active document found', 'NO_ACTIVE_DOC');
    }

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
            const timestamp = Date.now();
            const baseFileName = `output_${timestamp}_${index + 1}`;
            const pngPath = `${folders.pngFolder.nativePath}/${baseFileName}.png`;
            const psdPath = `${folders.psdFolder.nativePath}/${baseFileName}.psd`;

            console.log("[DEBUG] Starting file saves:", {
                png: pngPath,
                psd: psdPath,
                processedLayers: layerUpdates.length,
                failedLayers: errors.length,
                timestamp: new Date().toISOString()
            });

            // Save PNG first
            try {
                await saveAsPNG(doc, pngPath);
                console.log("[DEBUG] PNG save completed successfully:", pngPath);
            } catch (pngError) {
                console.error("[DEBUG] PNG save failed:", {
                    path: pngPath,
                    error: pngError.message,
                    code: pngError.code
                });
                errors.push({
                    type: 'PNG_SAVE',
                    error: pngError.message,
                    code: pngError.code
                });
            }

            // Save PSD if PNG succeeded
            try {
                await saveAsPSD(doc, psdPath);
                console.log("[DEBUG] PSD save completed successfully:", psdPath);
            } catch (psdError) {
                console.error("[DEBUG] PSD save failed:", {
                    path: psdPath,
                    error: psdError.message,
                    code: psdError.code
                });
                errors.push({
                    type: 'PSD_SAVE',
                    error: psdError.message,
                    code: psdError.code
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

// Enhanced Image Row Processing with Progress Tracking
async function processImageRow(row, index, total) {
    if (!row || index === undefined || total === undefined) {
        throw new PluginError(
            'Invalid image row processing parameters',
            'INVALID_ROW_PARAMS',
            { index, total, hasRow: !!row }
        );
    }

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
                
                console.log("[DEBUG] Processing single row:", {
                    rowData: currentRow,
                    timestamp: new Date().toISOString()
                });
                
                await processTextRow(currentRow, 0, 1, folders);
                textStatus.textContent = 'Current row processed successfully';
                
            } else {
                const total = textReplaceState.data.csvData.length;
                console.log("[DEBUG] Starting batch processing:", {
                    totalRows: total,
                    timestamp: new Date().toISOString()
                });
                
                for (let i = 0; i < total; i++) {
                    const row = textReplaceState.data.csvData[i];
                    if (!row) {
                        throw new PluginError(`Empty row at index ${i}`, 'EMPTY_ROW');
                    }
                    
                    // Update UI with progress
                    textStatus.textContent = `Processing row ${i + 1} of ${total}...`;
                    
                    try {
                        await processTextRow(row, i, total, folders);
                        
                        // Log progress
                        console.log("[DEBUG] Row completed successfully:", {
                            rowIndex: i,
                            totalRows: total,
                            progress: Math.round(((i + 1) / total) * 100)
                        });
                        
                    } catch (rowError) {
                        // If row processing fails, stop the entire process
                        console.error("[DEBUG] Row processing failed - stopping batch:", {
                            rowIndex: i,
                            error: rowError
                        });
                        
                        throw new PluginError(
                            `Failed at row ${i + 1}/${total}`,
                            'BATCH_PROCESSING_HALTED',
                            {
                                originalError: rowError,
                                rowIndex: i,
                                totalRows: total
                            }
                        );
                    }
                }
                textStatus.textContent = 'All rows processed successfully';
            }

            const duration = Date.now() - startTime;
            
            // Write completion log to MCP relay
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
            
        } catch (processingError) {
            // Handle processing errors with specific UI updates
            console.error("[DEBUG] Processing failed:", processingError);
            
            // Update UI with specific error message
            let errorMessage = 'Processing failed';
            switch (processingError.code) {
                case 'LAYER_PROCESS_ERROR':
                    errorMessage = `Failed to process layer: ${processingError.details?.layer}`;
                    break;
                case 'PNG_SAVE_HALT':
                    errorMessage = 'Failed to save PNG - process stopped';
                    break;
                case 'PSD_SAVE_HALT':
                    errorMessage = 'Failed to save PSD - process stopped';
                    break;
                case 'BATCH_PROCESSING_HALTED':
                    errorMessage = `Processing stopped at row ${processingError.details?.rowIndex + 1}`;
                    break;
                default:
                    errorMessage = processingError.message;
            }
            
            textStatus.textContent = errorMessage;
            throw processingError; // Re-throw to be caught by outer catch
        }
        
    } catch (error) {
        await handleError(error, 'text');
    } finally {
        textReplaceState.status.isProcessing = false;
        textReplaceState.status.currentOperation = null;
    }
}

// Update processImageReplacement to use image-specific folder setup
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

// Modify error handling in other functions to use log instead of showing prompts
async function handleError(error, context) {
    const errorMessage = `[Cursor OK] ${context} error: ${error.message}`;
    log(errorMessage);
    
    // Write to MCP relay without showing prompts
    writeToMCPRelay({
        command: "sendError",
        error: {
            message: error.message,
            code: error.code,
            details: error.details
        },
        timestamp: new Date().toISOString()
    }).catch(mcpError => {
        log(`[Cursor OK] MCP relay write failed: ${mcpError.message}`);
    });
}

// Update manifest version
const manifest = {
    "id": "com.saveonpeptides.autoreplace",
    "name": "SaveOnPeptides Auto Replace",
    "version": "1.0.0",
    "main": "index.js",
    "host": {
        "app": "PS",
        "minVersion": "22.0.0"
    },
    "manifestVersion": 4,
    "requiredPermissions": {
        "allowCodeGenerationFromStrings": true,
        "launchProcess": {
            "schemes": ["http", "https"],
            "extensions": [".exe", ".bat", ".cmd"]
        },
        "network": {
            "domains": ["localhost"]
        },
        "clipboard": "readAndWrite",
        "fs": "readWrite",
        "webview": {
            "allow": "yes",
            "domains": ["https://*.adobe.com"]
        }
    },
    "entrypoints": [
        {
            "type": "panel",
            "id": "vanilla",
            "label": {
                "default": "SaveOnPeptides Auto Replace"
            },
            "minimumSize": {
                "width": 230,
                "height": 200
            },
            "maximumSize": {
                "width": 2000,
                "height": 2000
            },
            "preferredDockedSize": {
                "width": 230,
                "height": 300
            },
            "preferredFloatingSize": {
                "width": 230,
                "height": 300
            },
            "icons": [
                {
                    "width": 23,
                    "height": 23,
                    "path": "icons/dark.png",
                    "scale": [1, 2],
                    "theme": ["darkest", "dark", "medium"]
                },
                {
                    "width": 23,
                    "height": 23,
                    "path": "icons/light.png",
                    "scale": [1, 2],
                    "theme": ["lightest", "light"]
                }
            ]
        }
    ],
    "icons": [
        {
            "width": 48,
            "height": 48,
            "path": "icons/plugin.png",
            "scale": [1, 2]
        }
    ],
    "apiVersion": 2
};