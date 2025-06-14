# PLAN-ACT Process: A Safe Guide to Fixing the PNG Save Issue

**Objective:** Correct the PNG saving mechanism without altering the main CSV processing loop or other working parts of the plugin. We will isolate the problem, apply a fix, test it in a safe environment, and then confirm its success.

---

## Phase 1: PLAN (Understanding and Strategy)

### Step 1: Analyze the Root Cause

After reviewing your latest `index.js`, the exact issue has been located. The problem lies within the `saveDocumentAsPNG` function.

Here is the relevant section of your current code:

```javascript
async function saveDocumentAsPNG(document, folder, fileName) {
    // ...
    const file = await folder.createFile(fileName, { overwrite: true });
    const descriptor = {
        // ...
        in: { _path: file, _kind: 'local' }, // <-- This is the precise error
        // ...
    };
    await batchPlay([descriptor], { /* ... */ });
}

The Diagnosis: The batchPlay command cannot understand the file object that folder.createFile() creates. It requires a special, secure "session token" that points to that file. The command fails silently because the _path it receives is not in the format it expects.

Step 2: Define the Correction
The fix is to add one line of code that generates this required token and then use that token in the batchPlay command.

The Corrected Logic:

async function saveDocumentAsPNG(document, folder, fileName) {
    // ...
    const file = await folder.createFile(fileName, { overwrite: true });

    // THE FIX: Generate a session token from the file object.
    const token = fs.createSessionToken(file);

    const descriptor = {
        // ...
        in: { _path: token, _kind: 'local' }, // <-- Use the token here.
        // ...
    };
    await batchPlay([descriptor], { /* ... */ });
}

Step 3: Create a Safe Test Plan
We will not modify the main processCSV function at first. Instead, we will:

Add a temporary "Test Save" button to the UI.
Use this button to call and test only the saveDocumentAsPNG function.
Apply the correction to the saveDocumentAsPNG function.
Verify the fix using the test button.
Once verified, the main "Run" button will automatically start working correctly with no further changes.
Remove the temporary test button.

Phase 2: ACT (Step-by-Step Instructions for Cursor)
Here are the precise prompts to give to Cursor to execute the plan safely.

Action 1: Create a Safe Test Environment
Your Prompt to Cursor:

In index.html, add a new temporary button below the "Stop Processing" button with the ID testSaveBtn and the text "Test Single Save".

Then, in index.js inside the setupUIEventListeners function, add a new event listener for this button. The event listener should call a new, empty async function named runSingleSaveTest.

Action 2: Replicate the Failure in Isolation
Your Prompt to Cursor:

Now, implement the runSingleSaveTest function. This function should:

Log "Running single save test..." to the log panel.
Get the activeDocument and the outputFolder from the textReplaceState.
Check if both the document and folder exist, and if not, log an error message.
If they exist, it should call the existing saveDocumentAsPNG function with the active document, the output folder, and a hardcoded filename like "test-output.png".
(After this step, run the plugin, select an output folder, and click the test button. You will see in the logs that it runs but no file is saved. This safely confirms the bug in an isolated environment.)

Action 3: Apply the Correction
Your Prompt to Cursor:

The test confirmed the bug. Now, please modify the saveDocumentAsPNG function. Inside it, right after the line const file = await folder.createFile(...), add this new line:
const token = fs.createSessionToken(file);

Then, in the descriptor object within the same function, change the in property's _path value from file to token.

Action 4: Verify the Fix
Your Action:

Reload the plugin in Photoshop (Plugins > Development > Reload...).
Open any document.
Use your UI to select an output folder.
Click the "Test Single Save" button again.
Expected Result: The file test-output.png should now appear successfully in your selected output folder. The logs should show a success message. Because you have fixed the core function, the main "Run" button for your CSV batch process will now also work correctly.

Action 5: Clean Up
Your Prompt to Cursor:

The fix is verified. Please remove the "Test Single Save" button from index.html and also remove the runSingleSaveTest function and its event listener from index.js.

This completes the process. The bug is fixed, the core functionality is preserved, and you have a reliable method for testing isolated functions in the future.