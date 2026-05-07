const STORAGE_KEY = "smartwriter_gemini_key";
const LAST_SELECTION = "smartwriter_last_selection";
const LAST_RESULT = "smartwriter_last_result";

Office.onReady(async () => {
    await loadSelection();
    loadSavedState();
    checkApiKey();
    bindEvents();
});

function bindEvents() {
    document.getElementById("saveApiKey").onclick = saveApiKey;
    document.getElementById("toggleSettings").onclick = toggleSettings;
    document.getElementById("improveBtn").onclick = improveWriting;
    document.getElementById("replaceBtn").onclick = replaceText;
    document.getElementById("insertBtn").onclick = insertBelow;
}

function checkApiKey() {
    const key = localStorage.getItem(STORAGE_KEY);
    if (!key) {
        document.getElementById("settingsPanel").classList.remove("hidden");
    }
}

function saveApiKey() {
    const key = document.getElementById("apiKeyInput").value.trim();
    localStorage.setItem(STORAGE_KEY, key);
    document.getElementById("settingsPanel").classList.add("hidden");
}

function toggleSettings() {
    document.getElementById("settingsPanel").classList.toggle("hidden");
}

async function loadSelection() {
    await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load("text");
        await context.sync();

        if (range.text) {
            document.getElementById("selectedText").value = range.text;
            localStorage.setItem(LAST_SELECTION, range.text);
        }
    });
}

function loadSavedState() {
    const lastSelection = localStorage.getItem(LAST_SELECTION);
    const lastResult = localStorage.getItem(LAST_RESULT);

    if (lastSelection) {
        document.getElementById("selectedText").value = lastSelection;
    }

    if (lastResult) {
        document.getElementById("resultText").value = lastResult;
    }
}

async function improveWriting() {
    const apiKey = localStorage.getItem(STORAGE_KEY);
    const text = document.getElementById("selectedText").value;
    const instructions = document.getElementById("instructions").value;

    if (!apiKey || !text) return;

    toggleSpinner(true);

    const prompt = `
}