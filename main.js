

import { initMuPDFWorker } from "./mupdf/mupdf-async.js";
import { MSGReader } from "./lib/msg.reader.js";
import { ZipReader, BlobReader, TextWriter } from "./lib/zip.js/index.js";
import { getAllFileEntries } from "./js/drag-and-drop.js";
import { config, sizeLimits } from "./js/config.js";

import Tesseract from './lib/tesseract.esm.min.js';

const fileListSuccessElem = document.getElementById('fileListSuccess');
const fileListFailedElem = document.getElementById('fileListFailed');
const fileListSkippedElem = document.getElementById('fileListSkipped');
const fileCountSuccessElem = document.getElementById('fileCountSuccess');
const fileCountFailedElem = document.getElementById('fileCountFailed');
const fileCountSkippedElem = document.getElementById('fileCountSkipped');

const matchListElem = document.getElementById('matchList');

// Replace documentation URL if a non-default option is set in config object
const docsURLElem = document.getElementById('docsURL');
if (config.docsURL) {
    docsURLElem.setAttribute("href", config.docsURL);
}

globalThis.docNames = {};
globalThis.docText = {};
globalThis.docTextHighlighted = {};

globalThis.zone = document.getElementById("uploadDropZone");

// Inserts an error message with contents `innerHTML` into `divId`.
// `error` argument determines is message is warning (yellow) or error (red).
export function insertAlertMessage(innerHTML, error = true, divId = "alertDiv") {
    const warningSVG = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" class="bi flex-shrink-0 me-2" viewBox=" 0 0 16 16">
    <path d="M8 15A7 7 0 1 1 8 1a7 7 0 0 1 0 14zm0 1A8 8 0 1 0 8 0a8 8 0 0 0 0 16z" />
    <path d="M7.002 11a1 1 0 1 1 2 0 1 1 0 0 1-2 0zM7.1 4.995a.905.905 0 1 1 1.8 0l-.35 3.507a.552.552 0 0 1-1.1 0L7.1 4.995z" />
  </svg>`;
  
    const errorSVG = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" class="bi flex-shrink-0 me-2" viewBox=" 0 0 16 16">
    <path
      d="M7.938 2.016A.13.13 0 0 1 8.002 2a.13.13 0 0 1 .063.016.146.146 0 0 1 .054.057l6.857 11.667c.036.06.035.124.002.183a.163.163 0 0 1-.054.06.116.116 0 0 1-.066.017H1.146a.115.115 0 0 1-.066-.017.163.163 0 0 1-.054-.06.176.176 0 0 1 .002-.183L7.884 2.073a.147.147 0 0 1 .054-.057zm1.044-.45a1.13 1.13 0 0 0-1.96 0L.165 13.233c-.457.778.091 1.767.98 1.767h13.713c.889 0 1.438-.99.98-1.767L8.982 1.566z" />
    <path d="M7.002 12a1 1 0 1 1 2 0 1 1 0 0 1-2 0zM7.1 5.995a.905.905 0 1 1 1.8 0l-.35 3.507a.552.552 0 0 1-1.1 0L7.1 5.995z" />
  </svg>`;
  
    const chosenSVG = error ? errorSVG : warningSVG;
  
    const htmlDiv = document.createElement("div");
  
    htmlDiv.innerHTML = `<div class="alert alert-dismissible ${error ? "alert-danger" : "alert-warning"} d-flex align-items-center show fade mb-1">
    <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    ${chosenSVG}
    <div class="mb-0"> ${innerHTML} </div>
  </div>`;
  
    document.getElementById(divId)?.appendChild(htmlDiv);
  
  }
  

const importProgressCollapseElem = document.getElementById("import-progress-collapse");
globalThis.progressCollapseObj = new bootstrap.Collapse(importProgressCollapseElem, { toggle: false });
const progressBar = importProgressCollapseElem.getElementsByClassName("progress-bar")[0];
const progress = {
    max: 0,
    value: 0,
    elem: progressBar,
    show: () => progressCollapseObj.show(),
    hide: () => progressCollapseObj.hide(),
    setMax: async function (max) {
        this.max = max;
        this.elem.setAttribute("aria-valuemax", this.max);
        this.elem.setAttribute("style", "width: " + ((this.value) / this.max * 100) + "%");
        await new Promise((r) => setTimeout(r, 0));
    },
    setValue: async function (value) {
        this.value++;
        if ((this.value) % 5 == 0 || this.value == this.max) {
            this.elem.setAttribute("aria-valuenow", this.value.toString());
            this.elem.setAttribute("style", "width: " + (this.value / this.max * 100) + "%");
            await new Promise((r) => setTimeout(r, 0));
        }
    }
}


async function initMuPDFScheduler(workers = 3) {
    const scheduler = Tesseract.createScheduler();
    scheduler["workers"] = new Array(workers);
    for (let i = 0; i < workers; i++) {
        const w = await initMuPDFWorker();
        w.id = `png-${Math.random().toString(16).slice(3, 8)}`;
        scheduler.addWorker(w);
        scheduler["workers"][i] = w;
    }
    return scheduler;
}

async function getMuPDFScheduler() {
    if (!globalThis.muPDFScheduler) {
        globalThis.muPDFScheduler = initMuPDFScheduler();
    }
    return globalThis.muPDFScheduler;
}

globalThis.muPDFScheduler = initMuPDFScheduler();

zone.addEventListener('dragover', (event) => {
    event.preventDefault();
    event.target.classList.add('highlight');
});

zone.addEventListener('dragleave', (event) => {
    event.preventDefault();
    event.target.classList.remove('highlight');
});


// This is where the drop is handled.
zone.addEventListener('drop', async (event) => {
    // Prevent navigation.
    event.preventDefault();
    let items = await getAllFileEntries(event.dataTransfer.items);

    const filesPromises = await Promise.allSettled(items.map((x) => new Promise((resolve, reject) => x.file(resolve, reject))));
    const files = filesPromises.map(x => x.value);

    const filePaths = items.map((x) => x.fullPath.replace(/^\//, ""));

    readFiles(files, filePaths);

    event.target.classList.remove('highlight');

});

document.getElementById("viewerSpacer").setAttribute("style", "height:" + document.getElementById("titleArea").offsetHeight + "px");

const highlight = event => event.target.classList.add('highlight');

const unhighlight = event => event.target.classList.remove('highlight');

zone.addEventListener(event, highlight, false);

['dragenter', 'dragover'].forEach(event => {
    zone.addEventListener(event, highlight, false);
})

    // Highlighting drop area when item is dragged over it
    ;['dragenter', 'dragover'].forEach(event => {
        zone.addEventListener(event, highlight, false);
    });
;['dragleave', 'drop'].forEach(event => {
    zone.addEventListener(event, unhighlight, false);
});


// https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Math/random
export function getRandomInt(min, max) {
    min = Math.ceil(min);
    max = Math.floor(max);
    return Math.floor(Math.random() * (max - min) + min); //The maximum is exclusive and the minimum is inclusive
}

function getRandomAlphanum(num) {
    let outArr = new Array(num);
    for (let i = 0; i < num; i++) {
        let intI = getRandomInt(1, 62);
        if (intI <= 10) {
            intI = intI + 47;
        } else if (intI <= 36) {
            intI = intI - 10 + 64;
        } else {
            intI = intI - 36 + 96;
        }
        outArr[i] = String.fromCharCode(intI);
    }
    return outArr.join('');
}

const w = await initMuPDFWorker();

const readMsg = async (file) => {
    const msgReader = new MSGReader(await file.arrayBuffer());
    const fileData = msgReader.getFileData();
    const text = fileData.body;

    const attachmentFiles = [];
    for (let i = 0; i < fileData.attachments.length; i++) {
        const attachmentObj = msgReader.getAttachment(i);
        const attachmentFile = new File([attachmentObj.content], attachmentObj.fileName, { type: attachmentObj.mimeType ? attachmentObj.mimeType : "application/octet-stream" });
        attachmentFiles.push(attachmentFile);
    }
    if (attachmentFiles.length > 0) await readFiles(attachmentFiles);
    return text;
}

const readPdf = async (file) => {
    const fileIArray = await file.arrayBuffer();
    const fileData = new Uint8Array(fileIArray);

    const scheduler = await getMuPDFScheduler();

    let text = await scheduler.addJob("openDocumentExtractText", [fileData, "file.pdf"]);

    // PDF portfolios are essentially archive files with a .pdf extension
    // Unfortunately, they do not throw an error when read as a standard PDF files, but rather (at least when created using Acrobat)
    // show a page advising the user to install Acrobat.  Therefore, we detect this page and thrown an error.
    if (text.length < 500 && /^For the best experience, open this PDF portfolio in/.test(text)) {
        text = "";
        throw "PDF portfolio detected (not supported)"
    }

    return text;
}

const readDocx = async (file) => {
    const zipFileReader = new BlobReader(file);
    const zipReader = new ZipReader(zipFileReader);
    const entries = await zipReader.getEntries();
    const textWriter = new TextWriter();
    let text = "";

    for (let i = 0; i < entries.length; i++) {
        if (['word/document.xml', 'word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml'].includes(entries[i].filename)) {
            const xmlStr = await entries[i].getData(new TextWriter());

            // Get array of paragraph ("p") elements
            // This step allows for inserting line breaks between paragraphs
            const pArr = xmlStr.match(/(?<=\<w:p[^>\/]{0,200}?\>)[\s\S]+?(?=\<\/w:p\>)/g);

            if (!pArr) continue;

            for (let j = 0; j < pArr.length; j++) {

                // This matches both (1) normal text and (2) text inserted in tracked changes.
                // Text deleted in tracked changes is not included, as it is in "<w:delText>" tags rather than "<w:t>"
                const textArr = pArr[j].match(/\<w:t[^>\/]{0,200}?\>[\s\S]+?(?=\<\/w:t\>)/g);
                if (!textArr) continue;

                for (let k = 0; k < textArr.length; k++) {
                    text += textArr[k].replace(/\<w:t[^>\/]{0,200}?\>/, "") + " ";
                }
                text += "\n";

            }

        }
    }

    await zipReader.close();

    return text;
}

const readXlsx = async (file) => {
    const zipFileReader = new BlobReader(file);
    const zipReader = new ZipReader(zipFileReader);
    const entries = await zipReader.getEntries();
    const textWriter = new TextWriter();
    let text = "";

    for (let i = 0; i < entries.length; i++) {
        if (['xl/workbook.xml', 'xl/sharedStrings.xml'].includes(entries[i].filename) || /xl\/worksheets\/[^\/]+.xml/.test(entries[i].filename)) {
            const xmlStr = await entries[i].getData(new TextWriter());
            // This matches both (1) normal text and (2) text inserted in tracked changes.
            // Text deleted in tracked changes is not included, as it is in "<w:delText>" tags rather than "<w:t>"

            // Note: At present (2023) lookbehinds come with a MAJOR performance penalty.
            // Therefore, we instead leade on the opening tags and remove them in a later step.  
            // This may change in the future if lookbehind performance improves.
            // const textArr = xmlStr.match(/(?<=\<t[^>\/]{0,200}?\>)[\s\S]+?(?=\<\/t\>)/g);

            const textArr = xmlStr.match(/\<t[^>\/]{0,200}?\>[\s\S]+?(?=\<\/t\>)/g);
            if (!textArr) continue;

            for (let j = 0; j < textArr.length; j++) {
                text += textArr[j].replace(/\<t[^>\/]{0,200}?\>/, "") + " ";
            }
            text += "\n";
        }
    }

    await zipReader.close();

    return text;
}

const readPptx = async (file) => {
    const zipFileReader = new BlobReader(file);
    const zipReader = new ZipReader(zipFileReader);
    const entries = await zipReader.getEntries();
    const textWriter = new TextWriter();
    let text = "";

    for (let i = 0; i < entries.length; i++) {
        if (/ppt\/slides\/[^\/]+.xml/.test(entries[i].filename) || /ppt\/notesSlides\/[^\/]+.xml/.test(entries[i].filename) || /ppt\/comments\/[^\/]+.xml/.test(entries[i].filename)) {
            const xmlStr = await entries[i].getData(new TextWriter());
            // This matches both (1) normal text and (2) text inserted in tracked changes.
            // Text deleted in tracked changes is not included, as it is in "<w:delText>" tags rather than "<w:t>"
            const textArr = xmlStr.match(/\<a:t[^>\/]{0,200}?\>[\s\S]+?(?=\<\/a:t\>)/g);
            if (!textArr) continue;

            for (let j = 0; j < textArr.length; j++) {
                text += textArr[j].replace(/\<a:t[^>\/]{0,200}?\>/, "") + " ";
            }
            text += "\n";
        }
    }

    await zipReader.close();

    return text;
}

function readTxt(file) {
    return new Promise((resolve, reject) => {
        let reader = new FileReader();
        reader.onload = () => {
            resolve(reader.result);
        };
        reader.onerror = reject;
        reader.readAsText(file);
    });
}


const readHtml = async (file) => {
    let fileStr = await readTxt(file);
    // Delete any embedded Javascript code
    fileStr = fileStr.replaceAll(/\<script[^>]*?\>[\s\S]*?\<\/script\>/gi, "");
    const parser = new DOMParser();
    const htmlDoc = parser.parseFromString(fileStr, "text/html");
    // The text content often has an excessive number of newlines
    const text = htmlDoc.body.textContent?.replaceAll(/\n{2,}/g, "\n");

    return text;
}

// This object contains the mapping between file extensions and read functions.
const read = {
    docx: readDocx,
    htm: readHtml,
    html: readHtml,
    msg: readMsg,
    pdf: readPdf,
    pptx: readPptx,
    txt: readTxt,
    xlsx: readXlsx,
}

// Opt-in to bootstrap tooltip feature
// https://getbootstrap.com/docs/5.0/components/tooltips/
var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'))
var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
    return new bootstrap.Tooltip(tooltipTriggerEl);
})

// Create list group HTML element with `text`
// Used for adding files to success/skipped/failed lists
function createListItem(text) {
    const li = document.createElement("a");
    li.innerText = text;
    li.setAttribute("class", "list-group-item list-group-item-action");
    return li;
}

// Add file to the list of failed files
export function addToFailed(fileName) {
    const elem = createListItem(fileName);
    fileListFailedElem?.appendChild(elem);
    fileCountFailedElem.textContent = String(parseInt(fileCountFailedElem.textContent) + 1);
    return;
}

// Add file to the list of skipped files
function addToSkipped(fileName, reason = null) {
    const elem = createListItem(fileName);
    if (reason) elem.innerHTML = elem.innerHTML + "<span style='right:0;position:absolute'>[" + reason + "]</span>";
    fileListSkippedElem?.appendChild(elem);
    fileCountSkippedElem.textContent = String(parseInt(fileCountSkippedElem.textContent) + 1);
    return;
}

// Add file to the list of successful files
function addToSuccess(fileName) {
    const elem = createListItem(fileName);
    elem.setAttribute("data-bs-toggle", "list");
    elem.addEventListener("click", () => viewDoc(fileName));

    fileListSuccessElem?.appendChild(elem);
    fileCountSuccessElem.textContent = String(parseInt(fileCountSuccessElem.textContent) + 1);
    return;
}

let fileNWarningThrown = false;
/**
 * @param {File[]} files - Name of file
 * @param {string[]} filePaths - when using the drag-and-drop interface, file paths must be passed manually as an array
 *      as the File objects lack valid directory info.  This argument can be left blank when using a standard file input.
 */
async function readFiles(files, filePaths = []) {

    const start = Date.now();

    // The files are added to the existing total.
    // This is necessary as this function may be called recursively if files contain additional embedded files. 
    progress.setMax(progress.max + files.length);
    progress.show();

    // If this set of files is above the warning thershold, and no warning has been thrown yet,
    // create a warning message. 
    if (config.fileNWarning && progress.max > config.fileNWarning && !fileNWarningThrown) {
        insertAlertMessage(config.warningHTML, false);
        fileNWarningThrown = true;
    }

    const promiseArr = [];
    for (let i = 0; i < files.length; i++) {

        const file = files[i];

        const key = filePaths[i] || file.webkitRelativePath || file.name;

        const ext = file.name.match(/\.(\w{1,5})$/)?.[1]?.toLowerCase();

        const startFile = Date.now();

        if (!read[ext]) {
            addToSkipped(key, "Unsupported Extension");
            progress.setValue(progress.value + 1);
        } else if (sizeLimits[ext] && file.size > sizeLimits[ext]) {
            addToSkipped(key, "Over Size Limit");
            progress.setValue(progress.value + 1);
        } else {
            // TODO: This should eventually use promises + workers for better performance, but this will require edits.
            // Notably, as the same mupdf worker is reused, if run in asyc the PDF may be replaced before readPdf is finished reading it.
            // The other functions are not set up to run in workers.
            promiseArr[i] = read[ext](file).then((text) => {

                // Set `globalThis.debugMode = true` in the console to print the runtimes for each file
                const endFile = Date.now();
                if (globalThis.debugMode) console.log(`${key}: ${endFile - startFile} ms`);

                // Remove excessive newline characters to improve readability
                text = text.replaceAll(/(\n\s*){3,}/g, "\n\n");

                const fileNameBase = key.match(/[^\/]+$/, "")?.[0];

                // If another file exists with (1) the same name and (2) the same content, then this file is skipped as a duplicate.
                // This frequently occurs when the same file occurs both independently and as an email attachment.
                if (globalThis.docNames[fileNameBase] && text === globalThis.docText[globalThis.docNames[fileNameBase]]) {
                    addToSkipped(key, "Duplicate");
                    return;
                }

                // In the case of .pdf files, the file is marked as "skipped" rather than "success" if no text was extracted.
                // This is because the PDF is assumed to be an image-native PDF that would require OCR to extract.
                if (ext == "pdf" && text.trim() === "") {
                    addToSkipped(key, "No Text Content");
                    return;
                }

                globalThis.docNames[fileNameBase] = key;
                globalThis.docText[key] = text;

                addToSuccess(key);

        }).catch((error) => {
                console.log(error);
                addToFailed(key);

            }).finally(() => {
                progress.setValue(progress.value + 1);
            });
        }

    }

    await Promise.allSettled(promiseArr);

    const end = Date.now();
    console.log(`Execution time: ${end - start} ms`);

}

/**
 * @param {string} fileName - Name of file
 * @param {number} index - Index of the start of the snippet
 * @param {string} search - Search string
 */
function searchMatch(fileName, index, search) {
    /** @type {string} */
    this.fileName = fileName;
    /** @type {number} */
    this.index = index;
    // Snippets are always `contextLength`*2 characters long, with `contextLength` characters coming before and after `index` when possible.
    // For example, if `contextLength` is 100 and `index` is `300`, the snippet indices will be `200` and `400`.
    // If `index` is `0` then the snippet indices will be `0` and `200`. 
    // This avoids situations where the same match ends up in multiple snippets. 
    /** @type {number} */
    this.snippetStartIndex = Math.min(Math.max(index - contextLength, 0), globalThis.docText[fileName].length - contextLength * 2);
    /** @type {number} */
    this.snippetEndIndex = Math.max(Math.min(index + contextLength, globalThis.docText[fileName].length), contextLength * 2);
    /** @type {string} */
    this.search = search;
    /** @type {string} */
    this.id = getRandomAlphanum(10);
}

function getSnippet(match) {
    const replaceRegex = new RegExp("(" + match.search + ")", "ig");
    const snippetText = document.createElement("p");
    snippetText.setAttribute("class", 'mb-1');
    snippetText.textContent = globalThis.docText[match.fileName].slice(match.snippetStartIndex, match.snippetEndIndex);
    snippetText.innerHTML = snippetText.innerHTML.replaceAll(replaceRegex, "<b>$1</b>");

    return snippetText;
}

const contextLength = 100;

/**
 * @param {string} fileName - Name of file
 * @param {string} search - Search term
 */
function searchText(fileName, search) {

    const text = globalThis.docText[fileName];

    const regex = new RegExp(search, "gi");
    let result;
    const indices = [];
    while ((result = regex.exec(text))) {
        indices.push(result.index);
    }

    const matches = [];
    let lastIndexIncluded = 0;
    for (let i = 0; i < indices.length; i++) {
        const match = new searchMatch(fileName, indices[i], search);

        if (i == 0 || indices[i] + search.length > lastIndexIncluded) {
            matches.push(match);
            lastIndexIncluded = match.snippetEndIndex;
        }
    }

    return matches;

}

globalThis.matches = [];

async function searchDocs(search) {
    matchListElem.innerHTML = "";
    globalThis.matches = [];
    globalThis.docTextHighlighted = {};
    for (const [key, value] of Object.entries(globalThis.docText)) {
        globalThis.matches.push(...searchText(key, search));
    }

    for (let j = 0; j < globalThis.matches.length; j++) {
        const entry = document.createElement('a');
        entry.setAttribute("class", "list-group-item list-group-item-action flex-column align-items-start");
        entry.setAttribute("data-bs-toggle", "list");

        entry.addEventListener("click", () => viewResult(globalThis.matches[j]));

        entry.appendChild(getSnippet(globalThis.matches[j]));

        const fileName = document.createElement("small");
        fileName.textContent = globalThis.matches[j].fileName;

        entry.appendChild(fileName);

        matchListElem.appendChild(entry);
    }


    if (matchListElem.innerHTML == "") {
        const entry = document.createElement('a');
        entry.setAttribute("class", "list-group-item list-group-item-action flex-column align-items-start");

        const elem = document.createElement("p");
        elem.setAttribute("class", "mb-1");
        elem.textContent = "[No Results]";

        entry.appendChild(elem);

        matchListElem.appendChild(entry);
    }


}

const state = {
    // Whether the viewer has been initialized yet
    initViewerBool: false,
    // Whether the viewer is showing matches (not just a document) 
    viewerMatchMode: false,
    // The current file being displayed
    viewerFileName: ""
}

/**
* Initializes the document viewer UI if it has not already been initialized.
*/
async function initViewer () {
    if (!state.initViewerBool) {
        document.getElementById("viewerCol").style.width = "50%";
        // The location of the highlighted text is not detected correctly without waiting for the animation
        await new Promise((r) => setTimeout(r, 250));
        state.initViewerBool = true;
    }
}

/**
 * Opens document in viewer. 
 * Used when user selects a document to view, not a search result. 
* @param {fileName} String - Name of file
*/
async function viewDoc(fileName) {
    await initViewer();

    // Return early if the file being selected is already in the viewer
    if (!state.viewerMatchMode && state.viewerFileName == fileName) return;

    state.viewerMatchMode = false;
    state.viewerFileName = fileName;

    const elem = document.createElement("span");
    elem.setAttribute("style", 'white-space: pre-line');
    elem.textContent = globalThis.docText[fileName];

    document.getElementById("viewerCard").replaceChild(elem, document.getElementById("viewerCard").firstChild);

    document.getElementById("viewerCard").scrollTop = 0;
}

/**
* Opens document in viewer, highlights all matches, and scolls to the location of selected match.
* @param {searchMatch} match - Match object to scoll to in document
*/
async function viewResult(match) {
    await initViewer();

    state.viewerMatchMode = true;
    state.viewerFileName = match.fileName;

    if (!globalThis.docTextHighlighted[match.fileName]) {
        globalThis.docTextHighlighted[match.fileName] = document.createElement("span");
        const replaceRegex = new RegExp("(" + match.search + ")", "ig");
        // Get all matches for the same file
        const matches = globalThis.matches.filter(x => x.fileName == match.fileName);
        let lastIndex = 0;
        for (let i = 0; i < matches.length; i++) {

            // Add text after last match and before match i
            const preText = document.createElement("span");
            preText.setAttribute("style", 'white-space: pre-line');
            preText.textContent = globalThis.docText[matches[i].fileName].slice(lastIndex, matches[i].snippetStartIndex);
            globalThis.docTextHighlighted[match.fileName].appendChild(preText);

            // Snippets are wrapped in <span> tags with unique ids so (1) to highlight them and (2) so they can be scrolled to
            const snippetText = document.createElement("span");
            snippetText.setAttribute("id", matches[i].id);
            snippetText.setAttribute("style", 'white-space: pre-line;background-color:yellow');
            snippetText.textContent = globalThis.docText[matches[i].fileName].slice(matches[i].snippetStartIndex, matches[i].snippetEndIndex);
            snippetText.innerHTML = snippetText.innerHTML.replaceAll(replaceRegex, "<strong>$1</strong>");
            globalThis.docTextHighlighted[match.fileName].appendChild(snippetText);

            lastIndex = matches[i].snippetEndIndex;
        }

        // Add text after final match
        const postText = document.createElement("span");
        postText.setAttribute("style", 'white-space: pre-line');
        postText.textContent = globalThis.docText[match.fileName].slice(lastIndex);
        globalThis.docTextHighlighted[match.fileName].appendChild(postText);

    }
      

    document.getElementById("viewerCard").replaceChild(globalThis.docTextHighlighted[match.fileName], document.getElementById("viewerCard").firstChild)


    // Position the match ~1/3 of the way down the viewer
    document.getElementById("viewerCard").scrollTop = document.getElementById(match.id).offsetTop - document.getElementById("viewerCard").offsetHeight / 3;


}

document.getElementById('openFileInput').addEventListener('change', (event) => readFiles(event.target.files));
document.getElementById('openDirInput').addEventListener('change', (event) => readFiles(event.target.files));


document.getElementById('searchTextInput').addEventListener('keyup', function (event) {
    if (event.keyCode === 13) {
        searchDocs(document.getElementById("searchTextInput").value);
    }
});



document.getElementById('searchTextBtn').addEventListener('click', (event) => searchDocs(document.getElementById("searchTextInput").value));

document.getElementById("supportedFormats").innerText = Object.keys(read).join(", ");