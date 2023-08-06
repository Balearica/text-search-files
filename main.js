

import { initMuPDFWorker } from "./mupdf/mupdf-async.js";
import { MSGReader } from "./lib/msg.reader.js";
import { ZipReader, BlobReader, TextWriter } from "./lib/zip.js/index.js";

const fileListSuccessElem = document.getElementById('fileListSuccess');
const fileListFailedElem = document.getElementById('fileListFailed');
const fileListSkippedElem = document.getElementById('fileListSkipped');
const fileCountSuccessElem = document.getElementById('fileCountSuccess');
const fileCountFailedElem = document.getElementById('fileCountFailed');
const fileCountSkippedElem = document.getElementById('fileCountSkipped');

const matchListElem = document.getElementById('matchList');

globalThis.docText = {};

const w = await initMuPDFWorker();

const readMsg = async (file) => {
    const msgReader = new MSGReader(await file.arrayBuffer());
    const fileData = msgReader.getFileData();
    globalThis.docText[file.name] += fileData.body;

    const attachmentFiles = [];
    for (let i=0; i<fileData.attachments.length; i++) {
        const attachmentObj = msgReader.getAttachment(i);
        const attachmentFile = new File([attachmentObj.content], attachmentObj.fileName, {type: attachmentObj.mimeType ? attachmentObj.mimeType : "application/octet-stream"});
        attachmentFiles.push(attachmentFile);
    }
    if (attachmentFiles.length > 0) await readFiles(attachmentFiles);
}

const readPdf = async (file) => {
    const fileIArray = await file.arrayBuffer();
    const fileData = new Uint8Array(fileIArray);


    const pdfDoc = await w.openDocument(fileData, "file.pdf");
    w["pdfDoc"] = pdfDoc;

    const pageCountImage = await w.countPages([]);

    for (let j = 0; j < pageCountImage; j++) {
        globalThis.docText[file.name] += await w.pageText([j + 1, 72, false]);
    }

}

const readDocx = async (file) => {
    const zipFileReader = new BlobReader(file);
    const zipReader = new ZipReader(zipFileReader);
    const entries = await zipReader.getEntries();
    const textWriter = new TextWriter();

    for (let i = 0; i < entries.length; i++) {
        if (['word/document.xml', 'word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml'].includes(entries[i].filename)) {
            const xmlStr = await entries[i].getData(new TextWriter());

            // Get array of paragraph ("p") elements
            // This step allows for inserting line breaks between paragraphs
            const pArr = xmlStr.match(/(?<=\<w:p[^>\/]{0,200}?\>)[\s\S]+?(?=\<\/w:p\>)/g);

            if (!pArr) continue;

            for (let j=0; j < pArr.length; j++) {

                // This matches both (1) normal text and (2) text inserted in tracked changes.
                // Text deleted in tracked changes is not included, as it is in "<w:delText>" tags rather than "<w:t>"
                const textArr = pArr[j].match(/(?<=\<w:t[^>\/]{0,200}?\>)[\s\S]+?(?=\<\/w:t\>)/g);
                if (!textArr) continue;

                for (let k = 0; k < textArr.length; k++) {
                    globalThis.docText[file.name] += textArr[k] + " ";
                }
                globalThis.docText[file.name] += "\n";

            }

        }
    }

    await zipReader.close();

}

const readXlsx = async (file) => {
    const zipFileReader = new BlobReader(file);
    const zipReader = new ZipReader(zipFileReader);
    const entries = await zipReader.getEntries();
    const textWriter = new TextWriter();

    for (let i = 0; i < entries.length; i++) {
        if (['xl/workbook.xml', 'xl/sharedStrings.xml'].includes(entries[i].filename)) {
            const xmlStr = await entries[i].getData(new TextWriter());
            // This matches both (1) normal text and (2) text inserted in tracked changes.
            // Text deleted in tracked changes is not included, as it is in "<w:delText>" tags rather than "<w:t>"
            const textArr = xmlStr.match(/(?<=\<t[^>\/]{0,30}?\>)[\s\S]+?(?=\<\/t\>)/g);
            if (!textArr) continue;

            for (let j = 0; j < textArr.length; j++) {
                globalThis.docText[file.name] += textArr[j] + " ";
            }
            globalThis.docText[file.name] += "\n";
        }
    }

    await zipReader.close();

}

const readPptx = async (file) => {
    const zipFileReader = new BlobReader(file);
    const zipReader = new ZipReader(zipFileReader);
    const entries = await zipReader.getEntries();
    const textWriter = new TextWriter();

    for (let i = 0; i < entries.length; i++) {
        if (/ppt\/slides\/[^\/]+.xml/.test(entries[i].filename) || /ppt\/notesSlides\/[^\/]+.xml/.test(entries[i].filename) || /ppt\/comments\/[^\/]+.xml/.test(entries[i].filename)) {
            const xmlStr = await entries[i].getData(new TextWriter());
            // This matches both (1) normal text and (2) text inserted in tracked changes.
            // Text deleted in tracked changes is not included, as it is in "<w:delText>" tags rather than "<w:t>"
            const textArr = xmlStr.match(/(?<=\<a:t[^>\/]{0,30}?\>)[\s\S]+?(?=\<\/a:t\>)/g);
            if (!textArr) continue;

            for (let j = 0; j < textArr.length; j++) {
                globalThis.docText[file.name] += textArr[j] + " ";
            }
            globalThis.docText[file.name] += "\n";
        }
    }

    await zipReader.close();

}

function readTextFile(file) {
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
    let fileStr = await readTextFile(file);
    // Delete any embedded Javascript code
    fileStr = fileStr.replaceAll(/\<script[^>]*?\>[\s\S]*?\<\/script\>/gi, "");
    const parser = new DOMParser();
    const htmlDoc = parser.parseFromString(fileStr, "text/html");
    // The text content often has an excessive number of newlines
    const htmlStr = htmlDoc.body.textContent?.replaceAll(/\n{2,}/g, "\n");

    globalThis.docText[file.name] =  htmlStr;

}

const readTxt = async (file) => {
    const fileStr = await readTextFile(file);

    globalThis.docText[file.name] =  fileStr;
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


async function readFiles(files) {
    const elemArr = [];
    for (let i = 0; i < files.length; i++) {
        const li = document.createElement("li");
        li.innerHTML = files[i].name;
        li.setAttribute("class", "list-group-item");
        elemArr.push(li);
    }

    for (let i = 0; i < files.length; i++) {

        const file = files[i];
        globalThis.docText[file.name] = "";

        const ext = file.name.match(/\.(\w{1,5})$/)?.[1]?.toLowerCase();

        if (!read[ext]) {
            fileListSkippedElem?.appendChild(elemArr[i]);
            fileCountSkippedElem.textContent = String(parseInt(fileCountSkippedElem.textContent) + 1);
        } else {
            try {
                // TODO: This should eventually use promises + workers for better performance, but this will require edits.
                // Notably, as the same mupdf worker is reused, if run in asyc the PDF may be replaced before readPdf is finished reading it.
                // The other functions are not set up to run in workers.
                await read[ext](file);
                // Remove excessive newline characters to improve readability
                globalThis.docText[file.name] = globalThis.docText[file.name].replaceAll(/(\n\s*){3,}/g, "\n\n");
                fileListSuccessElem?.appendChild(elemArr[i]);
                fileCountSuccessElem.textContent = String(parseInt(fileCountSuccessElem.textContent) + 1);
            } catch (error) {
                fileListFailedElem?.appendChild(elemArr[i]);
                fileCountFailedElem.textContent = String(parseInt(fileCountFailedElem.textContent) + 1);
            }
        }

    }
}

/**
 * @param {string} fileName - Name of file
 * @param {number} index - Index of the start of the snippet
 * @param {string} snippet - Short text with matching term and some context to display in the search results
 */
function searchMatch(fileName, index, snippet) {
    /** @type {string} */ 
    this.fileName = fileName;
    /** @type {number} */ 
    this.index = index;
    /** @type {string} */ 
    this.snippet = snippet;
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
        // Matches are omitted if they are already included in the context for another match
        if (i == 0 || indices[i] > (lastIndexIncluded + contextLength - 10)) {

            const matchText = text.slice(Math.max(0, indices[i] - contextLength), Math.min(text.length, indices[i] + contextLength));
            const replaceRegex = new RegExp("(" + search + ")", "ig");

            matches.push(new searchMatch(fileName, indices[i], matchText.replaceAll(replaceRegex, "<b>$1</b>")));
            lastIndexIncluded = indices[i];
        }
    }

    return matches;


}

globalThis.matches = [];

async function searchDocs(search) {
    matchListElem.innerHTML = "";
    globalThis.matches = [];
    for (const [key, value] of Object.entries(globalThis.docText)) {
        globalThis.matches.push(...searchText(key, search));
    }

    for (let j = 0; j < globalThis.matches.length; j++) {
        const entry = document.createElement('a');
        entry.setAttribute("class", "list-group-item list-group-item-action flex-column align-items-start");
        entry.setAttribute("data-bs-toggle", "list");

        entry.addEventListener("click", () => viewResult(globalThis.matches[j]));


        entry.innerHTML = `<p class="mb-1">${globalThis.matches[j].snippet}</p>
                <small>${globalThis.matches[j].fileName}</small>`;

        matchListElem.appendChild(entry);
    }


    if (matchListElem.innerHTML == "") {
        const entry = document.createElement('a');
        entry.setAttribute("class", "list-group-item list-group-item-action flex-column align-items-start");

        entry.innerHTML = `<p class="mb-1">[No Results]</p>`;

        matchListElem.appendChild(entry);
    }


}

globalThis.initViewer = false;

/**
* @param {searchMatch} match - Name of file
*/
async function viewResult(match) {
    if (!globalThis.initViewer) {
        document.getElementById("viewerCol").style.width = "50%";
        // The location of the highlighted text is not detected correctly without waiting for the animation
        await new Promise((r) => setTimeout(r, 100));
        globalThis.initViewer = true;

    }

    // The snippet is wrapped in <span> tags as this allows for detection of the height of the match in the viewer
    // and allows the scoll position to be set.
    let innerHTML = globalThis.docText[match.fileName].slice(0,match.index) + "<span id='snippetText' style='background-color:yellow'>" + globalThis.docText[match.fileName].slice(match.index,match.index+200) + "</span>" + globalThis.docText[match.fileName].slice(match.index+200);

    innerHTML = innerHTML.replaceAll(/\n/g, "<br/>");

    document.getElementById("viewerCard").innerHTML = "<span>" + innerHTML + "</span>";

    // Position the match ~1/3 of the way down the viewer
    document.getElementById("viewerCard").scrollTop = document.getElementById("snippetText").offsetTop - document.getElementById("viewerCard").offsetHeight / 3;


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