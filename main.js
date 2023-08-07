

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
globalThis.docTextHighlighted = {};

// https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Math/random
export function getRandomInt(min, max) {
    min = Math.ceil(min);
    max = Math.floor(max);
    return Math.floor(Math.random() * (max - min) + min); //The maximum is exclusive and the minimum is inclusive
  }
  
function getRandomAlphanum(num){
    let outArr = new Array(num);
    for(let i=0;i<num;i++){
      let intI = getRandomInt(1,62);
      if(intI <= 10){
        intI = intI + 47;
      } else if(intI <= 36){
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
    this.snippetStartIndex = Math.min(Math.max(index - contextLength, 0), globalThis.docText[fileName].length - contextLength*2);
    /** @type {number} */ 
    this.snippetEndIndex = Math.max(Math.min(index + contextLength, globalThis.docText[fileName].length), contextLength*2);
    /** @type {string} */ 
    this.search = search;
    /** @type {string} */ 
    this.id = getRandomAlphanum(10);
  }

function getSnippet(match) {
    const replaceRegex = new RegExp("(" + match.search + ")", "ig");
    return globalThis.docText[match.fileName].slice(match.snippetStartIndex, match.snippetEndIndex).replaceAll(replaceRegex, "<b>$1</b>")
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


        entry.innerHTML = `<p class="mb-1">${getSnippet(globalThis.matches[j])}</p>
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

    if (!globalThis.docTextHighlighted[match.fileName]) {
        globalThis.docTextHighlighted[match.fileName] = "";
        const replaceRegex = new RegExp("(" + match.search + ")", "ig");
        // Get all matches for the same file
        const matches = globalThis.matches.filter(x => x.fileName == match.fileName);
        let lastIndex = 0;
        for (let i=0; i<matches.length; i++) {

            // Snippets are wrapped in <span> tags with unique ids so (1) to highlight them and (2) so they can be scrolled to
            globalThis.docTextHighlighted[matches[i].fileName] += globalThis.docText[matches[i].fileName].slice(lastIndex, matches[i].snippetStartIndex);
            globalThis.docTextHighlighted[matches[i].fileName] += "<span id='" + matches[i].id + "' style='background-color:yellow'>" ;
            globalThis.docTextHighlighted[matches[i].fileName] += globalThis.docText[matches[i].fileName].slice(matches[i].snippetStartIndex,matches[i].snippetEndIndex).replaceAll(replaceRegex, "<b>$1</b>") + "</span>";

            lastIndex = matches[i].snippetEndIndex;
        }
        globalThis.docTextHighlighted[match.fileName] += globalThis.docText[match.fileName].slice(lastIndex);

        globalThis.docTextHighlighted[match.fileName] = globalThis.docTextHighlighted[match.fileName].replaceAll(/\n/g, "<br/>");
    }

    document.getElementById("viewerCard").innerHTML = "<span>" + globalThis.docTextHighlighted[match.fileName] + "</span>";

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