// Worker for reading .zip files (including Microsoft Office files such as .docx, .xlsx, and .pptx)
importScripts('msg.reader.js');
importScripts('DataStream.js');

// importScripts('zip-no-worker-inflate.min.js');
importScripts('zip-full.min.js');

zip.configure({
    useWebWorkers : false
  });
  
addEventListener('message', async e => {
    const func = e.data[0];
    const file = e.data[1];
    const id = e.data[2];

    try {
        const res = await read[func](file);
        postMessage({"status": 0, "data": res, "id": id });
    } catch (error) {
        postMessage({"status": 1, "data": "", "id": id });
        console.error(error);
    }
    
});

const readMsg = async (file) => {
    const startTime = Date.now();

    const msgReader = new MSGReader(await file.arrayBuffer());
    const fileData = msgReader.getFileData();
    const text = fileData.body;

    const attachmentFiles = [];
    for (let i = 0; i < fileData.attachments.length; i++) {
        const attachmentObj = msgReader.getAttachment(i);
        const attachmentFile = new File([attachmentObj.content], attachmentObj.fileName, { type: attachmentObj.mimeType ? attachmentObj.mimeType : "application/octet-stream" });
        attachmentFiles.push(attachmentFile);
    }

    // Set `globalThis.debugMode = true` in the console to print the runtimes for each file
    const runtime = Date.now() - startTime;
    if (globalThis.debugMode) console.log(`${file?.name}: ${runtime} ms`);

    return {text: text, attachmentFiles: attachmentFiles};
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

// TODO: Write a version of readHTML that runs in a worker.
// This version does not because DOMParser does not exist in workers.
// const readHtml = async (file) => {
//     let fileStr = await readTxt(file);
//     // Delete any embedded Javascript code
//     fileStr = fileStr.replaceAll(/\<script[^>]*?\>[\s\S]*?\<\/script\>/gi, "");
//     const parser = new DOMParser();
//     const htmlDoc = parser.parseFromString(fileStr, "text/html");
//     // The text content often has an excessive number of newlines
//     const text = htmlDoc.body.textContent?.replaceAll(/\n{2,}/g, "\n");

//     return text;
// }

const readDocx = async (file) => {
    const zipFileReader = new zip.BlobReader(file);
    const zipReader = new zip.ZipReader(zipFileReader);
    const entries = await zipReader.getEntries();
    let text = "";

    for (let i = 0; i < entries.length; i++) {
        if (['word/document.xml', 'word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml'].includes(entries[i].filename)) {
            const xmlStr = await entries[i].getData(new zip.TextWriter());

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
    const zipFileReader = new zip.BlobReader(file);
    const zipReader = new zip.ZipReader(zipFileReader);
    const entries = await zipReader.getEntries();
    let text = "";

    for (let i = 0; i < entries.length; i++) {
        if (['xl/workbook.xml', 'xl/sharedStrings.xml'].includes(entries[i].filename) || /xl\/worksheets\/[^\/]+.xml/.test(entries[i].filename)) {
            const xmlStr = await entries[i].getData(new zip.TextWriter());
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
    const zipFileReader = new zip.BlobReader(file);
    const zipReader = new zip.ZipReader(zipFileReader);
    const entries = await zipReader.getEntries();
    let text = "";

    for (let i = 0; i < entries.length; i++) {
        if (/ppt\/slides\/[^\/]+.xml/.test(entries[i].filename) || /ppt\/notesSlides\/[^\/]+.xml/.test(entries[i].filename) || /ppt\/comments\/[^\/]+.xml/.test(entries[i].filename)) {
            const xmlStr = await entries[i].getData(new zip.TextWriter());
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

const read = {
    "readMsg": readMsg,
    "readTxt": readTxt,
    // "readHtml": readHtml,
    "readXlsx": readXlsx,
    "readDocx": readDocx,
    "readPptx": readPptx
}