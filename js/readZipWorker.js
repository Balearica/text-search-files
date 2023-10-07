// Worker for reading .zip files (including Microsoft Office files such as .docx, .xlsx, and .pptx)
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
    "readXlsx": readXlsx,
    "readDocx": readDocx,
    "readPptx": readPptx
}