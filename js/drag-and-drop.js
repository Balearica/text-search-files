// Code for handling both files and directories uploaded through drag-and-drop interface
// Taken from: https://stackoverflow.com/questions/3590058/does-html5-allow-drag-drop-upload-of-folders-or-a-folder-tree/53058574#53058574

// Drop handler function to get all files
export async function getAllFileEntries(dataTransferItemList) {
    let fileEntries = [];
    // Use BFS to traverse entire directory/file structure
    let queue = [];
    // Unfortunately dataTransferItemList is not iterable i.e. no forEach
    for (let i = 0; i < dataTransferItemList.length; i++) {
      // Note webkitGetAsEntry a non-standard feature and may change
      // Usage is necessary for handling directories
      queue.push(dataTransferItemList[i].webkitGetAsEntry());
    }
    while (queue.length > 0) {
      let entry = queue.shift();
      if (entry.isFile) {
        fileEntries.push(entry);
      } else if (entry.isDirectory) {
        queue.push(...await readAllDirectoryEntries(entry.createReader()));
      }
    }
    return fileEntries;
  }
  
  // Get all the entries (files or sub-directories) in a directory 
  // by calling readEntries until it returns empty array
  async function readAllDirectoryEntries(directoryReader) {
    let entries = [];
    let readEntries = await readEntriesPromise(directoryReader);
    while (readEntries.length > 0) {
      entries.push(...readEntries);
      readEntries = await readEntriesPromise(directoryReader);
    }
    return entries;
  }
  
  // Wrap readEntries in a promise to make working with readEntries easier
  // readEntries will return only some of the entries in a directory
  // e.g. Chrome returns at most 100 entries at a time
  async function readEntriesPromise(directoryReader) {
    try {
      return await new Promise((resolve, reject) => {
        directoryReader.readEntries(resolve, reject);
      });
    } catch (err) {
      console.log(err);
    }
  }
