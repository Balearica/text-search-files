
const warningHTML = `Large number of files detected—performance may degrade.`;

export const config = {
    // Non-default documentation URL.
    docsURL: null,
    // Threshold for throwing warning that user is loading too many files
    fileNWarning: null,
    // Warning to show when user uploads more than `fileNWarning` files
    warningHTML: warningHTML,
}

// Size limits for files to be processed (in bytes).  Files over the limit will be skipped.
// Limits are necessary for file types such as .xlsx, which can contain human-readable content
// (which should be processed) or be used to store gigabytes of data (which should be skipped).
export const sizeLimits = {
    "xlsx": 250000
}