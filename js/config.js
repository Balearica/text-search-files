
export const config = {
    // Non-default documentation URL.
    docsURL: null
}

// Size limits for files to be processed (in bytes).  Files over the limit will be skipped.
// Limits are necessary for file types such as .xlsx, which can contain human-readable content
// (which should be processed) or be used to store gigabytes of data (which should be skipped).
export const sizeLimits = {
    "xlsx": 250000
}