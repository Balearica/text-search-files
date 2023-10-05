
const warningHTML = `Large number of files detected—performance may degrade.`;

export const config = {
    // Non-default documentation URL.
    docsURL: null,
    // Threshold for throwing warning that user is loading too many files
    fileNWarning: null,
    // Warning to show when user uploads more than `fileNWarning` files
    warningHTML: warningHTML,
}