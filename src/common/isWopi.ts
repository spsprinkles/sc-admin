// Determines if the document can be viewed in office online servers
export function isWopi(fileName: string) {
    // Determine if the file type is supported by WOPI
    let extension: any = fileName.split('.');
    extension = extension[extension.length - 1].toLowerCase();
    switch (extension) {
        // Office Doc Types
        case "csv":
        case "doc":
        case "docx":
        case "dot":
        case "dotx":
        case "pot":
        case "potx":
        case "pps":
        case "ppsx":
        case "ppt":
        case "pptx":
        case "xls":
        case "xlsx":
        case "xlt":
        case "xltx":
            return true;
        // Default
        default: {
            return false;
        }
    }
}
