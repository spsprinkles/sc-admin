import { ContextInfo } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the source url
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "sc-admin",
    FractionDigits: 2,
    GlobalVariable: "SCAdmin",
    ProjectName: "Site Admin Tool",
    ProjectDescription: "A tool to help manage SharePoint sites.",
    SearchFileTypes: "csv doc docx dot dotx pdf pot potx pps ppsx ppt pptx txt xls xlsx xlt xltx",
    SearchMonths: 18,
    SearchTerms: "PHI PII",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    TimeFormat: "YYYY-MMM-DD HH:mm:ss",
    Version: "0.0.9"
};
export default Strings;