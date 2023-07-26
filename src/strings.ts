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
    DateFormat: "YYYY-MMM-DD",
    GlobalVariable: "SCAdmin",
    ProjectName: "Site Admin Tool",
    ProjectDescription: "A tool to help manage SharePoint sites.",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.0.6"
};
export default Strings;