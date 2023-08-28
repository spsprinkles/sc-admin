import { ContextInfo, SPTypes } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, envType?: number, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the properties
    Strings.IsClassic = envType == SPTypes.EnvironmentType.ClassicSharePoint;
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "sc-admin",
    FractionDigits: 2,
    GlobalVariable: "SCAdmin",
    IsClassic: true,
    ProjectName: "Site Admin Tool",
    ProjectDescription: "A tool to help manage SharePoint sites.",
    SearchFileTypes: "csv doc docx dot dotx pdf pot potx pps ppsx ppt pptx txt xls xlsx xlt xltx",
    SearchMonths: 18,
    SearchTerms: "phi pii ssn",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    TimeFormat: "YYYY-MMM-DD HH:mm:ss",
    Version: "0.1.2"
};
export default Strings;