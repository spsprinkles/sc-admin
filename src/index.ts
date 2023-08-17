import { ContextInfo } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { Dashboard } from "./dashboard";
import Strings, { setContext } from "./strings";
import "./styles.scss";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    Dashboard: null,
    description: Strings.ProjectDescription,
    render: (el, context?, fractionDigits?: number, searchFileTypes?: string, searchMonths?: number, searchTerms?: string, timeFormat?: string, sourceUrl?: string) => {
        // See if the page context exists
        if (context) {
            // Set the context
            setContext(context, sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(sourceUrl || ContextInfo.webServerRelativeUrl);

            // See if SPFx string values are set
            fractionDigits ? Strings.FractionDigits = fractionDigits : null;
            searchFileTypes ? Strings.SearchFileTypes = searchFileTypes : null;
            searchMonths ? Strings.SearchMonths = searchMonths : null;
            searchTerms ? Strings.SearchTerms = searchTerms : null;
            timeFormat ? Strings.TimeFormat = timeFormat : null;
        }

        // Create the application
        GlobalVariable.Dashboard = new Dashboard(el);
    },
    version: Strings.Version
};

// Make is available in the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Render the application
    GlobalVariable.render(elApp);
}