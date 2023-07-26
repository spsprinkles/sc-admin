import { ContextInfo } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { register } from "./jslink";
import { Dashboard } from "./dashboard";
import { Menu } from "./menu";
import Strings, { setContext } from "./strings";
import "./styles.scss";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    Dashboard: null,
    description: Strings.ProjectDescription,
    Menu,
    register,
    render: (el, context?, timeFormat?: string, sourceUrl?: string) => {
        // See if the page context exists
        if (context) {
            // Set the context
            setContext(context, sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(sourceUrl || ContextInfo.webServerRelativeUrl);

            // See if the timeFormat is set
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
} else {
    // Register the JSLink
    register();
}