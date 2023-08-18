import { ContextInfo } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { Dashboard } from "./dashboard";
import Strings, { setContext } from "./strings";
import "./styles.scss";

// Properties
interface IProps {
    el: HTMLElement;
    context?: any;
    envType?: number;
    fractionDigits?: number;
    searchFileTypes?: string;
    searchMonths?: number;
    searchTerms?: string;
    timeFormat?: string;
    sourceUrl?: string;
}

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    Dashboard: null,
    description: Strings.ProjectDescription,
    render: (props: IProps) => {
        // See if the page context exists
        if (props.context) {
            // Set the context
            setContext(props.context, props.envType, props.sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(props.sourceUrl || ContextInfo.webServerRelativeUrl);

            // See if SPFx string values are set
            props.fractionDigits ? Strings.FractionDigits = props.fractionDigits : null;
            props.searchFileTypes ? Strings.SearchFileTypes = props.searchFileTypes : null;
            props.searchMonths ? Strings.SearchMonths = props.searchMonths : null;
            props.searchTerms ? Strings.SearchTerms = props.searchTerms : null;
            props.timeFormat ? Strings.TimeFormat = props.timeFormat : null;
        }

        // Create the application
        GlobalVariable.Dashboard = new Dashboard(props.el);
    },
    version: Strings.Version
};

// Make is available in the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Render the application
    GlobalVariable.render({ el: elApp });
}