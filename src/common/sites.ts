import { LoadingDialog } from "dattatable";
import { Types, Site } from "gd-sprest-bs";

// Properties
interface IProps {
    onComplete?: (webs: Types.SP.SiteOData[]) => void;
    onError?: (url: string) => void;
    onQuerySite?: (odata: Types.IODataQuery) => void;
    url: string;
}

/**
 * Sites
 * Gets the site collection information
 */
export class Sites {
    private _props: IProps = null;
    private _sites: Types.SP.SiteOData[] = null;

    // Constructor
    constructor(props: IProps) {
        // Save the properties
        this._props = props;

        // Display a loading dialog
        LoadingDialog.setHeader("Loading Site Collection Information");
        LoadingDialog.setBody("This will close after the site information is collected...");
        LoadingDialog.show();

        // Clear the sites
        this._sites = [];

        // Get the site collection information
        this.getSiteCollectionInfo(this._props.url).then(() => {
            // Close the loading dialog
            LoadingDialog.hide();

            // Call the event
            this._props.onComplete ? this._props.onComplete(this._sites) : null;
        });
    }

    // Get the site collection information
    private getSiteCollectionInfo(url: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Construct the odata query
            let odata: Types.IODataQuery = {
                Expand: ["RootWeb"],
                GetAllItems: true,
                OrderBy: [],
                Select: ["Id", "Title", "Url"],
                Top: 5000
            };

            // Execute the query event
            this._props.onQuerySite ? this._props.onQuerySite(odata) : null;

            // Update the body
            LoadingDialog.setBody("This will close after the webs information is collected.<br/>" + url);

            // Query the web
            Site(url).query(odata).execute(
                // Success
                site => {
                    // Add the web
                    this._sites.push(site);

                    // Resolve the request
                    resolve();
                },

                // Error
                () => {
                    // Call the event
                    this._props.onError ? this._props.onError(this._props.url) : null;

                    // Resolve the request
                    resolve();
                }
            );
        });
    }
}