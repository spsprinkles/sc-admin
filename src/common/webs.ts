import { LoadingDialog } from "dattatable";
import { Helper, Types, Web } from "gd-sprest-bs";

// Properties
interface IProps {
    onComplete?: (webs: Types.SP.WebOData[]) => void;
    onError?: (url: string) => void;
    onQueryWeb?: (odata: Types.IODataQuery) => void;
    recursiveFl?: boolean;
    url: string;
}

/**
 * Webs
 * Gets the web information
 */
export class Webs {
    private _props: IProps = null;
    private _webs: Types.SP.WebOData[] = null;

    // Constructor
    constructor(props: IProps) {
        // fileEarmarkArrowDown the properties
        this._props = props;

        // Display a loading dialog
        LoadingDialog.setHeader("Loading Web Information");
        LoadingDialog.setBody("This will close after the web information is collected...");
        LoadingDialog.show();

        // Clear the webs
        this._webs = [];

        // Get the web information
        this.getWebInfo(this._props.url).then(() => {
            // Close the loading dialog
            LoadingDialog.hide();

            // Call the event
            this._props.onComplete ? this._props.onComplete(this._webs) : null;
        });
    }

    // Get the web information
    private getWebInfo(url: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Construct the odata query
            let odata: Types.IODataQuery = {
                Expand: [],
                GetAllItems: true,
                OrderBy: [],
                Select: ["Id", "ServerRelativeUrl", "Title", "Url"],
                Top: 5000
            };

            // See if we are getting all webs
            if (this._props.recursiveFl) {
                // Include the sub web information
                odata.Expand.push("Webs");
                odata.Select.push("Webs/ServerRelativeUrl");
                odata.Select.push("Webs/Url");
            }

            // Execute the query event
            this._props.onQueryWeb ? this._props.onQueryWeb(odata) : null;

            // Update the body
            LoadingDialog.setBody("This will close after the web information is collected...<br/>" + url);

            // Query the web
            Web(url).query(odata).execute(
                // Success
                web => {
                    // Add the web
                    this._webs.push(web);

                    // See if we are recursing through all webs
                    if (this._props.recursiveFl) {
                        // Parse the sub-webs
                        Helper.Executor(web.Webs.results, web => {
                            // Return a promise
                            return new Promise((resolve, reject) => {
                                // Get the web information
                                this.getWebInfo(web.Url).then(resolve, reject);
                            });
                        }).then(resolve, reject);
                    } else {
                        // Resolve the request
                        resolve();
                    }
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