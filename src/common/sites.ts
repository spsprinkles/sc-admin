import { LoadingDialog } from "dattatable";
import { Types, Search, Site } from "gd-sprest-bs";

// Properties
interface IProps {
    onComplete?: (webs: Types.SP.SiteOData[]) => void;
    onError?: (url: string) => void;
    onQuerySite?: (odata: Types.IODataQuery) => void;
    url: string;
}

// Site Information
export interface ISiteInfo {
    WebId?: string;
    WebTitle: string;
    WebUrl: string;
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
        LoadingDialog.setHeader("Loading Site Information");
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
                Select: ["Id", "Title", "ServerRelativeUrl", "Url"],
                Top: 5000
            };

            // Execute the query event
            this._props.onQuerySite ? this._props.onQuerySite(odata) : null;

            // Update the body
            LoadingDialog.setBody("This will close after the web information is collected...<br/>" + url);

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

    // Static method to get all site collections the user has access to
    static getSites(): PromiseLike<Array<ISiteInfo>> {
        // Return a promise
        return new Promise((resolve) => {
            let sites: ISiteInfo[] = [];

            // Show a loading dialog
            LoadingDialog.setHeader("Searching Site Collections");
            LoadingDialog.setBody('Getting the sites you have access to...');
            LoadingDialog.show();

            // Search for the sites
            this.searchForSites(sites).then(sites => {
                // Hide the dialog
                LoadingDialog.hide();

                // Resolve the request
                resolve(sites);
            });
        });
    }

    // Static method for executing the search api request
    private static searchForSites(sites: ISiteInfo[]): PromiseLike<Array<ISiteInfo>> {
        // Return a promise
        return new Promise(resolve => {
            // Get the associated sites
            Search.postQuery({
                query: {
                    Querytext: `contentclass=sts_site`,
                    TrimDuplicates: true,
                    SelectProperties: {
                        results: [
                            "Title", "SPSiteUrl", "WebId"
                        ]
                    }
                }
            }).then(
                results => {
                    // Parse the results
                    for (let i = 0; i < results.PrimaryQueryResult.RelevantResults.RowCount; i++) {
                        let row = results.PrimaryQueryResult.RelevantResults.Table.Rows.results[i];
                        let siteInfo: ISiteInfo = {} as any;

                        // Parse the cells
                        for (let j = 0; j < row.Cells.results.length; j++) {
                            let cell = row.Cells.results[j];

                            // Set the values
                            switch (cell.Key) {
                                case "SPSiteUrl":
                                    siteInfo.WebUrl = cell.Value;
                                    break;
                                case "Title":
                                    siteInfo.WebTitle = cell.Value;
                                    break;
                                case "WebId":
                                    siteInfo.WebId = cell.Value;
                                    break;
                            }
                        }

                        // Add the site information
                        sites.push(siteInfo);
                    }

                    // Resolve the request
                    resolve(sites);
                },
                () => {
                    // Error executing the search
                    resolve(sites);
                }
            );
        });
    }
}