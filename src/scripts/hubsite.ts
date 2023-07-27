import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, HubSites, Search, Types, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { ExportCSV, Webs, IScript } from "../common";

// Row Information
interface IRowInfo {
    Owners: string;
    SCAs: string;
    WebDescription: string;
    WebId: string;
    WebTitle: string;
    WebUrl: string;
}

/** User Information */
interface IUserInfo {
    EMail: string;
    Id: number;
    Name: string;
    Title: string;
    UserName: string;
}

// CSV Export Fields
const CSVExportFields = [
    "WebId", "WebTitle", "WebUrl", "WebDescription", "Owners", "SCAs"
];

/**
 * Hub Site Info
 * Displays a dialog to get the site information.
 */
class HubSiteInfo {
    private _errors: string[] = null;
    private _rows: IRowInfo[] = null;
    private _urls: string[] = null;

    // Constructor
    constructor(urls: string[] = []) {
        // Set the urls
        this._urls = urls;

        // Render the modal dialog
        this.render();
    }

    // Analyzes the site
    private analyzeSites(webs: Types.SP.WebOData[]) {
        // Return a promise
        return new Promise(resolve => {
            // Parse the webs
            Helper.Executor(webs, web => {
                // Return a promise
                return new Promise(resolve => {
                    // Get the owners
                    this.getOwners(web.ServerRelativeUrl).then(owners => {
                        let siteOwners = [];
                        for (let i = 0; i < owners.length; i++) {
                            // Add the owner email
                            siteOwners.push(owners[i].EMail || owners[i].Name || owners[i].UserName || owners[i].Title);
                        }

                        // Get the scas
                        this.getSCAs(web.ServerRelativeUrl, web.ParentWeb).then(admins => {
                            let siteAdmins = [];
                            for (let i = 0; i < admins.length; i++) {
                                // Add the admin email
                                siteAdmins.push(admins[i].EMail || admins[i].Name || admins[i].UserName || admins[i].Title);
                            }

                            // Add a row for this entry
                            this._rows.push({
                                Owners: siteOwners.join(', '),
                                SCAs: siteAdmins.join(', '),
                                WebDescription: web.Description,
                                WebId: web.Id,
                                WebTitle: web.Title,
                                WebUrl: web.Url
                            });

                            // Check the next web
                            resolve(null);
                        });
                    });
                });
            }).then(() => {
                // Check the next site collection
                resolve(null);
            });
        });
    }

    // Gets the hub sites
    private getHubSites(): PromiseLike<string[]> {
        // Return a promise
        return new Promise(resolve => {
            // Show a loading dialog
            LoadingDialog.setHeader("Getting Hub Sites");
            LoadingDialog.setBody("Getting the hub sites you have access to...");
            LoadingDialog.show();

            // Get the hub sites
            HubSites().execute(sites => {
                let urls = [];

                // Parse the sites
                Helper.Executor(sites.results, site => {
                    // Append the hub site
                    urls.push(site.SiteUrl);

                    // Return a promise
                    return new Promise(resolve => {
                        // Get the associated sites
                        Search().postquery({
                            Querytext: `DepartmentId=${site.ID} contentclass=sts_site -SiteId:${site.ID}`
                        }).execute(results => {
                            // Parse the results
                            for (let i = 0; i < results.postquery.PrimaryQueryResult.RelevantResults.RowCount; i++) {
                                let row = results.postquery.PrimaryQueryResult.RelevantResults.Table.Rows.results[i];

                                // Parse the cells
                                for (let j = 0; j < row.Cells.results.length; j++) {
                                    let cell = row.Cells.results[j];

                                    // See if this is the url
                                    if (cell.Key == "SiteName") {
                                        // Add the url and break from the loop
                                        urls.push(cell.Value);
                                        break;
                                    }
                                }
                            }

                            // Check the next site
                            resolve(null);
                        });
                    });
                }).then(() => {
                    // Hide the dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    resolve(urls);
                });
            }, () => {
                // Hide the dialog
                LoadingDialog.hide();

                // Return nothing
                resolve([]);
            })
        });
    }

    // Gets the site owners
    private getOwners(siteUrl: string): PromiseLike<IUserInfo[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let users: IUserInfo[] = [];

            // Query the user information list for admins
            Web(siteUrl).AssociatedOwnerGroup().Users().query({
                GetAllItems: true,
                Top: 5000
            }).execute(items => {
                // Parse the items
                for (let i = 0; i < items.results.length; i++) {
                    let item = items.results[i];

                    // Add the user
                    users.push({
                        EMail: item.Email,
                        Id: item.Id,
                        Name: item.LoginName,
                        Title: item.Title,
                        UserName: item.UserPrincipalName
                    });
                }

                // Resolve the request
                resolve(users);
            }, reject);
        });
    }

    // Gets the site collection admins
    private getSCAs(siteUrl: string, parentWeb: Types.SP.WebInformation): PromiseLike<IUserInfo[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let users: IUserInfo[] = [];

            // See if this is a root web
            if (parentWeb.Id) {
                // Skip this web
                resolve(users);
                return;
            }

            // Query the user information list for admins
            Web(siteUrl).Lists("User Information List").Items().query({
                Filter: `IsSiteAdmin eq 1`,
                Select: ["Id", "Name", "EMail", "Title", "UserName"],
                GetAllItems: true,
                Top: 5000
            }).execute(items => {
                // Parse the items
                for (let i = 0; i < items.results.length; i++) {
                    let item = items.results[i];

                    // Add the user
                    users.push({
                        EMail: item["EMail"],
                        Id: item.Id,
                        Name: item["Name"],
                        Title: item.Title,
                        UserName: item["UserName"]
                    });
                }

                // Resolve the request
                resolve(users);
            }, reject);
        });
    }

    // Renders the dialog
    private render() {
        // Set the type
        Modal.setType(Components.ModalTypes.Large);

        // Set the header
        Modal.setHeader("Hub Site Information");

        // Render the form
        let form = Components.Form({
            controls: [
                {
                    label: "Site Url(s)",
                    name: "Urls",
                    description: "Enter the relative site url(s). (Ex: /sites/dev)",
                    errorMessage: "Please enter a site url.",
                    type: Components.FormControlTypes.TextArea,
                    required: true,
                    rows: 10,
                    value: this._urls.join('\n')
                } as Components.IFormControlPropsTextField
            ]
        });

        // Render the body
        Modal.setBody(form.el);

        // Render the footer
        Modal.setFooter(Components.ButtonGroup({
            buttons: [
                {
                    text: "Load Hub Sites",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Get the hub sites the user has access to
                        this.getHubSites().then(sites => {
                            // Set the web urls
                            form.getControl("Urls").setValue(sites.join('\n'));
                        });
                    }
                },
                {
                    text: "Analyze",
                    type: Components.ButtonTypes.OutlineSuccess,
                    onClick: () => {
                        // Ensure the form is valid
                        if (form.isValid()) {
                            let formValues = form.getValues();
                            let webUrls: string[] = formValues["Urls"].match(/[^\n]+/g);

                            // Clear the data
                            this._errors = [];
                            this._rows = [];

                            // Parse the webs
                            Helper.Executor(webUrls, webUrl => {
                                // Return a promise
                                return new Promise((resolve) => {
                                    new Webs({
                                        url: webUrl,
                                        onQueryWeb: (odata) => {
                                            // Include the parent web
                                            odata.Expand.push("ParentWeb");

                                            // Include the web description
                                            odata.Select.push("Description");
                                        },
                                        recursiveFl: true,
                                        onComplete: webs => {
                                            // Analyze the site
                                            this.analyzeSites(webs).then(resolve);
                                        },
                                        onError: () => {
                                            // Add the url to the errors list
                                            this._errors.push(webUrl);
                                            resolve(null);
                                        }
                                    })
                                });
                            }).then(() => {
                                // Render the summary
                                this.renderSummary();
                            });
                        }
                    }
                },
                {
                    text: "Cancel",
                    type: Components.ButtonTypes.OutlineDanger,
                    onClick: () => {
                        // Close the modal
                        Modal.hide();
                    }
                }
            ]
        }).el);

        // Show the modal
        Modal.show();
    }

    // Renders the summary dialog
    private renderSummary() {
        // Set the type
        Modal.setType(Components.ModalTypes.Full);

        // Show the modal dialog
        Modal.setHeader("Sites");

        // Render the table
        let elTable = document.createElement("div");
        new DataTable({
            el: elTable,
            rows: this._rows,
            dtProps: {
                dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                columnDefs: [
                    {
                        "targets": 3,
                        "orderable": false,
                        "searchable": false
                    }
                ],
                // Add some classes to the dataTable elements
                createdRow: function (row, data, index) {
                    jQuery('td', row).addClass('align-middle');
                },
                drawCallback: function (settings) {
                    let api = new jQuery.fn.dataTable.Api(settings) as any;
                    let div = api.table().container() as HTMLDivElement;
                    let table = api.table().node() as HTMLTableElement;
                    div.querySelector(".dataTables_info").classList.add("text-center");
                    div.querySelector(".dataTables_length").classList.add("pt-2");
                    div.querySelector(".dataTables_paginate").classList.add("pt-03");
                    table.classList.remove("no-footer");
                    table.classList.add("tbl-footer");
                    table.classList.add("table-striped");
                },
                headerCallback: function (thead, data, start, end, display) {
                    jQuery('th', thead).addClass('align-middle');
                },
                // Order by the 1st column by default; ascending
                order: [[1, "asc"]]
            },
            columns: [
                {
                    name: "WebTitle",
                    title: "Title"
                },
                {
                    name: "WebUrl",
                    title: "Url"
                },
                {
                    name: "WebDescription",
                    title: "Description"
                },
                {
                    name: "SCAs",
                    title: "Site Admins"
                },
                {
                    name: "Owners",
                    title: "Owners"
                },
                {
                    className: "text-end",
                    name: "",
                    title: "",
                    onRenderCell: (el, col, row: IRowInfo) => {
                        let btnDelete: Components.IButton = null;

                        // Render the buttons
                        Components.ButtonGroup({
                            el,
                            buttons: [
                                {
                                    text: "View",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Show the security group
                                        window.open(row.WebUrl, "_blank");
                                    }
                                }
                            ]
                        });
                    }
                }
            ]
        });

        // Set the body
        Modal.setBody(elTable)

        // Set the footer
        Modal.setFooter(Components.ButtonGroup({
            buttons: [
                {
                    text: "Export",
                    type: Components.ButtonTypes.OutlineSuccess,
                    onClick: () => {
                        // Export the CSV
                        new ExportCSV("hub_site_information.csv", CSVExportFields, this._rows);
                    }
                },
                {
                    text: "Cancel",
                    type: Components.ButtonTypes.OutlineDanger,
                    onClick: () => {
                        // Close the modal
                        Modal.hide();
                    }
                }
            ]
        }).el);

        // Show the modal
        Modal.show();
    }
}

// Script Information
export const HubSiteInfoModal: IScript = {
    init: HubSiteInfo,
    name: "Hub Site Information",
    description: "Scan the hub sites for details, admins, & owners."
};