import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
import { search } from "gd-sprest-bs/build/icons/svgs/search";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import { ExportCSV, GetIcon, IScript, ISiteInfo, Sites } from "../common";

// Row Information
interface IRowInfo extends ISiteInfo {
    Account: string;
    EMail: string;
    Id: number;
    IsSiteAdmin: boolean;
    Title: string;
}

// CSV Export Fields
const CSVExportFields = [
    "WebTitle", "WebUrl", "WebId",
    "Account", "EMail", "Id", "IsSiteAdmin", "Title"
];

// Script Constants
const ScriptDescription = "Scan site(s) for specified a site user.";
const ScriptFileName = "user_search_info.csv";
const ScriptName = "User Search";

/**
 * User Search
 * Displays a dialog to get the user information from sites.
 */
class UserSearch {
    private _rows: IRowInfo[] = null;

    // Constructor
    constructor() {
        // Render the modal dialog
        this.render();
    }

    // Analyzes the site
    private analyzeSites(sites: ISiteInfo[], userInfo: string | Types.SP.User) {
        // Return a promise
        return new Promise(resolve => {
            let counter = 0;

            // Show a loading dialog
            LoadingDialog.setHeader("Getting the User Information");
            LoadingDialog.setBody('Analyzing the Sites');
            LoadingDialog.show();

            // Parse the users
            Helper.Executor(sites, site => {
                // Update the loading dialog
                LoadingDialog.setBody(`Analyzing Site ${++counter} of ${sites.length}`);

                // Get the user information
                return this.getUserInfo(site, userInfo);
            }).then(() => {
                // Hide the loading dialog
                LoadingDialog.hide();

                // Resolve the request
                resolve(null);
            });
        });
    }

    // Gets the user information
    private getUserInfo(site: ISiteInfo, user: string | Types.SP.User): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // See if we are searching by a string
            if (typeof (user) === "string") {
                // Get the user information list
                Web(site.WebUrl).Lists("User Information List").Items().query({
                    Filter: `substringof('${user}', Name) or substringof('${user}', Title) or substringof('${user}', UserName)`,
                    Select: ["Id", "Name", "EMail", "IsSiteAdmin", "Title"],
                    GetAllItems: true,
                    Top: 5000
                }).execute(
                    items => {
                        // Parse the items
                        for (let i = 0; i < items.results.length; i++) {
                            let item = items.results[i];

                            // Add the row
                            this._rows.push({
                                Account: item["Name"],
                                EMail: item["EMail"],
                                Id: item.Id,
                                IsSiteAdmin: item["IsSiteAdmin"],
                                Title: item["Title"],
                                WebId: site.WebId,
                                WebTitle: site.WebTitle,
                                WebUrl: site.WebUrl
                            });
                        }

                        // Resolve the request
                        resolve();
                    },
                    () => {
                        // Resolve the request
                        resolve();
                    }
                );
            } else {
                // Get the users
                Web(site.WebUrl).Lists("User Information List").Items().query({
                    Filter: `EMail eq '${user.Email}' or UserName eq '${user.UserPrincipalName}'`,
                    GetAllItems: true,
                    Select: ["Id", "Name", "EMail", "IsSiteAdmin", "Title"],
                    Top: 1
                }).execute(
                    items => {
                        // See if the user is in this site
                        let item = items.results[0];
                        if (item) {
                            // Add the row
                            this._rows.push({
                                Account: item["Name"],
                                EMail: item["EMail"],
                                Id: item.Id,
                                IsSiteAdmin: item["IsSiteAdmin"],
                                Title: item["Title"],
                                WebId: site.WebId,
                                WebTitle: site.WebTitle,
                                WebUrl: site.WebUrl
                            });
                        }

                        // Resolve the request
                        resolve();
                    },
                    () => {
                        // Resolve the request
                        resolve();
                    }
                );
            }
        });
    }

    // Removes a user from a group
    private removeUser(row: IRowInfo) {
        // Display a loading dialog
        LoadingDialog.setHeader("Removing Site User");
        LoadingDialog.setBody(`Removing the site user '${row.Title}' from all groups. This will close after the request completes.`);
        LoadingDialog.show();

        // Get the web context
        ContextInfo.getWeb(row.WebUrl).execute(context => {
            // Remove the user from the site
            Web(row.WebUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).SiteUsers().removeById(row.Id).execute(
                // Success
                () => {
                    // Close the dialog
                    LoadingDialog.hide();
                },

                // Error
                () => {
                    // Close the dialog
                    LoadingDialog.hide();

                    // TODO
                }
            )
        });
    }

    // Renders the dialog
    private render() {
        // Set the type
        Modal.setType(Components.ModalTypes.Large);

        // Prevent auto close
        Modal.setAutoClose(false);

        // Set the header
        Modal.setHeader(ScriptName);

        // Render the form
        let form = Components.Form({
            controls: [
                {
                    label: "User or Group Search by Text",
                    name: "UserName",
                    className: "mb-3",
                    description: "Type a user or group display or login name for the search",
                    errorMessage: "Please enter the user or group information",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    onValidate: (ctrl, results) => {
                        // See if user has been entered
                        if (!results.isValid) {
                            // See if the people picker has a value
                            let selectedUser = form.getControl("PeoplePicker").getValue();
                            if (selectedUser.length > 0) {
                                // Set the results
                                results.isValid = true;
                            }
                        }

                        // Return the results
                        return results;
                    }
                },
                {
                    label: "People Search by Lookup",
                    name: "PeoplePicker",
                    className: "mb-3",
                    description: "Enter a minimum of 3 characters to search for a user",
                    errorMessage: "No user was selected...",
                    allowGroups: false,
                    type: Components.FormControlTypes.PeoplePicker,
                    required: true,
                    onValidate: (ctrl, results) => {
                        // See if user has been entered
                        if (!results.isValid) {
                            // See if the people picker has a value
                            let selectedUser = form.getControl("UserName").getValue();
                            if (selectedUser) {
                                // Set the results
                                results.isValid = true;
                            }
                        }

                        // Return the results
                        return results;
                    }
                } as Components.IFormControlPropsPeoplePicker
            ]
        });

        // Render the body
        Modal.setBody(form.el);

        // Render the footer
        Modal.setFooter(Components.TooltipGroup({
            tooltips: [
                {
                    content: "Search for Users",
                    btnProps: {
                        className: "pe-2 py-1",
                        iconClassName: "mx-1",
                        iconType: search,
                        iconSize: 24,
                        text: "Search",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Clear the data
                            this._rows = [];

                            // Ensure the form is valid
                            if (form.isValid()) {
                                // Search for the sites
                                Sites.getSites().then(sites => {
                                    let formValues = form.getValues();
                                    let userName: string = formValues["UserName"].trim();
                                    let user: Types.SP.User = formValues["PeoplePicker"][0];

                                    // Analyze this sites
                                    this.analyzeSites(sites, userName || user).then(() => {
                                        // Render the summary
                                        this.renderSummary();
                                    });
                                });
                            }
                        }
                    }
                },
                {
                    content: "Close Window",
                    btnProps: {
                        className: "pe-2 py-1",
                        iconClassName: "mx-1",
                        iconType: xSquare,
                        iconSize: 24,
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Close the modal
                            Modal.hide();
                        }
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

        // Prevent auto close
        Modal.setAutoClose(false);

        // Show the modal dialog
        Modal.setHeader(ScriptName);

        // Render the table
        let elTable = document.createElement("div");
        Modal.setBody(elTable);
        new DataTable({
            el: elTable,
            rows: this._rows,
            onRendering: dtProps => {
                dtProps.columnDefs = [
                    {
                        "targets": 6,
                        "orderable": false,
                        "searchable": false
                    }
                ];

                // Order by the 2nd & 3rd column by default; ascending
                dtProps.order = [[1, "asc"], [2, "asc"]];

                // Return the properties
                return dtProps;
            },
            columns: [
                {
                    name: "WebTitle",
                    title: "Web Title"
                },
                {
                    name: "WebUrl",
                    title: "Web Url"
                },
                {
                    name: "Account",
                    title: "Name"
                },
                {
                    name: "Title",
                    title: "Title"
                },
                {
                    name: "IsSiteAdmin",
                    title: "Is Site Admin"
                },
                {
                    name: "EMail",
                    title: "User Email"
                },
                {
                    className: "text-end",
                    name: "",
                    title: "",
                    onRenderCell: (el, col, row: IRowInfo) => {
                        let btnDelete: Components.IButton = null;

                        // Render the delete button
                        Components.Tooltip({
                            el,
                            content: "Remove User",
                            btnProps: {
                                assignTo: btn => { btnDelete = btn; },
                                className: "pe-2 py-1",
                                iconType: GetIcon(24, 24, "PersonDelete", "mx-1"),
                                text: "Remove",
                                type: Components.ButtonTypes.OutlineDanger,
                                onClick: () => {
                                    // Confirm the deletion of the group
                                    if (confirm("Are you sure you want to remove the user from this site?")) {
                                        // Disable this button
                                        btnDelete.disable();

                                        // Remove the user
                                        this.removeUser(row);
                                    }
                                }
                            }
                        });
                    }
                }
            ]
        });

        // Set the footer
        Modal.setFooter(Components.TooltipGroup({
            tooltips: [
                {
                    content: "Export to a CSV file",
                    btnProps: {
                        className: "pe-2 py-1",
                        iconType: GetIcon(24, 24, "ExcelDocument", "icon-svg mx-1"),
                        text: "Export",
                        type: Components.ButtonTypes.OutlineSuccess,
                        onClick: () => {
                            // Export the CSV
                            new ExportCSV(ScriptFileName, CSVExportFields, this._rows);
                        }
                    }
                },
                {
                    content: "Close Window",
                    btnProps: {
                        className: "pe-2 py-1",
                        iconClassName: "mx-1",
                        iconType: xSquare,
                        iconSize: 24,
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Close the modal
                            Modal.hide();
                        }
                    }
                }
            ]
        }).el);

        // Show the modal
        Modal.show();
    }
}

// Script Information
export const UserSearchModal: IScript = {
    init: UserSearch,
    name: ScriptName,
    description: ScriptDescription
};