import { CanvasForm, DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
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
 * Sites
 * Displays a dialog to get the site information.
 */
class SiteInfo {
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

    // Adds an SCA
    private addSCA(webInfo: IRowInfo) {
        // Set the header
        CanvasForm.clear();
        CanvasForm.setHeader("Add SCA");

        // Set the body
        let form = Components.Form({
            el: CanvasForm.BodyElement,
            controls: [{
                name: "User",
                title: "User:",
                description: "The user to add as a site collection administrator",
                type: Components.FormControlTypes.PeoplePicker,
                allowGroups: false,
                required: true
            } as Components.IFormControlPropsPeoplePicker]
        });

        // Add a button to add the user
        Components.Tooltip({
            el: CanvasForm.BodyElement,
            content: "Click to add the selected user as an admin.",
            btnProps: {
                text: "Add",
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => {
                    // Ensure the form is valid
                    if (form.isValid()) {
                        // Show a loading dialog
                        LoadingDialog.setHeader("Adding User");
                        LoadingDialog.setBody("This will close after the user is added.");
                        LoadingDialog.show();

                        // Get the user
                        let ctrl = form.getControl("User");
                        let userInfo = ctrl.getValue()[0] as Types.SP.User;

                        // Get the context of the target web
                        ContextInfo.getWeb(webInfo.WebUrl).execute(context => {
                            // Ensure they are added to the web user info list
                            Web(webInfo.WebUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).ensureUser(userInfo.LoginName).execute(user => {
                                // Update the user
                                user.update({
                                    IsSiteAdmin: true
                                }).execute(() => {
                                    // Successfully added the user
                                    ctrl.updateValidation(ctrl.el, {
                                        isValid: true,
                                        validMessage: "Successfully added the user as a site collection admin."
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();
                                }, () => {
                                    // Error adding the user
                                    ctrl.updateValidation(ctrl.el, {
                                        isValid: false,
                                        invalidMessage: "Error adding the user as a site collection admin."
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();
                                });
                            }, () => {
                                // Error adding the user
                                ctrl.updateValidation(ctrl.el, {
                                    isValid: false,
                                    invalidMessage: "Error adding the user to the web."
                                });

                                // Hide the loading dialog
                                LoadingDialog.hide();
                            });
                        }, () => {
                            // Error adding the user
                            ctrl.updateValidation(ctrl.el, {
                                isValid: false,
                                invalidMessage: "Error getting the context information of the web."
                            });

                            // Hide the loading dialog
                            LoadingDialog.hide();
                        });
                    }
                }
            }
        });

        // Show the canvas form
        CanvasForm.show();
    }

    // Analyzes the site
    private analyzeSites(webs: Types.SP.WebOData[]) {
        // Return a promise
        return new Promise(resolve => {
            // Show a loading dialog
            LoadingDialog.setHeader("Analyzing the Data");
            LoadingDialog.setBody("Getting additional user information...");
            LoadingDialog.show();

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
                // Hide the loading dialog
                LoadingDialog.hide();

                // Check the next site collection
                resolve(null);
            });
        });
    }

    // Deletes a web
    private deleteWeb(webUrl: string) {
        // Display a loading dialog
        LoadingDialog.setHeader("Deleting Web");
        LoadingDialog.setBody("Deleting the web: " + webUrl);
        LoadingDialog.show();

        // Get the web context
        ContextInfo.getWeb(webUrl).execute(context => {
            // Delete the site group
            Web(webUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).delete().execute(
                // Success
                () => {
                    // TODO - Display the confirmation

                    // Close the dialog
                    LoadingDialog.hide();
                },

                // Error
                () => {
                    // TODO - Display an error

                    // Close the dialog
                    LoadingDialog.hide();
                }
            );
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

    // Remove an SCA
    private removeSCA(webInfo: IRowInfo) {
        // Set the header
        CanvasForm.clear();
        CanvasForm.setHeader("Remove SCA");

        // Show a loading dialog
        LoadingDialog.setHeader("Loading Site Admins");
        LoadingDialog.setBody("This will close after the user information is loaded...");
        LoadingDialog.show();

        // Query the web's site users
        Web(webInfo.WebUrl).SiteUsers().query({
            Filter: "IsSiteAdmin eq true"
        }).execute(users => {
            let items: Components.IDropdownItem[] = [];

            // Parse the users
            for (let i = 0; i < users.results.length; i++) {
                let user = users.results[i];

                // Add the user
                items.push({
                    text: user.Title,
                    data: user,
                    value: user.Id.toString()
                });
            }

            // Set the body
            let form = Components.Form({
                el: CanvasForm.BodyElement,
                controls: [{
                    name: "User",
                    title: "User:",
                    description: "Select a site admin to remove.",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items
                } as Components.IFormControlPropsDropdown]
            });

            // Add a button to add the user
            Components.Tooltip({
                el: CanvasForm.BodyElement,
                content: "Click to remove the selected admin.",
                btnProps: {
                    text: "Remove",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Ensure the form is valid
                        if (form.isValid()) {
                            // Show a loading dialog
                            LoadingDialog.setHeader("Removing User");
                            LoadingDialog.setBody("This will close after the user is removed.");
                            LoadingDialog.show();

                            // Get the user
                            let ctrl = form.getControl("User");
                            let selectedItem = ctrl.dropdown.getValue() as Components.IDropdownItem;
                            let userInfo: Types.SP.User = selectedItem.data;

                            // Get the context of the target web
                            ContextInfo.getWeb(webInfo.WebUrl).execute(context => {
                                // Update the user
                                Web(webInfo.WebUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).SiteUsers(userInfo.Id).update({
                                    IsSiteAdmin: false
                                }).execute(user => {
                                    // Successfully added the user
                                    ctrl.updateValidation(ctrl.el, {
                                        isValid: true,
                                        validMessage: "Successfully removed the user as a site collection admin."
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();
                                }, () => {
                                    // Error adding the user
                                    ctrl.updateValidation(ctrl.el, {
                                        isValid: false,
                                        invalidMessage: "Error removing the user as a site collection admin."
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();
                                });
                            }, () => {
                                // Error adding the user
                                ctrl.updateValidation(ctrl.el, {
                                    isValid: false,
                                    invalidMessage: "Error getting the context information of the web."
                                });

                                // Hide the loading dialog
                                LoadingDialog.hide();
                            });
                        }
                    }
                }
            });

            // Hide the loading dialog
            LoadingDialog.hide();

            // Show the canvas form
            CanvasForm.show();
        });
    }

    // Renders the dialog
    private render() {
        // Set the type
        Modal.setType(Components.ModalTypes.Large);

        // Set the header
        Modal.setHeader("Site Information");

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
                                },
                                {
                                    assignTo: btn => { btnDelete = btn; },
                                    text: "Delete",
                                    type: Components.ButtonTypes.OutlineDanger,
                                    onClick: () => {
                                        // Confirm the deletion of the group
                                        if (confirm("Are you sure you want to delete this web?")) {
                                            // Disable this button
                                            btnDelete.disable();

                                            // Delete the site group
                                            this.deleteWeb(row.WebUrl);
                                        }
                                    }
                                },
                                {
                                    text: "Add SCA",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Show the add form
                                        this.addSCA(row);
                                    }
                                },
                                {
                                    text: "Remove SCA",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Show the remove form
                                        this.removeSCA(row);
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
                        new ExportCSV("site_information.csv", CSVExportFields, this._rows);
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
export const SiteInfoModal: IScript = {
    init: SiteInfo,
    name: "Site Information",
    description: "Scan for site details, admins, & owners."
};