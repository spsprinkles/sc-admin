import { CanvasForm, DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
import { search } from "gd-sprest-bs/build/icons/svgs/search";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import { ExportCSV, GetIcon, IScript, Webs } from "../common";
import Strings from "../strings";

// Row Information
interface IRowInfo {
    IsRootWeb: boolean;
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
    "IsRootWeb", "WebId", "WebTitle", "WebUrl", "WebDescription", "Owners", "SCAs"
];

// Script Constants
const ScriptDescription = "Scan for site details, admins, & owners.";
const ScriptFileName = "site_information.csv";
const ScriptName = "Site Information";

/**
 * Sites
 * Displays a dialog to get the site information.
 */
class SiteInfo {
    private _errors: string[] = null;
    private _rows: IRowInfo[] = null;
    private _urls: string[] = null;

    // Constructor
    constructor(urls: string[] = Strings.SiteUrls) {
        // Set the urls
        this._urls = urls;

        // Render the modal dialog
        this.render();
    }

    // Analyzes the site
    private analyzeSites(webs: Types.SP.WebOData[]) {
        // Return a promise
        return new Promise(resolve => {
            let counter = 0;

            // Show a loading dialog
            LoadingDialog.setHeader("Analyzing the data");
            LoadingDialog.setBody("Getting additional user information...");
            LoadingDialog.show();

            // Parse the webs
            Helper.Executor(webs, web => {
                // Return a promise
                return new Promise(resolve => {
                    // Update the loading dialog
                    LoadingDialog.setBody(`Getting user information (${++counter} of ${webs.length})`);

                    // Get the owners
                    this.getOwners(web.ServerRelativeUrl).then(
                        // Success
                        owners => {
                            let siteOwners = [];
                            for (let i = 0; i < owners.length; i++) {
                                // Add the owner email
                                siteOwners.push(owners[i].EMail || owners[i].Name || owners[i].UserName || owners[i].Title);
                            }

                            // Get the scas
                            this.getSCAs(web.ServerRelativeUrl, web.ParentWeb).then(
                                // Success
                                admins => {
                                    let siteAdmins = [];
                                    for (let i = 0; i < admins.length; i++) {
                                        // Add the admin email
                                        siteAdmins.push(admins[i].EMail || admins[i].Name || admins[i].UserName || admins[i].Title);
                                    }

                                    // Add a row for this entry
                                    this._rows.push({
                                        IsRootWeb: web.ParentWeb && web.ParentWeb.Id ? false : true,
                                        Owners: siteOwners.join(', '),
                                        SCAs: siteAdmins.join(', '),
                                        WebDescription: web.Description,
                                        WebId: web.Id,
                                        WebTitle: web.Title,
                                        WebUrl: web.Url
                                    });

                                    // Check the next web
                                    resolve(null);
                                },
                                // Error getting the admin information
                                () => {
                                    // Add a row for this entry
                                    this._rows.push({
                                        IsRootWeb: web.ParentWeb && web.ParentWeb.Id ? false : true,
                                        Owners: siteOwners.join(', '),
                                        SCAs: "Unable to get admin information",
                                        WebDescription: web.Description,
                                        WebId: web.Id,
                                        WebTitle: web.Title,
                                        WebUrl: web.Url
                                    });

                                    // Check the next web
                                    resolve(null);
                                }
                            );
                        },
                        // Error getting the owner information
                        () => {
                            // Add a row for this entry
                            this._rows.push({
                                IsRootWeb: web.ParentWeb && web.ParentWeb.Id ? false : true,
                                Owners: "Unable to get owner information",
                                SCAs: "Unable to get admin information",
                                WebDescription: web.Description,
                                WebId: web.Id,
                                WebTitle: web.Title,
                                WebUrl: web.Url
                            });

                            // Check the next web
                            resolve(null);
                        }
                    );
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

    // Manages the Owners
    private manageOwners(webInfo: IRowInfo) {
        // Set the header
        CanvasForm.clear();

        // Prevent auto close
        CanvasForm.setAutoClose(false);

        CanvasForm.setHeader("Site Owners");

        // Show a loading dialog
        LoadingDialog.setHeader("Loading Site Owners");
        LoadingDialog.setBody("This will close after the user information is loaded...");
        LoadingDialog.show();

        // Render the add form
        let formAdd = Components.Form({
            el: CanvasForm.BodyElement,
            controls: [{
                name: "User",
                title: "User:",
                description: "Search for a Site Owner to add",
                type: Components.FormControlTypes.PeoplePicker,
                allowGroups: false,
                required: true
            } as Components.IFormControlPropsPeoplePicker]
        });

        let label = document.createElement("label");
        label.className = "mb-3";
        label.innerHTML = "Manage Site:<br/>" + webInfo.WebUrl;
        CanvasForm.BodyElement.prepend(label);

        // Add a button to add the user
        Components.Tooltip({
            el: CanvasForm.BodyElement,
            content: "Click to add the user as a Site Owner",
            placement: Components.TooltipPlacements.Left,
            btnProps: {
                className: "float-end mb-3 mw-6 pe-2 py-1",
                iconType: GetIcon(24, 24, "PersonAdd", "mx-1"),
                text: "Add User",
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => {
                    // Ensure the form is valid
                    if (formAdd.isValid()) {
                        // Show a loading dialog
                        LoadingDialog.setHeader("Adding User");
                        LoadingDialog.setBody("This will close after the user is added");
                        LoadingDialog.show();

                        // Get the user
                        let ctrl = formAdd.getControl("User");
                        let userInfo = ctrl.getValue()[0] as Types.SP.User;

                        // Get the context of the target web
                        ContextInfo.getWeb(webInfo.WebUrl).execute(context => {
                            // Ensure they are added to the web user info list
                            let web = Web(webInfo.WebUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue });
                            web.ensureUser(userInfo.LoginName).execute(user => {
                                // Add the user to the owners group
                                web.AssociatedOwnerGroup().Users().addUserById(user.Id).execute(() => {
                                    // Successfully added the user
                                    ctrl.updateValidation(ctrl.el, {
                                        isValid: true,
                                        validMessage: "Successfully added the user as a Site Admin"
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();

                                    // Reload the form
                                    this.manageOwners(webInfo);
                                }, () => {
                                    // Error adding the user
                                    ctrl.updateValidation(ctrl.el, {
                                        isValid: false,
                                        invalidMessage: "Error adding the user as a Site Owner"
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();
                                });
                            }, () => {
                                // Error adding the user
                                ctrl.updateValidation(ctrl.el, {
                                    isValid: false,
                                    invalidMessage: "Error adding the user to the web"
                                });

                                // Hide the loading dialog
                                LoadingDialog.hide();
                            });
                        }, () => {
                            // Error adding the user
                            ctrl.updateValidation(ctrl.el, {
                                isValid: false,
                                invalidMessage: "Error getting the context information of the web"
                            });

                            // Hide the loading dialog
                            LoadingDialog.hide();
                        });
                    }
                }
            }
        });

        // Query the web's site owners
        Web(webInfo.WebUrl).AssociatedOwnerGroup().Users().execute(users => {
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
            let formRemove = Components.Form({
                el: CanvasForm.BodyElement,
                controls: [{
                    name: "User",
                    title: "User:",
                    description: "Select a Site Owner to remove",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items
                } as Components.IFormControlPropsDropdown]
            });

            // Add a button to add the user
            Components.Tooltip({
                el: CanvasForm.BodyElement,
                content: "Click to remove the selected Site Owner",
                placement: Components.TooltipPlacements.Left,
                btnProps: {
                    className: "float-end mw-6 pe-2 py-1",
                    iconType: GetIcon(24, 24, "PersonDelete", "mx-1"),
                    text: "Remove",
                    type: Components.ButtonTypes.OutlineDanger,
                    onClick: () => {
                        // Ensure the form is valid
                        if (formRemove.isValid()) {
                            // Show a loading dialog
                            LoadingDialog.setHeader("Removing User");
                            LoadingDialog.setBody("This will close after the user is removed");
                            LoadingDialog.show();

                            // Get the user
                            let ctrl = formRemove.getControl("User");
                            let selectedItem = ctrl.dropdown.getValue() as Components.IDropdownItem;
                            let userInfo: Types.SP.User = selectedItem.data;

                            // Get the context of the target web
                            ContextInfo.getWeb(webInfo.WebUrl).execute(context => {
                                // Update the user
                                Web(webInfo.WebUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).AssociatedOwnerGroup().Users().removeById(userInfo.Id).execute(() => {
                                    // Successfully added the user
                                    ctrl.updateValidation(ctrl.el, {
                                        isValid: true,
                                        validMessage: "Successfully removed the user as a Site Owner"
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();

                                    // Reload the form
                                    this.manageSCAs(webInfo);
                                }, () => {
                                    // Error adding the user
                                    ctrl.updateValidation(ctrl.el, {
                                        isValid: false,
                                        invalidMessage: "Error removing the user as a Site Owner"
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();
                                });
                            }, () => {
                                // Error adding the user
                                ctrl.updateValidation(ctrl.el, {
                                    isValid: false,
                                    invalidMessage: "Error getting the context information of the web"
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

    // Manages the SCAs
    private manageSCAs(webInfo: IRowInfo) {
        // Set the header
        CanvasForm.clear();

        // Prevent auto close
        CanvasForm.setAutoClose(false);

        CanvasForm.setHeader("Site Admins");

        // Show a loading dialog
        LoadingDialog.setHeader("Loading Site Admins");
        LoadingDialog.setBody("This will close after the user information is loaded...");
        LoadingDialog.show();

        // Render the add form
        let formAdd = Components.Form({
            el: CanvasForm.BodyElement,
            controls: [{
                name: "User",
                title: "User:",
                description: "Search for a Site Admin to add",
                type: Components.FormControlTypes.PeoplePicker,
                allowGroups: false,
                required: true
            } as Components.IFormControlPropsPeoplePicker]
        });

        let label = document.createElement("label");
        label.className = "mb-3";
        label.innerHTML = "Manage Site:<br/>" + webInfo.WebUrl;
        CanvasForm.BodyElement.prepend(label);

        // Add a button to add the user
        Components.Tooltip({
            el: CanvasForm.BodyElement,
            content: "Click to add the user as a Site Admin",
            placement: Components.TooltipPlacements.Left,
            btnProps: {
                className: "float-end mb-3 mw-6 pe-2 py-1",
                iconType: GetIcon(24, 24, "PersonAdd", "mx-1"),
                text: "Add User",
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => {
                    // Ensure the form is valid
                    if (formAdd.isValid()) {
                        // Show a loading dialog
                        LoadingDialog.setHeader("Adding User");
                        LoadingDialog.setBody("This will close after the user is added");
                        LoadingDialog.show();

                        // Get the user
                        let ctrl = formAdd.getControl("User");
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
                                        validMessage: "Successfully added the user as a Site Admin"
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();

                                    // Reload the form
                                    this.manageSCAs(webInfo);
                                }, () => {
                                    // Error adding the user
                                    ctrl.updateValidation(ctrl.el, {
                                        isValid: false,
                                        invalidMessage: "Error adding the user as a Site Admin"
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();
                                });
                            }, () => {
                                // Error adding the user
                                ctrl.updateValidation(ctrl.el, {
                                    isValid: false,
                                    invalidMessage: "Error adding the user to the web"
                                });

                                // Hide the loading dialog
                                LoadingDialog.hide();
                            });
                        }, () => {
                            // Error adding the user
                            ctrl.updateValidation(ctrl.el, {
                                isValid: false,
                                invalidMessage: "Error getting the context information of the web"
                            });

                            // Hide the loading dialog
                            LoadingDialog.hide();
                        });
                    }
                }
            }
        });

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
            let formRemove = Components.Form({
                el: CanvasForm.BodyElement,
                controls: [{
                    name: "User",
                    title: "User:",
                    description: "Select a Site Admin to remove",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items
                } as Components.IFormControlPropsDropdown]
            });

            // Add a button to add the user
            Components.Tooltip({
                el: CanvasForm.BodyElement,
                content: "Click to remove the selected Site Admin",
                placement: Components.TooltipPlacements.Left,
                btnProps: {
                    className: "float-end mw-6 pe-2 py-1",
                    iconType: GetIcon(24, 24, "PersonDelete", "mx-1"),
                    text: "Remove",
                    type: Components.ButtonTypes.OutlineDanger,
                    onClick: () => {
                        // Ensure the form is valid
                        if (formRemove.isValid()) {
                            // Show a loading dialog
                            LoadingDialog.setHeader("Removing User");
                            LoadingDialog.setBody("This will close after the user is removed");
                            LoadingDialog.show();

                            // Get the user
                            let ctrl = formRemove.getControl("User");
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
                                        validMessage: "Successfully removed the user as a Site Admin"
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();

                                    // Reload the form
                                    this.manageSCAs(webInfo);
                                }, () => {
                                    // Error adding the user
                                    ctrl.updateValidation(ctrl.el, {
                                        isValid: false,
                                        invalidMessage: "Error removing the user as a Site Admin"
                                    });

                                    // Hide the loading dialog
                                    LoadingDialog.hide();
                                });
                            }, () => {
                                // Error adding the user
                                ctrl.updateValidation(ctrl.el, {
                                    isValid: false,
                                    invalidMessage: "Error getting the context information of the web"
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

        // Prevent auto close
        Modal.setAutoClose(false);

        // Set the header
        Modal.setHeader(ScriptName);

        // Render the form
        let form = Components.Form({
            controls: [
                {
                    label: "Site Url(s)",
                    name: "Urls",
                    description: "Enter the relative site url(s) [Ex: /sites/dev]",
                    errorMessage: "Please enter a site url",
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
        Modal.setFooter(Components.TooltipGroup({
            tooltips: [
                {
                    content: "Search for Site Information",
                    btnProps: {
                        className: "pe-2 py-1",
                        iconClassName: "mx-1",
                        iconType: search,
                        iconSize: 24,
                        text: "Search",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let formValues = form.getValues();
                                Strings.SiteUrls = formValues["Urls"].match(/[^\n]+/g);

                                // Clear the data
                                this._errors = [];
                                this._rows = [];

                                // Parse the webs
                                Helper.Executor(Strings.SiteUrls, webUrl => {
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
                        "targets": 5,
                        "orderable": false,
                        "searchable": false
                    }
                ];

                // Order by the 1st column by default; ascending
                dtProps.order = [[1, "asc"]];

                // Return the properties
                return dtProps;
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
                    title: "Description",
                    onRenderCell: (el) => {
                        // Add the data-filter attribute for searching notes properly
                        el.setAttribute("data-filter", el.innerHTML);
                        // Add the data-order attribute for sorting notes properly
                        el.setAttribute("data-order", el.innerHTML);

                        // Declare a span element
                        let span = document.createElement("span");

                        // Return the plain text if less than 50 chars
                        if (el.innerHTML.length < 50) {
                            span.innerHTML = el.innerHTML;
                        } else {
                            // Truncate to the last white space character in the text after 50 chars and add an ellipsis
                            span.innerHTML = el.innerHTML.substring(0, 50).replace(/\s([^\s]*)$/, '') + '&#8230';

                            // Add a tooltip containing the text
                            Components.Tooltip({
                                content: "<small>" + el.innerHTML + "</small>",
                                target: span
                            });
                        }

                        // Clear the element
                        el.innerHTML = "";
                        // Append the span
                        el.appendChild(span);
                    }
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
                        Components.TooltipGroup({
                            el,
                            tooltips: [
                                {
                                    content: "View Site",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconType: GetIcon(24, 24, "LiveSite", "mx-1"),
                                        text: "View",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the security group
                                            window.open(row.WebUrl, "_blank");
                                        }
                                    }
                                },
                                {
                                    content: "Delete Site",
                                    btnProps: {
                                        assignTo: btn => { btnDelete = btn; },
                                        className: "pe-2 py-1",
                                        iconClassName: "mx-1",
                                        iconType: trash,
                                        iconSize: 24,
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
                                    }
                                },
                                {
                                    content: "Manage " + row.IsRootWeb ? "Site Admins" : "Site Owners",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconType: GetIcon(24, 24, "PeopleTeam", "mx-1"),
                                        text: row.IsRootWeb ? "Admins" : "Owners",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the add form
                                            row.IsRootWeb ? this.manageSCAs(row) : this.manageOwners(row);
                                        }
                                    }
                                }
                            ]
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
                        iconType: GetIcon(24, 24, "ExcelDocument", "mx-1"),
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
export const SiteInfoModal: IScript = {
    init: SiteInfo,
    name: ScriptName,
    description: ScriptDescription
};