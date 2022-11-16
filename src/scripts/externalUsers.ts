import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { ExportCSV, Webs, IScript } from "../common";

// Row Information
interface IRowInfo {
    Role?: string;
    RoleInfo?: string;
    Name: string;
    Email?: string;
    Group?: string;
    GroupId?: number;
    GroupInfo?: string;
    WebTitle: string;
    WebUrl: string;
}

/** User Information */
interface IUserInfo {
    EMail: string;
    Id: number;
    Name: string;
    Title: string;
}

// CSV Export Fields
const CSVExportFields = [
    "WebTitle", "WebUrl",
    "Name", "Email", "Group", "GroupInfo",
    "Role", "RoleInfo"
];

/**
 * Security Group Information
 * Displays a dialog to get the site information.
 */
class ExternalUsers {
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
    private analyzeSites(rootWeb: Types.SP.WebOData, webs: Types.SP.WebOData[]) {
        // Return a promise
        return new Promise(resolve => {
            // Show a loading dialog
            LoadingDialog.setHeader("Getting User Information");
            LoadingDialog.setBody(rootWeb.ServerRelativeUrl);
            LoadingDialog.show();

            // Get the users
            this.getUsers(rootWeb.ServerRelativeUrl).then(users => {
                let counter = 0;

                // Parse the users
                Helper.Executor(users, user => {
                    // Update the loading dialog
                    LoadingDialog.setBody(`Analyzing User ${++counter} of ${users.length}`);

                    // Get the user information
                    return this.getUserInfo(rootWeb, webs, user);
                }).then(() => {
                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    resolve(null);
                });
            }, resolve);
        });
    }

    // Deletes a group
    private deleteGroup(webUrl: string, groupName: string, groupId: number) {
        // Display a loading dialog
        LoadingDialog.setHeader("Deleting Site Group");
        LoadingDialog.setBody("Deleting the site group '" + groupName + "'. This will close after the request completes.");
        LoadingDialog.show();

        // Get the web context
        ContextInfo.getWeb(webUrl).execute(context => {
            // Delete the site group
            Web(webUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).SiteGroups().removeById(groupId).execute(
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
            )
        });
    }

    // Deletes a user
    private deleteUser(webUrl: string, roleAssignmentId: number, userName: string) {
        // Display a loading dialog
        LoadingDialog.setHeader("Deleting Site User");
        LoadingDialog.setBody("Deleting the site user '" + userName + "'. This will close after the request completes.");
        LoadingDialog.show();

        // Get the web context
        ContextInfo.getWeb(webUrl).execute(context => {
            // Delete the site group
            Web(webUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).RoleAssignments().removeRoleAssignment(roleAssignmentId).execute(
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
            )
        });
    }

    // Get the user information
    private getUserInfo(rootWeb: Types.SP.WebOData, webs: Types.SP.WebOData[], userInfo: IUserInfo) {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the user information
            rootWeb.getUserById(userInfo.Id).query({ Expand: ["Groups"] }).execute(
                // Success
                user => {
                    // Parse the groups
                    Helper.Executor(user.Groups.results, group => {
                        // Return a promise
                        return new Promise((resolve, reject) => {
                            // Parse the roles
                            for (let i = 0; i < rootWeb.RoleAssignments.results.length; i++) {
                                let role: Types.SP.RoleAssignmentOData = rootWeb.RoleAssignments.results[i] as any;

                                // See if the user belongs to this role
                                if (role.Member.LoginName == group.LoginName) {
                                    // Add the user information
                                    this._rows.push({
                                        WebUrl: rootWeb.ServerRelativeUrl,
                                        WebTitle: rootWeb.Title,
                                        Name: userInfo.Title || userInfo.Name,
                                        Email: userInfo.EMail,
                                        Group: group.Title,
                                        GroupId: group.Id,
                                        GroupInfo: group.Description || "",
                                        Role: role.RoleDefinitionBindings.results[0].Name,
                                        RoleInfo: role.RoleDefinitionBindings.results[0].Description || ""
                                    });

                                    // Resolve the request and return
                                    resolve(null);
                                    return;
                                }
                            }

                            // See if this is not a sharing link
                            if (!group.Title.startsWith("SharingLinks.")) {
                                // Add the user information
                                this._rows.push({
                                    WebUrl: rootWeb.ServerRelativeUrl,
                                    WebTitle: rootWeb.Title,
                                    Name: userInfo.Title || userInfo.Name,
                                    Email: userInfo.EMail,
                                    Group: group.Title,
                                    GroupId: group.Id,
                                    GroupInfo: group.Description || "",
                                    Role: "Unknown",
                                    RoleInfo: "Unable to determine the role for this group"
                                });

                                // Resolve the promise
                                resolve(null);
                                return;
                            }

                            // Set the role name
                            let info = group.Title.split('.');
                            let docId = info.length > 2 ? info[1] : "";
                            let roleName = info.length > 3 ? info[2] : "";

                            // See if the doc id exists
                            if (docId) {
                                let doc: Types.SP.File = null;
                                let webTitle = rootWeb.Title;
                                let webUrl = rootWeb.ServerRelativeUrl;

                                // Parse the webs
                                Helper.Executor(webs, web => {
                                    // Return a promise
                                    return new Promise((resolve) => {
                                        // See if the document was found
                                        if (doc == null) {
                                            // Find the document
                                            web.getFileById(docId).execute(
                                                // Success
                                                file => {
                                                    // Set the document and web url containing the document
                                                    doc = file;
                                                    webTitle = web.Title;
                                                    webUrl = web.ServerRelativeUrl;

                                                    // Resolve the promise
                                                    resolve(null);
                                                },
                                                // Error
                                                resolve
                                            )
                                        } else {
                                            // Resolve the promise
                                            resolve(null);
                                        }
                                    });
                                }).then(() => {
                                    let roleInfo = "";

                                    // See if the document was found
                                    if (doc) {
                                        // Set the role information
                                        roleInfo = "Has '" + roleName + "' access to the file <a target='_blank' onclick='javascript:SCAdmin.ShowFile(this);' " +
                                            "href='" + doc.ServerRelativeUrl + "'>" + doc.ServerRelativeUrl + "</a>.";
                                    }

                                    // Add the user information
                                    this._rows.push({
                                        WebUrl: webUrl,
                                        WebTitle: webTitle,
                                        Name: userInfo.Title || userInfo.Name,
                                        Email: userInfo.EMail,
                                        Group: group.Title,
                                        GroupId: group.Id,
                                        GroupInfo: group.Description,
                                        Role: roleName,
                                        RoleInfo: roleInfo
                                    });

                                    // Resolve the promise
                                    resolve(null);
                                });
                            }
                        });
                    }).then(() => {
                        // Resolve the promise
                        resolve(null);
                    });
                },
                // Check the next user
                resolve
            );
        });
    }

    // Gets the external users
    private getUsers(siteUrl: string): PromiseLike<IUserInfo[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let users: IUserInfo[] = [];

            // Get the user information list
            Web(siteUrl).Lists("User Information List").Items().query({
                Filter: "substringof('%23ext%23', Name)",
                Select: ["Id", "Name", "EMail", "Title"],
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
                        Title: item.Title
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
                            let webUrls: string[] = formValues["Urls"].split('\n');

                            // Clear the data
                            this._errors = [];
                            this._rows = [];

                            // Parse the webs
                            Helper.Executor(webUrls, webUrl => {
                                // Return a promise
                                return new Promise((resolve) => {
                                    // Get the webs
                                    new Webs({
                                        url: webUrl,
                                        onQueryWeb: (odata) => {
                                            // Include the site group information
                                            odata.Expand.push("RoleAssignments");
                                            odata.Expand.push("RoleAssignments/Groups");
                                            odata.Expand.push("RoleAssignments/Member");
                                            odata.Expand.push("RoleAssignments/Member/Users");
                                            odata.Expand.push("RoleAssignments/RoleDefinitionBindings");
                                        },
                                        recursiveFl: false,
                                        onComplete: webs => {
                                            // Analyze the site
                                            this.analyzeSites(webs[0], webs).then(resolve);
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
        Modal.setHeader("External Users");

        // Render the table
        let elTable = document.createElement("div");
        new DataTable({
            el: elTable,
            rows: this._rows,
            dtProps: {
                dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                columnDefs: [
                    {
                        "targets": [5],
                        "orderable": false,
                        "searchable": false
                    }
                ],
                // Add some classes to the dataTable elements
                drawCallback: function () {
                    jQuery('.table', this._table).removeClass('no-footer');
                    jQuery('.table', this._table).addClass('tbl-footer');
                    jQuery('.table', this._table).addClass('table-striped');
                    jQuery('.table thead th', this._table).addClass('align-middle');
                    jQuery('.table tbody td', this._table).addClass('align-middle');
                    jQuery('.dataTables_info', this._table).addClass('text-center');
                    jQuery('.dataTables_length', this._table).addClass('pt-2');
                    jQuery('.dataTables_paginate', this._table).addClass('pt-03');
                },
                // Order by the 1st column by default; ascending
                order: [[0, "asc"]]
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
                    name: "Name",
                    title: "User Name"
                },
                {
                    name: "Email",
                    title: "User Email"
                },
                {
                    name: "Group",
                    title: "Group Name"
                },
                {
                    name: "GroupInfo",
                    title: "Group Information"
                },
                {
                    name: "Role",
                    title: "Role"
                },
                {
                    name: "RoleInfo",
                    title: "Role/Share Information"
                },
                {
                    name: "",
                    title: "",
                    onRenderCell: (el, col, row: IRowInfo) => {
                        let btnDelete: Components.IButton = null;

                        // Ensure this is a group
                        if (row.Group) {
                            // Render the buttons
                            Components.ButtonGroup({
                                el,
                                buttons: [
                                    {
                                        text: "View",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // TODO
                                        }
                                    },
                                    {
                                        assignTo: btn => { btnDelete = btn; },
                                        text: "Delete",
                                        type: Components.ButtonTypes.OutlineDanger,
                                        onClick: () => {
                                            // Confirm the deletion of the group
                                            if (confirm("Are you sure you want to delete this site group?")) {
                                                // Disable this button
                                                btnDelete.disable();

                                                // Delete the site group
                                                this.deleteGroup(row.WebUrl, row.Group, row.GroupId);
                                            }
                                        }
                                    }
                                ]
                            });
                        } else {
                            // Render the delete button
                            Components.Button({
                                assignTo: btn => { btnDelete = btn; },
                                el,
                                text: "Delete",
                                type: Components.ButtonTypes.OutlineDanger,
                                onClick: () => {
                                    // Confirm the deletion of the group
                                    if (confirm("Are you sure you want to delete this user?")) {
                                        // Disable this button
                                        btnDelete.disable();

                                        // Delete the site group
                                        // TODO
                                        //this.deleteUser(row.WebUrl, row.RoleAssignmentId, row.SiteGroupName);
                                    }
                                }
                            });
                        }
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
                        new ExportCSV("security_groups.csv", CSVExportFields, this._rows);
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
export const ExternalUsersModal: IScript = {
    init: ExternalUsers,
    name: "External Users Information",
    description: "Scan site(s) for external users information."
};