import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { ExportCSV, Webs, IScript } from "../common";

// Row Information
interface IRowInfo {
    RoleAssignmentId?: number;
    SiteGroupId?: number;
    SiteGroupName: string;
    SiteGroupPermission: string;
    SiteGroupUrl?: string;
    SiteGroupUsers?: string;
    WebTitle: string;
    WebUrl: string;
}

// CSV Export Fields
const CSVExportFields = [
    "RoleAssignmentId", "SiteGroupId", "SiteGroupName",
    "SiteGroupPermission", "SiteGroupUrl", "SiteGroupUsers",
    "WebTitle", "WebUrl"
];

// Script Constants
const ScriptDescription = "Scan site(s) for security group information.";
const ScriptFileName = "security_groups.csv";
const ScriptName = "Security Group Information";

/**
 * Security Group Information
 * Displays a dialog to get the site information.
 */
class SecurityGroups {
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
                // Parse the role assignments
                Helper.Executor(web.RoleAssignments.results, (role) => {
                    let roleAssignment: Types.SP.RoleAssignmentOData = role as any;

                    // Parse the role definitions and create a list of permissions
                    let roleDefinitions = [];
                    for (let i = 0; i < roleAssignment.RoleDefinitionBindings.results.length; i++) {
                        // Add the permission name
                        roleDefinitions.push(roleAssignment.RoleDefinitionBindings.results[i].Name);
                    }

                    // See if this is a group
                    if (roleAssignment.Member["Users"] != null) {
                        let group: Types.SP.GroupOData = role.Member as any;

                        // Parse the users and create a list of members
                        let members = [];
                        for (let i = 0; i < group.Users.results.length; i++) {
                            let user = group.Users.results[i];

                            // Add the user information
                            members.push(user.Email || user.UserPrincipalName || user.Title);
                        }

                        // Add a row for this entry
                        this._rows.push({
                            SiteGroupId: group.Id,
                            SiteGroupName: group.LoginName,
                            SiteGroupPermission: roleDefinitions.join(', '),
                            SiteGroupUrl: web.Url + "/_layouts/15/people.aspx?MembershipGroupId=" + group.Id,
                            SiteGroupUsers: members.join(', '),
                            WebTitle: web.Title,
                            WebUrl: web.Url
                        });
                    } else {
                        let user: Types.SP.User = role.Member as any;

                        // Add a row for this entry
                        this._rows.push({
                            RoleAssignmentId: roleAssignment.PrincipalId,
                            SiteGroupId: user.Id,
                            SiteGroupName: user.Email || user.UserPrincipalName || user.Title,
                            SiteGroupPermission: roleDefinitions.join(', '),
                            WebTitle: web.Title,
                            WebUrl: web.Url
                        });
                    }
                });
            }).then(() => {
                // Check the next site collection
                resolve(null);
            });
        });
    }

    // Deletes a site group
    private deleteSiteGroup(webUrl: string, groupName: string, groupId: number) {
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

    // Deletes a site user
    private deleteSiteUser(webUrl: string, roleAssignmentId: number, userName: string) {
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
                    label: "Search Sub-Sites?",
                    name: "RecurseWebs",
                    className: "mb-2",
                    type: Components.FormControlTypes.Switch
                },
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
                                            // Include the site group information
                                            odata.Expand.push("RoleAssignments");
                                            odata.Expand.push("RoleAssignments/Member");
                                            odata.Expand.push("RoleAssignments/Member/Users");
                                            odata.Expand.push("RoleAssignments/RoleDefinitionBindings");
                                        },
                                        recursiveFl: formValues["RecurseWebs"],
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

        // Prevent auto close
        Modal.setAutoClose(false);

        // Show the modal dialog
        Modal.setHeader(ScriptName);

        // Render the table
        let elTable = document.createElement("div");
        new DataTable({
            el: elTable,
            rows: this._rows,
            dtProps: {
                dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                columnDefs: [
                    {
                        "targets": 5,
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
                // Order by the 2nd & 3rd column by default; ascending
                order: [[1, "asc"],[2, "asc"]]
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
                    name: "SiteGroupName",
                    title: "Site Group"
                },
                {
                    name: "SiteGroupPermission",
                    title: "Permission"
                },
                {
                    name: "SiteGroupUsers",
                    title: "User(s)"
                },
                {
                    className: "text-end",
                    name: "",
                    title: "",
                    onRenderCell: (el, col, row: IRowInfo) => {
                        let btnDelete: Components.IButton = null;

                        // Ensure this is a group
                        if (row.SiteGroupUrl) {
                            // Render the buttons
                            Components.ButtonGroup({
                                el,
                                buttons: [
                                    {
                                        text: "View",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the security group
                                            window.open(row.SiteGroupUrl, "_blank");
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
                                                this.deleteSiteGroup(row.WebUrl, row.SiteGroupName, row.SiteGroupId);
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
                                        this.deleteSiteUser(row.WebUrl, row.RoleAssignmentId, row.SiteGroupName);
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
                        new ExportCSV(ScriptFileName, CSVExportFields, this._rows);
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
export const SecurityGroupsModal: IScript = {
    init: SecurityGroups,
    name: ScriptName,
    description: ScriptDescription
};