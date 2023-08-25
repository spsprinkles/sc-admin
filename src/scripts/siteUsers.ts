import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
import { personX } from "gd-sprest-bs/build/icons/svgs/personX";
import { search } from "gd-sprest-bs/build/icons/svgs/search";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import * as jQuery from "jquery";
import { ExportCSV, GetIcon, IScript, Webs } from "../common";

// Row Information
interface IRowInfo {
    Role?: string;
    RoleInfo?: string;
    LoginName: string;
    Name: string;
    Email?: string;
    Group?: string;
    GroupId?: number;
    GroupInfo?: string;
    Id?: number;
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
    "WebTitle", "WebUrl",
    "Name", "UserName", "Email", "Group", "GroupInfo",
    "Role", "RoleInfo"
];

// Script Constants
const ScriptDescription = "Scan site(s) for specified site users.";
const ScriptFileName = "site_users_info.csv";
const ScriptName = "Site Users Information";

/**
 * Security Group Information
 * Displays a dialog to get the site user information.
 */
class SiteUsers {
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
    private analyzeSites(rootWeb: Types.SP.WebOData, webs: Types.SP.WebOData[], user: string | Types.SP.User) {
        // Return a promise
        return new Promise(resolve => {
            // Show a loading dialog
            LoadingDialog.setHeader("Getting User Information");
            LoadingDialog.setBody(rootWeb.Url);
            LoadingDialog.show();

            // Get the users
            this.getUsers(rootWeb.Url, user, rootWeb.ParentWeb).then(users => {
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

    // Get the user information
    private getUserInfo(rootWeb: Types.SP.WebOData, webs: Types.SP.WebOData[], userInfo: IUserInfo) {
        // Return a promise
        return new Promise((resolve) => {
            // Get the user information
            rootWeb.getUserById(userInfo.Id).query({ Expand: ["Groups"] }).execute(
                // Success
                user => {
                    // Parse the groups
                    Helper.Executor(user.Groups.results, group => {
                        // Parse the roles
                        for (let i = 0; i < rootWeb.RoleAssignments.results.length; i++) {
                            let role: Types.SP.RoleAssignmentOData = rootWeb.RoleAssignments.results[i] as any;

                            // See if the user belongs to this role
                            if (role.Member.LoginName == group.LoginName) {
                                // Add the user information
                                this._rows.push({
                                    WebUrl: rootWeb.Url,
                                    WebTitle: rootWeb.Title,
                                    Id: userInfo.Id,
                                    LoginName: userInfo.Name,
                                    Name: userInfo.Title || userInfo.Name,
                                    Email: userInfo.EMail,
                                    Group: group.Title,
                                    GroupId: group.Id,
                                    GroupInfo: group.Description || "",
                                    Role: role.RoleDefinitionBindings.results[0].Name,
                                    RoleInfo: role.RoleDefinitionBindings.results[0].Description || ""
                                });
                            }
                        }
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
    private getUsers(siteUrl: string, user: string | Types.SP.User, parentWeb: Types.SP.WebInformation): PromiseLike<IUserInfo[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let users: IUserInfo[] = [];

            // See if this is a root web
            if (parentWeb.Id) {
                // Skip this web
                resolve(users);
                return;
            }

            // See if we are searching by a string
            if (typeof (user) === "string") {
                // Get the user information list
                Web(siteUrl).Lists("User Information List").Items().query({
                    Filter: `substringof('${user}', Name) or substringof('${user}', Title) or substringof('${user}', UserName)`,
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
            } else {
                // Get the users
                Web(siteUrl).Lists("User Information List").Items().query({
                    GetAllItems: true,
                    Select: ["Id", "Name", "EMail", "Title", "UserName"],
                    Top: 5000
                }).execute(items => {
                    // Parse the items
                    for (let i = 0; i < items.results.length; i++) {
                        let item = items.results[i];

                        // See if this is the target user
                        if (item["EMail"] == user.Email || item["Name"] == user.LoginName || item["UserName"] == user.UserPrincipalName) {
                            // Add the user
                            users.push({
                                EMail: item["EMail"],
                                Id: item.Id,
                                Name: item["Name"],
                                Title: item.Title,
                                UserName: item["UserName"]
                            });
                        }
                    }

                    // Resolve the request
                    resolve(users);
                }, reject);
            }
        });
    }

    // Removes a user from a group
    private removeUser(webUrl: string, user: string, userId: number, group: string) {
        // Display a loading dialog
        LoadingDialog.setHeader("Removing Site User");
        LoadingDialog.setBody(`Removing the site user '${user}' from all groups. This will close after the request completes.`);
        LoadingDialog.show();

        // Get the web context
        ContextInfo.getWeb(webUrl).execute(context => {
            // Remove the user from the site
            Web(webUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).SiteUsers().removeById(userId).execute(
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
                } as Components.IFormControlPropsPeoplePicker,
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
                            let userName: string = formValues["UserName"].trim();
                            let user: Types.SP.User = formValues["PeoplePicker"][0];
                            let webUrls: string[] = formValues["Urls"].match(/[^\n]+/g);

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
                                            odata.Expand.push("ParentWeb");
                                            odata.Expand.push("RoleAssignments");
                                            odata.Expand.push("RoleAssignments/Groups");
                                            odata.Expand.push("RoleAssignments/Member");
                                            odata.Expand.push("RoleAssignments/Member/Users");
                                            odata.Expand.push("RoleAssignments/RoleDefinitionBindings");
                                        },
                                        recursiveFl: false,
                                        onComplete: webs => {
                                            // Analyze the site
                                            this.analyzeSites(webs[0], webs, userName || user).then(resolve);
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
                    className: "pe-2 py-1",
                    iconClassName: "mx-1",
                    iconType: xSquare,
                    iconSize: 24,
                    text: "Cancel",
                    type: Components.ButtonTypes.OutlineSecondary,
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
                        "targets": 9,
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
                    name: "Name",
                    title: "User Name"
                },
                {
                    name: "LoginName",
                    title: "Login"
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
                    title: "Group Detail",
                    onRenderCell: (el) => {
                        // Add the data-filter attribute for searching notes properly
                        el.setAttribute("data-filter", el.innerHTML);
                        // Add the data-order attribute for sorting notes properly
                        el.setAttribute("data-order", el.innerHTML);

                        // Declare a span element
                        let span = document.createElement("span");
                        span.className = "notes";

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
                    name: "Role",
                    title: "Permission"
                },
                {
                    name: "RoleInfo",
                    title: "Permission Detail",
                    onRenderCell: (el) => {
                        // Add the data-filter attribute for searching notes properly
                        el.setAttribute("data-filter", el.innerHTML);
                        // Add the data-order attribute for sorting notes properly
                        el.setAttribute("data-order", el.innerHTML);

                        // Declare a span element
                        let span = document.createElement("span");
                        span.className = "notes";

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
                    className: "text-end",
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
                                        className: "pe-2 py-1",
                                        iconType: GetIcon(24, 24, "EntryView", "mx-1"),
                                        text: "View",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        isDisabled: !(row.GroupId > 0),
                                        onClick: () => {
                                            // View the group
                                            window.open(`${row.WebUrl}/${ContextInfo.layoutsUrl}/people.aspx?MembershipGroupId=${row.GroupId}`);
                                        }
                                    },
                                    {
                                        assignTo: btn => { btnDelete = btn; },
                                        iconClassName: "mx-1",
                                        iconType: personX,
                                        iconSize: 24,
                                        text: "Remove User",
                                        type: Components.ButtonTypes.OutlineDanger,
                                        isDisabled: !(row.Id > 0),
                                        onClick: () => {
                                            // Confirm the removal of the user
                                            if (confirm("Are you sure you want to remove the user from this site?")) {
                                                // Disable this button
                                                btnDelete.disable();

                                                // Remove the user
                                                this.removeUser(row.WebUrl, row.Name, row.Id, row.Group);
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
                                    if (confirm("Are you sure you want to remove the user from this site?")) {
                                        // Disable this button
                                        btnDelete.disable();

                                        // Remove the user
                                        this.removeUser(row.WebUrl, row.Name, row.Id, row.Group);
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
                    className: "pe-2 py-1",
                    iconType: GetIcon(24, 24, "ExcelDocument", "icon-svg mx-1"),
                    text: "Export",
                    type: Components.ButtonTypes.OutlineSuccess,
                    onClick: () => {
                        // Export the CSV
                        new ExportCSV(ScriptFileName, CSVExportFields, this._rows);
                    }
                },
                {
                    className: "pe-2 py-1",
                    iconClassName: "mx-1",
                    iconType: xSquare,
                    iconSize: 24,
                    text: "Cancel",
                    type: Components.ButtonTypes.OutlineSecondary,
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
export const SiteUsersModal: IScript = {
    init: SiteUsers,
    name: ScriptName,
    description: ScriptDescription
};