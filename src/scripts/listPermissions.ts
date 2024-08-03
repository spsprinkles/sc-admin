import { DataTable, Documents, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { search } from "gd-sprest-bs/build/icons/svgs/search";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import { ExportCSV, GetIcon, IScript, Webs } from "../common";

// Row Information
interface IRowInfo {
    FileName?: string;
    FileUrl?: string;
    ItemId: number;
    ListName: string;
    ListType: number;
    ListUrl: string;
    ListViewUrl: string;
    RoleAssignmentId?: number;
    SiteGroupId: number;
    SiteGroupName: string;
    SiteGroupPermission: string;
    SiteGroupUrl?: string;
    SiteGroupUsers?: string;
    WebTitle: string;
    WebUrl: string;
}

// CSV Export Fields
const CSVExportFields = [
    "WebTitle", "WebUrl", "ListType", "ListName", "ListUrl", "ItemId", "FileName", "FileUrl", "Permissions", "RoleAssignmentId",
    "SiteGroupId", "SiteGroupName", "SiteGroupPermission", "SiteGroupUrl", "SiteGroupUsers"
];

// Script Constants
const ScriptDescription = "Scan for list permissions and broken inheritance.";
const ScriptFileName = "list_permissions.csv";
const ScriptName = "List Permissions";

/**
 * List Permissions
 * Displays a dialog to get the list permissions.
 */
class ListPermissions {
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

    // Analyzes the list
    private analyzeList(context: Types.SP.ContextWebInformation, web: Types.SP.WebOData, list: Types.SP.ListOData): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            let Select = ["Id", "HasUniqueRoleAssignments"];

            // See if this is a document library
            if (list.BaseTemplate == SPTypes.ListTemplateType.DocumentLibrary || list.BaseTemplate == SPTypes.ListTemplateType.PageLibrary) {
                // Get the file information
                Select.push("FileLeafRef");
                Select.push("FileRef");
            }

            // Get the items where it has broken inheritance
            Web(web.ServerRelativeUrl).Lists(list.Title).Items().query({
                GetAllItems: true,
                Select,
                Top: 5000
            }).execute(items => {
                // Create a batch job
                let batch = Web(web.ServerRelativeUrl, { requestDigest: context.FormDigestValue }).Lists(list.Title);

                // Parse the items
                Helper.Executor(items.results, item => {
                    // See if this item doesn't have unique permissions
                    if (!item.HasUniqueRoleAssignments) { return; }

                    // Get the permissions
                    batch.Items(item.Id).RoleAssignments().query({
                        Expand: [
                            "Member/Users", "RoleDefinitionBindings"
                        ]
                    }).batch(roles => {
                        // Parse the role assignments
                        Helper.Executor(roles.results, roleAssignment => {
                            // Parse the role definitions and create a list of permissions
                            let roleDefinitions = [];
                            for (let i = 0; i < roleAssignment.RoleDefinitionBindings.results.length; i++) {
                                // Add the permission name
                                roleDefinitions.push(roleAssignment.RoleDefinitionBindings.results[i].Name);
                            }

                            // See if this is a group
                            if (roleAssignment.Member["Users"] != null) {
                                let group: Types.SP.GroupOData = roleAssignment.Member as any;

                                // Parse the users and create a list of members
                                let members = [];
                                for (let i = 0; i < group.Users.results.length; i++) {
                                    let user = group.Users.results[i];

                                    // Add the user information
                                    members.push(user.Email || user.UserPrincipalName || user.Title);
                                }

                                // Add a row for this entry
                                this._rows.push({
                                    FileName: item["FileLeafRef"],
                                    FileUrl: item["FileRef"],
                                    ItemId: item.Id,
                                    ListName: list.Title,
                                    ListType: list.BaseTemplate,
                                    ListUrl: list.RootFolder.ServerRelativeUrl,
                                    ListViewUrl: list.DefaultDisplayFormUrl,
                                    SiteGroupId: group.Id,
                                    SiteGroupName: group.LoginName,
                                    SiteGroupPermission: roleDefinitions.join(', '),
                                    SiteGroupUrl: web.Url + "/_layouts/15/people.aspx?MembershipGroupId=" + group.Id,
                                    SiteGroupUsers: members.join(', '),
                                    WebTitle: web.Title,
                                    WebUrl: web.Url
                                });
                            } else {
                                let user: Types.SP.User = roleAssignment.Member as any;

                                // Add a row for this entry
                                this._rows.push({
                                    FileName: item["FileLeafRef"],
                                    FileUrl: item["FileRef"],
                                    ItemId: item.Id,
                                    ListName: list.Title,
                                    ListType: list.BaseTemplate,
                                    ListUrl: list.RootFolder.ServerRelativeUrl,
                                    ListViewUrl: list.DefaultDisplayFormUrl,
                                    RoleAssignmentId: roleAssignment.PrincipalId,
                                    SiteGroupId: user.Id,
                                    SiteGroupName: user.Email || user.UserPrincipalName || user.Title,
                                    SiteGroupPermission: roleDefinitions.join(', '),
                                    WebTitle: web.Title,
                                    WebUrl: web.Url
                                });
                            }
                        });
                    });
                }).then(() => {
                    // Execute the batch job
                    batch.execute(() => {
                        // Resolve the request
                        resolve();
                    });
                });
            });
        });
    }

    // Analyzes the site
    private analyzeSites(webs: Types.SP.WebOData[]) {
        // Return a promise
        return new Promise(resolve => {
            let counter = 0;

            // Show a loading dialog
            LoadingDialog.setHeader("Analyzing the data");
            LoadingDialog.setBody("Getting lists with broken inheritence...");
            LoadingDialog.show();

            // Parse the webs
            Helper.Executor(webs, web => {
                // Return a promise
                return new Promise(resolve => {
                    // Update the loading dialog
                    LoadingDialog.setBody(`Getting list information (${++counter} of ${webs.length})`);

                    // Get the context for this web
                    ContextInfo.getWeb(web.ServerRelativeUrl).execute(context => {
                        // Get the lists w/ broken inheritance
                        Web(web.ServerRelativeUrl).Lists().query({
                            Filter: "Hidden eq false and HasUniqueRoleAssignments eq true",
                            Expand: ["DefaultDisplayFormUrl", "DefaultViewFormUrl", "RootFolder"],
                            Select: ["BaseTemplate", "Id", "Title", "HasUniqueRoleAssignments", "RootFolder/ServerRelativeUrl"]
                        }).execute(lists => {
                            // Parse the lists
                            Helper.Executor(lists.results, list => {
                                // Analyze the list
                                return this.analyzeList(context.GetContextWebInformation, web, list);
                            }).then(resolve);
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
                    name: "ListName",
                    title: "List Name"
                },
                {
                    name: "ItemId",
                    title: "Item Id"
                },
                {
                    name: "FileName",
                    title: "File Name"
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

                        // Render the buttons
                        Components.TooltipGroup({
                            el,
                            tooltips: [
                                {
                                    content: "Click to view the item properties.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconType: GetIcon(24, 24, "CustomList", "mx-1"),
                                        text: "View",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the security group
                                            window.open(row.ListViewUrl + "?Id=" + row.ItemId, "_blank");
                                        }
                                    }
                                },
                                {
                                    content: "Click to view the file.",
                                    className: row.FileName ? "" : "d-none",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconType: GetIcon(24, 24, "ExcelDocument", "mx-1"),
                                        text: "File",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Open the document in view mode
                                            Documents.open(row.FileUrl, false, row.WebUrl);
                                        }
                                    }
                                },
                                {
                                    content: "Click to restore permissions to inherit.",
                                    btnProps: {
                                        assignTo: btn => { btnDelete = btn; },
                                        className: "pe-2 py-1",
                                        iconType: GetIcon(24, 24, "PeopleTeamDelete", "mx-1"),
                                        text: "Restore",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Confirm the deletion of the group
                                            if (confirm("Are you sure you restore the permissions to inherit?")) {
                                                // Revert the permissions
                                                this.revertPermissions(row);
                                            }
                                        }
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

    // Reverts the item permissions
    private revertPermissions(row: IRowInfo) {
        // Show a loading dialog
        LoadingDialog.setHeader("Restoring Permissions");
        LoadingDialog.setBody("This window will close after the item permissions are restored...");
        LoadingDialog.show();

        // Get the context
        ContextInfo.getWeb(row.WebUrl).execute(context => {
            // Restore the permissions
            Web(context.GetContextWebInformation.WebFullUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue })
                .Lists(row.ListName).Items(row.ItemId).resetRoleInheritance().execute(() => {
                    // Close the loading dialog
                    LoadingDialog.hide();
                });
        });
    }
}

// Script Information
export const ListPermissionsModal: IScript = {
    init: ListPermissions,
    name: ScriptName,
    description: ScriptDescription
};