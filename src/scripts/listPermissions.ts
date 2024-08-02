import { CanvasForm, DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
import { search } from "gd-sprest-bs/build/icons/svgs/search";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import { ExportCSV, GetIcon, IScript, Webs } from "../common";

// Row Information
interface IRowInfo {
    Permissions: string;
    ItemId: number;
    ListName: string;
    ListType: number;
    ListUrl: string;
    WebTitle: string;
    WebUrl: string;
}

// CSV Export Fields
const CSVExportFields = [
    "WebTitle", "WebUrl", "ListType", "ListName", "ListUrl", "ItemId", "Permissions"
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
    private analyzeList(web: Types.SP.WebOData, list: Types.SP.ListOData) {
        // Return a promise
        return new Promise(resolve => {
            // Get the items where it has broken inheritance
            Web(web.ServerRelativeUrl).Lists(list.Title).Items().query({
                GetAllItems: true,
                Top: 5000,
                Select: ["Id", "HasUniqueRoleAssignments"]
            }).execute(items => {
                // Parse the items
                Helper.Executor(items.results, item => {
                    // See if this item doesn't have unique permissions
                    if (!item.HasUniqueRoleAssignments) { return; }

                    // Log the result
                    this._rows.push({
                        ItemId: item.Id,
                        ListName: list.Title,
                        ListType: list.BaseTemplate,
                        ListUrl: list.RootFolder.ServerRelativeUrl,
                        Permissions: "",
                        WebTitle: web.Title,
                        WebUrl: web.ServerRelativeUrl
                    });
                }).then(resolve);
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

                    // Get the lists w/ broken inheritance
                    Web(web.ServerRelativeUrl).Lists().query({
                        Filter: "HasUniqueRoleAssignments eq true",
                        Expand: ["RootFolder"],
                        Select: ["Id", "Title", "HasUniqueRoleAssignments", "RootFolder/ServerRelativeUrl"]
                    }).execute(lists => {
                        // Parse the lists
                        Helper.Executor(lists.results, list => {
                            // Analyze the list
                            return this.analyzeList(web, list);
                        }).then(resolve);
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
                    name: "ListUrl",
                    title: "List Url"
                },
                {
                    name: "ItemId",
                    title: "Item Id"
                },
                {
                    name: "Permissions",
                    title: "Permissions"
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
}

// Script Information
export const ListPermissionsModal: IScript = {
    init: ListPermissions,
    name: ScriptName,
    description: ScriptDescription
};