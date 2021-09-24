import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { ExportCSV } from "./exportCSV";
import { Webs } from "./webs";

// Row Information
interface IRowInfo {
    ListDescription: string;
    ListId: string;
    ListName: string;
    ListType: string;
    ListUrl: string;
    WebTitle: string;
    WebUrl: string;
}

/**
 * Lists
 * Displays a dialog to get the site information.
 */
export class Lists {
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
                // Parse the lists
                Helper.Executor(web.Lists.results, list => {
                    // Add a row for this entry
                    this._rows.push({
                        ListDescription: list.Description,
                        ListId: list.Id,
                        ListName: list.Title,
                        ListType: list.BaseType + "",
                        ListUrl: (list.RootFolder as any as Types.SP.Folder).ServerRelativeUrl,
                        WebTitle: web.Title,
                        WebUrl: web.Url
                    });
                });
            }).then(() => {
                // Check the next site collection
                resolve(null);
            });
        });
    }

    // Deletes a list
    private deleteList(webUrl: string, listName: string) {
        // Display a loading dialog
        LoadingDialog.setHeader("Deleting List");
        LoadingDialog.setBody("Deleting the list: " + listName);
        LoadingDialog.show();

        // Get the web context
        ContextInfo.getWeb(webUrl).execute(context => {
            // Delete the site group
            Web(webUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).getList(listName).delete().execute(
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
                                    new Webs({
                                        url: webUrl,
                                        onQueryWeb: (odata) => {
                                            // Include the list information
                                            odata.Expand.push("Lists/RootFolder");
                                            odata.Select.push("Lists/BaseType");
                                            odata.Select.push("Lists/Description");
                                            odata.Select.push("Lists/Id");
                                            odata.Select.push("Lists/RootFolder/ServerRelativeUrl");
                                            odata.Select.push("Lists/Title");
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
                        "targets": [6],
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
                    name: "ListType",
                    title: "List Type"
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
                    name: "ListDescription",
                    title: "List Description"
                },
                {
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
                                        window.open(row.ListUrl, "_blank");
                                    }
                                },
                                {
                                    assignTo: btn => { btnDelete = btn; },
                                    text: "Delete",
                                    type: Components.ButtonTypes.OutlineDanger,
                                    onClick: () => {
                                        // Confirm the deletion of the group
                                        if (confirm("Are you sure you want to delete this list?")) {
                                            // Disable this button
                                            btnDelete.disable();

                                            // Delete the site group
                                            this.deleteList(row.WebUrl, row.ListName);
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
        Modal.setFooter(Components.ButtonGroup({
            buttons: [
                {
                    text: "Export",
                    type: Components.ButtonTypes.OutlineSuccess,
                    onClick: () => {
                        // Export the CSV
                        new ExportCSV("security_groups.csv", [
                            "ListDescription", "ListId", "ListName",
                            "ListType", "ListUrl", "WebTitle", "WebUrl"
                        ], this._rows);
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