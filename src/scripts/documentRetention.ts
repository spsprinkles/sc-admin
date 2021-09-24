import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import * as moment from "moment";
import { ExportCSV } from "./exportCSV";
import { Webs } from "./webs";

// Row Information
interface IRowInfo {
    DocumentName: string;
    DocumentUrl: string;
    LastModifiedDate?: string;
    ListName: string;
    WebTitle: string;
    WebUrl: string;
}

/**
 * Document Retention
 * Displays a dialog to get the site information.
 */
export class DocumentRetention {
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

    // Analyzes the site collection
    private analyzeSites(webs: Types.SP.WebOData[], date: string) {
        // Display a loading dialog
        LoadingDialog.setHeader("Analyzing the Site Collection");
        LoadingDialog.setBody("This will close after the information is analyzed...");
        LoadingDialog.show();

        // Return a promise
        return new Promise(resolve => {
            // Create the date
            let dt = moment(date).toISOString();

            // Parse the webs
            Helper.Executor(webs, web => {
                // Update the loading dialog
                LoadingDialog.setBody("Analyzing the document libraries for site:<br/>" + web.Url);

                // Return a promise
                return new Promise(resolve => {
                    // Parse the lists
                    Helper.Executor(web.Lists.results, list => {
                        // Ensure this is a document library
                        if (list.BaseTemplate != SPTypes.ListTemplateType.DocumentLibrary) { return; }

                        // Return a promise
                        return new Promise((resolve, reject) => {
                            // Query the list
                            list.Items().query({
                                Filter: "Modified lt '" + dt + "'",
                                Select: ["Id", "FileLeafRef", "FileRef", "FileSizeDisplay", "Modified"],
                                GetAllItems: true,
                                Top: 5000
                            }).execute(items => {
                                // Parse the items
                                Helper.Executor(items.results, item => {
                                    // Ensure this is an item
                                    if (item["FileSizeDisplay"]) {
                                        // Add a row for this entry
                                        this._rows.push({
                                            DocumentName: item["FileLeafRef"],
                                            DocumentUrl: item["FileRef"],
                                            LastModifiedDate: item["Modified"],
                                            ListName: list.Title,
                                            WebTitle: web.Title,
                                            WebUrl: web.Url
                                        });
                                    }
                                });

                                // Check the next list
                                resolve(null);
                            }, reject);
                        });
                    }).then(() => {
                        // Check the next web
                        resolve(null);
                    });
                });
            }).then(() => {
                // Hide the dialog
                LoadingDialog.hide();

                // Check the next site collection
                resolve(null);
            });
        });
    }

    // Deletes a document
    private deleteDocument(webUrl: string, docTitle: string, docUrl: string) {
        // Display a loading dialog
        LoadingDialog.setHeader("Deleting Document");
        LoadingDialog.setBody("Deleting the document '" + docTitle + "'. This will close after the request completes.");
        LoadingDialog.show();

        // Get the web context
        ContextInfo.getWeb(webUrl).execute(context => {
            // Delete the document
            Web(webUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).getFileByServerRelativeUrl(docUrl).delete().execute(
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

    // Determines if the document can be viewed in office online servers
    private isWopi(fileName: string) {
        // Determine if the file type is supported by WOPI
        let extension: any = fileName.split('.');
        extension = extension[extension.length - 1].toLowerCase();
        switch (extension) {
            // Excel
            case "csv":
            case "doc":
            case "docx":
            case "dot":
            case "dotx":
            case "pot":
            case "potx":
            case "pps":
            case "ppsx":
            case "ppt":
            case "pptx":
            case "xls":
            case "xlsx":
            case "xlt":
            case "xltx":
                return true;
            // Default
            default: {
                return false;
            }
        }
    }

    // Renders the dialog
    private render() {
        // Set the type
        Modal.setType(Components.ModalTypes.Large);

        // Set the header
        Modal.setHeader("Site Information");

        // Set the default date for retention
        let defaultDate = moment(Date.now()).subtract("months", "24").toISOString();

        // Render the form
        let form = Components.Form({
            controls: [
                {
                    label: "Search Sub-Sites?",
                    name: "RecurseWebs",
                    type: Components.FormControlTypes.Switch
                },
                {
                    label: "Document Date",
                    name: "DocumentDate",
                    type: Components.FormControlTypes.DateTime,
                    required: true,
                    value: defaultDate
                },
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
                                        recursiveFl: formValues["RecurseWebs"],
                                        onQueryWeb: odata => {
                                            // Include the lsits
                                            odata.Expand.push("Lists");

                                            // Select the title and type only
                                            odata.Select.push("Lists/BaseTemplate");
                                            odata.Select.push("Lists/Title");

                                            return odata;
                                        },
                                        onComplete: webs => {
                                            // Analyze the site
                                            this.analyzeSites(webs, formValues["DocumentDate"]).then(resolve);
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
        Modal.setHeader("Security Groups");

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
                    name: "DocumentName",
                    title: "Document Name"
                },
                {
                    name: "DocumentUrl",
                    title: "Document Url"
                },
                {
                    name: "LastModifiedDate",
                    title: "Last Modified Date"
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
                                        window.open(this.isWopi(row.DocumentName) ? row.WebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + row.DocumentUrl + "&action=view" : row.DocumentUrl, "_blank");
                                    }
                                },
                                {
                                    assignTo: btn => { btnDelete = btn; },
                                    text: "Delete",
                                    type: Components.ButtonTypes.OutlineDanger,
                                    onClick: () => {
                                        // Confirm the deletion of the group
                                        if (confirm("Are you sure you want to delete this document?")) {
                                            // Disable this button
                                            btnDelete.disable();

                                            // Delete the document
                                            this.deleteDocument(row.WebUrl, row.DocumentName, row.DocumentUrl);
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
                            "DocumentName", "DocumentUrl", "LastModifiedDate",
                            "ListName", "WebTitle", "WebUrl"
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