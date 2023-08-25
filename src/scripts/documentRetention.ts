import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Search, Types, Web } from "gd-sprest-bs";
import { fileEarmarkArrowDown } from "gd-sprest-bs/build/icons/svgs/fileEarmarkArrowDown";
import { search } from "gd-sprest-bs/build/icons/svgs/search";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import * as jQuery from "jquery";
import * as moment from "moment";
import { ExportCSV, GetIcon, isWopi, IScript } from "../common";
import Strings from "../strings";

// Row Information
interface IRowInfo {
    Author: string;
    DocumentExt: string;
    DocumentName: string;
    DocumentUrl: string;
    LastModifiedDate: string;
    ListId: string;
    SiteUrl: string;
    WebId: string;
    WebUrl: string;
}

// CSV Export Fields
const CSVExportFields = [
    "DocumentName", "DocumentExt", "DocumentUrl", "Author",
    "LastModifiedDate", "ListId", "SiteUrl", "WebId", "WebUrl"
];

// Script Constants
const ScriptDescription = "Scans for files older than a specified date.";
const ScriptFileName = "document_retention.csv";
const ScriptName = "Document Retention";

/**
 * Document Retention
 * Displays a dialog to get the site information.
 */
class DocumentRetention {
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

    // Analyzes the results
    private analyzeResult(results: Types.Microsoft.Office.Server.Search.REST.SearchResult) {
        // Ensure results exist
        if (results.PrimaryQueryResult == null) { return; }

        // Parse the results
        for (let i = 0; i < results.PrimaryQueryResult.RelevantResults.RowCount; i++) {
            let result = results.PrimaryQueryResult.RelevantResults.Table.Rows.results[i];
            let rowInfo: IRowInfo = {} as any;

            // Parse the cells
            for (let j = 0; j < result.Cells.results.length; j++) {
                let cell = result.Cells.results[j];

                // See if this is a target value
                switch (cell.Key) {
                    case "Author":
                        rowInfo.Author = cell.Value;
                        break;
                    case "FileExtension":
                        rowInfo.DocumentExt = cell.Value;
                        break;
                    case "LastModifiedTime":
                        rowInfo.LastModifiedDate = cell.Value;
                        break;
                    case "ListId":
                        rowInfo.ListId = cell.Value;
                        break;
                    case "Path":
                        rowInfo.DocumentUrl = cell.Value;
                        break;
                    case "SPSiteUrl":
                        rowInfo.SiteUrl = cell.Value;
                        break;
                    case "SPWebUrl":
                        rowInfo.WebUrl = cell.Value;
                        break;
                    case "Title":
                        rowInfo.DocumentName = cell.Value;
                        break;
                    case "WebId":
                        rowInfo.WebId = cell.Value;
                        break;
                }
            }

            // Append the row
            this._rows.push(rowInfo);
        }
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

    // Renders the dialog
    private render() {
        // Set the type
        Modal.setType(Components.ModalTypes.Large);

        // Prevent auto close
        Modal.setAutoClose(false);

        // Set the header
        Modal.setHeader(ScriptName);

        // Set the default date for retention
        let defaultDate = moment(Date.now()).subtract(Strings.SearchMonths, "months").toISOString();

        // Render the form
        let form = Components.Form({
            controls: [
                {
                    label: "Document Date",
                    name: "DocumentDate",
                    className: "mb-3",
                    type: Components.FormControlTypes.DateTime,
                    required: true,
                    value: defaultDate
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
        Modal.setFooter(Components.TooltipGroup({
            tooltips: [
                {
                    content: "Search for Documents older than the Document Date",
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

                                // Display a loading dialog
                                LoadingDialog.setHeader("Searching Webs");
                                LoadingDialog.setBody("This will close after the searches are completed...");
                                LoadingDialog.show();

                            // Parse the webs
                            Helper.Executor(webUrls, webUrl => {
                                // Return a promise
                                return new Promise((resolve) => {
                                    // Get the context information of the web
                                    ContextInfo.getWeb(webUrl).execute(
                                        // Success
                                        (context) => {
                                            // Search the site
                                            Search(webUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).postquery({
                                                Querytext: `IsDocument: true LastModifiedTime<${moment(formValues["DocumentDate"]).format("YYYY-MM-DD")} path: ${context.GetContextWebInformation.WebFullUrl}`,
                                                RowLimit: 5000,
                                                SelectProperties: {
                                                    results: [
                                                        "Author", "FileExtension", "HitHighlightedSummary", "LastModifiedTime",
                                                        "ListId", "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId"
                                                    ]
                                                }
                                            }).execute(results => {
                                                // Analyze the results
                                                this.analyzeResult(results.postquery);

                                                    // Check the next web
                                                    resolve(null);
                                                }, () => {
                                                    // Error getting the search results
                                                    this._errors.push(webUrl);
                                                    resolve(null);
                                                });
                                            },

                                            // Error
                                            () => {
                                                // Append the error and check the next web
                                                this._errors.push(webUrl);
                                                resolve(null);
                                            }
                                        );
                                    });
                                }).then(() => {
                                    // Render the summary
                                    this.renderSummary();

                                    // Hide the dialog
                                    LoadingDialog.hide();
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
            dtProps: {
                dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                columnDefs: [
                    {
                        "targets": 7,
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
                // Order by the 4th column by default; ascending
                order: [[3, "asc"]]
            },
            columns: [
                {
                    name: "WebId",
                    title: "Web Id"
                },
                {
                    name: "WebUrl",
                    title: "Web Url"
                },
                {
                    name: "ListId",
                    title: "List Id"
                },
                {
                    name: "DocumentUrl",
                    title: "Document Url"
                },
                {
                    name: "DocumentName",
                    title: "File Name",
                    onRenderCell: (el, col, item: IRowInfo) => {
                        el.innerHTML = item.DocumentName + "." + item.DocumentExt;
                    }
                },
                {
                    name: "Author",
                    title: "Author(s)",
                    onRenderCell: (el, col, item: IRowInfo) => {
                        // Clear the cell
                        el.innerHTML = "";

                        // Validate Author exists & split by ;
                        let authors = (item.Author && item.Author.split(";")) || [item.Author];

                        // Parse the Authors
                        authors.forEach(author => {
                            // Append the Author
                            el.innerHTML += (author + "<br/>");
                        });
                    }
                },
                {
                    name: "LastModifiedDate",
                    title: "Modified",
                    onRenderCell: (el, col, item: IRowInfo) => {
                        el.innerHTML = item.LastModifiedDate ? moment(item.LastModifiedDate).format(Strings.TimeFormat) : "";
                    }
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
                                    content: "View Document",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconType: GetIcon(24, 24, "EntryView", "mx-1"),
                                        text: "View",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // View the document
                                            window.open(isWopi(`${row.DocumentName}.${row.DocumentExt}`) ? row.WebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + row.DocumentUrl + "&action=view" : row.DocumentUrl, "_blank");
                                        }
                                    }
                                },
                                {
                                    content: "Download Document",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconClassName: "mx-1",
                                        iconType: fileEarmarkArrowDown,
                                        iconSize: 24,
                                        text: "Download",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Download the document
                                            window.open(`${row.WebUrl}/_layouts/15/download.aspx?SourceUrl=${row.DocumentUrl}`, "_blank");
                                        }
                                    }
                                },
                                {
                                    content: "Delete Document",
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
                                            if (confirm("Are you sure you want to delete this document?")) {
                                                // Disable this button
                                                btnDelete.disable();

                                                // Delete the document
                                                this.deleteDocument(row.WebUrl, row.DocumentName, row.DocumentUrl);
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
export const DocumentRetentionModal: IScript = {
    init: DocumentRetention,
    name: ScriptName,
    description: ScriptDescription
};