import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Search, Types, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import * as moment from "moment";
import { ExportCSV, isWopi, IScript } from "../common";
import Strings from "../strings";

// Row Information
interface IRowInfo {
    Author: string;
    DocumentExt: string;
    DocumentName: string;
    DocumentUrl: string;
    LastModifiedDate: string;
    ListId: string;
    SearchResult: string;
    WebId: string;
    WebUrl: string;
}

// CSV Export Fields
const CSVExportFields = [
    "DocumentName", "DocumentExt", "DocumentUrl", "Author",
    "LastModifiedDate", "ListId", "WebId", "WebUrl", "SearchResult"
];

// Script Constants
const ScriptDescription = "Searches for documents by key word(s) within a site.";
const ScriptFileName = "document_search.csv";
const ScriptName = "Document Search";

/**
 * Document Search
 * Displays a dialog to search documents by key words.
 */
class DocumentSearch {
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
                    case "HitHighlightedSummary":
                        rowInfo.SearchResult = cell.Value;
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
                    case "SiteName":
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
        LoadingDialog.setBody("Deleting the document: '" + docTitle + "'. This will close after the request completes.");
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

        // Render the form
        let form = Components.Form({
            controls: [
                {
                    label: "Search Terms",
                    name: "SearchTerms",
                    className: "mb-3",
                    type: Components.FormControlTypes.TextField,
                    required: true
                },
                {
                    label: "File Types",
                    name: "FileTypes",
                    className: "mb-3",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    value: "csv doc docx dot dotx pdf pot potx pps ppsx ppt pptx txt xls xlsx xlt xltx"
                },
                {
                    label: "Site Url(s)",
                    name: "Urls",
                    description: "Enter the relative site url(s) [Ex: /sites/dev]",
                    errorMessage: "Please enter a site url",
                    type: Components.FormControlTypes.TextArea,
                    required: true,
                    rows: 5,
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
                    text: "Search",
                    type: Components.ButtonTypes.OutlineSuccess,
                    onClick: () => {
                        // Ensure the form is valid
                        if (form.isValid()) {
                            let formValues = form.getValues();
                            let fileExtensions = formValues["FileTypes"].split(' ').join('", "');
                            let queryText = formValues["SearchTerms"].split(' ').join(" OR ");
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
                                    // Update the dialog
                                    LoadingDialog.setBody("Searching " + webUrl);

                                    // Get the context information of the web
                                    ContextInfo.getWeb(webUrl).execute(
                                        // Success
                                        (context) => {
                                            // Search the site
                                            Search(webUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).postquery({
                                                Querytext: `${queryText} IsDocument: true path: ${context.GetContextWebInformation.WebFullUrl}`,
                                                RefinementFilters: {
                                                    results: [`fileExtension:or("${fileExtensions}")`]
                                                },
                                                SelectProperties: {
                                                    results: [
                                                        "Author", "FileExtension", "HitHighlightedSummary", "LastModifiedTime",
                                                        "ListId", "Path", "SiteName", "Title", "WebId"
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
                        "targets": 8,
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
                    name: "SearchResult",
                    title: "Search Result",
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
                                        window.open(isWopi(`${row.DocumentName}.${row.DocumentExt}`) ? row.WebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + row.DocumentUrl + "&action=view" : row.DocumentUrl, "_blank");
                                    }
                                },
                                {
                                    text: "Download",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Download the document
                                        window.open(`${row.WebUrl}/_layouts/15/download.aspx?SourceUrl=${row.DocumentUrl}`, "_blank");
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
export const DocumentSearchModal: IScript = {
    init: DocumentSearch,
    name: ScriptName,
    description: ScriptDescription
};