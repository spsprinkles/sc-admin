import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Search, Types, Web } from "gd-sprest-bs";
import { fileEarmark } from "gd-sprest-bs/build/icons/svgs/fileEarmark";
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
    SearchResult: string;
    SiteUrl: string;
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
                    description: "Enter the search terms using quotes for phrases [Ex: movie \"social media\" show]",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    value: Strings.SearchTerms
                },
                {
                    label: "File Types",
                    name: "FileTypes",
                    className: "mb-3",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    value: Strings.SearchFileTypes
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
        Modal.setFooter(Components.TooltipGroup({
            tooltips: [
                {
                    content: "Search for Documents based on the Search Terms",
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
                                let fileExtensions = formValues["FileTypes"].split(' ').join('", "');
                                let webUrls: string[] = formValues["Urls"].match(/[^\n]+/g);

                                // Determine the query text/phrases to search for
                                let searchTerms = formValues["SearchTerms"] || "";
                                let queryPhrases = [];

                                // See if the search terms contains phrases
                                let idxStart = searchTerms.indexOf('"');
                                if (idxStart < 0) {
                                    queryPhrases = searchTerms.split(' ');
                                } else {
                                    // Find all of the phrases
                                    let idx = 0;
                                    while (idxStart > 0) {
                                        // Add the words before the phrase
                                        let terms = searchTerms.substring(idx, idxStart).trim().split(' ');
                                        queryPhrases = queryPhrases.concat(terms);

                                        // Add the phrase
                                        let idxEnd = searchTerms.indexOf('"', idxStart + 1);
                                        queryPhrases.push(searchTerms.substring(idxStart, idxEnd + 1));

                                        // Find the next phrase
                                        idx = idxEnd + 1;
                                        idxStart = searchTerms.indexOf('"', idxEnd + 1);
                                    }

                                    // See if there are more words after the last phrase
                                    if (idx < searchTerms.length) {
                                        queryPhrases = queryPhrases.concat(searchTerms.substring(idx).trim().split(' '));
                                    }
                                }

                                // Remove the empty values and join them w/ an OR clause
                                let queryText = queryPhrases.filter(String).join(" OR ");

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
                                                Search.postQuery({
                                                    url: webUrl,
                                                    targetInfo: { requestDigest: context.GetContextWebInformation.FormDigestValue },
                                                    query: {
                                                        Querytext: `${queryText} IsDocument: true path: ${context.GetContextWebInformation.WebFullUrl}`,
                                                        RefinementFilters: {
                                                            results: [`fileExtension:or("${fileExtensions}")`]
                                                        },
                                                        RowLimit: 500,
                                                        SelectProperties: {
                                                            results: [
                                                                "Author", "FileExtension", "HitHighlightedSummary", "LastModifiedTime",
                                                                "ListId", "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId"
                                                            ]
                                                        }
                                                    },
                                                    onQueryCompleted: results => {
                                                        // Analyze the results
                                                        this.analyzeResult(results);
                                                    }
                                                }).then(() => {
                                                    // Check the next web
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
            onRendering: dtProps => {
                dtProps.columnDefs = [
                    {
                        "targets": 8,
                        "orderable": false,
                        "searchable": false
                    }
                ];

                // Order by the 4th column by default; ascending
                dtProps.order = [[3, "asc"]];

                // Return the properties
                return dtProps;
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
                        Components.TooltipGroup({
                            el,
                            tooltips: [
                                {
                                    content: "View Document",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconClassName: "mx-1",
                                        iconType: fileEarmark,
                                        iconSize: 24,
                                        text: "View",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the security group
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
export const DocumentSearchModal: IScript = {
    init: DocumentSearch,
    name: ScriptName,
    description: ScriptDescription
};