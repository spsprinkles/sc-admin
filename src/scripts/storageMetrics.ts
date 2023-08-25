import { DataTable, Modal } from "dattatable";
import { Components, Helper, Types } from "gd-sprest-bs";
import { search } from "gd-sprest-bs/build/icons/svgs/search";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import * as jQuery from "jquery";
import * as moment from "moment";
import { ExportCSV, GetIcon, IScript, Webs } from "../common";
import Strings from "../strings";

// Row Information
interface IRowInfo {
    LastModified: string;
    TotalFileCount: number;
    TotalFileStreamSize: number;
    TotalSize: number;
    WebDescription: string;
    WebTitle: string;
    WebUrl: string;
}

// CSV Export Fields
const CSVExportFields = [
    "WebTitle", "WebUrl", "WebDescription", "LastModified",
    "TotalFileCount", "TotalFileStreamSize", "TotalSize"
];

// Script Constants
const ScriptDescription = "Scans site(s) for storage metrics.";
const ScriptFileName = "storage_metrics.csv";
const ScriptName = "Storage Metrics";

/**
 * Site Storage Metrics
 * Displays a dialog to get the site collection storage metrics.
 */
class StorageMetrics {
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

    // Analyzes the web
    private analyzeWebs(webs: Types.SP.WebOData[]) {
        // Return a promise
        return new Promise(resolve => {
            // Parse the webs
            Helper.Executor(webs, web => {
                let rootFolder: Types.SP.FolderOData = web.RootFolder as any;

                // Calculate the total file stream size
                let totalFileStreamSize = rootFolder.StorageMetrics.TotalFileStreamSize;
                totalFileStreamSize = totalFileStreamSize > 0 ? (totalFileStreamSize / Math.pow(1024, 3)).toFixed(Strings.FractionDigits) + " GB" as any : totalFileStreamSize;

                // Calculate the total size
                let totalSize = rootFolder.StorageMetrics.TotalSize;
                totalSize = totalSize > 0 ? (totalSize / Math.pow(1024, 3)).toFixed(Strings.FractionDigits) + " GB" as any : totalSize;

                // Add a row for this entry
                this._rows.push({
                    LastModified: rootFolder.StorageMetrics.LastModified,
                    TotalFileCount: rootFolder.StorageMetrics.TotalFileCount,
                    TotalFileStreamSize: totalFileStreamSize,
                    TotalSize: totalSize,
                    WebDescription: web.Description,
                    WebTitle: web.Title,
                    WebUrl: web.ServerRelativeUrl
                });
            }).then(() => {
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
                                        recursiveFl: formValues["RecurseWebs"],
                                        onQueryWeb: (odata) => {
                                            // Include the web information
                                            odata.Select.push("Description");
                                            odata.Select.push("Title");
                                            odata.Select.push("ServerRelativeUrl");

                                            // Include the storage metrics
                                            odata.Expand.push("RootFolder/StorageMetrics")
                                        },
                                        onComplete: webs => {
                                            // Analyze the site
                                            this.analyzeWebs(webs).then(resolve);
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
                // Order by the 2nd column by default; ascending
                order: [[1, "asc"]]
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
                    name: "WebDescription",
                    title: "Description",
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
                    name: "LastModified",
                    title: "Modified",
                    onRenderCell: (el, col, item: IRowInfo) => {
                        el.innerHTML = item.LastModified ? moment(item.LastModified).format(Strings.TimeFormat) : "";
                    }
                },
                {
                    name: "TotalFileCount",
                    title: "Total File Count"
                },
                {
                    name: "TotalFileStreamSize",
                    title: "Total Stream Size"
                },
                {
                    name: "TotalSize",
                    title: "Total Size"
                },
                {
                    className: "text-end",
                    name: "",
                    title: "",
                    onRenderCell: (el, col, row: IRowInfo) => {
                        // Render the buttons
                        Components.ButtonGroup({
                            el,
                            buttons: [
                                {
                                    className: "pe-2 py-1",
                                    iconType: GetIcon(24, 24, "EntryView", "mx-1"),
                                    text: "View",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Show the security group
                                        window.open(row.WebUrl, "_blank");
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
export const StorageMetricsModal: IScript = {
    init: StorageMetrics,
    name: ScriptName,
    description: ScriptDescription
};