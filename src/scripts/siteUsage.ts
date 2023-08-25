import { DataTable, Modal } from "dattatable";
import { Components, Helper, Types } from "gd-sprest-bs";
import { search } from "gd-sprest-bs/build/icons/svgs/search";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import * as jQuery from "jquery";
import { ExportCSV, GetIcon, IScript, Sites } from "../common";
import Strings from "../strings";

// Row Information
interface IRowInfo {
    SiteBandwidth: number;
    SiteDescription: string;
    SiteDiscussionStorage: number;
    SiteHits: number;
    SitePercentageUsed: number;
    SiteStorage: number;
    SiteTitle: string;
    SiteUrl: string;
    SiteVisits: number;
}

// CSV Export Fields
const CSVExportFields = [
    "SiteId", "SiteTitle", "SiteUrl", "SiteDescription",
    "SiteStorage", "SitePercentageUsed", "SiteDiscussionStorage",
    "SiteHits", "SiteVisits"
];

// Script Constants
const ScriptDescription = "Scans site collection(s) for usage information.";
const ScriptFileName = "site_usage.csv";
const ScriptName = "Site Usage";

/**
 * Site Collection Usage
 * Displays a dialog to get the site collection usage information.
 */
class SiteUsage {
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
    private analyzeSites(sites: Types.SP.SiteOData[]) {
        // Return a promise
        return new Promise(resolve => {
            // Parse the webs
            Helper.Executor(sites, site => {
                // Calculate the site storage
                let siteStorage = site.Usage.Storage;
                siteStorage = siteStorage > 0 ? (siteStorage / Math.pow(1024, 3)).toFixed(Strings.FractionDigits) + " GB" as any : siteStorage;

                // Calculate the percent used
                let percentUsed = site.Usage.StoragePercentageUsed;
                percentUsed = percentUsed > 0 ? (percentUsed * 100).toFixed(Strings.FractionDigits) + "%" as any : percentUsed;

                // Add a row for this entry
                this._rows.push({
                    SiteBandwidth: site.Usage.Bandwidth,
                    SiteDescription: site.RootWeb.Description,
                    SiteDiscussionStorage: site.Usage.DiscussionStorage,
                    SiteHits: site.Usage.Hits,
                    SitePercentageUsed: percentUsed,
                    SiteStorage: siteStorage,
                    SiteTitle: site.RootWeb.Title,
                    SiteUrl: site.RootWeb.Url,
                    SiteVisits: site.Usage.Visits
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
                    label: "Site Collection Url(s)",
                    name: "Urls",
                    description: "Enter the relative site url(s) [Ex: /sites/dev]",
                    errorMessage: "Please enter a site collection url",
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
                    content: "Search for Site Usage",
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
                                let siteUrls: string[] = formValues["Urls"].match(/[^\n]+/g);

                                // Clear the data
                                this._errors = [];
                                this._rows = [];

                                // Parse the webs
                                Helper.Executor(siteUrls, siteUrl => {
                                    // Return a promise
                                    return new Promise((resolve) => {
                                        new Sites({
                                            url: siteUrl,
                                            onQuerySite: (odata) => {
                                                // Include the web description
                                                odata.Select.push("Usage");
                                            },
                                            onComplete: sites => {
                                                // Analyze the site
                                                this.analyzeSites(sites).then(resolve);
                                            },
                                            onError: () => {
                                                // Add the url to the errors list
                                                this._errors.push(siteUrl);
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
                // Order by the 2nd column by default; ascending
                order: [[1, "asc"]]
            },
            columns: [
                {
                    name: "SiteTitle",
                    title: "Title"
                },
                {
                    name: "SiteUrl",
                    title: "Url"
                },
                {
                    name: "SiteDescription",
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
                    name: "SiteStorage",
                    title: "Storage"
                },
                {
                    name: "SitePercentageUsed",
                    title: "Percentage Used"
                },
                {
                    name: "SiteDiscussionStorage",
                    title: "Discussion Storage"
                },
                {
                    name: "SiteHits",
                    title: "Hits"
                },
                {
                    name: "SiteVisits",
                    title: "Visits"
                },
                {
                    className: "text-end",
                    name: "",
                    title: "",
                    onRenderCell: (el, col, row: IRowInfo) => {
                        // Render the buttons
                        Components.TooltipGroup({
                            el,
                            tooltips: [
                                {
                                    content: "View Site",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconType: GetIcon(24, 24, "EntryView", "mx-1"),
                                        text: "View",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the security group
                                            window.open(row.SiteUrl, "_blank");
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
export const SiteUsageModal: IScript = {
    init: SiteUsage,
    name: ScriptName,
    description: ScriptDescription
};