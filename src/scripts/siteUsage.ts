import { DataTable, Modal } from "dattatable";
import { Components, Helper, Types } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { ExportCSV, Sites, IScript } from "../common";

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
                siteStorage = siteStorage > 0 ? (siteStorage / Math.pow(1024, 3)).toFixed(4) + " GB" as any : siteStorage;

                // Calculate the percent used
                let percentUsed = site.Usage.StoragePercentageUsed;
                percentUsed = percentUsed > 0 ? (percentUsed * 100).toFixed(4) + "%" as any : percentUsed;

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

        // Set the header
        Modal.setHeader("Site Usage");

        // Render the form
        let form = Components.Form({
            controls: [
                {
                    label: "Site Collection Url(s)",
                    name: "Urls",
                    description: "Enter the relative site url(s). (Ex: /sites/dev)",
                    errorMessage: "Please enter a site collection url.",
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
                            let siteUrls: string[] = formValues["Urls"].split('\n');

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
        Modal.setHeader("Site Usage");

        // Render the table
        let elTable = document.createElement("div");
        new DataTable({
            el: elTable,
            rows: this._rows,
            dtProps: {
                dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                columnDefs: [
                    {
                        "targets": [3],
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
                    name: "SiteTitle",
                    title: "Title"
                },
                {
                    name: "SiteUrl",
                    title: "Url"
                },
                {
                    name: "SiteDescription",
                    title: "Description"
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
                    name: "",
                    title: "",
                    onRenderCell: (el, col, row: IRowInfo) => {
                        // Render the buttons
                        Components.ButtonGroup({
                            el,
                            buttons: [
                                {
                                    text: "View",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Show the security group
                                        window.open(row.SiteUrl, "_blank");
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
                        new ExportCSV("site_usage.csv", CSVExportFields, this._rows);
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
export const SiteUsageModal: IScript = {
    init: SiteUsage,
    name: "Site Usage",
    description: "Scans for site collection(s) for usage information."
};