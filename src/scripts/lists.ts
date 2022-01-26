import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { ExportCSV } from "./exportCSV";
import { Webs } from "./webs";

// Row Information
interface IRowInfo {
    ListDescription: string;
    ListId: string;
    ListItemCount: string;
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
                        ListItemCount: list.ItemCount + "",
                        ListName: list.Title,
                        ListType: this.getListType(list.BaseTemplate),
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

    // Returns the string value for a list type
    private getListType(template: number): string {
        let listType = "Unknown Type (" + template + ")";

        // Set the type
        switch (template) {
            case SPTypes.ListTemplateType.AccessApp:
                listType = "Access App";
                break;
            case SPTypes.ListTemplateType.AccessRequest:
                listType = "Access Request";
                break;
            case SPTypes.ListTemplateType.AdminTasks:
                listType = "Admin Tasks";
                break;
            case SPTypes.ListTemplateType.Agenda:
                listType = "Agenda";
                break;
            case SPTypes.ListTemplateType.AlchemyApprovalWorkflow:
                listType = "Alchemy Approval Workflow";
                break;
            case SPTypes.ListTemplateType.AlchemyMobileForm:
                listType = "Alchemy Mobile Form";
                break;
            case SPTypes.ListTemplateType.Announcements:
                listType = "Announcements";
                break;
            case SPTypes.ListTemplateType.AppCatalog:
                listType = "App Catalog";
                break;
            case SPTypes.ListTemplateType.AppDataCatalog:
                listType = "App Data Catalog";
                break;
            case SPTypes.ListTemplateType.AssetLibrary:
                listType = "Asset Library";
                break;
            case SPTypes.ListTemplateType.CallTrack:
                listType = "Call Track";
                break;
            case SPTypes.ListTemplateType.Categories:
                listType = "Categories";
                break;
            case SPTypes.ListTemplateType.Circulation:
                listType = "Circulation";
                break;
            case SPTypes.ListTemplateType.ClientSideAssets:
                listType = "Client Side Assets";
                break;
            case SPTypes.ListTemplateType.ClientSideComponentManifests:
                listType = "Client Side Component Manifests";
                break;
            case SPTypes.ListTemplateType.Comments:
                listType = "Comments";
                break;
            case SPTypes.ListTemplateType.Contacts:
                listType = "Contacts";
                break;
            case SPTypes.ListTemplateType.CustomGrid:
                listType = "Custom Grid";
                break;
            case SPTypes.ListTemplateType.DataConnectionLibrary:
                listType = "Data Connection Library";
                break;
            case SPTypes.ListTemplateType.DataSources:
                listType = "Data Sources";
                break;
            case SPTypes.ListTemplateType.Decision:
                listType = "Decision";
                break;
            case SPTypes.ListTemplateType.DesignCatalog:
                listType = "Design Catalog";
                break;
            case SPTypes.ListTemplateType.DeveloperSiteDraftApps:
                listType = "Developer Site Draft Apps";
                break;
            case SPTypes.ListTemplateType.DiscussionBoard:
                listType = "Discussion Board";
                break;
            case SPTypes.ListTemplateType.DocumentLibrary:
                listType = "Document Library";
                break;
            case SPTypes.ListTemplateType.Events:
                listType = "Events";
                break;
            case SPTypes.ListTemplateType.ExternalList:
                listType = "External List";
                break;
            case SPTypes.ListTemplateType.Facility:
                listType = "Facility";
                break;
            case SPTypes.ListTemplateType.GanttTasks:
                listType = "Gantt Tasks";
                break;
            case SPTypes.ListTemplateType.GenericList:
                listType = "Generic List";
                break;
            case SPTypes.ListTemplateType.HealthReports:
                listType = "Health Reports";
                break;
            case SPTypes.ListTemplateType.HealthRules:
                listType = "Health Rules";
                break;
            case SPTypes.ListTemplateType.HelpLibrary:
                listType = "Help Library";
                break;
            case SPTypes.ListTemplateType.Holidays:
                listType = "Holidays";
                break;
            case SPTypes.ListTemplateType.HomePageLibrary:
                listType = "Home Page Library";
                break;
            case SPTypes.ListTemplateType.IMEDic:
                listType = "IMEDic";
                break;
            case SPTypes.ListTemplateType.IssueTracking:
                listType = "Issue Tracking";
                break;
            case SPTypes.ListTemplateType.KPIStatusList:
                listType = "KPI Status List";
                break;
            case SPTypes.ListTemplateType.LanguagesAndTranslatorsList:
                listType = "Languages And Translators List";
                break;
            case SPTypes.ListTemplateType.Links:
                listType = "Links";
                break;
            case SPTypes.ListTemplateType.ListTemplateCatalog:
                listType = "List Template Catalog";
                break;
            case SPTypes.ListTemplateType.MaintenanceLogs:
                listType = "Maintenance Logs";
                break;
            case SPTypes.ListTemplateType.MasterPageCatalog:
                listType = "Master Page Catalog";
                break;
            case SPTypes.ListTemplateType.MeetingObjective:
                listType = "Meeting Objective";
                break;
            case SPTypes.ListTemplateType.Meetings:
                listType = "Meetings";
                break;
            case SPTypes.ListTemplateType.MeetingUser:
                listType = "Meeting User";
                break;
            case SPTypes.ListTemplateType.MicroFeed:
                listType = "Micro Feed";
                break;
            case SPTypes.ListTemplateType.MySiteDocumentLibrary:
                listType = "My Site Document Library";
                break;
            case SPTypes.ListTemplateType.NoCodePublic:
                listType = "No Code Public";
                break;
            case SPTypes.ListTemplateType.NoCodeWorkflows:
                listType = "No Code Workflows";
                break;
            case SPTypes.ListTemplateType.PageLibrary:
                listType = "Page Library";
                break;
            case SPTypes.ListTemplateType.PerformancePointContentList:
                listType = "Performance Point Content List";
                break;
            case SPTypes.ListTemplateType.PerformancePointDashboardsLibrary:
                listType = "Performance Point Dashboards Library";
                break;
            case SPTypes.ListTemplateType.PerformancePointDataConnectionsLibrary:
                listType = "Performance Point Data Connections Library";
                break;
            case SPTypes.ListTemplateType.PerformancePointDataSourceLibrary:
                listType = "Performance Point Data Source Library";
                break;
            case SPTypes.ListTemplateType.PersonalDocumentLibrary:
                listType = "Personal Document Library";
                break;
            case SPTypes.ListTemplateType.PictureLibrary:
                listType = "Picture Library";
                break;
            case SPTypes.ListTemplateType.Posts:
                listType = "Posts";
                break;
            case SPTypes.ListTemplateType.PrivateDocumentLibrary:
                listType = "Private Document Library";
                break;
            case SPTypes.ListTemplateType.RecordLibrary:
                listType = "Record Library";
                break;
            case SPTypes.ListTemplateType.ReportLibrary:
                listType = "Report Library";
                break;
            case SPTypes.ListTemplateType.SharingLinks:
                listType = "Sharing Links";
                break;
            case SPTypes.ListTemplateType.SolutionCatalog:
                listType = "Solution Catalog";
                break;
            case SPTypes.ListTemplateType.Survey:
                listType = "Survey";
                break;
            case SPTypes.ListTemplateType.Tasks:
                listType = "Tasks";
                break;
            case SPTypes.ListTemplateType.TasksWithTimelineAndHierarchy:
                listType = "Tasks With Timeline And Hierarchy";
                break;
            case SPTypes.ListTemplateType.TenantWideExtensions:
                listType = "Tenant Wide Extensions";
                break;
            case SPTypes.ListTemplateType.TextBox:
                listType = "Text Box";
                break;
            case SPTypes.ListTemplateType.ThemeCatalog:
                listType = "Theme Catalog";
                break;
            case SPTypes.ListTemplateType.ThingsToBring:
                listType = "Things To Bring";
                break;
            case SPTypes.ListTemplateType.Timecard:
                listType = "Timecard";
                break;
            case SPTypes.ListTemplateType.TranslationManagementLibrary:
                listType = "Translation Management Library";
                break;
            case SPTypes.ListTemplateType.UserInformation:
                listType = "User Information";
                break;
            case SPTypes.ListTemplateType.VisioProcessDiagramMetricLibrary:
                listType = "Visio Process Diagram Metric Library";
                break;
            case SPTypes.ListTemplateType.VisioProcessDiagramUSUnitsLibrary:
                listType = "Visio Process Diagram US Units Library";
                break;
            case SPTypes.ListTemplateType.WebPageLibrary:
                listType = "Web Page Library";
                break;
            case SPTypes.ListTemplateType.WebPartCatalog:
                listType = "Web Part Catalog";
                break;
            case SPTypes.ListTemplateType.WebTemplateCatalog:
                listType = "Web Template Catalog";
                break;
            case SPTypes.ListTemplateType.Whereabouts:
                listType = "Whereabouts";
                break;
            case SPTypes.ListTemplateType.WorkflowHistory:
                listType = "Workflow History";
                break;
            case SPTypes.ListTemplateType.WorkflowProcess:
                listType = "Workflow Process";
                break;
            case SPTypes.ListTemplateType.XMLForm:
                listType = "XML Form";
                break;
        }

        return listType;
    }

    // Renders the dialog
    private render() {
        // Set the type
        Modal.setType(Components.ModalTypes.Large);

        // Set the header
        Modal.setHeader("List Information");

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
                                            odata.Select.push("Lists/BaseTemplate");
                                            odata.Select.push("Lists/Description");
                                            odata.Select.push("Lists/Id");
                                            odata.Select.push("Lists/ItemCount");
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
                    name: "ListItemCount",
                    title: "List Item Count"
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
                        new ExportCSV("list_information.csv", [
                            "ListDescription", "ListId", "ListItemCount", "ListName",
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