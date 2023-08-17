import { CanvasForm, DataTable, List, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { ExportCSV, Webs, IScript } from "../common";

// Row Information
interface IRowInfo {
    ListDescription: string;
    ListId: string;
    ListItemCount: string;
    ListName: string;
    ListType: string;
    ListUrl: string;
    ListViewCount: number;
    ListViewHiddenCount: number;
    WebTitle: string;
    WebUrl: string;
}

// CSV Export Fields
const CSVExportFields = [
    "ListDescription", "ListId", "ListItemCount", "ListViewCount",
    "ListViewHiddenCount", "ListName", "ListType", "ListUrl", "WebTitle", "WebUrl"
];

// Script Constants
const ScriptDescription = "Scan site(s) for list & library information. Ability to copy a list structure to another web.";
const ScriptFileName = "list_information.csv";
const ScriptName = "List Information";

/**
 * List Information
 * Displays a dialog to get the site information.
 */
class ListInfo {
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
                    // Parse the list views
                    let hiddenCount = 0;
                    let views = list.Views as any as Types.SP.IViewCollection;
                    for (let i = 0; i < views.results.length; i++) {
                        if (views.results[i].Hidden) { hiddenCount++; }
                    }

                    // Add a row for this entry
                    this._rows.push({
                        ListDescription: list.Description,
                        ListId: list.Id,
                        ListItemCount: list.ItemCount + "",
                        ListName: list.Title,
                        ListType: this.getListType(list.BaseTemplate),
                        ListUrl: (list.RootFolder as any as Types.SP.Folder).ServerRelativeUrl,
                        ListViewCount: views.results.length,
                        ListViewHiddenCount: hiddenCount,
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

    // Copies a list
    private copyList(listInfo: IRowInfo, form: Components.IForm): PromiseLike<string> {
        return new Promise(resolve => {
            // Show a loading Dialog
            LoadingDialog.setHeader("Reading Target Web");
            LoadingDialog.setBody("Getting the destination web information...");
            LoadingDialog.show();

            // Get the target web
            let formValues = form.getValues();
            let url = formValues["WebUrl"];
            Web(url).query({ Expand: ["EffectiveBasePermissions"] }).execute(
                // Exists
                (web) => {
                    // Ensure the user doesn't have permission to manage lists
                    if (!Helper.hasPermissions(web.EffectiveBasePermissions, [SPTypes.BasePermissionTypes.ManageLists])) {
                        // Update the validation
                        let ctrl = form.getControl("WebUrl");
                        ctrl.updateValidation(ctrl.el, {
                            isValid: false,
                            invalidMessage: "You do not have permission to create lists on this web."
                        });

                        // Hide the loading dialog
                        LoadingDialog.hide();
                        return;
                    }

                    // Get the list information
                    var list = new List({
                        listName: formValues["ListName"],
                        webUrl: listInfo.WebUrl,
                        itemQuery: { Filter: "Id eq 0" },
                        onInitError: () => {
                            // Update the validation
                            let ctrl = form.getControl("WebUrl");
                            ctrl.updateValidation(ctrl.el, {
                                isValid: false,
                                invalidMessage: "Error loading the list information. Please check your permission to this list."
                            });

                            // Hide the loading dialog
                            LoadingDialog.hide();
                        },
                        onInitialized: () => {
                            // Update the loading dialog
                            LoadingDialog.setHeader("Analyzing the List");
                            LoadingDialog.setBody("Getting the list information...");

                            // Create the configuration
                            let cfgProps: Helper.ISPConfigProps = {
                                ListCfg: [{
                                    ListInformation: {
                                        BaseTemplate: list.ListInfo.BaseTemplate,
                                        Title: list.ListInfo.Title,
                                        AllowContentTypes: list.ListInfo.AllowContentTypes,
                                        Hidden: list.ListInfo.Hidden,
                                        NoCrawl: list.ListInfo.NoCrawl
                                    },
                                    CustomFields: [],
                                    ViewInformation: []
                                }]
                            };

                            // Parse the content type fields
                            let lookupFields: Types.SP.FieldLookup[] = [];
                            for (let i = 0; i < list.ListContentTypes[0].Fields.results.length; i++) {
                                let fldInfo = list.ListContentTypes[0].Fields.results[i];

                                // Skip internal fields
                                if (fldInfo.InternalName == "ContentType" || fldInfo.InternalName == "Title") { continue; }

                                // See if this is a lookup field
                                if (fldInfo.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                    // Add the field
                                    lookupFields.push(fldInfo);
                                }

                                // Add the field information
                                cfgProps.ListCfg[0].CustomFields.push({
                                    name: fldInfo.InternalName,
                                    schemaXml: fldInfo.SchemaXml
                                });
                            }

                            // Parse the views
                            for (let i = 0; i < list.ListViews.length; i++) {
                                let viewInfo = list.ListViews[i];

                                // Add the view
                                cfgProps.ListCfg[0].ViewInformation.push({
                                    Default: true,
                                    ViewName: viewInfo.Title,
                                    ViewFields: viewInfo.ViewFields.Items.results,
                                    ViewQuery: viewInfo.ViewQuery
                                });
                            }

                            // Update the loading dialog
                            LoadingDialog.setHeader("Creating the List");
                            LoadingDialog.setBody("Creating the destination list...");

                            // Create the list
                            let cfg = Helper.SPConfig(cfgProps);
                            cfg.setWebUrl(web.ServerRelativeUrl);
                            cfg.install().then(() => {
                                // Update the loading dialog
                                LoadingDialog.setHeader("Reading Destination Web");
                                LoadingDialog.setBody("Getting the context information of the destination web...");

                                // Get the digest value for the destination web
                                ContextInfo.getWeb(web.ServerRelativeUrl).execute(context => {
                                    // Update the loading dialog
                                    LoadingDialog.setHeader("Updating the List");
                                    LoadingDialog.setBody("Updating the lookup field(s)...");

                                    // Parse the lookup fields
                                    Helper.Executor(lookupFields, lookupField => {
                                        // Return a promise
                                        return new Promise(resolve => {
                                            // Get the lookup field source list
                                            Web(listInfo.WebUrl).Lists().getById(lookupField.LookupList).execute(srcList => {
                                                // Get the lookup list in the destination site
                                                Web(web.ServerRelativeUrl).Lists(srcList.Title).execute(dstList => {
                                                    // Update the field schema xml
                                                    let fieldDef = lookupField.SchemaXml.replace(`List="${lookupField.LookupList}"`, `List="{${dstList.Id}}"`);
                                                    Web(web.ServerRelativeUrl, {
                                                        requestDigest: context.GetContextWebInformation.FormDigestValue
                                                    }).Lists(list.ListInfo.Title).Fields(lookupField.InternalName).update({
                                                        SchemaXml: fieldDef
                                                    }).execute(() => {
                                                        // Updated the lookup list
                                                        console.log(`Updated the lookup field '${lookupField.InternalName}' in lookup list successfully.`);

                                                        // Check the next field
                                                        resolve(null);
                                                    })
                                                }, () => {
                                                    // Error getting the lookup list
                                                    console.error(`Error getting the lookup list '${lookupField.LookupList}' from web '${web.ServerRelativeUrl}'.`);

                                                    // Check the next field
                                                    resolve(null);
                                                });
                                            }, () => {
                                                // Error getting the lookup list
                                                console.error(`Error getting the lookup list '${lookupField.LookupList}' from web '${listInfo.WebUrl}'.`);

                                                // Check the next field
                                                resolve(null);
                                            });
                                        });
                                    }).then(() => {
                                        // Update the validation
                                        let ctrl = form.getControl("WebUrl");
                                        ctrl.updateValidation(ctrl.el, {
                                            isValid: true,
                                            validMessage: "The list was copied successfully to the destination url."
                                        });

                                        // Get the new list's url
                                        Web(listInfo.WebUrl).Lists(listInfo.ListName).RootFolder().execute(folder => {
                                            // Hide the loading dialog
                                            LoadingDialog.hide();

                                            // Resolve the request
                                            resolve(folder.ServerRelativeUrl)
                                        });
                                    });
                                });
                            });
                        }
                    });
                },

                // Doesn't exist
                () => {
                    // Update the validation
                    let ctrl = form.getControl("WebUrl");
                    ctrl.updateValidation(ctrl.el, {
                        isValid: false,
                        invalidMessage: "The target web doesn't exist, or you do not have access to it."
                    });

                    // Hide the loading dialog
                    LoadingDialog.hide();
                }
            );
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
        Modal.setFooter(Components.ButtonGroup({
            buttons: [
                {
                    text: "Analyze",
                    type: Components.ButtonTypes.OutlineSuccess,
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
                                        onQueryWeb: (odata) => {
                                            // Include the list information
                                            odata.Expand.push("Lists/RootFolder");
                                            odata.Expand.push("Lists/Views");
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
                        "targets": 9,
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
                // Order by the 5th column by default; ascending
                order: [[4, "asc"]]
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
                    name: "ListType",
                    title: "List Type"
                },
                {
                    name: "ListUrl",
                    title: "List Path"
                },
                {
                    name: "ListDescription",
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
                    name: "ListItemCount",
                    title: "Item Count"
                },
                {
                    name: "ListViewCount",
                    title: "Total View Count"
                },
                {
                    name: "ListViewHiddenCount",
                    title: "Hidden View Count"
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
                                },
                                {
                                    text: "Copy",
                                    type: Components.ButtonTypes.OutlineSuccess,
                                    onClick: () => {
                                        // Display the copy form
                                        this.showCopyListForm(row);
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

    // Shows the copy list form
    private showCopyListForm(listInfo: IRowInfo) {
        // Clear the canvas
        CanvasForm.clear();

        // Prevent auto close
        CanvasForm.setAutoClose(false);

        // Set the header
        CanvasForm.setHeader("Copy List");

        // Set the body
        let form = Components.Form({
            el: CanvasForm.BodyElement,
            controls: [
                {
                    label: "Source List",
                    name: "SourceList",
                    className: "mb-3",
                    description: "The url of the source list",
                    type: Components.FormControlTypes.Readonly,
                    value: listInfo.ListUrl
                },
                {
                    label: "List Name",
                    name: "ListName",
                    className: "mb-3",
                    type: Components.FormControlTypes.TextField,
                    value: listInfo.ListName,
                    description: "The list name to create",
                    errorMessage: "A list name is required"
                },
                {
                    label: "Web Url",
                    name: "WebUrl",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    description: "The destination url of the site to copy the list to",
                    errorMessage: "The destination url is required"
                }
            ]
        });

        // Set the footer
        let btnView: Components.IButton = null;
        let newListUrl: string = null;
        Components.TooltipGroup({
            el: CanvasForm.BodyElement,
            className: "float-end mt-3",
            tooltips: [
                {
                    content: "Click to copy the list",
                    btnProps: {
                        text: "Copy",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Disable the view button
                            btnView.disable();

                            // Ensure the form is valid
                            if (form.isValid()) {
                                // Copy the list
                                this.copyList(listInfo, form).then(url => {
                                    // Set the url
                                    newListUrl = url;

                                    // Enable the button
                                    btnView.enable();
                                });
                            }
                        }
                    }
                },
                {
                    content: "View the new list",
                    btnProps: {
                        assignTo: btn => {
                            btnView = btn;
                        },
                        text: "View",
                        type: Components.ButtonTypes.OutlineSuccess,
                        isDisabled: true,
                        onClick: () => {
                            // Show the list in a new tab
                            newListUrl ? window.open(newListUrl, "_blank") : null;
                        }
                    }
                }
            ]
        });

        // Show the canvas form
        CanvasForm.show();
    }
}

// Script Information
export const ListInfoModal: IScript = {
    init: ListInfo,
    name: ScriptName,
    description: ScriptDescription
};