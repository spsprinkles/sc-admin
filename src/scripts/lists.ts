import { CanvasForm, DataTable, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, SPTypes, ThemeManager, Types, Web } from "gd-sprest-bs";
import { search } from "gd-sprest-bs/build/icons/svgs/search";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
import { CopyList, ExportCSV, GetIcon, IScript, Webs } from "../common";
import Strings from "../strings";

// List Information
export interface IListInfo {
    ListDescription: string;
    ListExperience: string;
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

// List Settings
interface IListSettings {
    Description: string;
    ListExperienceOptions: string;
    Title: string;
}

// CSV Export Fields
const CSVExportFields = [
    "ListDescription", "ListId", "ListItemCount", "ListViewCount", "ListViewHiddenCount",
    "ListName", "ListType", "ListExperience", "ListUrl", "WebTitle", "WebUrl"
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
    private _rows: IListInfo[] = null;
    private _urls: string[] = null;

    // Constructor
    constructor(urls: string[] = Strings.SiteUrls) {
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

                    // Set the list experience
                    let listExperience = "";
                    switch (list.ListExperienceOptions) {
                        case SPTypes.ListExperienceOptions.Auto:
                            listExperience = "Default experience for the site"
                            break;
                        case SPTypes.ListExperienceOptions.ClassicExperience:
                            listExperience = "Classic experience"
                            break;
                        case SPTypes.ListExperienceOptions.NewExperience:
                            listExperience = "New experience"
                            break;
                    }

                    // Add a row for this entry
                    this._rows.push({
                        ListDescription: list.Description,
                        ListExperience: listExperience,
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
        Modal.setFooter(Components.TooltipGroup({
            tooltips: [
                {
                    content: "Search for List Information",
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
                                Strings.SiteUrls = formValues["Urls"].match(/[^\n]+/g);

                                // Clear the data
                                this._errors = [];
                                this._rows = [];

                                // Parse the webs
                                Helper.Executor(Strings.SiteUrls, webUrl => {
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
                                                odata.Select.push("Lists/ListExperienceOptions");
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
        Modal.setBody(elTable);
        new DataTable({
            el: elTable,
            rows: this._rows,
            onRendering: dtProps => {
                dtProps.columnDefs = [
                    {
                        "targets": 10,
                        "orderable": false,
                        "searchable": false
                    }
                ];

                // Order by the 6th column by default; ascending
                dtProps.order = [[5, "asc"]];

                // Return the properties
                return dtProps;
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
                    name: "ListExperience",
                    title: "List Experience"
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
                    onRenderCell: (el, col, row: IListInfo) => {
                        let btnDelete: Components.IButton = null;

                        // Render the buttons
                        Components.TooltipGroup({
                            el,
                            tooltips: [
                                {
                                    content: "View List",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconType: GetIcon(24, 24, "CustomList", "mx-1"),
                                        text: "View",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the security group
                                            window.open(row.ListUrl, "_blank");
                                        }
                                    }
                                },
                                {
                                    content: "Delete List",
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
                                            if (confirm("Are you sure you want to delete this list?")) {
                                                // Disable this button
                                                btnDelete.disable();

                                                // Delete the site group
                                                this.deleteList(row.WebUrl, row.ListName);
                                            }
                                        }
                                    }
                                },
                                {
                                    content: "Copy List",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconType: GetIcon(24, 24, "CopyAdd", "mx-1"),
                                        text: "Copy",
                                        type: Components.ButtonTypes.OutlineSuccess,
                                        onClick: () => {
                                            // Display the copy form
                                            this.showCopyListForm(row);
                                        }
                                    }
                                },
                                {
                                    content: "Settings",
                                    btnProps: {
                                        iconType: GetIcon(24, 24, "TextBulletListSquare", "mx-1"),
                                        text: "Settings",
                                        type: Components.ButtonTypes.OutlineSecondary,
                                        onClick: () => {
                                            // Display the list settings form
                                            this.showListSettingsForm(row);
                                        }
                                    }
                                }
                            ]
                        });
                    }
                }
            ]
        });

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

    // Shows the copy list form
    private showCopyListForm(listInfo: IListInfo) {
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
                    required: true,
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

        // Create the results element
        let elResults = document.createElement("div");
        elResults.id = "results";
        elResults.classList.add("mt-3");
        CanvasForm.BodyElement.appendChild(elResults);

        // Add the disclaimer about not copying the content
        let disclaimer = document.createElement("p");
        disclaimer.className = "fst-italic mb-0 mt-3 small";
        disclaimer.innerHTML = "* The new list that is created will not contain any content. Only the structure of the list is preserved. Lookup list associations are preserved if the target lookup list exists.";
        CanvasForm.BodyElement.appendChild(disclaimer);

        // Set the footer
        Components.TooltipGroup({
            el: CanvasForm.BodyElement,
            className: "float-end mt-3",
            tooltips: [
                {
                    content: "Start the list copy",
                    btnProps: {
                        className: "pe-2 py-1",
                        iconType: GetIcon(24, 24, "CopyArrowRight", "mx-1"),
                        text: "Copy",
                        type: Components.ButtonTypes.OutlineSuccess,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let formValues = form.getValues();
                                let dstListName = formValues["ListName"];
                                let dstWebUrl = formValues["WebUrl"];

                                // Copy the list
                                CopyList.createListConfiguration(elResults, listInfo, dstWebUrl, dstListName).then(
                                    // The list configuration as a JSON string
                                    cfg => {
                                        // Ensure the control is valid
                                        let ctrl = form.getControl("WebUrl");
                                        ctrl.updateValidation(ctrl.el, {
                                            isValid: true,
                                            validMessage: "The list was copied successfully."
                                        });
                                    },
                                    // Error
                                    err => {
                                        if (typeof (err) === "string") {
                                            // Show the error message
                                            let ctrl = form.getControl("WebUrl");
                                            ctrl.updateValidation(ctrl.el, {
                                                isValid: false,
                                                invalidMessage: err
                                            });

                                            // Hide the loading dialog
                                            LoadingDialog.hide();
                                        }
                                    }
                                );
                            }
                        }
                    }
                }
            ]
        });

        // Show the canvas form
        CanvasForm.show();
    }

    // Shows the list settings form
    private showListSettingsForm(listInfo: IListInfo) {
        // Clear the canvas
        CanvasForm.clear();

        // Prevent auto close
        CanvasForm.setAutoClose(false);

        // Set the header
        CanvasForm.setHeader(listInfo.ListName + " List Settings");

        // Render an alert
        let elAlert = document.createElement("div");
        CanvasForm.BodyElement.appendChild(elAlert);

        // Set the body
        let form = Components.Form({
            el: CanvasForm.BodyElement,
            controls: [
                {
                    label: "The title of the list.",
                    name: "ListTitle",
                    className: "mb-3",
                    errorMessage: "A title is required",
                    required: true,
                    type: Components.FormControlTypes.TextField,
                    value: listInfo.ListName
                },
                {
                    label: "The description of the list.",
                    name: "ListDescription",
                    className: "mb-3",
                    type: Components.FormControlTypes.TextArea,
                    value: listInfo.ListDescription
                },
                {
                    label: "Display this list using the new or classic experience?",
                    name: "ListExperience",
                    className: "mb-3",
                    errorMessage: "A selection a required",
                    required: true,
                    type: Components.FormControlTypes.Switch,
                    value: listInfo.ListExperience,
                    items: [
                        {
                            data: SPTypes.ListExperienceOptions.Auto,
                            name: "Auto",
                            label: "Default experience for the site"
                        },
                        {
                            data: SPTypes.ListExperienceOptions.NewExperience,
                            name: "NewExperience",
                            label: "New experience"
                        },
                        {
                            data: SPTypes.ListExperienceOptions.ClassicExperience,
                            name: "ClassicExperience",
                            label: "Classic experience"
                        }
                    ],
                    onControlRendered: (ctrl) => {
                        // Add the dark class if theme is inverted
                        if (ThemeManager.IsInverted) {
                            ctrl.el.querySelectorAll("div.form-check.form-switch input[type=checkbox].form-check-input").forEach((el: HTMLElement) => {
                                el.classList.add("dark");
                            });
                        }
                    }
                } as Components.IFormControlPropsSwitch
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: CanvasForm.BodyElement,
            className: "float-end mt-3",
            tooltips: [
                {
                    content: "Updates the list settings",
                    btnProps: {
                        className: "pe-2 py-1",
                        iconType: GetIcon(24, 24, "TextBulletListSquareEdit", "mx-1"),
                        text: "Update",
                        type: Components.ButtonTypes.OutlineSuccess,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let formValues = form.getValues();

                                // Clear the alert
                                while (elAlert.firstChild) { elAlert.removeChild(elAlert.firstChild); }

                                // Update the list
                                this.updateListSettings(listInfo, {
                                    Description: formValues["ListDescription"],
                                    ListExperienceOptions: (formValues["ListExperience"] as Components.ICheckboxGroupItem).data,
                                    Title: formValues["ListTitle"]
                                }).then(
                                    // Success
                                    () => {
                                        // Show an alert
                                        Components.Alert({
                                            el: elAlert,
                                            content: "The list was successfully updated.",
                                            type: Components.AlertTypes.Primary
                                        });
                                    },

                                    // Error
                                    msg => {
                                        // Show an alert
                                        Components.Alert({
                                            el: elAlert,
                                            content: msg,
                                            type: Components.AlertTypes.Danger
                                        });
                                    }
                                );
                            }
                        }
                    }
                }
            ]
        });

        // Show the canvas form
        CanvasForm.show();
    }

    // Updates the list settings
    private updateListSettings(listInfo: IListInfo, values: IListSettings): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Show a loading Dialog
            LoadingDialog.setHeader("Reading Target List");
            LoadingDialog.setBody("Validating the permissions...");
            LoadingDialog.show();

            // Get the target list permissions
            Web(listInfo.WebUrl).Lists().getById(listInfo.ListId).query({ Expand: ["EffectiveBasePermissions"] }).execute(
                // Exists
                (list) => {
                    // Ensure the user doesn't have permission to manage lists
                    if (!Helper.hasPermissions(list.EffectiveBasePermissions, [SPTypes.BasePermissionTypes.ManageLists])) {
                        // Hide the loading dialog
                        LoadingDialog.hide();

                        // Reject the request
                        reject("You do not have permission to manage lists on this web.");
                        return;
                    }

                    // Get the digest value for the destination web
                    ContextInfo.getWeb(listInfo.WebUrl).execute(context => {
                        // Update the loading dialog
                        LoadingDialog.setHeader("Updating the List");
                        LoadingDialog.setBody("Updating the list experience...");

                        // Update the list experience value
                        Web(listInfo.WebUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue })
                            .Lists().getById(listInfo.ListId).update(values).execute(
                                // Success
                                () => {
                                    // Hide the loading dialog
                                    LoadingDialog.hide();

                                    // Resolve the request
                                    resolve();
                                },

                                // Error
                                () => {
                                    // Hide the loading dialog
                                    LoadingDialog.hide();

                                    // Reject the request
                                    reject("Error updating the list");
                                }
                            )
                    });
                },

                // Doesn't exist
                () => {
                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Reject the request
                    reject("The target web doesn't exist, or you do not have access to it.");
                }
            );
        });
    }
}

// Script Information
export const ListInfoModal: IScript = {
    init: ListInfo,
    name: ScriptName,
    description: ScriptDescription
};