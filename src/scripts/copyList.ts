import { CanvasForm, List, LoadingDialog } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { IListInfo } from "./lists";

/**
 * Copy List
 */
export class CopyList {
    // Method to create the list configuration
    static createListConfiguration(elResults: HTMLElement, elLog: HTMLElement, srcList: IListInfo, dstWebUrl: string, dstListName: string): PromiseLike<string> {
        // Show a loading dialog
        LoadingDialog.setHeader("Copying the List");
        LoadingDialog.setBody("Initializing the request...");
        LoadingDialog.show();

        // Return a promise
        return new Promise((resolve, reject) => {
            // Ensure the user has the correct permissions to create the list
            Web(dstWebUrl).query({ Expand: ["EffectiveBasePermissions"] }).execute(
                // Exists
                web => {
                    // Ensure the user doesn't have permission to manage lists
                    if (!Helper.hasPermissions(web.EffectiveBasePermissions, [SPTypes.BasePermissionTypes.ManageLists])) {
                        // Resolve the request
                        console.error("You do not have permission to copy templates on this web.");
                        LoadingDialog.hide();
                        resolve(null);
                        return;
                    }

                    // Generate the list configuration
                    this.generateListConfiguration(srcList.WebUrl, srcList, dstListName).then(listCfg => {
                        // Save a copy of the configuration
                        let strConfig = JSON.stringify(listCfg.cfg);

                        // Validate the lookup fields
                        this.validateLookups(srcList.WebUrl, web.ServerRelativeUrl, srcList, listCfg.lookupFields).then(() => {
                            // Test the configuration
                            this.installConfiguration(listCfg.cfg, web.ServerRelativeUrl, elLog).then(lists => {
                                // Show the results
                                this.renderResults(elResults, listCfg.cfg, web.ServerRelativeUrl, lists, true);

                                // Resolve the request
                                resolve(strConfig);
                            }, reject);
                        }, reject);
                    }, reject);
                },

                // Doesn't exist
                () => {
                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    console.error("Error getting the target web. You may not have permissions or it doesn't exist.");
                    resolve(null);
                }
            );
        });
    }

    // Method to create the lists
    private static createLists(cfgProps: Helper.ISPConfigProps, webUrl: string, elLog: HTMLElement): PromiseLike<List[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Set the log event
            cfgProps.onLogMessage = msg => {
                // Append the message
                let elMessage = document.createElement("p");
                elMessage.innerHTML = msg;
                elLog.appendChild(elMessage);

                // Focus on the message
                elMessage.focus();
                elMessage.scrollIntoView();
            };

            // Clear the log
            while (elLog.firstChild) { elLog.removeChild(elLog.firstChild); }

            // Show the log
            elLog.classList.remove("d-none");

            // Create the configuration
            let cfg = Helper.SPConfig(cfgProps, webUrl);

            // Update the loading dialog
            LoadingDialog.setBody("Creating the list(s)...");

            // Install the solution
            cfg.install().then(() => {
                let lists: List[] = [];

                // Update the loading dialog
                LoadingDialog.setBody("Validating the list(s)...");

                // Parse the lists
                Helper.Executor(cfgProps.ListCfg, listCfg => {
                    // Return a promise
                    return new Promise(resolve => {
                        // Test the list
                        this.testList(listCfg, webUrl).then(list => {
                            // Append the list
                            lists.push(list);

                            // Check the next list
                            resolve(null);
                        });
                    });
                }).then(() => {
                    // Hide the log
                    elLog.classList.add("d-none");

                    // Resolve the request
                    resolve(lists);
                });
            }, reject);
        });
    }

    // Generates the list configuration
    private static generateListConfiguration(srcWebUrl: string, srcList: IListInfo, dstListName?: string): PromiseLike<{ cfg: Helper.ISPConfigProps, lookupFields: Types.SP.FieldLookup[] }> {
        // Getting the source list information
        LoadingDialog.setBody("Loading source list...");

        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the list information
            var list = new List({
                listName: srcList.ListName,
                webUrl: srcWebUrl,
                itemQuery: { Filter: "Id eq 0" },
                onInitError: () => {
                    // Reject the request
                    reject("Error loading the list information. Please check your permission to the source list.");
                },
                onInitialized: () => {
                    let calcFields: Types.SP.Field[] = [];
                    let fields: { [key: string]: boolean } = {};
                    let lookupFields: Types.SP.FieldLookup[] = [];

                    // Update the loading dialog
                    LoadingDialog.setBody("Analyzing the list information...");

                    // Create the configuration
                    let cfgProps: Helper.ISPConfigProps = {
                        ContentTypes: [],
                        ListCfg: [{
                            ListInformation: {
                                AllowContentTypes: list.ListInfo.AllowContentTypes,
                                BaseTemplate: list.ListInfo.BaseTemplate,
                                ContentTypesEnabled: list.ListInfo.ContentTypesEnabled,
                                Title: dstListName || srcList.ListName,
                                Hidden: list.ListInfo.Hidden,
                                NoCrawl: list.ListInfo.NoCrawl
                            },
                            ContentTypes: [],
                            CustomFields: [],
                            ViewInformation: []
                        }]
                    };

                    // Parse the content types
                    for (let i = 0; i < list.ListContentTypes.length; i++) {
                        let ct = list.ListContentTypes[i];

                        // Skip sealed content types
                        if (ct.Sealed) { continue; }

                        // Skip the internal content types
                        if (ct.Name != "Document" && ct.Name != "Event" && ct.Name != "Item" && ct.Name != "Task") {
                            // Add the content type
                            cfgProps.ContentTypes.push({
                                Name: ct.Name,
                                ParentName: "Item"
                            });
                        }

                        // Parse the content type fields
                        let fieldRefs = [];
                        for (let j = 0; j < ct.FieldLinks.results.length; j++) {
                            let fldInfo: Types.SP.Field = list.getField(ct.FieldLinks.results[j].Name);

                            // See if this is a lookup field
                            if (fldInfo.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                // Ensure this isn't an associated lookup field
                                if ((fldInfo as Types.SP.FieldLookup).IsDependentLookup != true) {
                                    // Append the field ref
                                    fieldRefs.push(fldInfo.InternalName);
                                }
                            } else {
                                // Append the field ref
                                fieldRefs.push(fldInfo.InternalName);
                            }

                            // Skip internal fields
                            if (fldInfo.InternalName == "ContentType" || fldInfo.InternalName == "Title") { continue; }

                            // See if this is a calculated field
                            if (fldInfo.FieldTypeKind == SPTypes.FieldType.Calculated) {
                                // Add the field and continue the loop
                                calcFields.push(fldInfo);
                            }
                            // Else, see if this is a lookup field
                            else if (fldInfo.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                // Add the field
                                lookupFields.push(fldInfo);
                            }
                            // Else, ensure the field hasn't been added
                            else if (fields[fldInfo.InternalName] == null) {
                                // Add the field information
                                fields[fldInfo.InternalName] = true;
                                cfgProps.ListCfg[0].CustomFields.push({
                                    name: fldInfo.InternalName,
                                    schemaXml: fldInfo.SchemaXml
                                });
                            }
                        }

                        // Add the list content type
                        cfgProps.ListCfg[0].ContentTypes.push({
                            Name: ct.Name,
                            Description: ct.Description,
                            ParentName: ct.Name,
                            FieldRefs: fieldRefs
                        });
                    }

                    // Parse the views
                    for (let i = 0; i < list.ListViews.length; i++) {
                        let viewInfo = list.ListViews[i];

                        // Parse the fields
                        for (let j = 0; j < viewInfo.ViewFields.Items.results.length; j++) {
                            let field = list.getField(viewInfo.ViewFields.Items.results[j]);

                            // Ensure the field exists
                            if (fields[field.InternalName] == null) {
                                // See if this is a calculated field
                                if (field.FieldTypeKind == SPTypes.FieldType.Calculated) {
                                    // Add the field and continue the loop
                                    calcFields.push(field);
                                }
                                // Else, see if this is a lookup field
                                else if (field.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                    // Add the field
                                    lookupFields.push(field);
                                } else {
                                    // Append the field
                                    fields[field.InternalName] = true;
                                    cfgProps.ListCfg[0].CustomFields.push({
                                        name: field.InternalName,
                                        schemaXml: field.SchemaXml
                                    });
                                }
                            }
                        }

                        // Add the view
                        cfgProps.ListCfg[0].ViewInformation.push({
                            Default: viewInfo.DefaultView,
                            Hidden: viewInfo.Hidden,
                            JSLink: viewInfo.JSLink,
                            MobileDefaultView: viewInfo.MobileDefaultView,
                            MobileView: viewInfo.MobileView,
                            RowLimit: viewInfo.RowLimit,
                            Tabular: viewInfo.TabularView,
                            ViewName: viewInfo.Title,
                            ViewFields: viewInfo.ViewFields.Items.results,
                            ViewQuery: viewInfo.ViewQuery
                        });
                    }

                    // Update the loading dialog
                    LoadingDialog.setBody("Analyzing the lookup fields...");

                    // Parse the lookup fields
                    Helper.Executor(lookupFields, lookupField => {
                        // Skip the field, if it was already added
                        if (fields[lookupField.InternalName]) { return; }

                        // Return a promise
                        return new Promise((resolve) => {
                            // Get the lookup list
                            Web(srcWebUrl).Lists().getById(lookupField.LookupList).execute(
                                list => {
                                    // Add the lookup list field
                                    fields[lookupField.InternalName] = true;
                                    cfgProps.ListCfg[0].CustomFields.push({
                                        description: lookupField.Description,
                                        fieldRef: lookupField.PrimaryFieldId,
                                        hidden: lookupField.Hidden,
                                        id: lookupField.Id,
                                        indexed: lookupField.Indexed,
                                        listName: list.Title,
                                        multi: lookupField.AllowMultipleValues,
                                        name: lookupField.InternalName,
                                        readOnly: lookupField.ReadOnlyField,
                                        relationshipBehavior: lookupField.RelationshipDeleteBehavior,
                                        required: lookupField.Required,
                                        showField: lookupField.LookupField,
                                        title: lookupField.Title,
                                        type: Helper.SPCfgFieldType.Lookup
                                    } as Helper.IFieldInfoLookup);

                                    // Check the next field
                                    resolve(null);
                                },

                                err => {
                                    // Broken lookup field, don't add it
                                    console.log("Error trying to find lookup list for field '" + lookupField.InternalName + "' with id: " + lookupField.LookupList);
                                    resolve(null);
                                }
                            )
                        });
                    }).then(() => {
                        // Parse the calculated fields
                        for (let i = 0; i < calcFields.length; i++) {
                            let calcField = calcFields[i];

                            if (fields[calcField.InternalName] == null) {
                                let parser = new DOMParser();
                                let schemaXml = parser.parseFromString(calcField.SchemaXml, "application/xml");

                                // Get the formula
                                let formula = schemaXml.querySelector("Formula");

                                // Parse the field refs
                                let fieldRefs = schemaXml.querySelectorAll("FieldRef");
                                for (let j = 0; j < fieldRefs.length; j++) {
                                    let fieldRef = fieldRefs[j].getAttribute("Name");

                                    // Ensure the field exists
                                    let field = list.getField(fieldRef);
                                    if (field) {
                                        // Calculated formulas are supposed to contain the display name
                                        // Replace any instance of the internal field w/ the correct format
                                        let regexp = new RegExp(fieldRef, "g");
                                        formula.innerHTML = formula.innerHTML.replace(regexp, "[" + field.Title + "]");
                                    }
                                }

                                // Append the field
                                fields[calcField.InternalName] = true;
                                cfgProps.ListCfg[0].CustomFields.push({
                                    name: calcField.InternalName,
                                    schemaXml: schemaXml.querySelector("Field").outerHTML
                                });
                            }
                        }

                        // Resolve the request
                        resolve({
                            cfg: cfgProps,
                            lookupFields
                        });
                    });
                }
            });
        });
    }

    // Installs the configuration
    private static installConfiguration(cfg: Helper.ISPConfigProps, webUrl: string, elLog: HTMLElement): PromiseLike<List[]> {
        // Show a loading dialog
        LoadingDialog.setHeader("Creating the List");
        LoadingDialog.setBody("Initializing the request...");
        LoadingDialog.show();

        // Return a promise
        return new Promise((resolve, reject) => {
            // Create the list(s)
            this.createLists(cfg, webUrl, elLog).then(lists => {
                // Hide the dialog
                LoadingDialog.hide();

                // Resolve the request
                resolve(lists);
            }, reject);
        });
    }

    // Renders the results
    static renderResults(el: HTMLElement, cfgProps: Helper.ISPConfigProps, webUrl: string, lists: List[], showDeleteFl: boolean = false) {
        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Set the body
        el.innerHTML = "<p>Click on the link(s) below to access the list settings for validation.</p>";

        // Parse the lists
        let items: Components.IListGroupItem[] = [];
        for (let i = 0; i < lists.length; i++) {
            let list = lists[i];

            // Add the list links
            items.push({
                data: list,
                onRender: (el, item) => {
                    // Render the list view button
                    Components.Button({
                        el,
                        isSmall: true,
                        text: "View List",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Go to the list
                            window.open(list.ListUrl, "_blank");
                        }
                    });

                    // Render the list settings button
                    Components.Button({
                        el,
                        className: "ms-2",
                        isSmall: true,
                        text: "List Settings",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Go to the list
                            window.open(list.ListSettingsUrl, "_blank");
                        }
                    });
                }
            });
        }

        Components.ListGroup({
            el: CanvasForm.BodyElement,
            items
        });

        // See if we are deleting the list
        if (showDeleteFl) {
            // Add the footer
            let footer = document.createElement("div");
            footer.classList.add("d-flex");
            footer.classList.add("justify-content-end");
            footer.classList.add("mt-2");
            CanvasForm.BodyElement.appendChild(footer);

            // Render a delete button
            Components.Tooltip({
                el: footer,
                content: "Click to delete the test lists.",
                btnProps: {
                    text: "Delete List",
                    type: Components.ButtonTypes.OutlineDanger,
                    onClick: () => {
                        // Show a loading dialog
                        LoadingDialog.setHeader("Deleting List");
                        LoadingDialog.setBody("This will close after the lists are removed...");
                        LoadingDialog.show();

                        // Uninstall the configuration
                        Helper.SPConfig(cfgProps, webUrl).uninstall().then(() => {
                            // Hide the dialogs
                            LoadingDialog.hide();
                            CanvasForm.hide();
                        });
                    }
                }
            });
        }

        // Show the modal
        CanvasForm.show();
    }

    // Fixes the lookup fields
    private static testList(listCfg: Helper.ISPCfgListInfo, webUrl: string): PromiseLike<List> {
        // Update the loading dialog
        LoadingDialog.setBody("Validating the list configuration");

        // Return a promise
        return new Promise((resolve, reject) => {
            // Create the list
            let list = new List({
                listName: listCfg.ListInformation.Title,
                webUrl,
                onInitialized: () => {
                    // Resolve the list
                    resolve(list);
                },
                onInitError: reject
            });
        });
    }

    // Validates the lookup fields
    private static validateLookups(srcUrl: string, dstUrl: string, srcList: IListInfo, lookups: Types.SP.FieldLookup[]) {
        // Parse the lookup fields
        return Helper.Executor(lookups, lookup => {
            // Ensure this lookup isn't to the source list
            if (lookup.LookupList?.indexOf(srcList.ListId) >= 0) { return; }

            // Return a promise
            return new Promise((resolve, reject) => {
                // Get the source list
                Web(srcUrl).Lists().getById(lookup.LookupList).execute(list => {
                    // Ensure the list exists in the destination
                    Web(dstUrl).Lists(list.Title).execute(resolve, () => {
                        // Reject the reqeust
                        reject("Lookup list for field '" + lookup.InternalName + "' does not exist in the configuration. Please add the lists in the appropriate order.");
                    });

                }, () => {
                    // Reject the reqeust
                    reject("Lookup list for field '" + lookup.InternalName + "' does not exist in the source web. Please review the source list for any issues.");
                });
            });
        });
    }
}