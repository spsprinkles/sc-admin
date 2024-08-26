import { CanvasForm, ILookupData, List, ListConfig, LoadingDialog } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { IListInfo } from "../scripts/lists";

/**
 * Copy List
 */
export class CopyList {
    // Method to create the list configuration
    static createListConfiguration(elResults: HTMLElement, srcList: IListInfo, dstWebUrl: string, dstListName: string): PromiseLike<string> {
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
                    ListConfig.generate({
                        showDialog: true,
                        srcWebUrl: srcList.WebUrl,
                        srcList: srcList.ListName
                    }).then(srcListCfg => {
                        // Validate the lookup fields
                        ListConfig.validateLookups({
                            cfg: srcListCfg.cfg,
                            dstUrl: web.ServerRelativeUrl,
                            lookupFields: srcListCfg.lookupFields,
                            showDialog: true,
                            srcListId: srcList.ListId,
                            srcWebUrl: srcList.WebUrl,
                        }).then((listCfg) => {
                            // See if the destination list name exists
                            if (dstListName) {
                                // Set the destination list name
                                listCfg.ListCfg[listCfg.ListCfg.length - 1].ListInformation.Title = dstListName;
                            }

                            // Save a copy of the configuration
                            let strConfig = JSON.stringify(listCfg);

                            // Get the lookup list data
                            this.getLookupData(true, srcList.WebUrl, srcList.ListId, srcListCfg.lookupFields).then(lookupData => {
                                // Test the configuration
                                this.installConfiguration(listCfg, web.ServerRelativeUrl, JSON.stringify(lookupData)).then(lists => {
                                    // Hide the loading dialog
                                    LoadingDialog.hide();

                                    // Show the results
                                    this.renderResults(elResults, listCfg, web.ServerRelativeUrl, lists);

                                    // Resolve the request
                                    resolve(strConfig);
                                });
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

    // Method to create the list
    private static createLists(cfgProps: Helper.ISPConfigProps, webUrl: string, lookupData: ILookupData[]): PromiseLike<List[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Initialize the logging form
            this.initLoggingForm();

            // Set the log event
            cfgProps.onLogMessage = (msg, isError) => {
                // Append the message
                let elMessage = document.createElement("p");
                elMessage.innerHTML = msg;
                isError ? elMessage.style.color = "red" : null;
                CanvasForm.BodyElement.appendChild(elMessage);

                // Focus on the message
                elMessage.focus();
                elMessage.scrollIntoView();
            };

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
                    // Create the lookup list data
                    ListConfig.createLookupListData({
                        lookupData,
                        webUrl,
                        showDialog: true
                    }).then(() => {
                        // Resolve the request
                        resolve(lists);
                    }, reject);
                });
            }, reject);
        });
    }

    // Gets the lookup data
    private static getLookupData(includeLookupData: boolean, srcWebUrl: string, srcListId: string, lookupFields: Types.SP.FieldLookup[]): PromiseLike<ILookupData[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if we are not getting the lookup data
            if (!includeLookupData) { resolve([]); return; }

            // Get the lookup list data
            ListConfig.generateLookupListData({
                lookupFields,
                srcListId,
                srcWebUrl,
                showDialog: true
            }).then(lookupData => {
                // Resolve the request
                resolve(lookupData);
            }, reject);
        });
    }

    // Initializes the logging form
    private static initLoggingForm() {
        // Clear the canvas form
        CanvasForm.clear();

        // Disable auto-close
        CanvasForm.setAutoClose(false);

        // Set the size
        CanvasForm.setSize(Components.OffcanvasSize.Medium2);

        // Set the header
        CanvasForm.setHeader("Create List Logging");

        // Show the logging
        CanvasForm.show();
    }

    // Installs the configuration
    private static installConfiguration(cfg: Helper.ISPConfigProps, webUrl: string, strLookupData: string): PromiseLike<List[]> {
        // Show a loading dialog
        LoadingDialog.setHeader("Creating the List");
        LoadingDialog.setBody("Initializing the request...");
        LoadingDialog.show();

        // Return a promise
        return new Promise((resolve, reject) => {
            // Try to convert the data
            let lookupData: ILookupData[] = null;
            try { lookupData = JSON.parse(strLookupData); }
            catch { lookupData = []; }

            // Create the list(s)
            this.createLists(cfg, webUrl, lookupData).then(lists => {
                // Hide the dialog
                LoadingDialog.hide();

                // Resolve the request
                resolve(lists);
            }, reject);
        });
    }

    // Renders the results
    static renderResults(el: HTMLElement, cfgProps: Helper.ISPConfigProps, webUrl: string, lists: List[]) {
        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Set the body
        el.innerHTML = "<p>Click on the link(s) below to access the list settings for validation.</p>";

        // Parse the lists
        let items: Components.IListGroupItem[] = [];
        for (let i = 0; i < lists.length; i++) {
            let list = lists[i];

            // Create the list element
            let elList = document.createElement("div");
            elList.innerHTML = `<h6>${list.ListName}:</h6>`;

            // Create the buttons
            Components.ButtonGroup({
                el: elList,
                isSmall: true,
                buttons: [
                    {
                        data: list,
                        text: "View List",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: (button) => {
                            // Go to the list
                            window.open((button.data as List).ListUrl, "_blank");
                        }
                    },
                    {
                        data: list,
                        text: "Settings",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: (button) => {
                            // Go to the list
                            window.open((button.data as List).ListSettingsUrl, "_blank");
                        }
                    }
                ]
            });

            // Add the view list link
            items.push({ content: elList });
        }

        // Render the list group
        Components.ListGroup({
            el: CanvasForm.BodyElement,
            items
        });

        // Render a button to close the form
        Components.ButtonGroup({
            el: CanvasForm.BodyElement,
            buttons: [
                {
                    text: "Delete Lists",
                    type: Components.ButtonTypes.OutlineDanger,
                    onClick: () => {
                        // Show a loading dialog
                        LoadingDialog.setHeader("Deleting List");
                        LoadingDialog.setBody("This will close after the lists are removed...");
                        LoadingDialog.show();

                        // Uninstall the configuration
                        Helper.SPConfig(cfgProps, webUrl).uninstall().then(() => {
                            // Hide the dialog and form
                            LoadingDialog.hide();
                            CanvasForm.hide();
                        });
                    }
                },
                {
                    text: "Close",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Close the form
                        CanvasForm.hide();
                    }
                }
            ]
        })

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
}