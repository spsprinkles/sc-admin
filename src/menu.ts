import { Components, ContextInfo } from "gd-sprest-bs";
import * as Scripts from "./scripts";
declare var SP;

/**
 * Menu
 */
export class Menu {
    private _ctx: any = null;

    // Constructor
    constructor(elMenu: Element, ctx: any) {
        // Save the context
        this._ctx = ctx;

        // See if this is coming from SPFx
        if (this._ctx.pageContext) {
            // Set the context
            ContextInfo.setPageContext(this._ctx.pageContext);

            // Render the menu
            this.render(elMenu);
        } else {
            // Create the link element
            let elLink = document.createElement("a");
            elLink.classList.add("ms-commandLink");
            elLink.id = "sc-admin-link";
            elLink.innerHTML = "Admin";
            elMenu.appendChild(elLink);

            // Render the menu
            this.render(elLink);
        }
    }

    // Sees if the admin menu exists
    static Exists(): boolean {
        // Query for the link
        return document.querySelector("#sc-admin-link") != null;
    }

    // Gets the selected item urls
    private getSelectedUrls() {
        let urls = [];

        // Get the selected items from the internal library
        let items = SP.ListOperation.Selection.getSelectedItems();

        // Parse the items
        for (let i = 0; i < this._ctx.ListData.Row.length; i++) {
            let item = this._ctx.ListData.Row[i];

            // Parse the selected items
            for (let j = 0; j < items.length; j++) {
                let itemId = items[j].id;

                // See if this is the item
                if (itemId == item.ID || itemId == item.Id) {
                    // Add the url
                    urls.push(item.SCUrl);
                    break;
                }
            }
        }

        // Return the urls
        return urls;
    }

    // Renders the menu
    private render(el: Element) {
        // Create the menu
        let menu = Components.Dropdown({
            autoSelect: false,
            className: "sc-admin-menu",
            menuOnly: true,
            items: [
                {
                    text: Scripts.DocumentRetentionModal.name,
                    onClick: () => {
                        // Hide the popover
                        popover.hide();

                        // Display the dialog
                        new Scripts.DocumentRetentionModal.init(this.getSelectedUrls());
                    }
                },
                {
                    text: Scripts.ExternalUsersModal.name,
                    onClick: () => {
                        // Hide the popover
                        popover.hide();

                        // Display the dialog
                        new Scripts.ExternalUsersModal.init(this.getSelectedUrls());
                    }
                },
                {
                    text: Scripts.ListInfoModal.name,
                    onClick: () => {
                        // Hide the popover
                        popover.hide();

                        // Display the dialog
                        new Scripts.ListInfoModal.init(this.getSelectedUrls());
                    }
                },
                {
                    text: Scripts.SecurityGroupsModal.name,
                    onClick: () => {
                        // Hide the popover
                        popover.hide();

                        // Display the dialog
                        new Scripts.SecurityGroupsModal.init(this.getSelectedUrls());
                    }
                },
                {
                    text: Scripts.SiteInfoModal.name,
                    onClick: () => {
                        // Hide the popover
                        popover.hide();

                        // Display the dialog
                        new Scripts.SiteInfoModal.init(this.getSelectedUrls());
                    }
                },
                {
                    text: Scripts.SiteUsageModal.name,
                    onClick: () => {
                        // Hide the popover
                        popover.hide();

                        // Display the dialog
                        new Scripts.SiteUsageModal.init(this.getSelectedUrls());
                    }
                }
            ]
        });

        let popover = Components.Popover({
            target: el,
            placement: Components.PopoverPlacements.Bottom,
            options: {
                trigger: "focus",
                content: menu.el
            }
        });
    }
}