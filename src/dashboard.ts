import { Components, Types, Web } from "gd-sprest-bs";
import * as Scripts from "./scripts";
import Strings from "./strings";

export class Dashboard {
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Set the properties
        this._el = el;

        // Ensure the user is an Owner or Admin
        this.isOwnerOrAdmin().then(
            // Render the solution if the user is an owner/admin
            this.render,

            // Not Owner/Admin
            () => {
                // Render an alert
                Components.Alert({
                    el: this._el,
                    content: "You are not an admin or owner of the site",
                    type: Components.AlertTypes.Danger
                });
            }
        )

        // Render the dashboard
        this.render();
    }

    // Sees if the user is an owner or admin
    private isOwnerOrAdmin(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the web information
            Web().query({
                Expand: ["AssociatedOwnerGroup/Users", "CurrentUser"]
            }).execute(web => {
                // See if the user is an admin
                if (web.CurrentUser.IsSiteAdmin) { resolve(); }
                else {
                    // Parse the users
                    let users = (web.AssociatedOwnerGroup as any as Types.SP.GroupOData).Users.results || [];
                    for (let i = 0; i < users.length; i++) {
                        let user = users[i];

                        // See if the user is in the group
                        if (user.Id == web.CurrentUser.Id) {
                            // Resolve the request and return
                            resolve();
                            return;
                        }
                    }

                    // Reject the request
                    reject();
                }
            }, reject);
        });
    }

    // Renders the scripts as a card
    private render() {
        let cards1: Components.ICardProps[] = [];
        let cards2: Components.ICardProps[] = [];
        let cards3: Components.ICardProps[] = [];

        // Parse the scripts
        [
            Scripts.DocumentRetentionModal,
            Scripts.DocumentSearchModal,
            Scripts.ExternalUsersModal,
            Scripts.HubSiteInfoModal,
        ].forEach(script => {
            cards1.push({
                header: {
                    className: "h6",
                    onRender: (el) => {
                        let div = document.createElement("div");
                        div.classList.add("mt-1");
                        div.innerText = script.name;
                        el.appendChild(div);
                    }
                },
                body: [{
                    onRender: (el) => {
                        el.classList.add("d-flex");
                        el.classList.add("flex-column");
                        el.classList.add("justify-content-between");
                        // Render the description
                        let elContent = document.createElement("p");
                        elContent.innerHTML = script.description;
                        el.appendChild(elContent);

                        // Render a tooltip
                        Components.Tooltip({
                            el,
                            content: "Run this report",
                            placement: Components.TooltipPlacements.Bottom,
                            btnProps: {
                                className: "mt-3",
                                text: "Run",
                                type: Components.ButtonTypes.OutlinePrimary,
                                onClick: () => {
                                    // Initialize the script
                                    new script.init([Strings.SourceUrl]);
                                }
                            }
                        });
                    }
                }]
            })
        });

        // Parse the scripts
        [
            Scripts.ListInfoModal,
            Scripts.SecurityGroupsModal,
            Scripts.SiteInfoModal,
            Scripts.SiteUsageModal,
        ].forEach(script => {
            cards2.push({
                header: {
                    className: "h6",
                    onRender: (el) => {
                        let div = document.createElement("div");
                        div.classList.add("mt-1");
                        div.innerText = script.name;
                        el.appendChild(div);
                    }
                },
                body: [{
                    onRender: (el) => {
                        el.classList.add("d-flex");
                        el.classList.add("flex-column");
                        el.classList.add("justify-content-between");
                        // Render the description
                        let elContent = document.createElement("p");
                        elContent.innerHTML = script.description;
                        el.appendChild(elContent);

                        // Render a tooltip
                        Components.Tooltip({
                            el,
                            content: "Run this report",
                            placement: Components.TooltipPlacements.Bottom,
                            btnProps: {
                                className: "mt-3",
                                text: "Run",
                                type: Components.ButtonTypes.OutlinePrimary,
                                onClick: () => {
                                    // Initialize the script
                                    new script.init([Strings.SourceUrl]);
                                }
                            }
                        });
                    }
                }]
            })
        });

        // Parse the scripts
        [
            Scripts.SiteUsersModal,
            Scripts.StorageMetricsModal,
        ].forEach(script => {
            cards3.push({
                header: {
                    className: "h6",
                    onRender: (el) => {
                        let div = document.createElement("div");
                        div.classList.add("mt-1");
                        div.innerText = script.name;
                        el.appendChild(div);
                    }
                },
                body: [{
                    onRender: (el) => {
                        el.classList.add("d-flex");
                        el.classList.add("flex-column");
                        el.classList.add("justify-content-between");
                        // Render the description
                        let elContent = document.createElement("p");
                        elContent.innerHTML = script.description;
                        el.appendChild(elContent);

                        // Render a tooltip
                        Components.Tooltip({
                            el,
                            content: "Run this report",
                            placement: Components.TooltipPlacements.Bottom,
                            btnProps: {
                                className: "mt-3",
                                text: "Run",
                                type: Components.ButtonTypes.OutlinePrimary,
                                onClick: () => {
                                    // Initialize the script
                                    new script.init([Strings.SourceUrl]);
                                }
                            }
                        });
                    }
                }]
            })
        });

        // Render the cards
        Components.CardGroup({
            el: this._el,
            cards: cards1,
            className: "cg-1"
        });

        // Render the cards
        Components.CardGroup({
            el: this._el,
            cards: cards2,
            className: "cg-2"
        });

        // Render the cards
        Components.CardGroup({
            el: this._el,
            cards: cards3,
            className: "cg-3"
        });

        // Render the footer
        let footer = document.createElement("div");
        footer.className = "d-flex justify-content-end pe-1";
        footer.id = "footer";
        footer.innerHTML = `<label class="text-dark">v${Strings.Version}</label>`;
        this._el.appendChild(footer);
    }
}