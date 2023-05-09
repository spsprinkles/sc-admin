import { Components, ContextInfo, Types, Web } from "gd-sprest-bs";
import * as Scripts from "./scripts";

export class Dashboard {
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement, context?) {
        // Set the properties
        this._el = el;

        // See if the context exists
        if (context) {
            // Set the context
            ContextInfo.setPageContext(context.pageContext);
        }

        // Ensure the user is an Owner or Admin
        this.isOwnerOrAdmin().then(
            // Render the solution if the user is an owner/admin
            this.render,

            // Not Owner/Admin
            () => {
                // Render an alert
                Components.Alert({
                    el: this._el,
                    content: "You are not an owner/admin of the site.",
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
            },)
        });
    }

    // Renders the scripts as a card
    private render() {
        let cards1: Components.ICardProps[] = [];
        let cards2: Components.ICardProps[] = [];

        // Parse the scripts
        [
            Scripts.DocumentRetentionModal,
            Scripts.ExternalUsersModal,
            Scripts.ListInfoModal,
            Scripts.SecurityGroupsModal
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
                            content: "Click to run this script",
                            placement: Components.TooltipPlacements.Bottom,
                            btnProps: {
                                className: "my-2",
                                text: "Run",
                                type: Components.ButtonTypes.OutlinePrimary,
                                onClick: () => {
                                    // Initialize the script
                                    new script.init([ContextInfo.siteServerRelativeUrl]);
                                }
                            }
                        });
                    }
                }]
            })
        });

        // Parse the scripts
        [
            Scripts.SiteInfoModal,
            Scripts.SiteUsageModal,
            Scripts.SiteUsersModal,
            Scripts.StorageMetricsModal
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
                            content: "Click to run this script",
                            placement: Components.TooltipPlacements.Bottom,
                            btnProps: {
                                className: "my-2",
                                text: "Run",
                                type: Components.ButtonTypes.OutlinePrimary,
                                onClick: () => {
                                    // Initialize the script
                                    new script.init([ContextInfo.siteServerRelativeUrl]);
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
            cards: cards1
        });

        // Render the cards
        Components.CardGroup({
            el: this._el,
            cards: cards2,
            className: "mt-3"
        });
    }
}