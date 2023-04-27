import { Components, ContextInfo } from "gd-sprest-bs";
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

        // Render the dashboard
        this.render();
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