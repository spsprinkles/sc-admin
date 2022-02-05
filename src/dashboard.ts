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
        let cards: Components.ICardProps[] = [];

        // Parse the scripts
        [
            Scripts.DocumentRetentionModal,
            Scripts.ListInfoModal,
            Scripts.SecurityGroupsModal,
            Scripts.SiteInfoModal
        ].forEach(script => {
            cards.push({
                header: {
                    content: script.name
                },
                body: [{
                    onRender: (el) => {
                        // Render the description
                        let elContent = document.createElement("p");
                        elContent.innerHTML = script.description;
                        el.appendChild(elContent);

                        // Render a popover
                        Components.Popover({
                            el,
                            options: {
                                content: "Click to run this script"
                            },
                            type: Components.PopoverTypes.Primary,
                            placement: Components.PopoverPlacements.Auto,
                            btnProps: {
                                text: "Run",
                                type: Components.ButtonTypes.OutlinePrimary,
                                onClick: () => {
                                    // Initialize the script
                                    new script.init();
                                }
                            }
                        })
                    }
                }]
            })
        });

        // Render the cards
        Components.CardGroup({
            el: this._el,
            className: "sc-admin",
            cards
        });
    }
}