import { Components, Types, Web } from "gd-sprest-bs";
import * as Scripts from "./scripts";
import Strings from "./strings";
import { IScript } from "./common";

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
                    content: "You are not an admin or an owner of this site.",
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
        let items: Components.IDropdownItem[] = [];
        let scripts = new Map<string, IScript>([
            [Scripts.DocumentRetentionModal.name, Scripts.DocumentRetentionModal],
            [Scripts.DocumentSearchModal.name, Scripts.DocumentSearchModal],
            [Scripts.ExternalUsersModal.name, Scripts.ExternalUsersModal],
            [Scripts.HubSiteInfoModal.name, Scripts.HubSiteInfoModal],
            [Scripts.ListInfoModal.name, Scripts.ListInfoModal],
            [Scripts.SecurityGroupsModal.name, Scripts.SecurityGroupsModal],
            [Scripts.SiteInfoModal.name, Scripts.SiteInfoModal],
            [Scripts.SiteUsageModal.name, Scripts.SiteUsageModal],
            [Scripts.SiteUsersModal.name, Scripts.SiteUsersModal],
            [Scripts.StorageMetricsModal.name, Scripts.StorageMetricsModal]
        ]);

        // Parse the scripts
        scripts.forEach(script => {
            items.push({
                text: script.name,
                value: script.description
            });
        });

        Components.Card({
            el: this._el,
            header: {
                className: "h6",
                onRender: (el) => {
                    let div = document.createElement("div");
                    div.classList.add("title");
                    div.innerText = Strings.ProjectName;
                    el.appendChild(div);
                }
            },
            body: [{
                onRender: (el) => {
                    el.classList.add("d-flex");
                    el.classList.add("flex-column");
                    el.classList.add("justify-content-between");

                    // Render the div
                    let elDiv = document.createElement("div");
                    elDiv.className = "d-flex";
                    el.appendChild(elDiv);

                    let ddl = Components.Dropdown({
                        el: elDiv,
                        btnClassName: "w-100",
                        items,
                        label: scripts.keys().next().value,
                        title: "Select a script to run",
                        type: Components.DropdownTypes.OutlinePrimary,
                        onChange: (item: Components.IDropdownItem) => {
                            ddl.setLabel(item.text);
                            elDescription.innerHTML = item.value;
                            ttp.button.el.setAttribute("data-script", item.text);
                        }
                    });

                    // Render a tooltip
                    let ttp = Components.Tooltip({
                        el: elDiv,
                        content: "Run this report",
                        placement: Components.TooltipPlacements.Bottom,
                        btnProps: {
                            className: "run",
                            text: "Run",
                            type: Components.ButtonTypes.OutlinePrimary,
                            onClick: (b, e) => {
                                let scriptName = (e.target as HTMLButtonElement).dataset.script;

                                // Initialize the script
                                new (scripts.get(scriptName) as IScript).init([Strings.SourceUrl]);
                            }
                        }
                    });
                    ttp.button.el.setAttribute("data-script", scripts.get(scripts.keys().next().value).name);

                    // Render the description
                    let elDescription = document.createElement("p");
                    elDescription.className = "description mt-2";
                    elDescription.innerHTML = scripts.get(scripts.keys().next().value).description;
                    el.appendChild(elDescription);
                }
            }]
        });

        // Render the footer
        let footer = document.createElement("div");
        footer.className = "d-flex justify-content-end pe-1";
        footer.id = "footer";
        footer.innerHTML = `<label class="text-dark">v${Strings.Version}</label>`;
        this._el.appendChild(footer);
    }
}