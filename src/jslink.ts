import { Helper } from "gd-sprest-bs";
import { Menu } from "./menu";

declare var RenderHeaderTemplate, RenderFooterTemplate;

// Registers a JSLink w/ a list
export function register() {
    // See if the TextDecoder global variable exists
    if (window["TextDecoder"] == null) {
        // Load the text-encoder
        var s = document.createElement("script");
        s.src = "https://unpkg.com/text-encoding@0.6.4/lib/encoding-indexes.js";
        document.head.appendChild(s);
        s = document.createElement("script");
        s.src = "https://unpkg.com/text-encoding@0.6.4/lib/encoding.js";
        document.head.appendChild(s);
    }

    // Create the JSLink
    Helper.JSLink.register({
        Templates: {
            Header: ctx => {
                // Render the default header
                return "<div id='sc-admin-header'>" + RenderHeaderTemplate(ctx) + "</div>";
            },
            Footer: ctx => {
                // See if the menu item doesn't exist
                if (!Menu.Exists()) {
                    // Get the header element
                    let elViewControl = document.querySelector("#sc-admin-header .ms-csrlistview-controldiv");
                    if (elViewControl) {
                        // Create the menu element
                        let elMenu = document.createElement("div");
                        elMenu.classList.add("ms-InlineSearch-DivBaseline");
                        elMenu.style.paddingRight = "7px";

                        // Get the inline search element
                        let elSearch = elViewControl.querySelector(".ms-InlineSearch-DivBaseline");
                        if (elSearch) {
                            // Add the link after this
                            elViewControl.insertBefore(elMenu, elSearch);
                        } else {
                            // Append the link
                            elViewControl.appendChild(elMenu);
                        }

                        // Create the menu
                        new Menu(elMenu, ctx);
                    }
                }

                // Render the default footer
                return RenderFooterTemplate(ctx);
            }
        }
    });
}