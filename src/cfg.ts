import { Helper, SPTypes } from "gd-sprest-bs";

/**
 * Configuration
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: "Site Collections",
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            CustomFields: [
                {
                    name: "SCUrl",
                    title: "Site Collection Url",
                    type: Helper.SPCfgFieldType.Url
                }
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    JSLink: "~site/siteassets/sc-admin.min.js",
                    ViewFields: ["LinkTitle", "SCUrl"]
                }
            ]
        }
    ]
});