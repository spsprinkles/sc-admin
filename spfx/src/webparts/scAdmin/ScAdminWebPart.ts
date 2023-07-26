import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneLabel } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import * as strings from 'ScAdminWebPartStrings';

// Reference the solution
import "../../../../dist/sc-admin.min.js";
declare const SCAdmin: {
  description: string;
  render: (el: HTMLElement, context: WebPartContext) => void;
  version: string;
};

export interface IScAdminWebPartProps {
  description: string;
}

export default class ScAdminWebPart extends BaseClientSideWebPart<IScAdminWebPartProps> {

  public render(): void {
    // Create the dashboard
    SCAdmin.render(this.domElement, this.context);
  }

  protected get dataVersion(): Version {
    return Version.parse(SCAdmin.version);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneLabel('version', {
                  text: "v" + SCAdmin.version
                })
              ]
            }
          ],
          header: {
            description: SCAdmin.description
          },
        }
      ]
    };
  }
}
