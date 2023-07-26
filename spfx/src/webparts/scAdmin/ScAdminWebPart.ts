import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneLabel, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import * as strings from 'ScAdminWebPartStrings';

export interface IScAdminWebPartProps {
  timeFormat: string;
}

// Reference the solution
import "../../../../dist/sc-admin.min.js";
declare const SCAdmin: {
  description: string;
  render: new (el: HTMLElement, context: WebPartContext, timeFormat: string) => void;
  version: string;
};

export default class ScAdminWebPart extends BaseClientSideWebPart<IScAdminWebPartProps> {
  private _hasRendered: boolean = false;

  public render(): void {
    // See if have rendered the solution
    if (this._hasRendered) {
      // Clear the element
      while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }
    }
    
    // Set the default property values
    if (!this.properties.timeFormat) { this.properties.timeFormat = strings.TimeFormatFieldValue; }

    // Render the solution
    new SCAdmin.render(this.domElement, this.context, this.properties.timeFormat);

    // Set the flag
    this._hasRendered = true;
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
                PropertyPaneTextField('timeFormat', {
                  label: strings.TimeFormatFieldLabel,
                  description: strings.TimeFormatFieldDescription
                }),
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
