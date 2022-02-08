import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ScAdminWebPart.module.scss';
import * as strings from 'ScAdminWebPartStrings';

export interface IScAdminWebPartProps {
  description: string;
}

// Reference the solution
import "../../../../dist/sc-admin.min.js";
declare var SCAdmin;

export default class ScAdminWebPart extends BaseClientSideWebPart<IScAdminWebPartProps> {

  public render(): void {
    // Create the dashboard
    SCAdmin.render(this.domElement, this.context);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
