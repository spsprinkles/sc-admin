import { Environment, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneLabel, PropertyPaneSlider, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import * as strings from 'ScAdminWebPartStrings';

export interface IScAdminWebPartProps {
  fractionDigits: number;
  searchFileTypes: string;
  searchMonths: number;
  searchTerms: string;
  timeFormat: string;
}

// Reference the solution
import "../../../../dist/sc-admin.min.js";
declare const SCAdmin: {
  description: string;
  render: new (props: {
    el: HTMLElement;
    context?: WebPartContext;
    envType?: number;
    fractionDigits?: number;
    searchFileTypes?: string;
    searchMonths?: number;
    searchTerms?: string;
    timeFormat?: string;
    sourceUrl?: string;
  }) => void;
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
    if (!this.properties.fractionDigits) { this.properties.fractionDigits = strings.FractionDigitsFieldValue; }
    if (!this.properties.searchFileTypes) { this.properties.searchFileTypes = strings.SearchFileTypesFieldValue; }
    if (!this.properties.searchMonths) { this.properties.searchMonths = strings.SearchMonthsFieldValue; }
    if (!this.properties.searchTerms) { this.properties.searchTerms = strings.SearchTermsFieldValue; }
    if (!this.properties.timeFormat) { this.properties.timeFormat = strings.TimeFormatFieldValue; }

    // Render the solution
    new SCAdmin.render({
      el: this.domElement,
      context: this.context,
      envType: Environment.type,
      fractionDigits: this.properties.fractionDigits,
      searchFileTypes: this.properties.searchFileTypes,
      searchMonths: this.properties.searchMonths,
      searchTerms: this.properties.searchTerms,
      timeFormat: this.properties.timeFormat
    });

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
                PropertyPaneSlider('fractionDigits', {
                  label: strings.FractionDigitsFieldLabel,
                  max: 6,
                  min: 1,
                  showValue: true,
                  value: strings.FractionDigitsFieldValue
                }),
                PropertyPaneSlider('searchMonths', {
                  label: strings.SearchMonthsFieldLabel,
                  max: 36,
                  min: 1,
                  showValue: true,
                  value: strings.SearchMonthsFieldValue
                }),
                PropertyPaneTextField('searchTerms', {
                  label: strings.SearchTermsFieldLabel,
                  description: strings.SearchTermsFieldDescription
                }),
                PropertyPaneTextField('searchFileTypes', {
                  label: strings.SearchFileTypesFieldLabel,
                  description: strings.SearchFileTypesFieldDescription
                }),
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
