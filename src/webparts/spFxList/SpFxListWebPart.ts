import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'SpFxListWebPartStrings';
import SpFxList from './components/SpFxList';
import { ISpFxListProps } from './components/ISpFxListProps';

export interface ISpFxListWebPartProps {
  description: string;
  selectedListId: string;
}

export default class SpFxListWebPart extends BaseClientSideWebPart<ISpFxListWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _listsDropdownDisabled: boolean = false;

  public render(): void {
    console.log('SpFxListWebPart render called');
    console.log('selectedListId:', this.properties.selectedListId);
    console.log('domElement:', this.domElement);
    
    try {
      const element: React.ReactElement<ISpFxListProps> = React.createElement(
        SpFxList,
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          selectedListId: this.properties.selectedListId || '',
          spHttpClient: this.context.spHttpClient,
          webUrl: this.context.pageContext.web.absoluteUrl
        }
      );

      console.log('Element created, rendering...');
      ReactDom.render(element, this.domElement);
      console.log('ReactDom.render completed');
    } catch (error) {
      console.error('Error in render:', error);
      this.domElement.innerHTML = `<div style="padding: 20px; color: red;">Error rendering web part: ${error.message}</div>`;
    }
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    }).then(() => {
      return this._loadLists();
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private _loadLists(): Promise<void> {
    this._listsDropdownDisabled = true;
    this.context.propertyPane.refresh();

    const webUrl = this.context.pageContext.web.absoluteUrl;
    // Using a simpler query that's more reliable
    const apiUrl = `${webUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,BaseTemplate&$orderby=Title`;

    console.log('Loading lists from:', apiUrl);

    // Use SPHttpClient which handles SharePoint authentication and headers automatically
    return this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        console.log('Response status:', response.status, response.statusText);
        
        if (!response.ok) {
          // Try to get error details from response
          return response.text().then(text => {
            console.error('Error response body:', text);
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
          });
        }

        return response.json();
      })
      .then((data: any) => {
        console.log('Lists data received:', data);
        
        // Handle both odata=verbose (d.results) and odata=nometadata (value) formats
        let lists: Array<{ Id: string; Title: string; BaseTemplate: number }> = [];
        
        if (data && data.d && data.d.results) {
          // odata=verbose format
          lists = data.d.results;
        } else if (data && data.value) {
          // odata=nometadata format
          lists = data.value;
        }
        
        console.log('Number of lists:', lists.length);

        if (lists.length === 0) {
          console.warn('No lists found in response:', data);
          this._listOptions = [];
        } else {
          // Filter to only show lists (BaseTemplate 100) and document libraries (BaseTemplate 101)
          const filteredLists = lists.filter((list: { Id: string; Title: string; BaseTemplate: number }) => 
            list.BaseTemplate === 100 || list.BaseTemplate === 101
          );

          this._listOptions = filteredLists.map(list => ({
            key: list.Id,
            text: `${list.Title}${list.BaseTemplate === 101 ? ' (Document Library)' : ''}`
          }));

          console.log('List options created:', this._listOptions.length);

          // If no list is selected and we have options, select the first one
          if (!this.properties.selectedListId && this._listOptions.length > 0) {
            this.properties.selectedListId = this._listOptions[0].key.toString();
          }
        }

        this._listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
      })
      .catch((error: Error) => {
        console.error('Error loading lists:', error);
        console.error('Error details:', JSON.stringify(error, Object.getOwnPropertyNames(error)));
        this._listOptions = [];
        this._listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    // Load lists when property pane opens if not already loaded
    if (this._listOptions.length === 0) {
      this._loadLists();
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    // Re-render the web part when the selected list changes
    if (propertyPath === 'selectedListId' && oldValue !== newValue) {
      this.render();
    }
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
                }),
                PropertyPaneDropdown('selectedListId', {
                  label: 'Select List or Document Library',
                  options: this._listOptions,
                  disabled: this._listsDropdownDisabled
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
