import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MultipleGroupsDemoWebPart.module.scss';
import * as strings from 'MultipleGroupsDemoWebPartStrings';

export interface IMultipleGroupsDemoWebPartProps {
  description: string;
  productname: string;
  isCertified: boolean;
}

export default class MultipleGroupsDemoWebPart extends BaseClientSideWebPart<IMultipleGroupsDemoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.multipleGroupsDemo} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
        <p class="${styles.welcome}">${ this.properties.productname}</p>
        <p class="${styles.welcome}">${ this.properties.isCertified}</p>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
        The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
          <ul class="${styles.links}">
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>
      </div>
    </section>`;
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: {
          description: "Page 1"
        },
        groups: [{
          groupName: "First Group",
          groupFields: [
            PropertyPaneTextField('productname', {
              label: "Prodcut Name 1"
            })
          ]
        },{
          groupName: "Second Group",
          groupFields: [
            PropertyPaneToggle('isCertified', {
              label: "Is certified? 1"
            })
          ]
        }],
        // Al colocar la siguiente linea, en false, se quita el modo acordeon
        displayGroupsAsAccordion: true
      },
      {
        header: {
          description: "Page 2"
        },
        groups: [{
          groupName: "First Group",
          groupFields: [
            PropertyPaneTextField('productname', {
              label: "Prodcut Name 2"
            })
          ]
        },{
          groupName: "Second Group",
          groupFields: [
            PropertyPaneToggle('isCertified', {
              label: "Is certified? 2"
            })
          ]
        }],
        // Al colocar la siguiente linea, en false, se quita el modo acordeon
        displayGroupsAsAccordion: true
      },
      {
        header: {
          description: "Page 3"
        },
        groups: [{
          groupName: "First Group",
          groupFields: [
            PropertyPaneTextField('productname', {
              label: "Prodcut Name 3"
            })
          ]
        },{
          groupName: "Second Group",
          groupFields: [
            PropertyPaneToggle('isCertified', {
              label: "Is certified? 3"
            })
          ]
        }],
        // Al colocar la siguiente linea, en false, se quita el modo acordeon
        displayGroupsAsAccordion: true
      },
      {
        header: {
          description: "Page 4"
        },
        groups: [{
          groupName: "First Group",
          groupFields: [
            PropertyPaneTextField('productname', {
              label: "Prodcut Name 4"
            })
          ]
        },{
          groupName: "Second Group",
          groupFields: [
            PropertyPaneToggle('isCertified', {
              label: "Is certified? 4"
            })
          ]
        }],
        // Al colocar la siguiente linea, en false, se quita el modo acordeon
        displayGroupsAsAccordion: true
      }
    ]
    };
  }
}
