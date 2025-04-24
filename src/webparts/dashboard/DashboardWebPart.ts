import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'DashboardWebPartStrings';
import Dashboard from './components/Dashboard';
import { IDashboardProps } from './components/IDashboardProps';
import { getSP } from '../../pnpjsConfig';


export interface IMenuItem {
  nome: string;
  link: string;
  imagemUrl: string;
  isVisible: boolean;
}

export interface IDashboardWebPartProps {
  description: string;
  menuItems: IMenuItem[];
}

export default class DashboardWebPart extends BaseClientSideWebPart<IDashboardWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _userDisplayName: string = 'Colaborador';

  public async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);
    this._userDisplayName = await this.getCurrentUser();

    if (!this.properties.menuItems || this.properties.menuItems.length === 0) {
      this.properties.menuItems = [
        { nome: "Diretoria", link: "https://lacerdapar.sharepoint.com/sites/Diretoria2", imagemUrl: require('./assets/diretoria.png'), isVisible: true },
        { nome: "Jurídico", link: "https://lacerdapar.sharepoint.com/sites/Jurdico", imagemUrl: require('./assets/juridico.png'), isVisible: true },
        { nome: "Financeiro", link: "https://lacerdapar.sharepoint.com/sites/Financeiro", imagemUrl: require('./assets/financeiro.png'), isVisible: true },
        { nome: "Suporte TI", link: "https://lacerdapar.sharepoint.com/sites/SuporteTI2", imagemUrl: require('./assets/itsupport.png'), isVisible: true },
        { nome: "Compras", link: "https://lacerdapar.sharepoint.com/sites/ComprasBraslia", imagemUrl: require('./assets/compras.png'), isVisible: true },
        { nome: "Departamento Pessoal", link: "https://lacerdapar.sharepoint.com/sites/DP-Braslia/SitePages/ProjectHome.aspx", imagemUrl: require('./assets/dp.png'), isVisible: true },
        { nome: "Comercial", link: "https://lacerdapar.sharepoint.com/sites/Comercial", imagemUrl: require('./assets/comercial.png'), isVisible: true }
      ];
    }

    this._environmentMessage = await this._getEnvironmentMessage();
  }

  public render(): void {
    const element: React.ReactElement<IDashboardProps> = React.createElement(
      Dashboard,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks?.microsoftTeams,
        userDisplayName: this._userDisplayName,
        menuItems: this.properties.menuItems
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private async getCurrentUser(): Promise<string> {
    try {
      const user = await getSP().web.currentUser();
      const fullName = user.Title || "Colaborador";
      const nameParts = fullName.split(" ");
      const firstName = nameParts[0];
      const lastName = nameParts.length > 1 ? nameParts[nameParts.length - 1] : "";
      return `${firstName} ${lastName}`;
    } catch (error) {
      console.error("Erro ao buscar usuário:", error);
      return "Colaborador";
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;

    const semanticColors = currentTheme.semanticColors;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  private async _getEnvironmentMessage(): Promise<string> {
    if (this.context.sdks?.microsoftTeams) {
      return this.context.isServedFromLocalhost
        ? "App rodando local no Teams"
        : "App rodando no Teams";
    }
  
    return this.context.isServedFromLocalhost
      ? "App rodando local no SharePoint"
      : "App rodando no SharePoint";
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
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
