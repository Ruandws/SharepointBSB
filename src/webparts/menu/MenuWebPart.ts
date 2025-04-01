import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './MenuWebPart.module.scss';
import * as strings from 'MenuWebPartStrings';

export interface IMenuWebPartProps {
  description: string;
}

interface IMenuItem {
  nome: string;
  link: string;
  imagemUrl: string;
}

export default class MenuWebPart extends BaseClientSideWebPart<IMenuWebPartProps> {

  private menuItems: IMenuItem[] = [
    { nome: "Diretoria", link: "https://lacerdapar.sharepoint.com/sites/Diretoria2", imagemUrl: require('./assets/diretoria.png') },
    { nome: "Jurídico", link: "https://lacerdapar.sharepoint.com/sites/Jurdico", imagemUrl: require('./assets/juridico.png') },
    { nome: "Financeiro", link: "https://lacerdapar.sharepoint.com/sites/Financeiro", imagemUrl: require('./assets/financeiro.png') },
    { nome: "Suporte TI", link: "https://lacerdapar.sharepoint.com/sites/SuporteTI2", imagemUrl: require('./assets/itsupport.png') },
    { nome: "Compras", link: "https://lacerdapar.sharepoint.com/sites/ComprasBraslia", imagemUrl: require('./assets/compras.png') },
    { nome: "Departamento Pessoal", link: "https://lacerdapar.sharepoint.com/sites/DP-Braslia/SitePages/ProjectHome.aspx", imagemUrl: require('./assets/dp.png') },
    { nome: "Comercial", link: "https://lacerdapar.sharepoint.com/sites/Comercial", imagemUrl: require('./assets/comercial.png') }
  ];

  public render(): void {
    const menuHtml = this.menuItems
      .map(item => this.renderMenuItem(item))
      .join('');

    this.domElement.innerHTML = `
      <div class="${styles.menuContainer}">
        ${menuHtml}
      </div>
    `;
  }

  private renderMenuItem(item: IMenuItem): string {
    return `
      <a href="${item.link}" class="${styles.menuItemCard}" target="_blank">
        <img src="${item.imagemUrl}" alt="${item.nome}" class="${styles.menuItemImagem}">
        <div class="${styles.menuItemNome}">${item.nome}</div>
      </a>
    `;
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