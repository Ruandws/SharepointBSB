import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './MenuWebPart.module.scss';

export interface IMenuWebPartProps {
  menuItems: IMenuItem[];
  ocultarSetorAtual?: boolean;
}

export interface IMenuItem {
  nome: string;
  link: string;
  imagemUrl?: string;
  isVisible: boolean;
}

export default class MenuWebPart extends BaseClientSideWebPart<IMenuWebPartProps> {

  protected onInit(): Promise<void> {
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

    this.properties.menuItems.forEach(item => {
      item.isVisible = item.isVisible !== false;
    });

    return Promise.resolve();
  }

  public render(): void {
    const currentUrl = this.context.pageContext.web.absoluteUrl.toLowerCase();
    const visibleItems = this.properties.menuItems.filter(item => {
      const isCurrent = currentUrl.startsWith(item.link.toLowerCase());
      return item.isVisible && (!this.properties.ocultarSetorAtual || !isCurrent);
    });
  
    this.domElement.innerHTML = `
      <div class="${styles.menuWrapper}">
        <div class="${styles.menuContainer}">
          ${visibleItems.map(item => this.renderMenuItem(item)).join('')}
        </div>
      </div>`;
  }

  private renderMenuItem(item: IMenuItem): string {
    const imageSrc = item.imagemUrl || '';
    return `
      <a href="${item.link}" class="${styles.menuItemCard}" target="_blank">
        <img src="${imageSrc}" alt="${item.nome}" class="${styles.menuItemImagem}">
        <div class="${styles.menuItemNome}">${item.nome}</div>
      </a>
    `;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Configuração do Menu' },
          groups: [
            {
              groupName: 'Gerenciar Itens Existentes',
              groupFields: [
                ...this.properties.menuItems.map((item, index) => (
                  PropertyPaneToggle(`itemVisible_${index}`, {
                    label: item.nome,
                    checked: item.isVisible,
                    onText: 'Visível',
                    offText: 'Oculto'
                  })
                )),
                PropertyPaneToggle('ocultarSetorAtual', {
                  label: 'Ocultar setor atual',
                  checked: this.properties.ocultarSetorAtual,
                  onText: 'Sim',
                  offText: 'Não'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    if (propertyPath.startsWith('itemVisible_')) {
      const index = parseInt(propertyPath.split('_')[1], 10);
      if (this.properties.menuItems[index]) {
        this.properties.menuItems[index].isVisible = Boolean(newValue);
        this.render();
      }
    } else if (propertyPath === 'ocultarSetorAtual') {
      this.properties.ocultarSetorAtual = Boolean(newValue);
      this.render();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
