import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import * as strings from 'Aniversariantes2WebPartStrings';
import styles from './Aniversariantes2WebPart.module.scss';

export interface IAniversariantes2WebPartProps {
  title: string;
  listName: string;
  maxResults: number;
}

interface IAniversariante {
  nome: string;
  cargo: string;
  fotoUrl: string;
  dataAniversario: Date;
  isBirthdayToday: boolean;
  mensagemAniversario: string;
  iconeAniversario: string;
  diasRestantes: number;
}

interface ISPListItem {
  Title: string;
  field_1: number;
  field_2: number;
  field_3?: string;
}

export default class Aniversariantes2WebPart extends BaseClientSideWebPart<IAniversariantes2WebPartProps> {
  private defaultPhotoUrl = require('./assets/pessoa.png');
  private birthdayIconUrl = require('./assets/birthdaycake.png');
  private ellipsisIconUrl = require('./assets/ellipsis.png');
  private today = new Date();
  private usuariosSemFoto: string[] = [];

  public async onInit(): Promise<void> {
    await super.onInit();
    this.today = new Date(this.today.getFullYear(), this.today.getMonth(), this.today.getDate());
  }

  public async render(): Promise<void> {
    this.domElement.innerHTML = `
      <div class="${styles.aniversariantes}">
        <div class="${styles.header}">
          <h2 class="${styles.title}">${this.properties.title || 'Aniversariantes do Mês'}</h2>
          <p class="${styles.subtitle}">Brasília Segurança 2025</p>
        </div>
        <div id="aniversariantesContainer" class="${styles.container}"></div>
      </div>`;

    try {
      const aniversariantes = await this.carregarAniversariantes();
      this.renderizarAniversariantes(aniversariantes);

      if (this.usuariosSemFoto.length > 0) {
        console.warn(`⚠️ ${this.usuariosSemFoto.length} usuário(s) sem foto no AD.`);
        console.table(this.usuariosSemFoto);
      }
    } catch (error) {
      console.error("Erro ao carregar aniversariantes:", error);
      this.mostrarErro(strings.ErrorMessage);
    }
  }

  private async carregarAniversariantes(): Promise<IAniversariante[]> {
    const fromSharePoint = await this.carregarDoSharePoint().catch(() => []);
    return this.processarAniversariantes(fromSharePoint);
  }

  private async carregarDoSharePoint(): Promise<IAniversariante[]> {
    if (!this.properties.listName) return [];

    try {
      const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.properties.listName)}')`;
      const listExists = await this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);

      if (!listExists.ok) throw new Error("Lista não encontrada");

      const url = `${listUrl}/items?$select=Title,field_1,field_2,field_3`;
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);

      if (!response.ok) throw new Error("Erro ao buscar dados da lista");

      const data = await response.json();
      return await Promise.all(data.value.map((item: ISPListItem) => this.criarAniversarianteSP(item)));
    } catch (error) {
      console.error("Erro ao acessar SharePoint:", error);
      return [];
    }
  }

  private async criarAniversarianteSP(item: ISPListItem): Promise<IAniversariante> {
    const birthDate = new Date(this.today.getFullYear(), item.field_2 - 1, item.field_1);
    const diasRestantes = this.calcularDiferencaDias(birthDate);
    const { mensagem, icone } = this.obterMensagemEIcone(diasRestantes, birthDate);

    const fotoUrl = await this.obterFotoDoUsuario(item.Title);

    return {
      nome: item.Title,
      cargo: item.field_3 || strings.DefaultJobTitle,
      fotoUrl: fotoUrl || this.defaultPhotoUrl,
      dataAniversario: birthDate,
      isBirthdayToday: diasRestantes === 0,
      mensagemAniversario: mensagem,
      iconeAniversario: icone,
      diasRestantes: diasRestantes
    };
  }

  private async obterFotoDoUsuario(userId: string): Promise<string | null> {
    try {
      const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      const token = await tokenProvider.getToken("https://graph.microsoft.com");
  
      const response = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}/photo/$value`, {
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });

      if (!response.ok) {
        if (response.status === 404) {
          this.usuariosSemFoto.push(userId);
          return null;
        } else {
          console.error(`Erro ao buscar foto de ${userId}: ${response.statusText}`);
          return null;
        }
      }

      const arrayBuffer = await response.arrayBuffer();
      const base64 = this.arrayBufferToBase64(arrayBuffer);
      return `data:image/jpeg;base64,${base64}`;
    } catch (error) {
      console.error(`Erro ao obter foto de ${userId}:`, error);
      return null;
    }
  }

  private arrayBufferToBase64(buffer: ArrayBuffer): string {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
  }

  private calcularDiferencaDias(dataAniversario: Date): number {
    const aniv = new Date(dataAniversario);
    aniv.setFullYear(this.today.getFullYear());

    if (aniv < this.today) {
      aniv.setFullYear(this.today.getFullYear() + 1);
    }

    const diffTime = aniv.getTime() - this.today.getTime();
    return Math.floor(diffTime / (1000 * 60 * 60 * 24));
  }

  private obterMensagemEIcone(diasRestantes: number, dataAniversario: Date): { mensagem: string, icone: string } {
    if (diasRestantes === 0) {
      return {
        mensagem: "🎉 Feliz Aniversário!",
        icone: this.birthdayIconUrl
      };
    }

    const jaPassouEsteMes = dataAniversario.getMonth() === this.today.getMonth() && 
                            dataAniversario.getDate() < this.today.getDate();

    if (jaPassouEsteMes) {
      return {
        mensagem: "Aproveite seu mês! 😊",
        icone: this.ellipsisIconUrl
      };
    }

    if (diasRestantes === 1) {
      return {
        mensagem: "Amanhã é o grande dia!",
        icone: this.ellipsisIconUrl
      };
    } else if (diasRestantes > 1 && diasRestantes <= 7) {
      return {
        mensagem: `Faltam ${diasRestantes} dias!`,
        icone: this.ellipsisIconUrl
      };
    } else if (diasRestantes > 7 && dataAniversario.getMonth() === this.today.getMonth()) {
      return {
        mensagem: "Ainda este mês...",
        icone: this.ellipsisIconUrl
      };
    }

    return {
      mensagem: "",
      icone: ""
    };
  }

  private async processarAniversariantes(aniversariantes: IAniversariante[]): Promise<IAniversariante[]> {
    const comDatasValidas = aniversariantes.filter(a =>
      a.dataAniversario instanceof Date && !isNaN(a.dataAniversario.getTime())
    );

    const nomesUnicos = new Set<string>();
    const unique = comDatasValidas.filter(a => {
      if (!nomesUnicos.has(a.nome)) {
        nomesUnicos.add(a.nome);
        return true;
      }
      return false;
    });

    return unique
      .filter(a => a.dataAniversario.getMonth() === this.today.getMonth())
      .sort((a, b) => a.dataAniversario.getDate() - b.dataAniversario.getDate())
      .slice(0, Math.max(1, Number(this.properties.maxResults) || 5));
  }

  private renderizarAniversariantes(aniversariantes: IAniversariante[]): void {
    const container = this.domElement.querySelector('#aniversariantesContainer');
    if (!container) return;

    if (aniversariantes.length === 0) {
      container.innerHTML = `<div class="${styles.empty}">${strings.NoBirthdays}</div>`;
      return;
    }

    container.innerHTML = aniversariantes.map(pessoa => this.renderCard(pessoa)).join('');
  }

  private renderCard(pessoa: IAniversariante): string {
    const dataFormatada = `${pessoa.dataAniversario.getDate().toString().padStart(2, '0')}/$${
      (pessoa.dataAniversario.getMonth() + 1).toString().padStart(2, '0')
    }`;

    return `
      <div class="${styles.card} ${pessoa.isBirthdayToday ? styles.highlight : ''}">
        <div class="${styles.foto}">
          <img src="${pessoa.fotoUrl}" alt="${pessoa.nome}" class="${styles.photo}" onerror="this.src='${this.defaultPhotoUrl}'">
        </div>

        <div class="${styles.info}">
          <h3 class="${styles.name}">${pessoa.nome}</h3>
          <p class="${styles.jobTitle}">${pessoa.cargo}</p>
          <div class="${styles.messageRow}">
            <p class="${styles.message}">${pessoa.mensagemAniversario}</p>
            <p class="${styles.date}">${dataFormatada}</p>
          </div>
        </div>

        <div class="${styles.lateral}">
          ${pessoa.iconeAniversario
            ? `<img src="${pessoa.iconeAniversario}" class="${styles.badge}" alt="Ícone">`
            : `<span class="${styles.ellipsis}">•••</span>`}
        </div>
      </div>`;
  }

  private mostrarErro(mensagem: string): void {
    const container = this.domElement.querySelector('#aniversariantesContainer') || this.domElement;
    container.innerHTML = `<div class="${styles.error}">${mensagem}</div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Configuração da Lista" },
          groups: [
            {
              groupName: "Fonte de dados",
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: "Nome da lista de aniversariantes",
                  placeholder: "Digite o nome da lista"
                }),
                PropertyPaneTextField('maxResults', {
                  label: "Número máximo de resultados",
                  value: '5',
                  description: "Quantidade máxima de aniversariantes a serem exibidos"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}