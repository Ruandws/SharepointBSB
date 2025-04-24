import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './DiretorioWebPart.module.scss';
import * as strings from 'DiretorioWebPartStrings';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import defaultPhotoUrl from './assets/pessoa.png';

export interface IDiretorioPessoasWebPartProps {
  description: string;
}

export default class DiretorioPessoasWebPart extends BaseClientSideWebPart<IDiretorioPessoasWebPartProps> {
  private graphClient: MSGraphClientV3;
  private people: MicrosoftGraph.User[] = [];
  private filteredPeople: MicrosoftGraph.User[] = [];
  private readonly scrollBatchSize: number = 20;
  private scrollIndex: number = 0;

  public async render(): Promise<void> {
    this.domElement.innerHTML = `
      <section class="${styles.diretorio}">
        <div class="${styles.welcome}">
          <h2 class="${styles.welcomeTitle}">Diretório de Pessoas</h2>
          <p class="${styles.welcomeSubtitle}">Procurando por alguém?</p>
          <div class="${styles.searchContainer}">
            <input type="text" placeholder="Insira um nome, departamento ou cargo" id="searchInput" class="${styles.searchInput}" />
          </div>
        </div>
        <div id="peopleCards" class="${styles.peopleCards}">
          <!-- Cards de pessoas serão inseridos aqui via JavaScript -->
        </div>
      </section>`;

    try {
      await this.getPeopleWithLicenses();
      await this.renderPeopleCards();
      this.setupEventListeners();
    } catch (error) {
      console.error('Erro ao carregar pessoas:', error);
      this.showError();
    }
  }

  private showError(): void {
    this.domElement.innerHTML = `
      <div class="${styles.error}">
        Erro ao carregar o diretório. Por favor, tente novamente mais tarde.
        <button id="retryButton" class="${styles.retryButton}">Tentar novamente</button>
      </div>`;

    this.domElement.querySelector('#retryButton')?.addEventListener('click', () => this.render());
  }

  private setupEventListeners(): void {
    const searchInput = this.domElement.querySelector('#searchInput');

    searchInput?.addEventListener('input', this.debounce(async (e: Event) => {
      await this.filterPeople((e.target as HTMLInputElement).value);
    }, 300));

    const container = this.domElement.querySelector('#peopleCards');
    if (container) {
      container.addEventListener('scroll', async () => {
        const nearBottom = container.scrollTop + container.clientHeight >= container.scrollHeight - 50;
        if (nearBottom && this.scrollIndex < this.filteredPeople.length) {
          await this.renderPeopleCards(true);
        }
      });
    }
  }

  private debounce<F extends (...args: unknown[]) => void>(func: F, wait: number): (...args: Parameters<F>) => void {
    let timeout: number;
    return (...args: Parameters<F>) => {
      clearTimeout(timeout);
      timeout = window.setTimeout(() => func.apply(this, args), wait);
    };
  }

  private async filterPeople(searchTerm: string): Promise<void> {
    const container = this.domElement.querySelector('#peopleCards');
    if (!container) return;
  
    // Correção: interpolação de estilo usando crase (template literal)
    container.innerHTML = `<div class="${styles.loading}">Buscando...</div>`;
  
    try {
      const term = searchTerm.toLowerCase();
      this.filteredPeople = this.people.filter(person =>
        person.displayName?.toLowerCase().includes(term) ||
        person.jobTitle?.toLowerCase().includes(term) ||
        person.department?.toLowerCase().includes(term)
      );
  
      await this.renderPeopleCards();
    } catch (error) {
      console.error('Erro ao filtrar pessoas:', error);
      container.innerHTML = `<div class="${styles.error}">Erro na busca</div>`;
    }
  }

  private async renderPeopleCards(scrollLoad = false): Promise<void> {
    const container = this.domElement.querySelector('#peopleCards');
    if (!container) return;

    if (!scrollLoad) {
      container.innerHTML = '';
      this.scrollIndex = 0;
    }

    const end = this.scrollIndex + this.scrollBatchSize;
    const peopleToRender = this.filteredPeople.slice(this.scrollIndex, end);

    for (const person of peopleToRender) {
      const card = await this.createPersonCard(person);
      container.appendChild(card);
    }

    this.scrollIndex = end;
  }

  private async getUserPhotoBase64(userId: string): Promise<string | null> {
    try {
      // Verifica se o usuário tem foto antes de continuar
      const client = await this.context.msGraphClientFactory.getClient('3');
      try {
        await client.api(`/users/${userId}/photo`).get();
      } catch {
        return null; // Se não tiver foto, já retorna
      }
  
      // Obtém o token de autenticação corretamente
      const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      const token = await tokenProvider.getToken('https://graph.microsoft.com');
  
      if (!token) {
        console.error("Token de autenticação está vazio.");
        return null;
      }
  
      // Faz a requisição da foto em si
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/users/${userId}/photo/$value`,
        {
          headers: {
            'Authorization': `Bearer ${token}`
          }
        }
      );
  
      if (!response.ok) {
        throw new Error('Falha ao buscar a foto do usuário.');
      }
  
      const blob = await response.blob();
      return await this.blobToBase64(blob);
    } catch (error) {
      console.warn(`Erro ao obter a foto do usuário: ${error}`);
      return null;
    }
  }
  
  private blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.readAsDataURL(blob);
    });
  }

  private async createPersonCard(person: MicrosoftGraph.User): Promise<HTMLDivElement> {
    const card = document.createElement('div');
    card.className = styles.personCard;

    const email = person.mail || person.userPrincipalName || '';
    const userId = person.id || '';

    const photoSrc = await this.getUserPhotoBase64(userId) || defaultPhotoUrl;

    card.innerHTML = `
      <div>
        <img src="${photoSrc}" alt="Foto de ${person.displayName}" class="${styles.personPhoto}" />
        <h3>${person.displayName || 'Nome não disponível'}</h3>
        <p>${person.jobTitle || 'Não informado'}</p><br>
        <p>${person.department || 'Não informado'}</p><br>
        <a href="msteams:${email}" target="_blank" rael="noopener noreferrer" class="${styles.teamsButton}">
          <i class="ms-Icon ms-Icon--TeamsLogo"></i> Abrir Chat
        </a>
      </div>
    `;

    return card;
  }

  private async getPeopleWithLicenses(): Promise<void> {
    try {
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');

      let allUsers: MicrosoftGraph.User[] = [];
      let response = await this.graphClient
        .api('/users')
        .select('id,displayName,assignedLicenses,jobTitle,mail,department,officeLocation')
        .top(500)
        .get();

      allUsers = response.value;

      while (response['@odata.nextLink']) {
        response = await this.graphClient.api(response['@odata.nextLink']).get();
        allUsers = [...allUsers, ...response.value];
      }

      this.people = allUsers.filter(user =>
        user.assignedLicenses?.length &&
        user.department &&
        user.department.trim() !== '' &&
        user.department.toLowerCase().trim() !== 'não informado');
      this.filteredPeople = [...this.people];
    } catch (error) {
      console.error('Erro ao buscar usuários:', error);
      throw error;
    }
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyPaneTextField('description', {
              label: strings.DescriptionFieldLabel
            })
          ]
        }]
      }]
    };
  }
}