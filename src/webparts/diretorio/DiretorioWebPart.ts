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
  private currentPage: number = 0;
  private readonly cardsPerPage: number = 4;

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
        <div class="${styles.peopleContainer}">
          <button id="prevPage" class="${styles.navButton}">❮</button>
          <div id="peopleCards" class="${styles.peopleCards}"></div>
          <button id="nextPage" class="${styles.navButton}">❯</button>
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
    const prevPageButton = this.domElement.querySelector('#prevPage');
    const nextPageButton = this.domElement.querySelector('#nextPage');
    
    searchInput?.addEventListener('input', this.debounce(async (e: Event) => {
      await this.filterPeople((e.target as HTMLInputElement).value);
    }, 300));

    prevPageButton?.addEventListener('click', () => this.changePage(-1));
    nextPageButton?.addEventListener('click', () => this.changePage(1));
  }

  private async changePage(direction: number): Promise<void> {
    const newPage = this.currentPage + direction;
    const maxPage = Math.ceil(this.filteredPeople.length / this.cardsPerPage) - 1;
    
    if (newPage >= 0 && newPage <= maxPage) {
      this.currentPage = newPage;
      await this.renderPeopleCards();
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
    const term = searchTerm.toLowerCase();
    this.filteredPeople = this.people.filter(person => 
      person.displayName?.toLowerCase().includes(term) ||
      person.jobTitle?.toLowerCase().includes(term) ||
      person.department?.toLowerCase().includes(term)
    );
    this.currentPage = 0;
    await this.renderPeopleCards();
  }

  private async renderPeopleCards(): Promise<void> {
    const container = this.domElement.querySelector('#peopleCards');
    if (!container) return;

    container.innerHTML = '';
    const start = this.currentPage * this.cardsPerPage;
    const end = start + this.cardsPerPage;
    
    this.filteredPeople.slice(start, end).forEach(person => {
      container.appendChild(this.createPersonCard(person));
    });
  }

  private createPersonCard(person: MicrosoftGraph.User): HTMLDivElement {
    const card = document.createElement('div');
    card.className = styles.personCard;
    card.innerHTML = `
      <div>
        <img src="${defaultPhotoUrl}" alt="Foto de ${person.displayName}" class="${styles.personPhoto}" />
        <h3>${person.displayName || 'Nome não disponível'}</h3>
        <p><strong>Cargo:</strong> ${person.jobTitle || 'Não informado'}</p>
        <p><strong>Departamento:</strong> ${person.department || 'Não informado'}</p>
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

      this.people = allUsers.filter(user => user.assignedLicenses?.length);
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