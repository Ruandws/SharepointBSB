import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import styles from './Aniversariantes2WebPart.module.scss';
import * as strings from 'Aniversariantes2WebPartStrings';

export interface IAniversariantes2WebPartProps {
    description: string;
}

interface IAniversariante2 {
    nome: string;
    cargo: string;
    fotoUrl: string;
    dataAniversario: Date;
    nomeDoAD?: string;
}

interface User {
    displayName: string;
    jobTitle: string;
    birthday: string;
    id: string;
    userPrincipalName: string;
}

interface UsersResponse {
    value: User[];
}

export default class Aniversariantes2WebPart extends BaseClientSideWebPart<IAniversariantes2WebPartProps> {

    private aniversariantes: IAniversariante2[] = [];

    public async render(): Promise<void> {
        try {
            const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient("3");
            await this.buscarAniversariantesDoAD(graphClient);
        } catch (error) {
            console.error("Erro ao obter dados do Graph API:", error);
        }
    }

    private async buscarAniversariantesDoAD(graphClient: MSGraphClientV3): Promise<void> {
        try {
            const usersResponse: UsersResponse = await graphClient.api('/users').get();
            const aniversariantesDoAD = usersResponse.value.map((user: User): IAniversariante2 => {
                return {
                    nome: user.displayName,
                    cargo: user.jobTitle,
                    fotoUrl: require('./assets/pessoa.png'),
                    dataAniversario: new Date(user.birthday),
                    nomeDoAD: user.userPrincipalName,
                };
            });

            this.aniversariantes = aniversariantesDoAD;
            await this.atualizarFotosDosAniversariantes(graphClient);
        } catch (error) {
            console.error("Erro ao buscar aniversariantes do AD:", error);
        }
    }

    private async atualizarFotosDosAniversariantes(graphClient: MSGraphClientV3): Promise<void> {
        const fotosPromises = this.aniversariantes.map(async (aniversariante: IAniversariante2) => {
            try {
                const photo: ArrayBuffer = await graphClient.api(`/users/${aniversariante.nomeDoAD}/photo/$value`).get();
                const blob = new Blob([photo], { type: 'image/jpeg' });
                return URL.createObjectURL(blob);
            } catch (error) {
                console.error(`Erro ao buscar foto para ${aniversariante.nomeDoAD}:`, error);
                return aniversariante.fotoUrl; // Retorna a foto padrão em caso de erro
            }
        });

        const fotos = await Promise.all(fotosPromises);

        this.aniversariantes = this.aniversariantes.map((aniversariante, index) => {
            return { ...aniversariante, fotoUrl: fotos[index] };
        });

        this.renderizarAniversariantes();
    }

    private renderizarAniversariantes(): void {
        const aniversariantesHtml = this.aniversariantes
            .sort((a, b) => a.dataAniversario.getTime() - b.dataAniversario.getTime())
            .slice(0, 5)
            .map((aniversariante, index) => this.renderAniversariante(aniversariante, index))
            .join('');

        this.domElement.innerHTML = `
            <div class="${styles.aniversariantes}">
                <div class="${styles.tituloAniversariantes}">Aniversariantes do Mês</div>
                <div class="${styles.subTituloAniversariantes}">Brasília segurança</div>
                ${aniversariantesHtml}
            </div>
        `;
    }

    private obterDataSemAno(data: Date): Date {
        return new Date(new Date().getFullYear(), data.getMonth(), data.getDate());
    }

    private saoMesmoDia(data1: Date, data2: Date): boolean {
        return (
            data1.getDate() === data2.getDate() &&
            data1.getMonth() === data2.getMonth()
        );
    }

    private renderAniversariante(aniversariante: IAniversariante2, index: number): string {
      const hoje = this.obterDataSemAno(new Date());
      const amanha = new Date(hoje);
      amanha.setDate(hoje.getDate() + 1);
      const seteDiasDepois = new Date(hoje);
      seteDiasDepois.setDate(hoje.getDate() + 7);
  
      const dataAniversario = this.obterDataSemAno(aniversariante.dataAniversario);
  
      let mensagem = '';
      if (this.saoMesmoDia(dataAniversario, hoje)) {
          mensagem = ' Feliz Aniversário!!';
      } else if (this.saoMesmoDia(dataAniversario, amanha)) {
          mensagem = 'Será';
      } else if (dataAniversario > amanha && dataAniversario <= seteDiasDepois) {
          const dias = Math.ceil((dataAniversario.getTime() - hoje.getTime()) / (1000 * 3600 * 24));
          mensagem = ` Faltam ${dias} dias`;
      }
      const iconeUrl = this.saoMesmoDia(dataAniversario, hoje) ? require('./assets/birthdaycake.png') : require('./assets/ellipsis.png');
  
      // Lógica para extrair primeiro e último nome
      const nomes = aniversariante.nome.split(' ');
      const primeiroNome = nomes[0];
      const ultimoNome = nomes.length > 1 ? nomes[nomes.length - 1] : '';
      const nomeExibicao = ultimoNome ? `${primeiroNome} ${ultimoNome}` : primeiroNome;
  
      return `
          <div class="${styles.aniversarianteCard}">
              <img src="${aniversariante.fotoUrl}" alt="${nomeExibicao}" class="${styles.foto}">
              <div class="${styles.info}">
                  <div class="${styles.nome}">${nomeExibicao}</div>
                  <div class="${styles.cargo}">${aniversariante.cargo}</div>
                  ${mensagem ? `<div class="${styles.mensagem}">${mensagem}</div>` : ''}
              </div>
              <div class="${styles.staticon}">
                  <img src="${iconeUrl}" alt="Icone" class="${styles.icone}">
              </div>
          </div>
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