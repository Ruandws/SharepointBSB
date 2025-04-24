import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from "@pnp/sp";
import styles from './BoasvindasWebPart.module.scss';
import "@pnp/sp/webs";
import "@pnp/sp/site-users";

export default class BoasVindasWebPart extends BaseClientSideWebPart<{}> {
  private _sp: ReturnType<typeof spfi>;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context)); // Inicializa corretamente
  }

  private async getCurrentUser(): Promise<string> {
    try {
      const user = await this._sp.web.currentUser(); // Corrigido para usar _sp
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

  public async render(): Promise<void> {
    const userName = await this.getCurrentUser();
  
    const pageTitle = document.title || "Página Atual"; // ou defina manualmente
  
    this.domElement.innerHTML = `
      <section class="${styles.banner}">
        <div class="${styles.textContainer}">
          <h2 class="${styles.pageTitle}">${pageTitle}</h2>
          <h1 class="${styles.welcomeText}">Bem-vindo, ${userName}</h1>
          <p class="${styles.subText}">Brasília Segurança S/A <br> 2025</p>
        </div>
      </section>
    `;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
