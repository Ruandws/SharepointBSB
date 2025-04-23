📦 SharePoint WebParts - Projeto Corporativo
Este repositório contém a solução principal de WebParts para SharePoint Online, distribuída como um pacote .sppkg. As WebParts aqui desenvolvidas são utilizadas em diversos sites institucionais da empresa, oferecendo funcionalidades como diretório de pessoas, aniversariantes, comunicação interna e muito mais.

🛠️ Estrutura do Projeto
A raiz do projeto contém:

Código-fonte das WebParts em TypeScript (SPFx - SharePoint Framework)

Arquivos de configuração do Visual Studio Code

Scripts de build e deploy (gulp, npm)

Arquivo .sppkg gerado para publicação no App Catalog do SharePoint

⚙️ Configuração do Ambiente de Desenvolvimento
1. Pré-requisitos
Node.js (v14.x ou v16.x, compatível com SPFx)

npm ou Yarn

Visual Studio Code

Yeoman e SharePoint Generator

Instalação recomendada:

bash
Copiar
Editar
npm install -g yo gulp
npm install -g @microsoft/generator-sharepoint
2. Clonando o repositório
bash
Copiar
Editar
git clone https://github.com/seu-usuario/seu-repositorio.git
cd seu-repositorio
3. Instalando dependências
Com npm:

bash
Copiar
Editar
npm install
Ou com yarn:

bash
Copiar
Editar
yarn install
🚀 Comandos úteis de desenvolvimento

Comando	Descrição
gulp serve	Inicia o servidor de desenvolvimento (localhost:4321)
gulp build	Compila o projeto
gulp bundle --ship	Cria os arquivos prontos para produção
gulp package-solution --ship	Gera o arquivo .sppkg para deploy
gulp clean	Limpa artefatos da build
gulp trust-dev-cert	Confia no certificado de dev local
✅ Publicando no SharePoint Online
Após rodar:

bash
Copiar
Editar
gulp bundle --ship
gulp package-solution --ship
O arquivo será gerado na pasta:

bash
Copiar
Editar
sharepoint/solution/*.sppkg
Publique esse pacote no App Catalog do seu tenant SharePoint.

💡 Dicas de Git

Comando	Ação
git clone <repo>	Clona o repositório
git pull	Atualiza sua branch local com as últimas mudanças
git status	Mostra arquivos modificados
git add .	Adiciona todas as alterações
git commit -m "Mensagem"	Cria um commit
git push	Envia as mudanças para o repositório remoto
git checkout -b nome-da-branch	Cria uma nova branch
git merge nome-da-branch	Mescla branch especificada com a atual
