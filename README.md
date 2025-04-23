# 📦 SharePoint WebParts - Projeto Corporativo

Este repositório contém a **solução principal de WebParts para SharePoint Online**, distribuída como um pacote `.sppkg`.  
As WebParts aqui desenvolvidas são utilizadas em diversos sites institucionais da empresa, oferecendo funcionalidades como:

- Diretório de Pessoas
- Aniversariantes
- Comunicação Interna
- Outras soluções corporativas em SharePoint

---

## 🛠️ Estrutura do Projeto

A raiz do projeto contém:

- Código-fonte das WebParts em TypeScript (SPFx - SharePoint Framework)
- Arquivos de configuração para o Visual Studio Code
- Scripts de build e deploy (`gulp`, `npm`)
- Pacote `.sppkg` gerado para publicação no App Catalog do SharePoint

---

## ⚙️ Configuração do Ambiente de Desenvolvimento

### 1. Pré-requisitos

- Node.js (v14.x ou v16.x - compatível com SPFx)
- npm ou Yarn
- Visual Studio Code
- Yeoman e SharePoint Generator

#### Instalação recomendada

```bash
npm install -g yo gulp
npm install -g @microsoft/generator-sharepoint
