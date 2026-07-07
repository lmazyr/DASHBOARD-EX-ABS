# Dashboard Executivo SESE Logistica

Dashboard HTML/CSS/JS para acompanhar absenteismo a partir de uma planilha Excel.

## Como funciona

- No computador, rodando em `localhost`, o dashboard tenta ler o Excel da pasta do projeto.
- Publicado no GitHub Pages, o dashboard prioriza o link do OneDrive salvo no campo "Fonte dos Dados".
- A planilha Excel nao deve ir para o GitHub. Ela fica no computador sincronizando pelo OneDrive.

## Publicar no GitHub Pages

1. Crie um repositorio no GitHub.
2. Envie estes arquivos:
   - `index.html`
   - `style.css`
   - `app.js`
   - `logo.png`
   - `.gitignore`
   - `README.md`
3. Nao envie arquivos `.xlsx` ou `.xls`.
4. No GitHub, entre em `Settings > Pages`.
5. Em `Build and deployment`, selecione `Deploy from a branch`.
6. Escolha a branch `main` e a pasta `/root`.
7. Abra o link gerado pelo GitHub Pages.
8. Cole o link compartilhavel do Excel no campo "Fonte dos Dados" e clique em "Salvar Link".

## Atualizacao do Excel

Edite o Excel normalmente no computador dentro da pasta sincronizada do OneDrive. Quando o OneDrive concluir a sincronizacao, o dashboard publicado tentara buscar a versao atualizada no intervalo configurado.
