# Patrimônio Dashboard

Dashboard pessoal de investimentos e patrimônio, feito em React + Vite. Lê os
dados a partir de um arquivo `.xlsx` (planilha com aba "API") escolhido
manualmente na tela — nada é enviado a nenhum servidor, o arquivo é lido
direto no navegador.

Este projeto já está pronto para ser publicado no GitHub + Vercel e
instalado como app (PWA) no celular.

## 1. Testar localmente (opcional)

```bash
npm install
npm run dev
```

Abre em `http://localhost:5173`.

## 2. Publicar no GitHub

1. Crie uma conta em https://github.com (se ainda não tiver).
2. No GitHub, clique em **New repository**, dê um nome (ex: `patrimonio-dashboard`)
   e deixe como **Private** (recomendado, já que é um app financeiro pessoal).
3. Na sua máquina, dentro desta pasta, rode:

   ```bash
   git init
   git add .
   git commit -m "Primeira versão do dashboard"
   git branch -M main
   git remote add origin https://github.com/SEU-USUARIO/patrimonio-dashboard.git
   git push -u origin main
   ```

   (o GitHub mostra esses comandos exatos na página do repositório recém-criado)

## 3. Publicar na Vercel

1. Crie uma conta em https://vercel.com — o mais simples é clicar em **Continue with GitHub** e autorizar com a mesma conta do passo anterior.
2. Clique em **Add New… → Project**.
3. Selecione o repositório `patrimonio-dashboard`. A Vercel detecta automaticamente que é um projeto Vite — não precisa mudar nada nas configurações de build.
4. Clique em **Deploy**. Em ~1 minuto você recebe uma URL do tipo `https://patrimonio-dashboard.vercel.app`.

A partir daí, todo `git push` para `main` publica uma nova versão automaticamente.

## 4. Instalar como app no Galaxy S25+ (PWA)

1. Abra a URL da Vercel no **Chrome** do celular.
2. Toque no menu (⋮) no canto superior direito.
3. Toque em **Adicionar à tela inicial** (ou **Instalar app**, dependendo da versão do Chrome).
4. Confirme. O ícone aparece na tela inicial e abre em tela cheia, sem barra de endereço, como um app nativo.

## 5. Usando os dados da planilha

Na primeira vez que abrir o app (no celular ou no computador), toque no botão de
upload e selecione o arquivo `.xlsx` exportado da sua planilha (aba "API").
O app guarda os dados localmente no aparelho (`localStorage`), então nas
próximas vezes ele abre direto com os últimos dados carregados — só precisa
enviar a planilha de novo quando quiser atualizar os números.

Cada navegador/aparelho guarda seus próprios dados localmente — ou seja, se
carregar a planilha no celular, ela não aparece automaticamente no
computador (e vice-versa). É preciso enviar o arquivo em cada dispositivo.

## Estrutura do projeto

```
src/App.jsx        → o dashboard (toda a lógica e visual)
src/main.jsx        → ponto de entrada + compatibilidade de armazenamento local
vite.config.js       → build + configuração do PWA (ícones, manifest, service worker)
public/icons/         → ícones do app (192px e 512px)
```
