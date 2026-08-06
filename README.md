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

   Ou, sem usar terminal: no repositório vazio criado no GitHub, use
   **Add file → Upload files** e arraste para dentro do navegador todo o
   conteúdo desta pasta (as pastas `src` e `public`, e os arquivos
   `package.json`, `vite.config.js`, `index.html`, `vercel.json`,
   `.gitignore` e `README.md`) — não a pasta em si, e não a pasta
   `node_modules` (ela não existe aqui de propósito; é gerada
   automaticamente no build).

## 3. Publicar na Vercel

1. Crie uma conta em https://vercel.com — o mais simples é clicar em **Continue with GitHub** e autorizar com a mesma conta do passo anterior.
2. Clique em **Add New… → Project**.
3. Selecione o repositório `patrimonio-dashboard`. A Vercel detecta automaticamente que é um projeto Vite — não precisa mudar nada nas configurações de build.
4. Clique em **Deploy**. Em ~1 minuto você recebe uma URL do tipo `https://patrimonio-dashboard.vercel.app`.

A partir daí, todo `git push` (ou upload de arquivo novo) para `main` publica uma nova versão automaticamente.

## 4. Instalar como app no Galaxy S25+ (PWA)

1. Abra a URL da Vercel no **Chrome** do celular.
2. Toque no menu (⋮) no canto superior direito.
3. Toque em **Adicionar à tela inicial** (ou **Instalar app**, dependendo da versão do Chrome).
4. Confirme. O ícone aparece na tela inicial e abre em tela cheia, sem barra de endereço, como um app nativo.

Se você já tinha instalado uma versão anterior do app, é bom removê-la da
tela inicial e reinstalar depois de atualizar o site, para garantir que o
Chrome carregue a versão nova (o service worker atualiza sozinho, mas às
vezes o ícone antigo fica com um cache visual).

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

## v0.9.7 — donut de doações, patrimônio e gráficos

- Achado o bug real do gráfico de doações não aparecer: o código procurava por um cabeçalho de coluna "Campo"/"Mês" dentro da tabela `tbDoadosDistribuicao`, mas a tabela usa "Categoria" (igual à tabela de alocação de investimentos) — nunca batia. Corrigido para procurar "Categoria", testado direto contra sua planilha real.
- Card do patrimônio (Home): o indicador de variação (%) agora fica abaixo do valor, no mesmo padrão do card "Investido" da tela Wallet, em vez de flutuar ao lado.
- Gráficos de linha (patrimônio na Home, investido na Wallet): adicionada uma grade horizontal discreta, eixo X mostrando só 1 mês a cada 2 (evita a "embolação" dos 12 meses), e eixo Y com menos marcações — visual mais limpo e organizado.

## v0.9.6 — largura dos cards pareados

- Achada a causa real da largura desigual entre cards lado a lado (ex: "Em dólar" menor, "Dividendos no ano" maior, sobrando espaço à direita): o `flex: 1` estava sendo passado como propriedade de estilo para o componente `CardWrapper`, mas esse valor caía no elemento interno (o "cartão visível"), e não no elemento externo que de fato participa da divisão de espaço na linha (`display: flex`). Corrigido para o elemento externo do `CardWrapper` sempre dividir o espaço igualmente com seus vizinhos.

## v0.9.5 — fontes locais + cards desalinhados

- Fontes Tomorrow e Inter agora estão embutidas no projeto (`public/fonts`, `.woff2`), com `@font-face` declarado no `index.html`. Não depende mais do Google Fonts em runtime — carrega mais rápido e funciona offline (o service worker faz o cache delas).
- Corrigido: quando dois cards ficavam lado a lado com conteúdo de tamanhos diferentes (ex: "Investido" / "Faturamento"), o card com menos conteúdo não esticava para acompanhar a altura do vizinho, deixando as bordas inferiores desalinhadas. Agora ambos preenchem a altura da linha corretamente.

## v0.9.4 — ajustes de tela cheia mobile

- O layout deixou de ter altura fixa (era pensado para a prévia do artifact) e agora ocupa `100dvh`, a altura real da tela do celular.
- As bordas brancas/arredondadas eram do "card" de prévia — removidas para o app ocupar a tela inteira.
- Os dois cards lado a lado que estouravam a tela (ex: "Em dólar" / "Dividendos no ano") tinham largura fixa em pixels; agora se ajustam proporcionalmente à largura da tela.
