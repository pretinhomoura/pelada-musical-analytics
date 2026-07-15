# Pelada Musical Analytics — Instagram + YouTube + Spotify

Painel único para acompanhar as 3 redes todo mês com os sócios e mostrar aos patrocinadores.

## O que tem aqui
- `index.html`, `style.css`, `app.js` — o painel (site estático, sem servidor).
- `dados.xlsx` — a planilha com 3 abas: **Instagram**, **YouTube**, **Spotify** (mais "Resumo" e "Como Atualizar").
- `manifest.webmanifest` + `assets/logo.png` — para poder "instalar" o painel como app no celular/computador.

## Como colocar no ar (grátis, ~2 minutos) para abrir de qualquer lugar
1. Acesse **https://app.netlify.com/drop**.
2. Arraste esta pasta inteira (com os 5 arquivos + pasta `assets`) para a página.
3. A Netlify gera um link público (tipo `https://pelada-musical.netlify.app`) — esse é o endereço que você, seus sócios e os patrocinadores vão usar, de qualquer navegador, celular ou computador.
4. (Opcional) Em "Site settings" você troca o nome do link e pode conectar um domínio próprio (ex: `analytics.peladamusical.com.br`).

Se preferir GitHub Pages, o processo é parecido: suba estes arquivos num repositório e ative o Pages nas configurações do repositório.

## Como atualizar todo mês
1. Abra `dados.xlsx`.
2. Vá até a aba da rede (Instagram, YouTube ou Spotify) e adicione uma linha nova no final com o mês.
3. **Não** mude os nomes das colunas nem das abas.
4. Salve mantendo o nome `dados.xlsx`.
5. Suba o arquivo atualizado no mesmo lugar (Netlify: arraste a pasta de novo, ou ative deploy contínuo puxando de uma pasta/GitHub).
6. Abra o painel — ele já confere sozinho a cada 5 minutos, ou clique em "Atualizar agora".

Dica: para conferir os números antes de publicar, use "Abrir outra planilha" no painel — ele lê o arquivo direto do seu computador, sem precisar subir em lugar nenhum.

## Trocar a logo
Troque o arquivo `assets/logo.png` por um PNG quadrado com a logo da Pelada Musical (recomendado 256x256 ou maior).

## Estrutura da planilha
Cada aba tem sua própria régua de métricas, porque cada rede reporta coisas diferentes:

- **Instagram**: Visualizações, % Não seguidores, Seguidores, Novos seguidores, Variação de views, Observação.
- **YouTube**: Visualizações, Horas assistidas, Inscritos, Novos inscritos, Vídeos publicados, Receita estimada, Observação.
- **Spotify**: Streams, Ouvintes mensais, Seguidores, Novos seguidores, Horas ouvidas, Adições em playlists, Observação.

O painel calcula sozinho a variação % mês a mês quando ela não é informada, então mesmo preenchendo só os números básicos o gráfico funciona.

## Botão "Exportar mês"
Gera uma página de relatório (via impressão do navegador — escolha "Salvar como PDF" na janela de impressão) com o resumo do mês mais recente das 3 redes. Ótimo para mandar para patrocinadores.
