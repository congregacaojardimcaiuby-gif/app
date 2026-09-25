# Programação da congregação

App de programação de reuniões e escalas, com seis áreas:
reunião de meio de semana, reunião de fim de semana,
áudio/vídeo/volantes, indicadores, limpeza e relatórios mensais.

## Estrutura

| Arquivo | Função |
|---------|--------|
| `index.html` | O app completo (HTML + CSS + JS em um único arquivo) |
| `Codigo.gs` | Apps Script colado na planilha do Google — é a API |
| `manifest.webmanifest` | Metadados para instalação como PWA |
| `sw.js` | Service worker — cache e funcionamento offline |
| `icon-192.png` / `icon-512.png` | Ícones do app instalado |

## Como publicar

1. Faça o push deste repositório para o GitHub.
2. Em **Settings → Pages**, escolha Branch `main`, pasta `/ (root)`.
3. O app fica disponível em `https://seu-usuario.github.io/programacao-congregacao/`.

## Base de dados

Os dados ficam em uma planilha do Google. O endereço da planilha
está definido em `API_URL` dentro de `index.html`. Para trocar:
- Crie uma nova planilha
- Cole o conteúdo de `Codigo.gs` no Apps Script da planilha
- Publique como App da Web (Executar como: Eu; Acesso: Qualquer pessoa)
- Substitua o valor de `API_URL` em `index.html`

## Acesso de administrador

- **PIN mestre** (padrão `1234`): abre todas as áreas
- **PIN por área**: abre só a área correspondente
- Troque os PINs em Administração → Configurações gerais
