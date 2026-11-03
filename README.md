# 🌸 Festa da Primavera 2026 — Controle de Mesas

**E.M. Deputado Federal Ulisses Guimarães** · Festa em **12/11/2026**

Página para controlar a venda das 50 mesas da festa. As vendas ficam salvas na planilha
**FESTA DA PRIMAVERA 2026 - VENDAS DE MESAS** no Google Drive e aparecem para todos em poucos segundos.

- **Mesas** (`/#mesas`): quantas mesas foram vendidas, quantas restam e o mapa das mesas. Quando todas forem vendidas aparece o aviso **"Mesas esgotadas!"**.
- **Área de vendas** (`/#vendas`): registrar venda (mesa, valor, vendedor, comprador, aluno, série, turma, pagamento), total recebido por vendedor (Dinheiro / PIX), lista com busca, cancelamento, comprovante para WhatsApp e planilha.

## Arquivos
| Arquivo | Para quê |
|---|---|
| `index.html` | A página |
| `config.js` | Configuração (endereço do script, valor da mesa, vendedores, séries) |
| `store.js` | Leitura e gravação na planilha |
| `apps-script/Codigo.gs` | Script do Google que grava na planilha (não precisa ir para o GitHub) |

## Ligar a planilha
1. Abra a planilha no Drive → **Extensões → Apps Script**.
2. Apague o conteúdo e cole o arquivo `apps-script/Codigo.gs`. Salve.
3. **Implantar → Nova implantação → App da Web** — Executar como: *Eu* · Quem pode acessar: *Qualquer pessoa*. Autorize.
4. Copie a URL que termina em `/exec` e cole em `config.js`, no campo `APPS_SCRIPT_URL`.

## Publicar no GitHub
Repositório `festa-primavera/festa-primavera.github.io` — publicado em https://festa-primavera.github.io
