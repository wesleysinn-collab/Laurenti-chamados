# Portal de Abertura de Chamados

Site estático (sem servidor). Abre direto no navegador ou pelo GitHub Pages.

## Arquivos
| Arquivo | Para quê |
|---|---|
| `index.html` | O portal |
| `dados.js` | Lojas e selfs (gerado a partir das planilhas) |
| `correcoes.js` | Correções manuais que valem por cima da planilha (ex.: cidade que falta). **Pode editar direto no GitHub.** |
| `atualizar_dados.py` | Regera o `dados.js` quando as planilhas mudarem |

## Publicar no GitHub Pages
1. Crie um repositório (de preferência **privado** ou da organização) e suba os 4 arquivos.
2. *Settings → Pages → Branch: main / root* → Save.
3. Compartilhe o link com o time.

## Atualizar os dados
```
pip install openpyxl
python atualizar_dados.py "Cadastro de Lojas.xlsx" "Tabela selfs.xlsx"
```
Faça commit do `dados.js` gerado. O `correcoes.js` nunca é sobrescrito.
