## Resumo de Frete (WIDE e LONG)

Gera dois relatórios a partir de `BaseDadosFretes.xlsx` e `metas_fretes.xlsx`:
- `resumo_frete_metas.xlsx` (WIDE, com coluna `Meta` e meses)
- `resumo_fretes_long_pbi.csv` (LONG, ideal para Power BI)

### Instalação

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### Uso

Coloque os arquivos de entrada na pasta de trabalho (ex.: `~/Desktop/logística`):
- `BaseDadosFretes.xlsx` com colunas: `Departamento`, `Data`, `Valor`
- `metas_fretes.xlsx` com colunas: `Departamento`, `Meta`

Execute:

```bash
python resumo_frete_2025_final.py --pasta ~/Desktop/logística --ano 2025 --meta-ja-em-escala
# Se a meta estiver em porcentagem inteira (ex.: 3 -> 3%), use:
python resumo_frete_2025_final.py --pasta ~/Desktop/logística --ano 2025 --meta-nao-em-escala
```

Saídas serão criadas na mesma pasta:
- `resumo_frete_metas.xlsx`
- `resumo_fretes_long_pbi.csv`

Observações:
- O script tenta usar `xlsxwriter` (com formatação de porcentagens). Se indisponível, usa `openpyxl` sem formatação.
- CSV é salvo em `utf-8-sig` (BOM) para compatibilidade com Windows/Power BI.