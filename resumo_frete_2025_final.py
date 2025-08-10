#!/usr/bin/env python3

"""
Resumo de Frete por Departamento (WIDE e LONG)

- Lê dois arquivos Excel na pasta indicada:
  - BaseDadosFretes.xlsx: colunas esperadas -> Departamento, Data, Valor
  - metas_fretes.xlsx:   colunas esperadas -> Departamento, Meta (ex.: 3% lido como 0.03 se já em escala 0..1)

- Gera duas saídas:
  - resumo_frete_metas.xlsx (WIDE, com porcentagens por mês e coluna Meta)
  - resumo_fretes_long_pbi.csv (LONG, ideal para Power BI)

Uso por linha de comando (exemplos):
  python resumo_frete_2025_final.py --pasta ~/Desktop/logística --ano 2025 --meta-ja-em-escala
  python resumo_frete_2025_final.py --pasta /caminho/para/pasta --ano 2025 --meta-nao-em-escala

Observações:
- Se a engine xlsxwriter não estiver disponível, será usado openpyxl para gravar Excel (sem formatação de porcentagens).
- O script cria a pasta de saída se não existir.
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
from typing import Tuple

import numpy as np
import pandas as pd


# ================= CONFIG PADRÃO =================
DEFAULT_PASTA = os.path.expanduser(r"~/Desktop/logística")
DEFAULT_ANO = 2025
DEFAULT_META_JA_EM_ESCALA_0_A_1 = True  # meta vem como 3% -> 0.03
# ================================================


def escolher_engine_excel() -> Tuple[str, bool]:
    """Escolhe engine para Excel. Tenta xlsxwriter; se indisponível, usa openpyxl.

    Retorna:
      (engine, suporta_formatacao)
    """
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter", True
    except Exception:
        try:
            import openpyxl  # noqa: F401
            return "openpyxl", False
        except Exception:
            # Última tentativa: pandas tentará o padrão; sem garantia
            return "xlsxwriter", False


def validar_colunas(df: pd.DataFrame, colunas: list[str], nome: str) -> None:
    ausentes = [c for c in colunas if c not in df.columns]
    if ausentes:
        raise ValueError(
            f"Arquivo '{nome}' está faltando as colunas: {', '.join(ausentes)}. "
            f"Colunas encontradas: {list(df.columns)}"
        )


def processar_resumo_frete(
    pasta: str,
    ano: int,
    meta_ja_em_escala_0_a_1: bool,
) -> Tuple[pd.DataFrame, pd.DataFrame, str, str]:
    pasta_path = Path(os.path.expanduser(pasta)).resolve()
    pasta_path.mkdir(parents=True, exist_ok=True)

    arq_frete = pasta_path / "BaseDadosFretes.xlsx"
    arq_metas = pasta_path / "metas_fretes.xlsx"

    saida_wide = pasta_path / "resumo_frete_metas.xlsx"
    saida_long = pasta_path / "resumo_fretes_long_pbi.csv"

    if not arq_frete.exists():
        raise FileNotFoundError(f"Não encontrado: {arq_frete}")
    if not arq_metas.exists():
        raise FileNotFoundError(f"Não encontrado: {arq_metas}")

    # ---------- Leitura e limpeza ----------
    df = pd.read_excel(arq_frete)
    metas = pd.read_excel(arq_metas)

    validar_colunas(df, ["Departamento", "Data", "Valor"], arq_frete.name)
    validar_colunas(metas, ["Departamento", "Meta"], arq_metas.name)

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0.0)
    df = df.dropna(subset=["Data", "Departamento"])  # elimina linhas com Data/Departamento vazios

    # Ano/Mês
    df["Ano"] = df["Data"].dt.year
    df["MesNum"] = df["Data"].dt.month

    df = df[df["Ano"] == ano]
    if df.empty:
        raise ValueError(f"Não há lançamentos para o ano {ano} em {arq_frete}")

    # Totais do mês (para % de participação)
    totais_mes = (
        df.groupby(["Ano", "MesNum"], as_index=False)["Valor"].sum()
        .rename(columns={"Valor": "TotalMes"})
    )
    base = df.groupby(["Ano", "MesNum", "Departamento"], as_index=False)["Valor"].sum()
    base = base.merge(totais_mes, on=["Ano", "MesNum"], how="left")
    base["Percentual"] = np.where(base["TotalMes"] > 0, base["Valor"] / base["TotalMes"], 0.0)

    # Meses (nomes pt-BR)
    mapa_meses = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun", 7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}
    ordem = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    base["Mes"] = base["MesNum"].map(mapa_meses)

    # ---------- Metas ----------
    metas = metas[["Departamento", "Meta"]].copy()
    metas["Meta"] = pd.to_numeric(metas["Meta"], errors="coerce").fillna(0.0)
    if not meta_ja_em_escala_0_a_1:
        metas["Meta"] = metas["Meta"] / 100.0

    # =========================================================
    # 1) WIDE (igual ao modelo) -> Excel com porcentagem
    # =========================================================
    tabela = (
        base.pivot_table(
            index="Departamento",
            columns="Mes",
            values="Percentual",
            aggfunc="sum",
            fill_value=0.0,
        )
        .rename_axis(None, axis=1)
    )
    tabela = tabela.reindex(columns=[m for m in ordem if m in tabela.columns])

    resumo_wide = (
        metas.set_index("Departamento").join(tabela, how="outer").fillna(0.0).reset_index()
    )
    resumo_wide = resumo_wide.sort_values(by="Departamento")

    # Grava Excel
    engine, suporta_fmt = escolher_engine_excel()
    if suporta_fmt:
        with pd.ExcelWriter(saida_wide, engine=engine) as xw:
            resumo_wide.to_excel(xw, sheet_name="Resumo", index=False)
            wb = xw.book
            ws = xw.sheets["Resumo"]
            fmt_hdr = wb.add_format({"bold": True})
            fmt_pct = wb.add_format({"num_format": "0%"})
            ws.set_row(0, None, fmt_hdr)
            cols_pct = ["Meta"] + [c for c in ordem if c in resumo_wide.columns]
            for col in cols_pct:
                idx = resumo_wide.columns.get_loc(col)
                ws.set_column(idx, idx, 12, fmt_pct)
            ws.set_column(0, 0, 28)  # Departamento
    else:
        # Fallback sem formatação avançada
        resumo_wide.to_excel(saida_wide, index=False)

    print(f"✅ WIDE salvo: {saida_wide}")

    # =========================================================
    # 2) LONG (Power BI) -> CSV numérico para regras
    # =========================================================
    resumo_long = base[["Departamento", "Ano", "MesNum", "Percentual"]].copy()
    resumo_long = resumo_long.merge(metas, on="Departamento", how="left")
    resumo_long = resumo_long.sort_values(["Departamento", "MesNum"]).reset_index(drop=True)
    resumo_long.rename(columns={"MesNum": "Mes"}, inplace=True)

    # Garante tipos: Percentual/Meta em 0..1 (decimais), Mes inteiro
    resumo_long["Percentual"] = pd.to_numeric(resumo_long["Percentual"], errors="coerce").fillna(0.0)
    resumo_long["Meta"] = pd.to_numeric(resumo_long["Meta"], errors="coerce").fillna(0.0)
    resumo_long["Mes"] = resumo_long["Mes"].astype(int)

    # CSV com BOM para acentuação no Windows/Power BI
    resumo_long.to_csv(saida_long, index=False, encoding="utf-8-sig")
    print(f"✅ LONG salvo: {saida_long}")

    return resumo_wide, resumo_long, str(saida_wide), str(saida_long)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Gera resumos de frete (WIDE e LONG)")
    parser.add_argument(
        "--pasta",
        default=DEFAULT_PASTA,
        help=f"Pasta de trabalho contendo os arquivos de entrada (padrão: {DEFAULT_PASTA})",
    )
    parser.add_argument(
        "--ano",
        type=int,
        default=DEFAULT_ANO,
        help=f"Ano de referência (padrão: {DEFAULT_ANO})",
    )
    grupo_meta = parser.add_mutually_exclusive_group()
    grupo_meta.add_argument(
        "--meta-ja-em-escala",
        action="store_true",
        help="Indica que a coluna Meta já está em escala 0..1 (ex.: 3% -> 0.03)",
    )
    grupo_meta.add_argument(
        "--meta-nao-em-escala",
        action="store_true",
        help="Indica que a coluna Meta está em % inteiro (ex.: 3 -> 3%)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.meta_nao_em_escala:
        meta_em_escala = False
    elif args.meta_ja_em_escala:
        meta_em_escala = True
    else:
        meta_em_escala = DEFAULT_META_JA_EM_ESCALA_0_A_1

    print(
        f"Executando com parâmetros:\n"
        f"- pasta: {os.path.expanduser(args.pasta)}\n"
        f"- ano: {args.ano}\n"
        f"- meta_ja_em_escala_0_a_1: {meta_em_escala}"
    )

    resumo_wide, resumo_long, out_wide, out_long = processar_resumo_frete(
        pasta=args.pasta,
        ano=args.ano,
        meta_ja_em_escala_0_a_1=meta_em_escala,
    )

    # Pequeno preview no console
    with pd.option_context("display.max_columns", None, "display.width", 120):
        print("\nPrévia WIDE:")
        print(resumo_wide.head(10))
        print("\nPrévia LONG:")
        print(resumo_long.head(10))


if __name__ == "__main__":
    main()