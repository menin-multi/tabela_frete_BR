import os
import glob
import math
import pandas as pd
from openpyxl import Workbook

# ---------------- CONFIGURAÇÕES ----------------
FATOR_FIXO = 0.43  # <-- ajuste manualmente este valor conforme necessário

PESO_MAP = {
    "10":10,
    "20":10,
    "30":30,
    "40":30,
    "50":50,
    "60":50,
    "70":65,
    "80":80,
    "90":80,
    "100":105,
    "110":105,
    "120":105,
    "130":140,
    "140":140,
    "150":140,
    "160":140,
    "170":180,
    "180":180,
    "190":180,
    "200":180,
    "250":225,
    "300":275,
    "350":325,
    "400":375,
    "450":425,
    "500":475,

}

COLUNAS_ALVO = list(PESO_MAP.keys())

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_ORIGEM = os.path.normpath(
    os.path.join(BASE_DIR, "..", "..", "atualizacao_de_fretes_br", "create_sheet", "results")
)
PASTA_DESTINO = os.path.join(BASE_DIR, "results")
# ------------------------------------------------


def encontrar_csv_mais_recente(pasta_base):
    padrao = os.path.join(pasta_base, "**", "TABELA_INTELIPOST_*.csv")
    arquivos = glob.glob(padrao, recursive=True)
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo TABELA_INTELIPOST_*.csv encontrado em {pasta_base}")
    return max(arquivos, key=os.path.getmtime)


def arredondar_para_9_99(valor):
    # Sobe para o próximo múltiplo de 10 e subtrai 0.01 (ex: 872.05 -> 879.99)
    proximo_multiplo = math.ceil(round(valor, 2) / 10) * 10
    resultado = proximo_multiplo - 0.01
    # Caso o valor já termine em 9.99 exato, mantém (evita subir uma casa a mais)
    if abs(round(valor, 2) - resultado) < 1e-9:
        return round(valor, 2)
    return round(resultado, 2)


def aplicar_formula(valor, peso):
    calculo = (peso * FATOR_FIXO) + valor
    return arredondar_para_9_99(calculo)


def processar_planilha(caminho_csv):
    df = pd.read_csv(
        caminho_csv,
        sep=";",
        decimal=",",
        encoding="utf-8-sig",
        dtype={"CEP DE": "string", "CEP ATÉ": "string"},
        low_memory=False,
    )

    for coluna in COLUNAS_ALVO:
        if coluna not in df.columns:
            continue
        peso = PESO_MAP[coluna]

        def transformar(valor):
            if pd.isna(valor):
                return valor
            try:
                valor_num = float(valor)
            except (ValueError, TypeError):
                return valor
            if valor_num == 0:
                return valor
            return aplicar_formula(valor_num, peso)

        df[coluna] = df[coluna].apply(transformar)

    return df


def main():
    os.makedirs(PASTA_DESTINO, exist_ok=True)

    caminho_csv = encontrar_csv_mais_recente(PASTA_ORIGEM)
    print(f"Arquivo mais recente encontrado: {caminho_csv}")

    df = processar_planilha(caminho_csv)

    nome_base = os.path.splitext(os.path.basename(caminho_csv))[0]
    caminho_saida = os.path.join(PASTA_DESTINO, f"{nome_base}.xlsx")

    salvar_xlsx_preservando_cep(df, caminho_saida)
    print(f"Arquivo salvo em: {caminho_saida}")


def salvar_xlsx_preservando_cep(df, caminho_saida):
    colunas_texto = {"CEP DE", "CEP ATÉ"}
    wb = Workbook()
    ws = wb.active
    ws.append(list(df.columns))

    indices_texto = [i for i, c in enumerate(df.columns) if c in colunas_texto]

    for row in df.itertuples(index=False, name=None):
        ws.append(row)
        r = ws.max_row
        for i in indices_texto:
            cell = ws.cell(row=r, column=i + 1)
            cell.number_format = "@"
            if cell.value is not None:
                cell.value = str(cell.value)

    wb.save(caminho_saida)


if __name__ == "__main__":
    main()
    