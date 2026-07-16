from pathlib import Path
from datetime import datetime
from time import time
from functools import lru_cache
import re
import sys
from openpyxl import Workbook, load_workbook
import numpy as np
import pandas as pd
from math import ceil

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from logging_utils import configure_logging


logger = configure_logging(__name__)
print = logger.info

# configuração de data
dia_mes = datetime.now().strftime("%d-%m")

def identificar_faixa_de_peso(peso) -> int | None:
    if peso is None or peso == "":
        return None

    try:
        peso = float(peso)
    except (TypeError, ValueError):
        return None

    if peso <= 10:
        return 10

    if peso <= 200:
        # nearest upper 10, minimum 20
        return int(max(20, ceil(peso / 10.0) * 10))

    if peso <= 500:
        # nearest upper 50, minimum 250
        return int(max(250, ceil(peso / 50.0) * 50))

    # above 500: nearest upper 100, minimum 600
    return int(max(600, ceil(peso / 100.0) * 100))


def normalizar_cep(valor) -> str:
    if valor is None:
        return ""

    texto = "".join(caractere for caractere in str(valor) if caractere.isdigit())
    if not texto:
        return ""

    return texto.zfill(8)


@lru_cache(maxsize=None)
def _carregar_linhas_planilha(base_dir: str, sheet_name: str) -> tuple[tuple, ...]:
    arquivo_base = Path(base_dir) / "base.xlsx"

    if not arquivo_base.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_base}")

    workbook = load_workbook(arquivo_base, data_only=True, read_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Aba '{sheet_name}' não encontrada no arquivo: {arquivo_base}")

        planilha = workbook[sheet_name]
        return tuple(tuple(linha) for linha in planilha.iter_rows(values_only=True))
    finally:
        workbook.close()


def carregar_base_cidades(base_dir: Path) -> dict[str, str]:
    linhas = _carregar_linhas_planilha(str(Path(base_dir).resolve()), "cidades")
    base_cidades = {}

    # Mapear índices de coluna a partir do cabeçalho (linha 1)
    header_row = linhas[0] if linhas else None
    cep_idx = 0
    cidade_idx = 8
    if header_row:
        for i, h in enumerate(header_row):
            if h is None:
                continue
            key = ''.join(ch for ch in str(h).lower() if ch.isalpha() or ch.isdigit())
            if any(k in key for k in ("CEP")):
                cep_idx = i

    for linha in linhas[1:]:
            cep = linha[cep_idx] if len(linha) > cep_idx else None
            cidade = linha[cidade_idx] if len(linha) > cidade_idx else None

            if cep is None or cidade is None:
                continue

            base_cidades[normalizar_cep(cep)] = cidade

    return base_cidades


def carregar_base_venda(base_dir: Path) -> dict[str, str]:
    linhas = _carregar_linhas_planilha(str(Path(base_dir).resolve()), "venda")
    base_venda = {}

    for linha in linhas[1:]:
            cidade = linha[4] if len(linha) > 4 else None
            venda = linha[1] if len(linha) > 1 else None

            if cidade is None or venda is None:
                continue

            base_venda[str(cidade).strip()] = venda

    return base_venda


def carregar_base_frete(base_dir: Path) -> dict[str, str]:
    linhas = _carregar_linhas_planilha(str(Path(base_dir).resolve()), "frete")
    base_frete = {}

    for linha in linhas[1:]:
            chave = linha[3] if len(linha) > 3 else None
            valor = linha[4] if len(linha) > 4 else None

            if chave is None or valor is None:
                continue

            base_frete[str(chave).strip()] = valor

    return base_frete


def carregar_base_pesos(base_dir: Path) -> list[tuple]:
    linhas = _carregar_linhas_planilha(str(Path(base_dir).resolve()), "pesos")
    base_pesos = []
    header_row = linhas[0] if linhas else None
    peso_de_idx = 0
    peso_por_idx = 1
    range_peso_idx = 2
    if header_row:
        for i, h in enumerate(header_row):
            if h is None:
                continue
            key = ''.join(ch for ch in str(h).lower() if ch.isalpha() or ch.isdigit())
            if 'pesode' in key:
                peso_de_idx = i
            elif 'pesopor' in key:
                peso_por_idx = i
            elif 'range' in key:
                range_peso_idx = i

    for linha in linhas[1:]:
            peso_de = linha[peso_de_idx] if len(linha) > peso_de_idx else None
            peso_por = linha[peso_por_idx] if len(linha) > peso_por_idx else None
            range_peso = linha[range_peso_idx] if len(linha) > range_peso_idx else None
            base_pesos.append((peso_de, peso_por, range_peso))

    return base_pesos


def carregar_base_cidades_detalhadas(base_dir: Path) -> dict[str, tuple]:
    linhas = _carregar_linhas_planilha(str(Path(base_dir).resolve()), "cidades")
    base_cidades = {}

    for linha in linhas[1:]:
            cep = linha[0] if len(linha) > 0 else None
            valor_d = linha[3] if len(linha) > 3 else None
            valor_e = linha[4] if len(linha) > 4 else None
            localidade = linha[5] if len(linha) > 5 else None
            mesoregiao = linha[6] if len(linha) > 6 else None
            microrregiao = linha[7] if len(linha) > 7 else None
            cidade_col_i = linha[8] if len(linha) > 8 else None

            if cep is None:
                continue

            base_cidades[normalizar_cep(cep)] = (
                valor_d,
                valor_e,
                localidade,
                mesoregiao,
                microrregiao,
                cidade_col_i,
            )

    return base_cidades


def carregar_bases_unicas(base_dir: Path) -> tuple[dict, dict, dict]:
    arquivo_base = base_dir / "base.xlsx"

    if not arquivo_base.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_base}")

    workbook = load_workbook(arquivo_base, data_only=True, read_only=True)
    try:
        base_cidades = {}
        base_venda = {}
        base_cidades_detalhadas = {}

        # Ler aba cidades (preenche base_cidades e base_cidades_detalhadas)
        if "cidades" in workbook.sheetnames:
            planilha = workbook["cidades"]
            # identificar índices do cabeçalho (similar ao carregar_base_cidades)
            header_row = next(planilha.iter_rows(min_row=1, max_row=1, values_only=True), None)
            cep_idx = 0
            cidade_idx = 8
            if header_row:
                for i, h in enumerate(header_row):
                    if h is None:
                        continue
                    key = ''.join(ch for ch in str(h).lower() if ch.isalpha() or ch.isdigit())
                    if any(k in key for k in ("CEP")):
                        cep_idx = i

                for linha in planilha.iter_rows(min_row=2, values_only=True):
                    cep = linha[cep_idx] if len(linha) > cep_idx else None
                    cidade = linha[cidade_idx] if len(linha) > cidade_idx else None

                    if cep is not None and cidade is not None:
                        base_cidades[normalizar_cep(cep)] = cidade

                    if cep is None:
                        continue

                    base_cidades_detalhadas[normalizar_cep(cep)] = (
                        linha[3] if len(linha) > 3 else None,
                        linha[4] if len(linha) > 4 else None,
                        linha[5] if len(linha) > 5 else None,
                        linha[6] if len(linha) > 6 else None,
                        linha[7] if len(linha) > 7 else None,
                        linha[8] if len(linha) > 8 else None,
                    )

        # Ler aba venda (preenche base_venda)
        if "venda" in workbook.sheetnames:
            planilha = workbook["venda"]
            for linha in planilha.iter_rows(min_row=2, values_only=True):
                cidade = linha[4] if len(linha) > 4 else None
                venda = linha[1] if len(linha) > 1 else None

                if cidade is None or venda is None:
                    continue

                base_venda[str(cidade).strip()] = venda

    finally:
        workbook.close()

    return base_cidades, base_venda, base_cidades_detalhadas


def export_to_xlsxwriter(openpyxl_wb: Workbook, path: Path) -> None:
    # Salva diretamente usando openpyxl sem aplicar qualquer formatação.
    # Garante que a pasta de destino exista antes de salvar.
    path_parent = Path(path).resolve().parent
    path_parent.mkdir(parents=True, exist_ok=True)
    openpyxl_wb.save(path)


def _escrever_dataframe_em_planilha(planilha, dataframe: pd.DataFrame) -> None:
    dataframe = dataframe.astype(object).where(pd.notna(dataframe), None)
    for linha in dataframe.itertuples(index=False, name=None):
        planilha.append(list(linha))


def _carregar_ultimo_custo_atual(base_dir: Path) -> dict[str, float]:
    # Usa o último CSV gerado como snapshot para deslocar o valor atual para a coluna de valor antigo.
    results_dir = base_dir / "results"
    if not results_dir.exists():
        return {}

    arquivos_csv = [
        arquivo
        for arquivo in results_dir.rglob("TABELA_DE_FRETE_*.csv")
        if arquivo.is_file()
    ]
    if not arquivos_csv:
        return {}

    ultimo_arquivo = max(arquivos_csv, key=lambda arquivo: arquivo.stat().st_mtime)
    try:
        snapshot = pd.read_csv(
            ultimo_arquivo,
            sep=";",
            decimal=",",
            encoding="utf-8-sig",
            usecols=[
               "CEP DE + PESO POR",
               "FRETE CUSTO ATUAL"
            ],
            dtype={"CEP DE + PESO POR": "string"},
        )
    except Exception:
        return {}

    snapshot["CEP DE + PESO POR"] = snapshot["CEP DE + PESO POR"].astype("string")
    snapshot["FRETE CUSTO ATUAL"] = pd.to_numeric(snapshot["FRETE CUSTO ATUAL"], errors="coerce")
    return (
        snapshot.dropna(subset=["CEP DE + PESO POR"])
        .drop_duplicates(subset=["CEP DE + PESO POR"], keep="last")
        .set_index("CEP DE + PESO POR")["FRETE CUSTO ATUAL"]
        .to_dict()
    )


def _carregar_ultimo_frete_comercial_atual(base_dir: Path) -> dict[str, float]:
    # Usa o último CSV gerado como snapshot para deslocar o valor atual para a coluna de valor antigo.
    results_dir = base_dir / "results"
    if not results_dir.exists():
        return {}

    arquivos_csv = [
        arquivo
        for arquivo in results_dir.rglob("TABELA_DE_FRETE_*.csv")
        if arquivo.is_file()
    ]
    if not arquivos_csv:
        return {}

    ultimo_arquivo = max(arquivos_csv, key=lambda arquivo: arquivo.stat().st_mtime)
    try:
        snapshot = pd.read_csv(
            ultimo_arquivo,
            sep=";",
            decimal=",",
            encoding="utf-8-sig",
            usecols=[
    "CEP DE + PESO POR",
    "FRETE COMERCIAL ATUAL"
],
            dtype={"CEP DE + PESO POR": "string"},
        )
    except Exception:
        return {}

    snapshot["CEP DE + PESO POR"] = snapshot["CEP DE + PESO POR"].astype("string")
    snapshot["FRETE COMERCIAL ATUAL"] = pd.to_numeric(snapshot["FRETE COMERCIAL ATUAL"], errors="coerce")
    return (
        snapshot.dropna(subset=["CEP DE + PESO POR"])
        .drop_duplicates(subset=["CEP DE + PESO POR"], keep="last")
        .set_index("CEP DE + PESO POR")["FRETE COMERCIAL ATUAL"]
        .to_dict()
    )


def _carregar_nome_custo_antigo(base_dir: Path) -> str:
    results_dir = base_dir / "results"
    if not results_dir.exists():
        return "FRETE CUSTO {valor antigo}"

    arquivos_csv = [
        arquivo
        for arquivo in results_dir.rglob("TABELA_DE_FRETE_*.csv")
        if arquivo.is_file()
    ]
    if not arquivos_csv:
        return "FRETE CUSTO {valor antigo}"

    ultimo_arquivo = max(arquivos_csv, key=lambda arquivo: arquivo.stat().st_mtime)
    # tentar extrair a data diretamente do nome do arquivo, ex: TABELA_DE_FRETE_27-05-2026.csv
    padrao_data = re.search(r"TABELA_DE_FRETE_(\d{2}-\d{2}-\d{4})$", ultimo_arquivo.stem, flags=re.IGNORECASE)
    if padrao_data:
        return f"FRETE CUSTO {padrao_data.group(1)}"

    # fallback para o nome da pasta (ex: '04-2026') quando o arquivo não seguir o novo padrão
    mes_ano_dir = ultimo_arquivo.parent.name
    try:
        # ainda tenta ler o cabeçalho apenas para validar leitura do CSV
        _ = pd.read_csv(ultimo_arquivo, sep=";", encoding="utf-8-sig", nrows=0)
    except Exception:
        # se não conseguir ler, ainda assim retorna com o mês-ano da pasta
        return f"FRETE CUSTO {mes_ano_dir}"

    # retornar sempre com o mês-ano da pasta onde o último CSV foi encontrado
    return f"FRETE CUSTO {mes_ano_dir}"


def _carregar_nome_coluna_frete_antigo(base_dir: Path) -> str:
    results_dir = base_dir / "results"
    if not results_dir.exists():
        return "FRETE COMERCIAL {valor antigo}"

    arquivos_csv = [
        arquivo
        for arquivo in results_dir.rglob("TABELA_DE_FRETE_*.csv")
        if arquivo.is_file()
    ]
    if not arquivos_csv:
        return "FRETE COMERCIAL {valor antigo}"

    ultimo_arquivo = max(arquivos_csv, key=lambda arquivo: arquivo.stat().st_mtime)
    # tentar extrair a data diretamente do nome do arquivo, ex: TABELA_DE_FRETE_27-05-2026.csv
    padrao_data = re.search(r"TABELA_DE_FRETE_(\d{2}-\d{2}-\d{4})$", ultimo_arquivo.stem, flags=re.IGNORECASE)
    if padrao_data:
        return f"FRETE COMERCIAL {padrao_data.group(1)}"

    # fallback para o nome da pasta (ex: '04-2026') quando o arquivo não seguir o novo padrão
    mes_ano_dir = ultimo_arquivo.parent.name
    try:
        # ainda tenta ler o cabeçalho apenas para validar leitura do CSV
        _ = pd.read_csv(ultimo_arquivo, sep=";", encoding="utf-8-sig", nrows=0)
    except Exception:
        # se não conseguir ler, ainda assim retorna com o mês-ano da pasta
        return f"FRETE COMERCIAL {mes_ano_dir}"

    # retornar sempre com o mês-ano da pasta onde o último CSV foi encontrado
    return f"FRETE COMERCIAL {mes_ano_dir}"


def base_rodrigo(workbook_saida: Workbook, base_cidades: dict) -> pd.DataFrame:
    print("Processando arquivos de cotação")
    base_dir = Path(__file__).resolve().parent
    cotacoes_dir = base_dir / "cotacoes" / dia_mes
    worksheet_name = "BASE RODRIGO"
    out_headers = [
        "UF",
        "REGIÃO",
        "CEP COTDO",
        "SKU",
        "FAIXA DE PESO",
        "PESO DO PRODUTO",
        "FAIXA DE PESO",
        "VALOR DO PRODUTO",
        "ID COTAÇÃO",
        "ID TRANSPORTADOR",
        "TRANSPORTADOR",
        "PRAZO DE ENTREGA",
        "VALOR COMERCIAL",
        "VALOR TRANSPORTADOR",
        "uf-cidade-peso",
        "Cidade",
        "CEP_PESO_KEY",
    ]

    if not cotacoes_dir.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {cotacoes_dir}")

    planilha_saida = workbook_saida.create_sheet(title=worksheet_name)
    planilha_saida.append(out_headers)

    registros = []
    total_linhas_cotacoes = 0
    arquivos = sorted(
        arquivo for arquivo in cotacoes_dir.iterdir()
        if arquivo.is_file() and arquivo.suffix.lower() in {".xlsx", ".xlsm"}
    )

    for arquivo in arquivos:
        workbook_origem = load_workbook(arquivo, data_only=True, read_only=True)
        try:
            planilha_origem = workbook_origem.worksheets[0]
            for linha in planilha_origem.iter_rows(min_row=2, values_only=True):
                total_linhas_cotacoes += 1
                registros.append({
                    "UF": linha[0] if len(linha) > 0 else None,
                    "REGIAO": linha[1] if len(linha) > 1 else None,
                    "CEP_COTDO": linha[2] if len(linha) > 2 else None,
                    "SKU": linha[3] if len(linha) > 3 else None,
                    "FAIXA_PESO_ORIG": linha[4] if len(linha) > 4 else None,
                    "PESO_PRODUTO": linha[5] if len(linha) > 5 else None,
                    "VALOR_PRODUTO": linha[6] if len(linha) > 6 else None,
                    "ID_COTACAO": linha[7] if len(linha) > 7 else None,
                    "ID_TRANSPORTADOR": linha[8] if len(linha) > 8 else None,
                    "TRANSPORTADOR": linha[9] if len(linha) > 9 else None,
                    "PRAZO_ENTREGA": linha[10] if len(linha) > 10 else None,
                    "VALOR_COMERCIAL": linha[11] if len(linha) > 11 else None,
                    "VALOR_TRANSPORTADOR": linha[12] if len(linha) > 12 else None,
                })
        finally:
            workbook_origem.close()

    df = pd.DataFrame(registros)
    if df.empty:
        print("Base Rodrigo criada")
        return df

    colunas_numericas_base_rodrigo = [
        "PESO_PRODUTO",
        "VALOR_PRODUTO",
        "ID_TRANSPORTADOR",
        "PRAZO_ENTREGA",
        "VALOR_COMERCIAL",
        "VALOR_TRANSPORTADOR",
    ]
    for coluna in colunas_numericas_base_rodrigo:
        df[coluna] = pd.to_numeric(df[coluna], errors="coerce").astype("float64")

    df["CEP_NORMALIZADO"] = df["CEP_COTDO"].map(normalizar_cep)

    df["CIDADE"] = df["CEP_NORMALIZADO"].map(base_cidades).fillna("")
    peso_numerico = df["PESO_PRODUTO"]
    faixa_calc = pd.Series(pd.NA, index=df.index, dtype="Float64")
    mask_valida = peso_numerico.notna()
    faixa_calc.loc[mask_valida & (peso_numerico <= 10)] = 10
    mask_20 = mask_valida & (peso_numerico > 10) & (peso_numerico <= 200)
    faixa_calc.loc[mask_20] = np.maximum(20, np.ceil(peso_numerico.loc[mask_20] / 10.0) * 10)
    mask_250 = mask_valida & (peso_numerico > 200) & (peso_numerico <= 500)
    faixa_calc.loc[mask_250] = np.maximum(250, np.ceil(peso_numerico.loc[mask_250] / 50.0) * 50)
    mask_600 = mask_valida & (peso_numerico > 500)
    faixa_calc.loc[mask_600] = np.maximum(600, np.ceil(peso_numerico.loc[mask_600] / 100.0) * 100)
    df["FAIXA_PESO_CALC"] = faixa_calc.round().astype("Int64")
    df["UF_CIDADE_PESO"] = (
        df["UF"].fillna("").astype(str)
        + "-"
        + df["CIDADE"].fillna("").astype(str)
        + "-"
        + df["FAIXA_PESO_CALC"].astype("string").fillna("")
    )
#######################################################################################
    df["CEP_PESO_KEY"] = (
        df["CEP_NORMALIZADO"]
        + "-"
        + df["FAIXA_PESO_CALC"].astype("string").fillna("")
    )

    duplicados_cep_peso = (
        df.groupby("CEP_PESO_KEY")["TRANSPORTADOR"]
        .nunique()
    )

    duplicados_cep_peso = duplicados_cep_peso[duplicados_cep_peso > 1]

    print(
        f"CEP+PESO com mais de uma transportadora: {len(duplicados_cep_peso)}"
    )

    if not duplicados_cep_peso.empty:
        print(duplicados_cep_peso.head(20))
#######################################################################################
    base_rodrigo_saida = pd.concat(
        [
            df["UF"].astype("string"),
            df["REGIAO"].astype("string"),
            df["CEP_COTDO"].astype("string"),
            df["SKU"].astype("string"),
            df["FAIXA_PESO_ORIG"].astype("string"),
            df["PESO_PRODUTO"].astype("float64"),
            df["FAIXA_PESO_CALC"].astype("float64"),
            df["VALOR_PRODUTO"].astype("float64"),
            df["ID_COTACAO"].astype("string"),
            df["ID_TRANSPORTADOR"].astype("float64"),
            df["TRANSPORTADOR"].astype("string"),
            df["PRAZO_ENTREGA"].astype("float64"),
            df["VALOR_COMERCIAL"].astype("float64"),
            df["VALOR_TRANSPORTADOR"].astype("float64"),
            df["UF_CIDADE_PESO"].astype("string"),
            df["CIDADE"].astype("string"),
            df["CEP_PESO_KEY"].astype("string"),
        ],
        axis=1,
    )
    base_rodrigo_saida.columns = out_headers

    _escrever_dataframe_em_planilha(planilha_saida, base_rodrigo_saida)

    # salvar CSV da aba BASE RODRIGO em results/{mes-ano}
    mes_ano = datetime.now().strftime("%m-%Y")
    results_dir = Path(__file__).resolve().parent / "results" / mes_ano
    results_dir.mkdir(parents=True, exist_ok=True)

    data_atual = datetime.now().strftime("%d-%m-%Y")
    csv_path = results_dir / f"BASE_RODRIGO_{data_atual}.csv"
    base_rodrigo_saida.to_csv(csv_path, index=False, sep=";", decimal=",", float_format="%.2f", encoding="utf-8-sig")
    total_linhas_path = results_dir / f"BASE_RODRIGO_{data_atual}_TOTAL_LINHAS.txt"
    total_linhas_path.write_text(str(total_linhas_cotacoes), encoding="utf-8")

    print("Base Rodrigo criada")
    return df


def tabela_frete(workbook_saida: Workbook, base_rodrigo_df: pd.DataFrame, base_venda: dict, base_cidades_detalhadas: dict, base_frete: dict[str, str], base_pesos: list[tuple]) -> None:
    print("Gerando tabela de frete")
    worksheet_name = "TABELA DE FRETE"
    mes_ano_atual = datetime.now().strftime("%m-%Y")
    ac3_valor = 12.5
    coluna_custo_antigo = _carregar_nome_custo_antigo(Path(__file__).resolve().parent)
    coluna_frete_antigo = _carregar_nome_coluna_frete_antigo(Path(__file__).resolve().parent)
    
    headers = [
        "UF+CIDADE",
        "VENDA 2025",
        "UF",
        "REGIÃO",
        "CEP DE",
        "CEP ATÉ",
        "PESO DE",
        "PESO POR",
        "RANGE",
        "LOCALIDADE",
        "MESOREGIÃO",
        "MICRORREGIÃO",
        "CIDADE",
        "FRETE COMERCIAL X FRETE REAL (R$)",
        "FRETE COMERCIAL X FRETE REAL (%)",
        "FRETE COMERCIAL ATUAL",
        "CEP DE + PESO POR",
        "CIDADE + PESO POR",
        coluna_custo_antigo,
        coluna_frete_antigo,
        "FRETE CUSTO ATUAL",
        "CUSTO MÉDIO / RANGE DE PESO",
        "DIF. FRETE CUSTO R$",
        "DIFERENÇA %",
        "TRANSPORTADORA",
        "REGIÃO + UF + LOCALIDADE",
        "% CUSTO FRETE",
    ]
    planilha_saida = workbook_saida.create_sheet(title=worksheet_name)
    planilha_saida.append([None] * len(headers))
    planilha_saida.append(headers)
    planilha_saida["ac3"] = ac3_valor

    df = base_rodrigo_df.copy().reset_index(drop=True)
    total_linhas_esperadas = len(df)
    df["UF_CIDADE_KEY"] = df["UF"].fillna("").astype(str) + ":" + df["CIDADE"].fillna("").astype(str)
    df["VENDA 2025"] = df["UF_CIDADE_KEY"].map(base_venda).fillna("")
    ultimo_custo_atual = _carregar_ultimo_custo_atual(Path(__file__).resolve().parent)
    ultimo_frete_comercial_atual = _carregar_ultimo_frete_comercial_atual(Path(__file__).resolve().parent)

    base_pesos_df = pd.DataFrame(base_pesos, columns=["PESO_DE", "PESO_POR", "RANGE_IDX"])
    base_pesos_df["PESO_POR"] = pd.to_numeric(
        base_pesos_df["PESO_POR"].astype("string").str.replace(",", ".", regex=False).str.strip(),
        errors="coerce",
    )
    base_pesos_df["PESO_DE"] = pd.to_numeric(
        base_pesos_df["PESO_DE"].astype("string").str.replace(",", ".", regex=False).str.strip(),
        errors="coerce",
    )
    base_pesos_df["RANGE_IDX"] = base_pesos_df["RANGE_IDX"].astype("string").str.strip()
    base_pesos_df = (
        base_pesos_df.dropna(subset=["PESO_POR"])
        .drop_duplicates(subset=["PESO_POR"], keep="last")
        .sort_values("PESO_POR")
    )
    df["FAIXA_PESO_CALC_NUM"] = pd.to_numeric(df["FAIXA_PESO_CALC"], errors="coerce")
    df = df.merge(base_pesos_df, left_on="FAIXA_PESO_CALC_NUM", right_on="PESO_POR", how="left")

    # Garantir que a coluna RANGE_IDX seja preenchida a partir do PESO_POR correspondente
    try:
        peso_por_to_range = (
            base_pesos_df[["PESO_POR", "RANGE_IDX"]]
            .dropna(subset=["PESO_POR"])
            .drop_duplicates(subset=["PESO_POR"], keep="last")
            .set_index("PESO_POR")["RANGE_IDX"]
            .to_dict()
        )
        df["RANGE_IDX"] = df["RANGE_IDX"].fillna(df["FAIXA_PESO_CALC_NUM"].map(peso_por_to_range))
    except Exception:
        # se algo falhar aqui, não interrompe o fluxo principal
        pass

    detalhes_df = pd.DataFrame.from_dict(
        base_cidades_detalhadas,
        orient="index",
        columns=["CEP_DE", "CEP_ATE", "LOCALIDADE", "MESOREGIÃO", "MICRORREGIÃO", "CIDADE_ACENTO"],
    ).reset_index(names="CEP_NORMALIZADO")
    df = df.merge(detalhes_df, on="CEP_NORMALIZADO", how="left")

    peso_por_chave = pd.to_numeric(
        df["PESO_POR"],
        errors="coerce"
    ).map(
        lambda valor:
            "" if pd.isna(valor)
            else str(int(valor))
            if float(valor).is_integer()
            else str(valor).rstrip("0").rstrip(".")
    )

    cep_de_chave = df["CEP_DE"].map(normalizar_cep)

    df["CEP DE + PESO POR"] = (
        cep_de_chave
        + "-"
        + peso_por_chave
    )

    base_rodrigo_lookup = (
        base_rodrigo_df[
            ["CEP_PESO_KEY", "VALOR_TRANSPORTADOR"]
        ]
        .dropna(subset=["CEP_PESO_KEY"])
        .drop_duplicates(
            subset=["CEP_PESO_KEY"],
            keep="last"
        )
        .rename(
            columns={
                "VALOR_TRANSPORTADOR":
                "VALOR_TRANSPORTADOR_BASE_RODRIGO"
            }
        )
    )
    df = df.merge(
    base_rodrigo_lookup,
    left_on="CEP DE + PESO POR",
    right_on="CEP_PESO_KEY",
    how="left"
    )
    if len(df) != total_linhas_esperadas:
        print(f"Erro: junção com base_rodrigo_lookup alterou quantidade de linhas ({len(df)} != {total_linhas_esperadas})")
        raise SystemExit(1)

    base_rodrigo_transportadora_lookup = (
        base_rodrigo_df[
            ["CEP_PESO_KEY", "TRANSPORTADOR"]
        ]
        .dropna(subset=["CEP_PESO_KEY"])
        .drop_duplicates(
            subset=["CEP_PESO_KEY"],
            keep="last"
        )
        .set_index("CEP_PESO_KEY")["TRANSPORTADOR"]
    )
    chaves_rodrigo = set(
    base_rodrigo_transportadora_lookup.index
)

    frete_df = pd.DataFrame(
        list(base_frete.items()),
        columns=["REGIAO_UF_LOCALIDADE", "ADICIONAL_FRETE"]
    )
    df["REGIAO_UF_LOCALIDADE"] = (
        df["REGIAO"].fillna("").astype(str) + " - " + df["UF"].fillna("").astype(str) + " - " + df["LOCALIDADE"].fillna("").astype(str)
    )
    df = df.merge(frete_df, on="REGIAO_UF_LOCALIDADE", how="left")

    df["PESO_DE"] = df["PESO_DE"]
    df["VALOR_TRANSPORTADOR_NUM"] = pd.to_numeric(df["VALOR_TRANSPORTADOR"], errors="coerce")
    df["ADICIONAL_FRETE_NUM"] = pd.to_numeric(df["ADICIONAL_FRETE"].astype(str).str.replace(",", ".", regex=False), errors="coerce")
    df[coluna_custo_antigo] = (
        df["CEP DE + PESO POR"]
        .map(ultimo_custo_atual)
    )

    df[coluna_frete_antigo] = (
        df["CEP DE + PESO POR"]
        .map(ultimo_frete_comercial_atual)
    )
    df["FRETE CUSTO ATUAL"] = pd.to_numeric(df["VALOR_TRANSPORTADOR_BASE_RODRIGO"], errors="coerce") + ac3_valor
    range_aux = df["RANGE_IDX"].fillna("__VAZIO__").astype(str)
    grupos_consecutivos_range = (range_aux.ne(range_aux.shift()).cumsum())
    
    df["CUSTO MÉDIO / RANGE DE PESO"] = df.groupby(grupos_consecutivos_range, dropna=False)[f"FRETE CUSTO ATUAL"].transform("mean").astype("float64")

    df["CIDADE + PESO POR"] = df["UF_CIDADE_PESO"]


    df["TRANSPORTADORA"] = (
        df["CEP DE + PESO POR"]
        .astype("string")
        .map(base_rodrigo_transportadora_lookup)
    )

    df["% CUSTO FRETE"] = df["ADICIONAL_FRETE_NUM"].astype("float64")

    mask_p = df["CUSTO MÉDIO / RANGE DE PESO"].notna() & df["ADICIONAL_FRETE_NUM"].notna()
    df["FRETE COMERCIAL ATUAL"] = pd.Series(np.nan, index=df.index, dtype="float64")
    df["FRETE COMERCIAL X FRETE REAL (R$)"] = pd.Series(np.nan, index=df.index, dtype="float64")
    df["FRETE COMERCIAL X FRETE REAL (%)"] = pd.Series(np.nan, index=df.index, dtype="float64")
    
    df.loc[mask_p, "FRETE COMERCIAL ATUAL"] = np.ceil((df.loc[mask_p, "CUSTO MÉDIO / RANGE DE PESO"] + (df.loc[mask_p, "CUSTO MÉDIO / RANGE DE PESO"] * df.loc[mask_p, "ADICIONAL_FRETE_NUM"])) / 10.0) * 10.0 - 0.01
    col_mes = "FRETE CUSTO ATUAL"
    mask_diff = df["FRETE COMERCIAL ATUAL"].notna() & df[col_mes].notna()
    df.loc[mask_diff, "FRETE COMERCIAL X FRETE REAL (R$)"] = df.loc[mask_diff, "FRETE COMERCIAL ATUAL"] - df.loc[mask_diff, col_mes]
    mask_perc = df["FRETE COMERCIAL ATUAL"].notna() & df[col_mes].notna() & (df[col_mes] != 0)
    df.loc[mask_perc, "FRETE COMERCIAL X FRETE REAL (%)"] = df.loc[mask_perc, "FRETE COMERCIAL ATUAL"] / df.loc[mask_perc, col_mes] - 1
    df["DIF. FRETE CUSTO R$"] = pd.Series(np.nan, index=df.index, dtype="float64")
    df["DIFERENÇA %"] = pd.Series(np.nan, index=df.index, dtype="float64")
    mask_dif_custo = df[f"FRETE CUSTO ATUAL"].notna() & df[coluna_custo_antigo].notna()
    df.loc[mask_dif_custo, "DIF. FRETE CUSTO R$"] = df.loc[mask_dif_custo, coluna_custo_antigo] - df.loc[mask_dif_custo, f"FRETE CUSTO ATUAL"]
    mask_dif_perc = mask_dif_custo & (df[coluna_custo_antigo] != 0)
    df.loc[mask_dif_perc, "DIFERENÇA %"] = (
        df.loc[mask_dif_perc, f"FRETE CUSTO ATUAL"] / df.loc[mask_dif_perc, coluna_custo_antigo] - 1
    ).round(4)

    resultado = pd.DataFrame({
        "UF+CIDADE": df["UF_CIDADE_KEY"],
        "VENDA 2025": df["VENDA 2025"],
        "UF": df["UF"],
        "REGIÃO": df["REGIAO"],
        "CEP DE": df["CEP_DE"],
        "CEP ATÉ": df["CEP_ATE"],
        "PESO DE": df["PESO_DE"],
        "PESO POR": df["PESO_POR"],
        "RANGE": df["RANGE_IDX"],
        "LOCALIDADE": df["LOCALIDADE"],
        "MESOREGIÃO": df["MESOREGIÃO"],
        "MICRORREGIÃO": df["MICRORREGIÃO"],
        "CIDADE": df["CIDADE_ACENTO"],
        "FRETE COMERCIAL X FRETE REAL (R$)": df["FRETE COMERCIAL X FRETE REAL (R$)"],
        "FRETE COMERCIAL X FRETE REAL (%)": df["FRETE COMERCIAL X FRETE REAL (%)"],
        "FRETE COMERCIAL ATUAL": df["FRETE COMERCIAL ATUAL"],
        "CEP DE + PESO POR": df["CEP DE + PESO POR"],
        "CIDADE + PESO POR": df["CIDADE + PESO POR"],
        coluna_custo_antigo: df[coluna_custo_antigo],
        coluna_frete_antigo: df[coluna_frete_antigo],
        f"FRETE CUSTO ATUAL": df["FRETE CUSTO ATUAL"],
        "CUSTO MÉDIO / RANGE DE PESO": df["CUSTO MÉDIO / RANGE DE PESO"],
        "DIF. FRETE CUSTO R$": df["DIF. FRETE CUSTO R$"],
        "DIFERENÇA %": df["DIFERENÇA %"],
        "TRANSPORTADORA": df["TRANSPORTADORA"],
        "REGIÃO + UF + LOCALIDADE": df["REGIAO_UF_LOCALIDADE"],
        "% CUSTO FRETE": df["% CUSTO FRETE"],
    })
    resultado["FRETE COMERCIAL X FRETE REAL (%)"] = resultado["FRETE COMERCIAL X FRETE REAL (%)"].round(4)
    resultado_csv = resultado.copy()
    resultado_csv["FRETE COMERCIAL X FRETE REAL (%)"] = resultado_csv["FRETE COMERCIAL X FRETE REAL (%)"].map(
        lambda valor: None if pd.isna(valor) else f"{valor:.4f}".replace(".", ",")
    )
    resultado_csv["DIFERENÇA %"] = resultado_csv["DIFERENÇA %"].map(
        lambda valor: None if pd.isna(valor) else f"{valor:.4f}".replace(".", ",")
    )

    colunas_texto_tabela_frete = [
        "UF+CIDADE",
        "UF",
        "REGIÃO",
        "CEP DE",
        "CEP ATÉ",
        "RANGE",
        "LOCALIDADE",
        "MESOREGIÃO",
        "MICRORREGIÃO",
        "CIDADE",
        "CEP DE + PESO POR",
        "CIDADE + PESO POR",
        "TRANSPORTADORA",
        "REGIÃO + UF + LOCALIDADE",
    ]
    for coluna in colunas_texto_tabela_frete:
        resultado[coluna] = resultado[coluna].astype("string")

    colunas_numericas_tabela_frete = [
        "VENDA 2025",
        "PESO DE",
        "PESO POR",
        "FRETE COMERCIAL X FRETE REAL (R$)",
        "FRETE COMERCIAL X FRETE REAL (%)",
        "FRETE COMERCIAL ATUAL",
        coluna_custo_antigo,
        coluna_frete_antigo,
        f"FRETE CUSTO ATUAL",
        "CUSTO MÉDIO / RANGE DE PESO",
        "DIF. FRETE CUSTO R$",
        "DIFERENÇA %",
        "% CUSTO FRETE",
    ]
    for coluna in colunas_numericas_tabela_frete:
        resultado[coluna] = pd.to_numeric(resultado[coluna], errors="coerce").astype("float64")
        if coluna == "VENDA 2025":
            resultado[coluna] = resultado[coluna].round(2)
            resultado_csv[coluna] = resultado[coluna]

    if len(resultado) != total_linhas_esperadas:
        print(f"Erro: tabela de frete final com quantidade de linhas incorreta ({len(resultado)} != {total_linhas_esperadas})")
        raise SystemExit(1)

    # salvar CSV da aba TABELA DE FRETE em results/{mes-ano}
    
    results_dir = Path(__file__).resolve().parent / "results" / mes_ano_atual
    results_dir.mkdir(parents=True, exist_ok=True)

    data_atual = datetime.now().strftime("%d-%m-%Y")
    csv_path = results_dir / f"TABELA_DE_FRETE_{data_atual}.csv"
    resultado_csv.to_csv(csv_path, index=False, sep=";", decimal=",", float_format="%.2f", encoding="utf-8-sig")

    _escrever_dataframe_em_planilha(planilha_saida, resultado)
    print("Tabela de frete criada")	


def tabela_intelipost(workbook_saida: Workbook) -> None:
    print("Gerando Tabela Intelipost")
    worksheet_name = "TABELA - INTELIPOST"
    headers = [
        "UF",
        "VENDA 2025",
        "UF+CIDADE",
        "CIDADE",
        "REGIÃO",
        "CEP DE",
        "CEP ATÉ",
        "10",
        "20",
        "30",
        "40",
        "50",
        "60",
        "70",
        "80",
        "90",
        "100",
        "110",
        "120",
        "130",
        "140",
        "150",
        "160",
        "170",
        "180",
        "190",
        "200",
        "250",
        "300",
        "350",
        "400",
        "450",
        "500",
    ]
    planilha_saida = workbook_saida.create_sheet(title=worksheet_name)
    planilha_saida.append(headers)
    mes_ano_atual = datetime.now().strftime("%m-%Y")
    base_dir = Path(__file__).resolve().parent
    data_atual = datetime.now().strftime("%d-%m-%Y")
    results_dir = base_dir / "results" / mes_ano_atual
    csv_frete_path = results_dir / f"TABELA_DE_FRETE_{data_atual}.csv"
    if csv_frete_path.exists():
        base_frete_df = pd.read_csv(
            csv_frete_path,
            sep=";",
            decimal=",",
            encoding="utf-8-sig",
            dtype={
                "CEP DE": "string",
                "CEP ATÉ": "string",
                "CEP DE + PESO POR": "string",
            },
            low_memory=False,
        )
    else:
        base_frete_df = pd.DataFrame(
            columns=[
                "UF",
                "VENDA 2025",
                "UF+CIDADE",
                "CIDADE",
                "REGIÃO",
                "CEP DE",
                "CEP ATÉ",
                "CEP DE + PESO POR",
                "FRETE COMERCIAL ATUAL",
            ]
        )
    resultado = pd.DataFrame({
        "UF": base_frete_df["UF"] if "UF" in base_frete_df.columns else pd.Series(dtype="string"),
        "VENDA 2025": base_frete_df["VENDA 2025"] if "VENDA 2025" in base_frete_df.columns else pd.Series(dtype="float64"),
        "UF+CIDADE": base_frete_df["UF+CIDADE"] if "UF+CIDADE" in base_frete_df.columns else pd.Series(dtype="string"),
        "CIDADE": base_frete_df["CIDADE"] if "CIDADE" in base_frete_df.columns else pd.Series(dtype="string"),
        "REGIÃO": base_frete_df["REGIÃO"] if "REGIÃO" in base_frete_df.columns else pd.Series(dtype="string"),
        "CEP DE": base_frete_df["CEP DE"] if "CEP DE" in base_frete_df.columns else pd.Series(dtype="string"),
        "CEP ATÉ": base_frete_df["CEP ATÉ"] if "CEP ATÉ" in base_frete_df.columns else pd.Series(dtype="string"),
    })
    resultado = resultado.drop_duplicates(
    subset=["CEP DE", "CEP ATÉ", "UF+CIDADE"],
    keep="first",
    ).reset_index(drop=True)
    base_frete_lookup = pd.DataFrame(
        {
            "CEP DE + PESO POR": base_frete_df["CEP DE + PESO POR"] if "CEP DE + PESO POR" in base_frete_df.columns else pd.Series(dtype="string"),
            "FRETE COMERCIAL ATUAL": base_frete_df["FRETE COMERCIAL ATUAL"] if "FRETE COMERCIAL ATUAL" in base_frete_df.columns else pd.Series(dtype="float64"),
        }
    )
    base_frete_lookup["CEP DE + PESO POR"] = (
        base_frete_lookup["CEP DE + PESO POR"]
        .astype("string")
        .str.strip()
        .str.replace(r"\.0+$", "", regex=True)
    )
    base_frete_lookup["FRETE COMERCIAL ATUAL"] = pd.to_numeric(base_frete_lookup["FRETE COMERCIAL ATUAL"], errors="coerce")
    base_frete_lookup = (
        base_frete_lookup.dropna(subset=["CEP DE + PESO POR"])
        .drop_duplicates(subset=["CEP DE + PESO POR"], keep="last")
        .set_index("CEP DE + PESO POR")["FRETE COMERCIAL ATUAL"]
    )
    cep_de_base = resultado["CEP DE"].map(normalizar_cep)
    for coluna in headers[7:]:
        resultado[coluna] = (cep_de_base + "-" + coluna).map(base_frete_lookup)
    resultado = resultado[headers]
    _escrever_dataframe_em_planilha(planilha_saida, resultado)
    results_dir.mkdir(parents=True, exist_ok=True)
    resultado_csv = resultado.copy()
    csv_path = results_dir / f"TABELA_INTELIPOST_{data_atual}.csv"
    resultado_csv.to_csv(csv_path, index=False, sep=";", decimal=",", encoding="utf-8-sig")
    print("Tabela Intelipost criada")
    print("Preparando salvamento do arquivo final")


def main() -> None:

    workbook = Workbook()
    if len(workbook.worksheets) == 1 and workbook.active.title == "Sheet":
        workbook.remove(workbook.active)

    base_dir = Path(__file__).resolve().parent
    base_cidades, base_venda, base_cidades_detalhadas = carregar_bases_unicas(base_dir)
    base_frete = carregar_base_frete(base_dir)
    base_pesos = carregar_base_pesos(base_dir)

    base_rodrigo_df = base_rodrigo(workbook, base_cidades)
    tabela_frete(workbook, base_rodrigo_df, base_venda, base_cidades_detalhadas, base_frete, base_pesos)
    tabela_intelipost(workbook)
    print("Planilhas criadas em: " + str(Path(__file__).resolve().parent / "results" / datetime.now().strftime("%m-%Y")))

if __name__ == "__main__":
    main()