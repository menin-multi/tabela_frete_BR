from datetime import datetime
from pathlib import Path
from typing import Optional
import re
import sys
ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))
from env_helper import get_main_dir
MAIN_DIR = get_main_dir()
import pandas as pd
from openpyxl import Workbook
from logging_utils import configure_logging
logger = configure_logging(__name__)
print = logger.info

# configuração de data
dia_mes = datetime.now().strftime("%d-%m")

COLUNAS_ORIGEM = [
    "CEP DE", "CEP ATÉ", "UF+CIDADE",
    "10", "20", "30", "40", "50", "60", "70", "80", "90", "100",
    "110", "120", "130", "140", "150", "160", "170", "180", "190",
    "200", "250", "300", "350", "400", "450", "500",
]

COLUNAS_SAIDA = [
    "ZipCodeStart", "ZipCodeEnd", "WeightStart", "WeightEnd",
    "AbsoluteMoneyCost", "TRANSPORTADORA", "UF+CIDADE", "TimeCost",
]


def _escrever_dataframe_em_planilha(planilha, dataframe: pd.DataFrame) -> None:
    for linha in dataframe.itertuples(index=False, name=None):
        planilha.append([None if pd.isna(valor) else valor for valor in linha])


def _calcular_peso_inicial(peso_faixa):
    mapa_peso_inicial = {
        10000: 0, 20000: 10001, 30000: 20001, 40000: 30001, 50000: 40001,
        60000: 50001, 70000: 60001, 80000: 70001, 90000: 80001, 100000: 90001,
        110000: 100001, 120000: 110001, 130000: 120001, 140000: 130001,
        150000: 140001, 160000: 150001, 170000: 160001, 180000: 170001,
        190000: 180001, 200000: 190001, 250000: 200001, 300000: 250001,
        350000: 300001, 400000: 350001, 450000: 400001, 500000: 450001,
    }
    if pd.isna(peso_faixa):
        return pd.NA
    try:
        peso_faixa = int(peso_faixa)
    except (TypeError, ValueError):
        return pd.NA
    return mapa_peso_inicial.get(peso_faixa, pd.NA)


def _normalizar_cep(valor) -> str:
    if pd.isna(valor):
        return pd.NA
    cep = re.sub(r"\D", "", str(valor))
    cep = cep.zfill(8)
    return cep if cep != "00000000" else pd.NA


def separar_transportadoras(arquivo_tabela_unificada: Path) -> list[Path]:
    if not arquivo_tabela_unificada.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_tabela_unificada}")
    df = pd.read_excel(arquivo_tabela_unificada)
    if "TRANSPORTADORA" not in df.columns:
        raise ValueError("Coluna 'TRANSPORTADORA' não encontrada na tabela unificada.")
    results_dir = MAIN_DIR / "atualizacao_vtex" / "results" / "precificacao_br"
    results_dir.mkdir(parents=True, exist_ok=True)
    arquivos_gerados: list[Path] = []
    for transportadora, grupo in df.groupby("TRANSPORTADORA", dropna=True):
        if pd.isna(transportadora):
            continue
        nome_transportadora = str(transportadora).strip().upper()
        if not nome_transportadora:
            continue
        nome_arquivo = f"TABELA {nome_transportadora}.xlsx"
        arquivo_saida = results_dir / nome_arquivo
        workbook_saida = Workbook()
        if len(workbook_saida.sheetnames) == 1 and workbook_saida.sheetnames[0] == "Sheet":
            del workbook_saida[workbook_saida.sheetnames[0]]
        planilha_saida = workbook_saida.create_sheet(title=nome_transportadora[:31])
        planilha_saida.append(list(grupo.columns))
        _escrever_dataframe_em_planilha(planilha_saida, grupo)
        workbook_saida.save(arquivo_saida)
        arquivos_gerados.append(arquivo_saida)
    return arquivos_gerados


def _salvar_erros(erros_df: pd.DataFrame) -> None:
    """Salva linhas com erro de cotação em results/erros/ para consulta."""
    if erros_df.empty:
        return
    erros_dir = MAIN_DIR / "atualizacao_vtex" / "results" / "erros"
    erros_dir.mkdir(parents=True, exist_ok=True)
    data_atual = datetime.now().strftime("%d-%m-%Y")
    arquivo_erros = erros_dir / f"ERROS_COTACAO_{data_atual}.xlsx"
    workbook_erros = Workbook()
    planilha_erros = workbook_erros.active
    planilha_erros.title = "Erros"
    planilha_erros.append(list(erros_df.columns))
    _escrever_dataframe_em_planilha(planilha_erros, erros_df)
    workbook_erros.save(arquivo_erros)
    print(f"{len(erros_df)} linha(s) com erro salvas em: {arquivo_erros}")


def _carregar_cotacoes(base_dir: Optional[Path] = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    base_dir = Path(base_dir) if base_dir is not None else MAIN_DIR
    quote_dir = base_dir / "atualizacao_vtex" / "results" / "quote_results" / dia_mes
    if not quote_dir.exists():
        raise FileNotFoundError(f"Diretório de resultados não encontrado: {quote_dir}")
    arquivos_quote = [
        arquivo for arquivo in quote_dir.iterdir()
        if arquivo.is_file()
        and not arquivo.name.startswith("~$")
        and arquivo.suffix.lower() in {".xlsx", ".xls", ".csv"}
    ]
    if not arquivos_quote:
        raise FileNotFoundError(f"Nenhum arquivo de cotação encontrado em: {quote_dir}")
    dataframes = []
    for arquivo in sorted(arquivos_quote, key=lambda a: a.name):
        if arquivo.suffix.lower() in {".xlsx", ".xls"}:
            dataframes.append(pd.read_excel(arquivo))
        else:
            dataframes.append(pd.read_csv(arquivo, sep=";", decimal=",", encoding="utf-8-sig", low_memory=False))
    if not dataframes:
        raise ValueError(f"Falha ao carregar arquivos de cotação em: {quote_dir}")
    df = pd.concat(dataframes, ignore_index=True, sort=False)

    tem_coluna_erro = "erro" in df.columns
    if tem_coluna_erro:
        erros_df = df[df["erro"].notna() & (df["erro"].astype(str).str.strip() != "")].copy()
        df = df[df["erro"].isna() | (df["erro"].astype(str).str.strip() == "")].copy()
    else:
        erros_df = pd.DataFrame(columns=df.columns)

    df["cep_normalizado"] = df["cep"].astype("string").map(_normalizar_cep)
    df = df.drop_duplicates(subset=["cep_normalizado", "transportadora id"], keep="first")
    return df, erros_df


def _carregar_cep_cotacao(base_dir: Optional[Path] = None) -> pd.DataFrame:
    base_dir = Path(base_dir) if base_dir is not None else MAIN_DIR
    cep_path = base_dir / "atualizacao_vtex" / "cotacao" / "cep_cotacao.xlsx"
    if not cep_path.exists():
        raise FileNotFoundError(f"Arquivo cep_cotacao não encontrado: {cep_path}")
    df = pd.read_excel(cep_path)
    if "CEP" not in df.columns or "CEP ATÉ" not in df.columns:
        raise ValueError("Arquivo cep_cotacao precisa ter as colunas 'CEP' e 'CEP ATÉ'.")
    df["CEP"] = df["CEP"].map(_normalizar_cep)
    df["CEP ATÉ"] = df["CEP ATÉ"].map(_normalizar_cep)
    if "UF+CIDADE" in df.columns:
        df["UF+CIDADE"] = df["UF+CIDADE"].astype("string").str.strip()
    elif "UF" in df.columns and "CIDADE" in df.columns:
        df["UF+CIDADE"] = df["UF"].astype("string").str.strip() + ":" + df["CIDADE"].astype("string").str.strip()
    else:
        raise ValueError("Arquivo cep_cotacao precisa ter 'UF+CIDADE' ou as colunas 'UF' e 'CIDADE'.")
    return df.sort_values("CEP").drop_duplicates(subset=["CEP"], keep="first")


def _carregar_tabela_intelipost(base_dir: Optional[Path] = None) -> pd.DataFrame:
    base_dir = Path(base_dir) if base_dir is not None else MAIN_DIR
    results_dir = base_dir / "create_sheet" / "results"
    if not results_dir.exists():
        raise FileNotFoundError(f"Diretório de resultados não encontrado: {results_dir}")
    arquivos_csv = [
        arquivo for arquivo in results_dir.rglob("TABELA_INTELIPOST_*.csv")
        if arquivo.is_file()
    ]
    if not arquivos_csv:
        raise FileNotFoundError(f"Nenhum arquivo TABELA_INTELIPOST_*.csv encontrado em: {results_dir}")
    ultimo_arquivo = max(arquivos_csv, key=lambda arquivo: arquivo.stat().st_mtime)
    df = pd.read_csv(ultimo_arquivo, sep=";", decimal=",", encoding="utf-8-sig", low_memory=False)
    for coluna in COLUNAS_ORIGEM:
        if coluna not in df.columns:
            df[coluna] = pd.NA
    df["CEP DE"] = df["CEP DE"].map(_normalizar_cep)
    df["CEP ATÉ"] = df["CEP ATÉ"].map(_normalizar_cep)
    df["UF+CIDADE"] = df["UF+CIDADE"].astype("string").str.strip()
    return df


def tabela_vtex(workbook_saida: Optional[Workbook] = None) -> Path:
    start_time = datetime.now()

    cotacoes_df, erros_cotacao_df = _carregar_cotacoes()
    cep_cotacao_df = _carregar_cep_cotacao()
    intelipost_df = _carregar_tabela_intelipost()
    
	# mapa ID → TRANSPORTADORA vindo do cep_cotacao (apenas para erros)
    mapa_transportador = (
        cep_cotacao_df[["ID", "TRANSPORTADORA"]]
        .dropna(subset=["ID", "TRANSPORTADORA"])
        .drop_duplicates(subset=["ID"], keep="first")
        .set_index("ID")["TRANSPORTADORA"]
        .to_dict()
    )
    if not erros_cotacao_df.empty and "transportadora id" in erros_cotacao_df.columns:
        erros_cotacao_df["transportador"] = erros_cotacao_df["transportadora id"].map(mapa_transportador)

    # mapa CEP → CIDADE vindo do cep_cotacao (apenas para erros)
    mapa_cidade = (
        cep_cotacao_df[["CEP", "CIDADE"]]
        .dropna(subset=["CEP", "CIDADE"])
        .drop_duplicates(subset=["CEP"], keep="first")
        .set_index("CEP")["CIDADE"]
        .to_dict()
    )
    if not erros_cotacao_df.empty and "cep" in erros_cotacao_df.columns:
        erros_cotacao_df["cidade"] = erros_cotacao_df["cep"].astype("string").map(_normalizar_cep).map(mapa_cidade)

    for coluna in ["cep", "transportadora id", "prazo_entrega"]:
        if coluna not in cotacoes_df.columns:
            raise ValueError(f"Arquivo de cotação precisa ter a coluna '{coluna}'.")

    merged = cotacoes_df.merge(
        cep_cotacao_df[["CEP", "CEP ATÉ", "UF+CIDADE"]],
        left_on="cep_normalizado",
        right_on="CEP",
        how="left",
    )
    if merged["CEP ATÉ"].isna().all():
        raise ValueError("Nenhuma correspondência entre quote_results e cep_cotacao.")

    colunas_faixa = [col for col in COLUNAS_ORIGEM[3:] if col in intelipost_df.columns]
    if not colunas_faixa:
        raise ValueError("Nenhuma coluna de faixa de peso encontrada na tabela Intelipost.")

    intelipost_long = intelipost_df.melt(
        id_vars=["CEP DE", "CEP ATÉ", "UF+CIDADE"],
        value_vars=colunas_faixa,
        var_name="FAIXA DE PESO",
        value_name="PREÇO",
    )
    intelipost_long["FAIXA DE PESO"] = (
        pd.to_numeric(intelipost_long["FAIXA DE PESO"], errors="coerce")
        .round()
        .astype("Int64")
    )

    faixas_df = pd.DataFrame({"FAIXA DE PESO": intelipost_long["FAIXA DE PESO"].unique()})
    merged["_tmp"] = 1
    faixas_df["_tmp"] = 1
    resultado = merged.merge(faixas_df, on="_tmp").drop(columns="_tmp")

    resultado = resultado.merge(
        intelipost_long[["CEP DE", "FAIXA DE PESO", "PREÇO"]],
        left_on=["cep_normalizado", "FAIXA DE PESO"],
        right_on=["CEP DE", "FAIXA DE PESO"],
        how="left",
    )

    # — consolidar erros —
    COLUNAS_ERRO = ["uf", "regiao", "cidade", "cep", "FAIXA DE PESO", "transportadora id", "id_cotacao", "transportador", "prazo_entrega", "erro"]

    sem_intelipost = resultado[resultado["PREÇO"].isna()].copy()
    if not sem_intelipost.empty:
        sem_intelipost["erro"] = "Sem cotação intelipost"
        # manter apenas colunas de cotação + FAIXA DE PESO
        for col in COLUNAS_ERRO:
            if col not in sem_intelipost.columns:
                sem_intelipost[col] = pd.NA
        sem_intelipost = sem_intelipost[COLUNAS_ERRO]

    # expandir erros de cotação pelas faixas de peso disponíveis
    faixas_erro = pd.DataFrame({"FAIXA DE PESO": intelipost_long["FAIXA DE PESO"].unique()})
    erros_cotacao_df["_tmp"] = 1
    faixas_erro["_tmp"] = 1
    erros_cotacao_df = erros_cotacao_df.merge(faixas_erro, on="_tmp").drop(columns="_tmp")
    for col in COLUNAS_ERRO:
        if col not in erros_cotacao_df.columns:
            erros_cotacao_df[col] = pd.NA
    erros_cotacao_df = erros_cotacao_df[COLUNAS_ERRO]

    todos_erros = pd.concat([erros_cotacao_df, sem_intelipost], ignore_index=True)
    _salvar_erros(todos_erros)

    resultado = resultado[resultado["PREÇO"].notna()].copy()

    # — montar colunas de saída —
    resultado["WeightEnd"] = resultado["FAIXA DE PESO"].astype("float64") * 1000
    resultado["WeightStart"] = resultado["WeightEnd"].map(_calcular_peso_inicial)
    resultado["AbsoluteMoneyCost"] = pd.to_numeric(resultado["PREÇO"], errors="coerce").astype("Float64")
    resultado["ZipCodeStart"] = resultado["cep_normalizado"].astype("string")
    resultado["ZipCodeEnd"] = resultado["CEP ATÉ"].astype("string")
    resultado["TRANSPORTADORA"] = resultado["transportador"].astype("string").str.strip()
    resultado["TimeCost"] = resultado["prazo_entrega"]
    resultado["UF+CIDADE"] = resultado["UF+CIDADE"].astype("string").str.strip()

    resultado = resultado[COLUNAS_SAIDA].sort_values(
        ["ZipCodeStart", "WeightEnd", "TRANSPORTADORA", "TimeCost"],
        na_position="last",
    ).reset_index(drop=True)

    if resultado.empty:
        raise ValueError("Nenhum registro foi gerado para a tabela VTEX.")

    if workbook_saida is None:
        workbook_saida = Workbook()
    if len(workbook_saida.sheetnames) == 1 and workbook_saida.sheetnames[0] == "Sheet":
        del workbook_saida[workbook_saida.sheetnames[0]]

    planilha_saida = workbook_saida.create_sheet(title="TABELA - VTEX")
    planilha_saida.append(COLUNAS_SAIDA)
    _escrever_dataframe_em_planilha(planilha_saida, resultado)

    results_dir = MAIN_DIR / "atualizacao_vtex" / "results"
    results_dir.mkdir(parents=True, exist_ok=True)
    data_atual = datetime.now().strftime("%d-%m-%Y")
    arquivo_saida = results_dir / f"TABELA_UNIFICADA_VTEX_{data_atual}.xlsx"
    workbook_saida.save(arquivo_saida)

    separar_transportadoras(arquivo_saida)

    elapsed_time = datetime.now() - start_time
    minutes = elapsed_time.total_seconds() // 60
    seconds = elapsed_time.total_seconds() % 60
    print(f"Tabela VTEX gerada com sucesso em {minutes:.0f}m{seconds:.0f}s.")

    return arquivo_saida


if __name__ == "__main__":
    tabela_vtex()