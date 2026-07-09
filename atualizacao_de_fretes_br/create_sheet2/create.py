from pathlib import Path
from datetime import datetime
import re
from time import time
from functools import lru_cache
import sys

from openpyxl import Workbook, load_workbook
import numpy as np

import pandas as pd
from math import ceil
import xlsxwriter

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from logging_utils import configure_logging


logger = configure_logging(__name__)
print = logger.info


def _carregar_total_linhas_esperadas(base_dir: Path) -> int | None:
    results_dir = base_dir.parent / "create_sheet" / "results"
    if not results_dir.exists():
        return None

    arquivos = sorted(
        arquivo for arquivo in results_dir.rglob("BASE_RODRIGO_*_TOTAL_LINHAS.txt")
        if arquivo.is_file()
    )
    if not arquivos:
        return None

    for arquivo in reversed(arquivos):
        try:
            return int(arquivo.read_text(encoding="utf-8").strip())
        except (OSError, ValueError):
            continue


    return None


def _resolver_diretorio_dados(base_dir: Path) -> Path:
    candidatos = [base_dir, base_dir.parent / "create_sheet"]
    for candidato in candidatos:
        if (candidato / "base.xlsx").exists():
            return candidato
    return base_dir


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
        return int(max(20, ceil(peso / 10.0) * 10))

    if peso <= 500:
        return int(max(250, ceil(peso / 50.0) * 50))

    return int(max(600, ceil(peso / 100.0) * 100))


def normalizar_cep(valor) -> str:
    if valor is None:
        return ""

    texto = "".join(caractere for caractere in str(valor) if caractere.isdigit())
    if not texto:
        return ""

    if texto.endswith(".0"):
        texto = texto[:-2]

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

def _localizar_ultimo_arquivo_intelipost(grupo_dir: Path) -> Path | None:
    if not grupo_dir.exists():
        return None
    arquivos = [a for a in grupo_dir.glob("TABELA_INTELIPOST_*.xlsx") if a.is_file()]
    if not arquivos:
        return None
    return max(arquivos, key=lambda a: a.stat().st_mtime)


def carregar_preco_atual_intelipost(grupo_dir: Path) -> tuple[dict[str, float], Path | None]:
    """
    Lê a planilha de origem (CEP DE na coluna G, pesos a partir da coluna H)
    e retorna um dict {"CEPNORM-PESO": preco}.
    """
    arquivo = _localizar_ultimo_arquivo_intelipost(grupo_dir)
    if arquivo is None:
        return {}, None

    workbook = load_workbook(arquivo, data_only=True, read_only=True)
    try:
        planilha = workbook[workbook.sheetnames[0]]
        linhas = list(planilha.iter_rows(values_only=True))
    finally:
        workbook.close()

    if not linhas:
        return {}, arquivo

    header = linhas[0]
    cep_idx = 5 
    pesos_por_indice: dict[int, int] = {}
    for i in range(7, len(header)):
        valor_header = header[i]
        if valor_header is None:
            continue
        texto = str(valor_header).strip().replace(",", ".")
        try:
            pesos_por_indice[i] = int(float(texto))
        except ValueError:
            continue

    lookup: dict[str, float] = {}
    for linha in linhas[1:]:
        if len(linha) <= cep_idx:
            continue
        cep_norm = normalizar_cep(linha[cep_idx])
        if not cep_norm:
            continue
        for indice_coluna, peso in pesos_por_indice.items():
            if indice_coluna >= len(linha):
                continue
            valor = linha[indice_coluna]
            if valor is None:
                continue
            lookup[f"{cep_norm}-{peso}"] = valor

    return lookup, arquivo

def carregar_base_cidades(base_dir: Path) -> dict[str, str]:
    linhas = _carregar_linhas_planilha(str(_resolver_diretorio_dados(base_dir)), "cidades")
    base_cidades = {}

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
    linhas = _carregar_linhas_planilha(str(_resolver_diretorio_dados(base_dir)), "venda")
    base_venda = {}

    for linha in linhas[1:]:
        cidade = linha[4] if len(linha) > 4 else None
        venda = linha[1] if len(linha) > 1 else None

        if cidade is None or venda is None:
            continue

        base_venda[str(cidade).strip()] = venda

    return base_venda


def carregar_base_frete(base_dir: Path) -> dict[str, str]:
    linhas = _carregar_linhas_planilha(str(_resolver_diretorio_dados(base_dir)), "frete")
    base_frete = {}

    for linha in linhas[1:]:
        chave = linha[3] if len(linha) > 3 else None
        valor = linha[4] if len(linha) > 4 else None

        if chave is None or valor is None:
            continue

        base_frete[str(chave).strip()] = valor

    return base_frete


def carregar_base_frete_detalhada(base_dir: Path) -> pd.DataFrame:
    linhas = _carregar_linhas_planilha(str(_resolver_diretorio_dados(base_dir)), "frete")
    registros = []

    for linha in linhas[1:]:
        regiao = linha[0] if len(linha) > 0 else None
        uf = linha[1] if len(linha) > 1 else None
        localidade = linha[2] if len(linha) > 2 else None
        regiao_uf_localidade = linha[3] if len(linha) > 3 else None
        percentual_frete = linha[4] if len(linha) > 4 else None

        if all(valor is None for valor in (regiao, uf, localidade, regiao_uf_localidade, percentual_frete)):
            continue

        registros.append((regiao, uf, localidade, regiao_uf_localidade, percentual_frete))

    return pd.DataFrame(
        registros,
        columns=["REGIÃO", "UF", "LOCALIDADE", "REGIÃO + UF + LOCALIDADE", "% FRETE"],
    )


def carregar_base_pesos(base_dir: Path) -> list[tuple]:
    linhas = _carregar_linhas_planilha(str(base_dir.parent / "create_sheet"), "pesos")
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
    linhas = _carregar_linhas_planilha(str(_resolver_diretorio_dados(base_dir)), "cidades")
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


def carregar_base_venda(base_dir: Path) -> dict[str, str]:
    linhas = _carregar_linhas_planilha(str(_resolver_diretorio_dados(base_dir)), "venda")
    base_venda = {}

    for linha in linhas[1:]:
        cidade = linha[4] if len(linha) > 4 else None
        venda = linha[1] if len(linha) > 1 else None

        if cidade is None or venda is None:
            continue

        base_venda[str(cidade).strip()] = venda

    return base_venda


def carregar_base_frete(base_dir: Path) -> dict[str, str]:
    linhas = _carregar_linhas_planilha(str(_resolver_diretorio_dados(base_dir)), "frete")
    base_frete = {}

    for linha in linhas[1:]:
        chave = linha[3] if len(linha) > 3 else None
        valor = linha[4] if len(linha) > 4 else None

        if chave is None or valor is None:
            continue

        base_frete[str(chave).strip()] = valor

    return base_frete


def carregar_base_frete_detalhada(base_dir: Path) -> pd.DataFrame:
    linhas = _carregar_linhas_planilha(str(_resolver_diretorio_dados(base_dir)), "frete")
    registros = []

    for linha in linhas[1:]:
        regiao = linha[0] if len(linha) > 0 else None
        uf = linha[1] if len(linha) > 1 else None
        localidade = linha[2] if len(linha) > 2 else None
        regiao_uf_localidade = linha[3] if len(linha) > 3 else None
        percentual_frete = linha[4] if len(linha) > 4 else None

        if all(valor is None for valor in (regiao, uf, localidade, regiao_uf_localidade, percentual_frete)):
            continue

        registros.append((regiao, uf, localidade, regiao_uf_localidade, percentual_frete))

    return pd.DataFrame(
        registros,
        columns=["REGIÃO", "UF", "LOCALIDADE", "REGIÃO + UF + LOCALIDADE", "% FRETE"],
    )


def carregar_base_pesos(base_dir: Path) -> list[tuple]:
    linhas = _carregar_linhas_planilha(str(base_dir.parent / "create_sheet"), "pesos")
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
    linhas = _carregar_linhas_planilha(str(_resolver_diretorio_dados(base_dir)), "cidades")
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
    diretorio_dados = _resolver_diretorio_dados(base_dir)
    arquivo_base = diretorio_dados / "base.xlsx"

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


def _exportar_tabela_frete_xlsx(
    resultado: pd.DataFrame,
    headers: list[str],
    path: Path,
    ac3_valor: float,
    aux_frete: pd.DataFrame | None = None,
) -> None:
    path_parent = Path(path).resolve().parent
    path_parent.mkdir(parents=True, exist_ok=True)
    resultado = resultado.astype(object).where(pd.notna(resultado), None)

    workbook_xlsx = xlsxwriter.Workbook(str(path), {"constant_memory": True})
    try:
        worksheet = workbook_xlsx.add_worksheet("TABELA DE FRETE")
        worksheet.write_row(0, 0, [None] * len(headers))
        worksheet.write_row(1, 0, headers)
        worksheet.write("AD2", "Adicional Custo")
        worksheet.write("AF2", "REGIÃO")
        worksheet.write("AG2", "UF")
        worksheet.write("AH2", "LOCALIDADE")
        worksheet.write("AI2", "REGIÃO + UF + LOCALIDADE")
        worksheet.write("AJ2", "% FRETE")
        worksheet.write("AD3", ac3_valor)
        aux_registros = []
        if aux_frete is not None and not aux_frete.empty:
            aux_frete = aux_frete.astype(object).where(pd.notna(aux_frete), None)
            aux_registros = list(aux_frete.itertuples(index=False, name=None))

        resultado_registros = list(resultado.itertuples(index=False, name=None))
        max_linhas = max(len(resultado_registros), len(aux_registros))
        for indice in range(max_linhas):
            linha_excel = indice + 2
            if indice < len(resultado_registros):
                worksheet.write_row(linha_excel, 0, resultado_registros[indice])
            if indice < len(aux_registros):
                worksheet.write_row(linha_excel, 31, aux_registros[indice])

        
    finally:
        workbook_xlsx.close()


def _linha_excel(indice_dataframe: int) -> int:
    return indice_dataframe + 3

def _carregar_ultimo_custo_atual(base_dir: Path) -> dict[str, float]:
    # Usa o último CSV gerado como snapshot para deslocar o valor atual para a coluna de valor antigo.
    results_dir = base_dir.parent / "create_sheet"/ "results"
    if not results_dir.exists():
        return {}

    arquivos_csv = [
        arquivo
        for arquivo in results_dir.rglob("TABELA_DE_FRETE_*.csv")
        if arquivo.is_file()
    ]
    if not arquivos_csv:
        return {}

    arquivos_csv_ordenados = sorted(arquivos_csv, key=lambda arquivo: arquivo.stat().st_mtime)
    if len(arquivos_csv_ordenados) < 2:
        return {}  # ou return "FRETE CUSTO {valor antigo}" nas funções de nome
    ultimo_arquivo = arquivos_csv_ordenados[-2]  # penúltimo
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
    results_dir = base_dir.parent / "create_sheet"/ "results"
    if not results_dir.exists():
        return {}

    arquivos_csv = [
        arquivo
        for arquivo in results_dir.rglob("TABELA_DE_FRETE_*.csv")
        if arquivo.is_file()
    ]
    if not arquivos_csv:
        return {}

    arquivos_csv_ordenados = sorted(arquivos_csv, key=lambda arquivo: arquivo.stat().st_mtime)
    if len(arquivos_csv_ordenados) < 2:
        return {}  # ou return "FRETE CUSTO {valor antigo}" nas funções de nome
    ultimo_arquivo = arquivos_csv_ordenados[-2]  # penúltimo
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
    results_dir = base_dir.parent / "create_sheet"/ "results"
    if not results_dir.exists():
        return "FRETE CUSTO {valor antigo}"

    arquivos_csv = [
        arquivo
        for arquivo in results_dir.rglob("TABELA_DE_FRETE_*.csv")
        if arquivo.is_file()
    ]
    if not arquivos_csv:
        return "FRETE CUSTO {valor antigo}"

    arquivos_csv_ordenados = sorted(arquivos_csv, key=lambda arquivo: arquivo.stat().st_mtime)
    if len(arquivos_csv_ordenados) < 2:
        return {}  # ou return "FRETE CUSTO {valor antigo}" nas funções de nome
    ultimo_arquivo = arquivos_csv_ordenados[-2]  # penúltimo
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
    results_dir = base_dir.parent / "create_sheet"/ "results"
    if not results_dir.exists():
        return "FRETE COMERCIAL {valor antigo}"

    arquivos_csv = [
        arquivo
        for arquivo in results_dir.rglob("TABELA_DE_FRETE_*.csv")
        if arquivo.is_file()
    ]
    if not arquivos_csv:
        return "FRETE COMERCIAL {valor antigo}"

    arquivos_csv_ordenados = sorted(arquivos_csv, key=lambda arquivo: arquivo.stat().st_mtime)
    if len(arquivos_csv_ordenados) < 2:
        return {}  # ou return "FRETE CUSTO {valor antigo}" nas funções de nome
    ultimo_arquivo = arquivos_csv_ordenados[-2]  # penúltimo
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


def base_rodrigo(base_dir: Path) -> pd.DataFrame:
    results_dir = base_dir.parent / "create_sheet" / "results"
    if not results_dir.exists():
        raise FileNotFoundError(f"Diretório de resultados não encontrado: {results_dir}")

    arquivos_csv = [
        arquivo
        for arquivo in results_dir.rglob("BASE_RODRIGO_*.csv")
        if arquivo.is_file()
    ]
    if not arquivos_csv:
        raise FileNotFoundError(f"Nenhum arquivo BASE_RODRIGO_*.csv encontrado no diretório: {results_dir}")

    base_rodrigo_csv = max(arquivos_csv, key=lambda arquivo: arquivo.stat().st_mtime)

    df = pd.read_csv(base_rodrigo_csv, sep=";", decimal=",", encoding="utf-8-sig", low_memory=False)
    df.columns = [
        "UF",
        "REGIAO",
        "CEP_COTDO",
        "SKU",
        "FAIXA_PESO_ORIG",
        "PESO_PRODUTO",
        "FAIXA_PESO_CALC",
        "VALOR_PRODUTO",
        "ID_COTACAO",
        "ID_TRANSPORTADOR",
        "TRANSPORTADOR",
        "PRAZO_ENTREGA",
        "VALOR_COMERCIAL",
        "VALOR_TRANSPORTADOR",
        "UF_CIDADE_PESO",
        "CIDADE",
        "CEP_PESO_KEY",
    ]
    df["CEP_NORMALIZADO"] = df["CEP_COTDO"].map(normalizar_cep)
    df["PESO_PRODUTO"] = pd.to_numeric(df["PESO_PRODUTO"], errors="coerce")
    df["VALOR_PRODUTO"] = pd.to_numeric(df["VALOR_PRODUTO"], errors="coerce")
    df["ID_TRANSPORTADOR"] = pd.to_numeric(df["ID_TRANSPORTADOR"], errors="coerce")
    df["PRAZO_ENTREGA"] = pd.to_numeric(df["PRAZO_ENTREGA"], errors="coerce")
    df["VALOR_COMERCIAL"] = pd.to_numeric(df["VALOR_COMERCIAL"], errors="coerce")
    df["VALOR_TRANSPORTADOR"] = pd.to_numeric(df["VALOR_TRANSPORTADOR"], errors="coerce")
    df["FAIXA_PESO_CALC"] = pd.to_numeric(df["FAIXA_PESO_CALC"], errors="coerce").round().astype("Int64")
    print("Base Rodrigo carregada")
    return df


def tabela_frete(base_rodrigo_df: pd.DataFrame, base_venda: dict, base_cidades_detalhadas: dict, base_frete: dict[str, str], base_pesos: list[tuple]) -> None:
    print("Gerando tabela de frete")
    worksheet_name = "TABELA DE FRETE"
    mes_ano_atual = datetime.now().strftime("%m-%Y")
    base_dir = Path(__file__).resolve().parent
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
        "FRETE ATUAL - GRUPO 1",
        "FRETE ATUAL - GRUPO 2",
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
    workbook = Workbook()
    if len(workbook.worksheets) == 1 and workbook.active.title == "Sheet":
        workbook.remove(workbook.active)
    planilha_saida = workbook.create_sheet(title=worksheet_name)
    planilha_saida.append([None] * len(headers))
    planilha_saida.append(headers)
    planilha_saida["ac3"] = ac3_valor

    df = base_rodrigo_df.copy().reset_index(drop=True)
    total_linhas_esperadas = _carregar_total_linhas_esperadas(base_dir) or len(df)
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
        lambda valor: ""
        if pd.isna(valor)
        else str(int(valor))
    )

    cep_de_chave = df["CEP_DE"].map(normalizar_cep)

    df["CEP DE + PESO POR"] = (
        cep_de_chave + "-" +peso_por_chave
    )

    grupo1_dir = ROOT_DIR / "separacao_intelipost" / "intelipost_add_value" / "grupo1" / "results"
    grupo2_dir = ROOT_DIR / "separacao_intelipost" / "intelipost_add_value" / "grupo2" / "results"
    lookup_grupo1, arquivo_grupo1 = carregar_preco_atual_intelipost(grupo1_dir)
    lookup_grupo2, arquivo_grupo2 = carregar_preco_atual_intelipost(grupo2_dir)

    df["FRETE ATUAL - GRUPO 1"] = df["CEP DE + PESO POR"].map(lookup_grupo1)
    df["FRETE ATUAL - GRUPO 2"] = df["CEP DE + PESO POR"].map(lookup_grupo2)

    sem_match_g1 = df["CEP DE + PESO POR"].notna().sum() - df["FRETE ATUAL - GRUPO 1"].notna().sum()
    sem_match_g2 = df["CEP DE + PESO POR"].notna().sum() - df["FRETE ATUAL - GRUPO 2"].notna().sum()
    print(f"FRETE ATUAL - GRUPO 1: {sem_match_g1} linha(s) sem correspondência")
    print(f"FRETE ATUAL - GRUPO 2: {sem_match_g2} linha(s) sem correspondência")


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
        print(
            f"Erro: junção com base_rodrigo_lookup alterou quantidade de linhas "
            f"({len(df)} != {total_linhas_esperadas})")
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


    frete_df = pd.DataFrame(list(base_frete.items()), columns=["REGIAO_UF_LOCALIDADE", "ADICIONAL_FRETE"])
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
        "FRETE ATUAL - GRUPO 1": df["FRETE ATUAL - GRUPO 1"],
        "FRETE ATUAL - GRUPO 2": df["FRETE ATUAL - GRUPO 2"],
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
        "FRETE ATUAL - GRUPO 1",
        "FRETE ATUAL - GRUPO 2",
        coluna_custo_antigo,
        coluna_frete_antigo,
        "FRETE CUSTO ATUAL",
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


    resultado_xlsx = resultado.copy().astype(object)
    resultado_xlsx["FRETE CUSTO ATUAL"] = [
        f"={valor}+$AD$3" if pd.notna(valor) else None
        for valor in pd.to_numeric(df["VALOR_TRANSPORTADOR_BASE_RODRIGO"], errors="coerce")
    ]
    resultado_xlsx["__ESPACO_AB"] = None
    resultado_xlsx["__ESPACO_AC"] = None
    headers = headers + ["", ""]

    linhas_excel = [_linha_excel(int(indice)) for indice in resultado_xlsx.index]
    resultado_xlsx["% CUSTO FRETE"] = [
        f"=XLOOKUP(AB{linha},AI:AI,AJ:AJ)"
        for linha in linhas_excel
    ]
    resultado_xlsx["FRETE COMERCIAL X FRETE REAL (R$)"] = [
        f"=P{linha}-W{linha}"
        for linha in linhas_excel
    ]
    resultado_xlsx["FRETE COMERCIAL X FRETE REAL (%)"] = [
        f"=P{linha}/W{linha}-1"
        for linha in linhas_excel
    ]
    resultado_xlsx["FRETE COMERCIAL ATUAL"] = [
        f"=ROUNDUP(X{linha}+(X{linha}*AC{linha}),-1)-0.01"
        for linha in linhas_excel
    ]
    resultado_xlsx["DIF. FRETE CUSTO R$"] = [
        f"=U{linha}-W{linha}"
        for linha in linhas_excel
    ]
    resultado_xlsx["DIFERENÇA %"] = [
        f"=W{linha}/U{linha}-1"
        for linha in linhas_excel
    ]

    aux_frete = carregar_base_frete_detalhada(Path(__file__).resolve().parent)

    mes_ano_atual = datetime.now().strftime("%m-%Y")
    data_atual = datetime.now().strftime("%d-%m-%Y")
    results_dir = Path(__file__).resolve().parent / "results" / mes_ano_atual
    results_dir.mkdir(parents=True, exist_ok=True)

    xlsx_path = results_dir / f"TABELA_DE_FRETE_{data_atual}.xlsx"
    _exportar_tabela_frete_xlsx(resultado_xlsx, headers, xlsx_path, ac3_valor, aux_frete=aux_frete)
    workbook_verificacao = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        planilha_verificacao = workbook_verificacao["TABELA DE FRETE"]
        total_linhas_arquivo = sum(
            1
            for linha in planilha_verificacao.iter_rows(min_row=3, values_only=True)
            if any(valor is not None for valor in linha)
        )
    finally:
        workbook_verificacao.close()

    if total_linhas_arquivo != total_linhas_esperadas:
        print(f"Erro: arquivo final com quantidade de linhas incorreta ({total_linhas_arquivo} != {total_linhas_esperadas})")
        raise SystemExit(1)

    print("Tabela de frete criada")  


def main() -> None:
    start_time = time()
    base_dir = Path(__file__).resolve().parent
    _, base_venda, base_cidades_detalhadas = carregar_bases_unicas(base_dir)
    base_frete = carregar_base_frete(base_dir)
    base_pesos = carregar_base_pesos(base_dir)

    base_rodrigo_df = base_rodrigo(base_dir)
    tabela_frete(base_rodrigo_df, base_venda, base_cidades_detalhadas, base_frete, base_pesos)
    
    elapsed_time = time() - start_time
    minutes = int(elapsed_time // 60)
    seconds = int(elapsed_time % 60)
    print(f"Tempo total de execução: {minutes}m e {seconds}s")


if __name__ == "__main__":
    main()