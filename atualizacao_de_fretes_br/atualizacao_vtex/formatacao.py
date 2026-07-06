from pathlib import Path
import importlib
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
	sys.path.insert(0, str(ROOT_DIR))

from env_helper import get_main_dir
MAIN_DIR = get_main_dir()
RESULTS_DIR = MAIN_DIR / 'atualizacao_vtex' /"results" / "precificacao_br"

import pandas as pd

COLUNAS_FINAIS = [
	"ZipCodeStart",
	"ZipCodeEnd",
	"PolygonName",
	"WeightStart",
	"WeightEnd",
	"AbsoluteMoneyCost",
	"PricePercent",
	"PriceByExtraWeight",
	"MaxVolume",
	"TimeCost",
	"Country",
	"MinimumValueInsurance",
]


def _number_format_decimal() -> str:
	return "0.000000"


def _escrever_dataframe_em_xls(
	arquivo_saida: Path,
	nome_aba: str,
	dados_dataframe: pd.DataFrame,
	colunas_decimais: set[str],
) -> None:
	xlwt = importlib.import_module("xlwt")
	workbook = xlwt.Workbook()
	worksheet = workbook.add_sheet(str(nome_aba)[:31] or "Sheet1")

	estilo_decimal = xlwt.easyxf(num_format_str=_number_format_decimal())

	for indice_coluna, nome_coluna in enumerate(dados_dataframe.columns):
		worksheet.write(0, indice_coluna, nome_coluna)

	for indice_linha, linha in enumerate(dados_dataframe.itertuples(index=False, name=None), start=1):
		for indice_coluna, valor in enumerate(linha):
			nome_coluna = dados_dataframe.columns[indice_coluna]
			if pd.isna(valor):
				worksheet.write(indice_linha, indice_coluna, "")
			elif nome_coluna in colunas_decimais:
				worksheet.write(indice_linha, indice_coluna, float(valor), estilo_decimal)
			else:
				worksheet.write(indice_linha, indice_coluna, valor)

	workbook.save(str(arquivo_saida))


def _normalizar_planilha(arquivo: Path) -> Path:
	with pd.ExcelFile(arquivo) as excel_file:
		nome_aba = excel_file.sheet_names[0]
		dados_dataframe = pd.read_excel(excel_file, sheet_name=nome_aba)

	if dados_dataframe.empty and not len(dados_dataframe.columns):
		arquivo_saida = arquivo.with_suffix(".xls")
		_escrever_dataframe_em_xls(arquivo_saida, nome_aba, dados_dataframe, set())
		return arquivo_saida

	dados_dataframe = dados_dataframe.loc[:, ~dados_dataframe.columns.duplicated()].copy()
	colunas_remover = {"PolygonName", "PricePercent", "PriceByExtraWeight", "MaxVolume", "Country", "MinimumValueInsurance"}
	dados_dataframe = dados_dataframe.drop(columns=[coluna for coluna in dados_dataframe.columns if coluna in colunas_remover], errors="ignore")

	if "TRANSPORTADORA" in dados_dataframe.columns:
		dados_dataframe = dados_dataframe.drop(columns=["TRANSPORTADORA"])

	dados_dataframe = dados_dataframe.loc[:, dados_dataframe.columns.intersection(COLUNAS_FINAIS)].copy()
	for coluna in COLUNAS_FINAIS:
		if coluna not in dados_dataframe.columns:
			dados_dataframe[coluna] = pd.NA

	dados_dataframe = dados_dataframe.reindex(columns=COLUNAS_FINAIS)
	dados_dataframe["PolygonName"] = ""
	dados_dataframe["PricePercent"] = 0.0
	dados_dataframe["PriceByExtraWeight"] = 0.0
	dados_dataframe["MaxVolume"] = 9999999999
	dados_dataframe["Country"] = "BRA"
	dados_dataframe["MinimumValueInsurance"] = 0.0
	dados_dataframe["TimeCost"] = dados_dataframe["TimeCost"].apply(
		lambda valor: None if pd.isna(valor) else f"{int(valor)}.00:00:00"
	)

	# Remover hífens dos CEPs
	for coluna in ["ZipCodeStart", "ZipCodeEnd"]:
		if coluna in dados_dataframe.columns:
			dados_dataframe[coluna] = (
				dados_dataframe[coluna]
				.astype(str)
				.str.replace("-", "", regex=False)
			)

	arquivo_saida = arquivo.with_suffix(".xls")
	colunas_decimais = {"PricePercent", "PriceByExtraWeight", "MinimumValueInsurance"}
	_escrever_dataframe_em_xls(arquivo_saida, nome_aba, dados_dataframe, colunas_decimais)
	if arquivo != arquivo_saida and arquivo.exists():
		arquivo.unlink()

	return arquivo_saida


def formatar_arquivos() -> list[Path]:
	if not RESULTS_DIR.exists():
		raise FileNotFoundError(f"Diretório não encontrado: {RESULTS_DIR}")

	arquivos_processados: list[Path] = []
	for arquivo in sorted(RESULTS_DIR.glob("*.xlsx")):
		if arquivo.name.startswith("~$"):
			continue
		arquivo_processado = _normalizar_planilha(arquivo)
		arquivos_processados.append(arquivo_processado)

	return arquivos_processados


if __name__ == "__main__":
	formatar_arquivos()
