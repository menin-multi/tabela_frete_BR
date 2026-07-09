from datetime import datetime
from pathlib import Path
from math import ceil
import re
from time import time
from copy import copy
import sys
import os
import glob

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
	sys.path.insert(0, str(ROOT_DIR))

from logging_utils import configure_logging

BASE_DIR = Path(__file__).resolve().parent
NOME_ABA = "TABELA DE FRETE"
logger = configure_logging(__name__)
print = logger.info

FILL_CABECALHO = PatternFill(fill_type="solid", fgColor="1F4E78")
FILL_VENDA = PatternFill(fill_type="solid", fgColor="00B050")
FILL_RANGE_PESO = PatternFill(fill_type="solid", fgColor="D9EAF7")
FILL_FRETES = PatternFill(fill_type="solid", fgColor="FCE4D6")
FILL_PRETO = PatternFill(fill_type="solid", fgColor="000000")
FILL_FAIXA_0_20 = PatternFill(fill_type="solid", fgColor="CC99FF")
FILL_FAIXA_20_40 = PatternFill(fill_type="solid", fgColor="FFC000")
FILL_FAIXA_40_60 = PatternFill(fill_type="solid", fgColor="92D050")
FILL_FAIXA_60_70 = PatternFill(fill_type="solid", fgColor="FFFFFF")
FILL_FAIXA_70_90 = PatternFill(fill_type="solid", fgColor="BDD7EE")
FILL_FAIXA_90_120 = PatternFill(fill_type="solid", fgColor="F4B084")
FILL_FAIXA_120_160 = PatternFill(fill_type="solid", fgColor="8497B0")
FILL_FAIXA_160_200 = PatternFill(fill_type="solid", fgColor="FFD966")
FILL_FAIXA_200_1500 = PatternFill(fill_type="solid", fgColor="FFFFFF")
FILL_AC_AG = PatternFill(fill_type="solid", fgColor="FFC000")
FILL_VENDA_CABECALHO = PatternFill(fill_type="solid", fgColor="FF0000")
FILL_FRETE_REAL_CABECALHO = PatternFill(fill_type="solid", fgColor="F4B084")
FILL_COND_POSITIVO = PatternFill(fill_type="solid", start_color="C6EFCE", end_color="C6EFCE")
FILL_COND_ZERO = PatternFill(fill_type="solid", start_color="FFEB9C", end_color="FFEB9C")
FILL_COND_NEGATIVO = PatternFill(fill_type="solid", start_color="FFC7CE", end_color="FFC7CE")

FONT_CABECALHO = Font(name="Calibri", size=10, color="FFFFFF", bold=True)
FONT_CABECALHO_PRETA = Font(name="Calibri", size=10, color="000000", bold=True)
THIN_SIDE = Side(style="thin", color="000000")
BORDER_PADRAO = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
ALINHAMENTO_CENTRO = Alignment(horizontal="center", vertical="center")


def encontrar_xlsx_mais_recente(pasta_base):
    padrao = os.path.join(pasta_base, "**", "TABELA_DE_FRETE*.xlsx")
    arquivos = glob.glob(padrao, recursive=True)
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo TABELA_DE_FRETE*.xlsx encontrado em {pasta_base}")
    return max(arquivos, key=os.path.getmtime)


def _aplicar_preenchimento_coluna(worksheet, coluna: int, fill: PatternFill, start_row: int = 3) -> None:
	for linha in range(start_row, worksheet.max_row + 1):
		cell = worksheet.cell(row=linha, column=coluna)
		if cell.value is None:
			continue
		cell.fill = fill
		cell.border = BORDER_PADRAO
		cell.alignment = ALINHAMENTO_CENTRO


def _aplicar_preenchimento_intervalo(worksheet, coluna_inicial: int, coluna_final: int, fill: PatternFill) -> None:
	for coluna in range(coluna_inicial, coluna_final + 1):
		_aplicar_preenchimento_coluna(worksheet, coluna, fill)


def largura_colunas(worksheet) -> None:
	larguras = {
		"A": 13.57,
		"B": 14,
		"C": 7.43,
		"D": 8.43,
		"E": 8.43,
		"F": 8.43,
		"G": 8.43,
		"H": 8.43,
		"I": 8.43,
		"J": 20.29,
		"K": 22,
		"L": 26.14,
		"M": 18.43,
		"N": 14,
		"O": 14,
		"P": 15.43,
		"Q": 15.43,
		"R": 15.43,
		"S": 11.86,
		"T": 20.14,
		"U": 12.3,
		"V": 15,
		"W": 12,
		"X": 14,
		"Y": 12,
		"Z": 11.30,
		"AA": 24,
		"AB": 30.71,
		"AC": 14.14,
		"AD": 9,
		"AE": 0.6,
		"AF": 11.57,
		"AG": 5.71,
		"AH": 13.57,
		"AI": 30.71,
		"AJ": 7.43,
	}

	for coluna, largura in larguras.items():
		worksheet.column_dimensions[coluna].width = largura


def cores_cabeçalhos(worksheet) -> None:
	for celula in worksheet[2]:
		if celula.value is None:
			continue
		if celula.value == "VENDA 2025":
			celula.fill = FILL_VENDA_CABECALHO
			celula.font = FONT_CABECALHO
		elif celula.value in {"FRETE COMERCIAL X FRETE REAL (R$)", "FRETE COMERCIAL X FRETE REAL (%)"}:
			celula.fill = FILL_FRETE_REAL_CABECALHO
			celula.font = FONT_CABECALHO_PRETA
		else:
			celula.fill = FILL_CABECALHO
			celula.font = FONT_CABECALHO
		celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
		celula.border = BORDER_PADRAO


def cores_range_peso(worksheet) -> None:
	_aplicar_preenchimento_intervalo(worksheet, 7, 9, FILL_RANGE_PESO)


def cores_fretes(worksheet) -> None:
	mapa_cores = {
		"0 - 20": FILL_FAIXA_0_20,
		"20 - 40": FILL_FAIXA_20_40,
		"40 - 60": FILL_FAIXA_40_60,
		"60 - 70": FILL_FAIXA_60_70,
		"70 - 90": FILL_FAIXA_70_90,
		"90 - 120": FILL_FAIXA_90_120,
		"120 - 160": FILL_FAIXA_120_160,
		"160 - 200": FILL_FAIXA_160_200,
		"200 - 250" : FILL_FAIXA_200_1500,
		"250 - 300" : FILL_FAIXA_200_1500,
		"300 - 350" : FILL_FAIXA_200_1500,
		"350 - 400" : FILL_FAIXA_200_1500,
		"400 - 450" : FILL_FAIXA_200_1500,
		"450 - 500" : FILL_FAIXA_200_1500,
	}

	for linha in range(3, worksheet.max_row + 1):
		cell = worksheet.cell(row=linha, column=9)
		if cell.value is None:
			continue

		texto = str(cell.value).strip()
		fill = mapa_cores.get(texto)
		if fill is None:
			continue

		for coluna in range(1, 28):
			celula = worksheet.cell(row=linha, column=coluna)
			if celula.value is None:
				continue
			celula.fill = fill
			celula.border = BORDER_PADRAO
			celula.alignment = ALINHAMENTO_CENTRO


def cores_vendas(worksheet) -> None:
	_aplicar_preenchimento_coluna(worksheet, 2, FILL_VENDA)
	for linha in range(1, worksheet.max_row + 1):
		cell = worksheet.cell(row=linha, column=2)
		if cell.value is None:
			continue
		fonte = copy(cell.font)
		fonte.bold = True
		cell.font = fonte


def cores_coluna_ab(worksheet) -> None:
	ultima_linha = 2
	for linha in range(worksheet.max_row, 1, -1):
		if worksheet.cell(row=linha, column=27).value is not None:
			ultima_linha = linha
			break

	for linha in range(2, ultima_linha + 1):
		cell = worksheet.cell(row=linha, column=29)
		cell.fill = FILL_PRETO
		cell.border = BORDER_PADRAO


def cores_ac_ag(worksheet) -> None:
	for linha in range(3, worksheet.max_row + 1):
		for coluna in range(30, 35):
			cell = worksheet.cell(row=linha, column=coluna)
			if cell.value is None:
				continue
			cell.fill = FILL_AC_AG
	ab3_cell = worksheet.cell(row=3, column=28)
	ab3_cell.fill = FILL_AC_AG


def _converter_para_numero(valor):
	if isinstance(valor, str):
		texto = valor.strip()
		if not texto:
			return valor
		try:
			return float(texto.replace(".", "").replace(",", "."))
		except ValueError:
			return valor
	return valor


def _normalizar_chave(valor) -> str:
	if valor is None:
		return ""
	return str(valor).strip()


def _extrair_base_frete_custo(formula) -> float | None:
	if not isinstance(formula, str) or not formula.startswith("="):
		return _converter_para_numero(formula) if formula is not None else None

	match = re.search(r"^=([+-]?(?:\d+(?:\.\d+)?|\.\d+))\+\$AB\$3$", formula.strip())
	if match is None:
		return None

	try:
		return float(match.group(1))
	except ValueError:
		return None


def _calcular_valor_frete_comercial_atual(worksheet, linha: int, lookup_percentuais: dict[str, float]) -> float | None:
	valor_u = _converter_para_numero(worksheet.cell(row=linha, column=21).value)
	if not isinstance(valor_u, (int, float)):
		return None

	valor_z = lookup_percentuais.get(_normalizar_chave(worksheet.cell(row=linha, column=25).value))
	if not isinstance(valor_z, (int, float)):
		return None

	return ceil((float(valor_u) + (float(valor_u) * float(valor_z))) / 10.0) * 10.0 - 0.01



def formatacao_tipos_texto(worksheet) -> None:
	formato_contabil = '_(* R$ #,##0.00_);_(* (R$ #,##0.00);_(* "-"??_);_(@_)'
	formato_percentual_2 = '0.00%'
	formato_percentual_0 = '0%'

	colunas_contabeis = (2, 14, 16, 19, 20, 21, 22, 23)
	for coluna in colunas_contabeis:
		for linha in range(3, worksheet.max_row + 1):
			cell = worksheet.cell(row=linha, column=coluna)
			if cell.value is None:
				continue
			cell.value = _converter_para_numero(cell.value)
			cell.number_format = formato_contabil
			

	for coluna in (15, 24):
		for linha in range(3, worksheet.max_row + 1):
			cell = worksheet.cell(row=linha, column=coluna)
			if cell.value is None:
				continue
			cell.value = _converter_para_numero(cell.value)
			cell.number_format = formato_percentual_2

	for coluna in (27, 34):
		for linha in range(3, worksheet.max_row + 1):
			cell = worksheet.cell(row=linha, column=coluna)
			if cell.value is None:
				continue
			cell.number_format = formato_percentual_0

	ab3_cell = worksheet.cell(row=3, column=28)
	if ab3_cell.value is not None:
		ab3_cell.value = _converter_para_numero(ab3_cell.value)
		ab3_cell.number_format = formato_contabil


def formatacao_condicional_no(worksheet, worksheet_valores) -> None:
	if worksheet.max_row < 3:
		return

	for coluna in ("N", "O"):
		intervalo = f"{coluna}3:{coluna}{worksheet.max_row}"
		worksheet.conditional_formatting.add(
			intervalo,
			FormulaRule(formula=[f"{coluna}3>0"], fill=FILL_COND_POSITIVO, font=Font(color="006100")),
		)
		worksheet.conditional_formatting.add(
			intervalo,
			FormulaRule(formula=[f"{coluna}3=0"], fill=FILL_COND_ZERO, font=Font(color="9C5700")),
		)
		worksheet.conditional_formatting.add(
			intervalo,
			FormulaRule(formula=[f"{coluna}3<0"], fill=FILL_COND_NEGATIVO, font=Font(color="9C0006")),
		)


def formatacao_letras(worksheet) -> None:
	for linha in worksheet.iter_rows():
		for cell in linha:
			if cell.value is None:
				continue
			fonte = copy(cell.font)
			fonte.name = "Calibri"
			fonte.size = 10
			if cell.column == 15:
				fonte.bold = True
			cell.font = fonte
			cell.border = BORDER_PADRAO


def _ajustar_cabecalhos_e_colunas_destaque(worksheet) -> None:
	for header_row in (1, 2):
		for cell in worksheet[header_row]:
			if cell.value is None:
				continue
			fonte = copy(cell.font)
			fonte.name = "Calibri"
			fonte.size = 10
			cell.font = fonte

	for coluna in (2, 15):
		for linha in range(1, worksheet.max_row + 1):
			cell = worksheet.cell(row=linha, column=coluna)
			if cell.value is None:
				continue
			fonte = copy(cell.font)
			fonte.bold = True
			cell.font = fonte


def optimized_main() -> None:
	start_time = time()
	print("Formatando planilha de visualização (formulas)")
	arquivo_saida = encontrar_xlsx_mais_recente(BASE_DIR)
	workbook = load_workbook(arquivo_saida)
	workbook_valores = load_workbook(arquivo_saida, data_only=True)
	try:
		worksheet = workbook[NOME_ABA]
		worksheet_valores = workbook_valores[NOME_ABA]

		largura_colunas(worksheet)
		cores_cabeçalhos(worksheet)

		max_row = worksheet.max_row
		max_col = max(worksheet.max_column, 34)

		ultima_linha = 2
		for r in range(max_row, 1, -1):
			if worksheet.cell(row=r, column=29).value is not None:
				ultima_linha = r
				break
		for r in range(2, ultima_linha + 1):
			c = worksheet.cell(row=r, column=31)
			c.fill = FILL_PRETO
			c.border = BORDER_PADRAO

		# mapa de cores para fretes (coluna 9)
		mapa_cores = {
			"0 - 20": FILL_FAIXA_0_20,
			"20 - 40": FILL_FAIXA_20_40,
			"40 - 60": FILL_FAIXA_40_60,
			"60 - 70": FILL_FAIXA_60_70,
			"70 - 90": FILL_FAIXA_70_90,
			"90 - 120": FILL_FAIXA_90_120,
			"120 - 160": FILL_FAIXA_120_160,
			"160 - 200": FILL_FAIXA_160_200,
			"200 - 250" : FILL_FAIXA_200_1500,
			"250 - 300" : FILL_FAIXA_200_1500,
			"300 - 350" : FILL_FAIXA_200_1500,
			"350 - 400" : FILL_FAIXA_200_1500,
			"400 - 450" : FILL_FAIXA_200_1500,
			"450 - 500" : FILL_FAIXA_200_1500,
			}

		formato_contabil = '_(* R$ #,##0.00_);_(* (R$ #,##0.00);_(* "-"??_);_(@_)'
		formato_percentual_2 = '0.00%'
		formato_percentual_0 = '0%'

		# fontes pré-criadas para reduzir instâncias
		font_default = Font(name="Calibri", size=10)
		font_bold = Font(name="Calibri", size=10, bold=True)

		# percorrer linhas em um único passe e aplicar formatações
		for row_idx, row in enumerate(worksheet.iter_rows(min_row=3, max_row=max_row, min_col=1, max_col=max_col), start=3):
			# cores range peso (colunas 7..9 => índices 6..8)
			for ci in (6, 7, 8):
				if ci < len(row):
					cell = row[ci]
					if cell.value is not None:
						cell.fill = FILL_RANGE_PESO
						cell.border = BORDER_PADRAO
						cell.alignment = ALINHAMENTO_CENTRO

			# vendas (coluna 2 => índice 1)
			if len(row) > 1:
				cell_venda = row[1]
				if cell_venda.value is not None:
					cell_venda.fill = FILL_VENDA
					cell_venda.font = font_bold
					cell_venda.border = BORDER_PADRAO

			# fretes (coluna 9 => índice 8)
			if len(row) > 8:
				c9 = row[8]
				if c9.value is not None:
					texto = str(c9.value).strip()
					fill = mapa_cores.get(texto)
					if fill is not None:
						for ci in range(0, 29):
							if ci < len(row):
								cel = row[ci]
								if cel.value is not None:
									cel.fill = fill
									cel.border = BORDER_PADRAO
									cel.alignment = ALINHAMENTO_CENTRO

				# AD..AH (colunas 30..34 => índices 29..33) e AB3 (col28 index 27)
				for ci in range(31, 36):
					if ci < len(row):
						cell = row[ci]
						if cell.value is not None:
							cell.fill = FILL_AC_AG
							cell.border = BORDER_PADRAO

				if row_idx == 3 and 29 < len(row):
					ab3 = row[29]
					if ab3.value is not None:
						ab3.value = _converter_para_numero(ab3.value)
						ab3.number_format = formato_contabil
						ab3.fill = FILL_AC_AG

				# formatações contábeis e percentuais
				col_contabeis = (1, 13, 15,16,17, 20, 21, 22, 23, 24)
				for ci in col_contabeis:
					if ci < len(row):
						cell = row[ci]
						if cell.value is not None:
							cell.value = _converter_para_numero(cell.value)
							cell.number_format = formato_contabil
							cell.font = font_default
							cell.border = BORDER_PADRAO
							cell.alignment = ALINHAMENTO_CENTRO

				for ci in (14, 25):  # (15, 24)
					if ci < len(row):
						cell = row[ci]
						if cell.value is not None:
							cell.value = _converter_para_numero(cell.value)
							cell.number_format = formato_percentual_2
							cell.font = font_default
							cell.border = BORDER_PADRAO
							cell.alignment = ALINHAMENTO_CENTRO

				for ci in (28, 35):  # (27, 34)
					if ci < len(row):
						cell = row[ci]
						if cell.value is not None:
							cell.number_format = formato_percentual_0
							cell.font = font_default
							cell.border = BORDER_PADRAO
							cell.alignment = ALINHAMENTO_CENTRO

			# garantir fonte padrão Calibri 10 para células não vazias (sem sobrescrever cores já aplicadas)
			for cell in row:
				if cell.value is None:
					continue
				# se já foi atribuída cor de fonte por formatação condicional, preserve-a
				has_color = False
				try:
					if getattr(cell.font, "color", None) and getattr(cell.font.color, "rgb", None):
						has_color = True
				except Exception:
					has_color = False

				if has_color:
					# preserve color but ensure name/size
					fonte = copy(cell.font)
					fonte.name = "Calibri"
					fonte.size = 10
					cell.font = fonte
				else:
					if cell.column == 15:
						cell.font = Font(name="Calibri", size=10, bold=True)
					else:
						cell.font = font_default
					cell.border = BORDER_PADRAO
					cell.alignment = ALINHAMENTO_CENTRO

		_ajustar_cabecalhos_e_colunas_destaque(worksheet)
		formatacao_condicional_no(worksheet, worksheet_valores)
		cores_vendas(worksheet)

		workbook.save(arquivo_saida)
	finally:
		workbook.close()
		workbook_valores.close()

	elapsed_time = time() - start_time
	minutes = int(elapsed_time // 60)
	seconds = int(elapsed_time % 60)
	print(f"Tempo total de execução: {minutes}m e {seconds}s")


if __name__ == "__main__":
	optimized_main()
