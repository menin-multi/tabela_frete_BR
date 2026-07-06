import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import os
import sys
import shutil
from pathlib import Path
import time
from datetime import datetime
from time import perf_counter
from dotenv import load_dotenv
ROOT_DIR = Path(__file__).resolve().parent.parent
API_DIR = ROOT_DIR / "api"
for path in (ROOT_DIR, API_DIR):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))
main_dir=os.path.dirname(os.path.abspath(__file__))

HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Calibri", size=11, bold=False, color="000000")
HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E78")
BODY_FILL = PatternFill(fill_type=None)
THIN_SIDE = Side(style="thin", color="000000")
STANDARD_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
STANDARD_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=False)


def aplicar_formatacao_fixa(ws, total_linhas, total_colunas):
    for linha in range(1, total_linhas + 1):
        for coluna in range(1, total_colunas + 1):
            cell = ws.cell(row=linha, column=coluna)
            # Aplicar formatação APENAS se a célula tem valor
            if cell.value is not None:
                cell.font = HEADER_FONT if linha == 1 else BODY_FONT
                cell.fill = HEADER_FILL if linha == 1 else BODY_FILL
                cell.border = STANDARD_BORDER
                cell.alignment = STANDARD_ALIGNMENT
                cell.number_format = "General"


def filtrar_e_copiar_dados(
    caminho_arquivo: str,
    nome_worksheet: str,
    coluna_filtro: str,
    valor_filtro: any,
    coluna_origem: str,
    caminho_saida: str = None
): 
    wb_origem = None
    
    try:        
        if not os.path.exists(caminho_arquivo):
            raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")
        
        df = pd.read_excel(caminho_arquivo, sheet_name=0)
        
        if coluna_filtro not in df.columns or coluna_origem not in df.columns:
            raise ValueError(f"Colunas não encontradas")
        
        df_filtrado = df[df[coluna_filtro].astype(str).str.strip() == str(valor_filtro).strip()]
        linhas_count = len(df_filtrado)
    
        idx_coluna = df_filtrado.columns.get_loc(coluna_origem)
        df_resultado = df_filtrado.iloc[:, idx_coluna:]
        
        if not caminho_saida:
            return linhas_count
        
        df_resultado.to_excel(caminho_saida, sheet_name=nome_worksheet, index=False)
        
        wb_novo = load_workbook(caminho_saida)
        ws_novo = wb_novo[nome_worksheet]
        aplicar_formatacao_fixa(ws_novo, ws_novo.max_row, ws_novo.max_column)
        
        wb_novo.save(caminho_saida)
        wb_novo.close()
        
        
    except FileNotFoundError as e:
        print(f"Erro: {e}")
        if wb_origem:
            wb_origem.close()
        return 0
    except ValueError as e:
        print(f"Erro: {e}")
        if wb_origem:
            wb_origem.close()
        return 0
    except Exception as e:
        print(f"Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
        if wb_origem:
            wb_origem.close()
        return 0


def separar_UFs(
    caminho_arquivo: str,
    nome_worksheet: str,
    coluna_filtro: str,
    coluna_origem: str,
    pasta_saida: str = None,
):
    try:

        if not os.path.exists(caminho_arquivo):
            raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

        df = pd.read_excel(caminho_arquivo, sheet_name=0)

        if coluna_filtro not in df.columns or coluna_origem not in df.columns:
            raise ValueError("Colunas não encontradas")

        valores = (
            df[coluna_filtro]
            .dropna()
            .astype(str)
            .str.strip()
        )
        valores_unicos = sorted(set(valor for valor in valores if valor))

        if not pasta_saida:
            pasta_saida = Path(caminho_arquivo).parent
        else:
            pasta_saida = Path(pasta_saida)

        for valor in valores_unicos:
            df_filtrado = df[df[coluna_filtro].astype(str).str.strip() == valor]
            if df_filtrado.empty:
                continue

            df_resultado = df_filtrado.iloc[:, df_filtrado.columns.get_loc(coluna_origem):]
            arquivo_saida = pasta_saida / f"{valor}.xlsx"

            df_resultado.to_excel(arquivo_saida, sheet_name=nome_worksheet_saida, index=False)

            wb_novo = load_workbook(arquivo_saida)
            ws_novo = wb_novo[nome_worksheet_saida]
            aplicar_formatacao_fixa(ws_novo, ws_novo.max_row, ws_novo.max_column)
            wb_novo.save(arquivo_saida)
            wb_novo.close()

        return len(valores_unicos)

    except FileNotFoundError as e:
        print(f"Erro: {e}")
        return 0
    except ValueError as e:
        print(f"Erro: {e}")
        return 0
    except Exception as e:
        print(f"Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
        return 0


def adicionar_TDA(
    caminho_arquivo_complementar: str,
    coluna_filtro: str,
    colunas_copia: list,
    pasta_saida: str = None,
):
    try:
        if not os.path.exists(caminho_arquivo_complementar):
            raise FileNotFoundError(f"Arquivo complementar não encontrado: {caminho_arquivo_complementar}")

        df = pd.read_excel(caminho_arquivo_complementar, sheet_name=0)

        if coluna_filtro not in df.columns:
            raise ValueError(f"Coluna '{coluna_filtro}' não encontrada no arquivo complementar")

        for col in colunas_copia:
            if col not in df.columns:
                raise ValueError(f"Coluna '{col}' não encontrada no arquivo complementar")

        valores = (
            df[coluna_filtro]
            .dropna()
            .astype(str)
            .str.strip()
        )
        valores_unicos = sorted(set(valor for valor in valores if valor))

        if not pasta_saida:
            pasta_saida = Path(caminho_arquivo_complementar).parent
        else:
            pasta_saida = Path(pasta_saida)


        for arquivo_destino in sorted(pasta_saida.glob("*.xlsx")):
            nome_sem_ext = arquivo_destino.stem.strip()

            if nome_sem_ext.upper().startswith("TDAS "):
                continue

            valor = nome_sem_ext.upper()
            df_filtrado = df[df[coluna_filtro].astype(str).str.strip() == valor]
            df_copia = df_filtrado[colunas_copia].copy()

            wb_destino = load_workbook(arquivo_destino)

            if "Planilha2" in wb_destino.sheetnames:
                ws_nova = wb_destino["Planilha2"]
                ws_nova.delete_rows(1, ws_nova.max_row)
            else:
                ws_nova = wb_destino.create_sheet("Planilha2")

            for col_idx, col_nome in enumerate(colunas_copia, 1):
                cell = ws_nova.cell(row=1, column=col_idx)
                cell.value = col_nome
                cell.font = HEADER_FONT
                cell.fill = HEADER_FILL
                cell.border = STANDARD_BORDER
                cell.alignment = STANDARD_ALIGNMENT


            # Escrever linhas — se não houver dados, este loop não executa e apenas o cabeçalho permanece
            for row_idx, row_data in enumerate(df_copia.values, 2):
                for col_idx, valor_cell in enumerate(row_data, 1):
                    cell = ws_nova.cell(row=row_idx, column=col_idx)
                    cell.value = valor_cell
                    cell.font = BODY_FONT
                    cell.border = STANDARD_BORDER
                    cell.alignment = STANDARD_ALIGNMENT

            for col_idx in range(1, len(colunas_copia) + 1):
                ws_nova.column_dimensions[get_column_letter(col_idx)].width = 15

            wb_destino.save(arquivo_destino)
            wb_destino.close()

        print(f"Dados complementares adicionados com sucesso em {len(valores_unicos)} arquivo(s)")
        return len(valores_unicos)

    except FileNotFoundError as e:
        print(f"Erro: {e}")
        return 0
    except ValueError as e:
        print(f"Erro: {e}")
        return 0
    except Exception as e:
        print(f"Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
        return 0


def criar_tabelas_tda(
    caminho_arquivo_complementar: str,
    coluna_filtro: str,
    colunas_copia: list,
    pasta_saida: str = None,
):
    try:
        if not os.path.exists(caminho_arquivo_complementar):
            raise FileNotFoundError(f"Arquivo complementar não encontrado: {caminho_arquivo_complementar}")

        df = pd.read_excel(caminho_arquivo_complementar, sheet_name=0)

        if coluna_filtro not in df.columns:
            raise ValueError(f"Coluna '{coluna_filtro}' não encontrada no arquivo complementar")

        for col in colunas_copia:
            if col not in df.columns:
                raise ValueError(f"Coluna '{col}' não encontrada no arquivo complementar")

        valores = (
            df[coluna_filtro]
            .dropna()
            .astype(str)
            .str.strip()
        )
        valores_unicos = sorted(set(valor for valor in valores if valor))

        if not pasta_saida:
            pasta_saida = Path(caminho_arquivo_complementar).parent
        else:
            pasta_saida = Path(pasta_saida)

        for valor in valores_unicos:
            df_filtrado = df[df[coluna_filtro].astype(str).str.strip() == valor]
            if df_filtrado.empty:
                continue

            df_copia = df_filtrado[colunas_copia].copy()
            arquivo_saida = pasta_saida / f"TDAS {valor}.xlsx"

            df_copia.to_excel(arquivo_saida, sheet_name="Planilha1", index=False)

            wb_novo = load_workbook(arquivo_saida)
            ws_novo = wb_novo["Planilha1"]
            aplicar_formatacao_fixa(ws_novo, ws_novo.max_row, ws_novo.max_column)
            wb_novo.save(arquivo_saida)
            wb_novo.close()

        print(f"Tabelas de TDA criadas com sucesso em {len(valores_unicos)} arquivo(s)")
        return len(valores_unicos)

    except FileNotFoundError as e:
        print(f"Erro: {e}")
        return 0
    except ValueError as e:
        print(f"Erro: {e}")
        return 0
    except Exception as e:
        print(f"Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
        return 0


def separar_arquivos_por_regiao(pasta_resultados: str):
    try:
        pasta = Path(pasta_resultados)
        if not pasta.exists():
            raise FileNotFoundError(f"Pasta de resultados nao encontrada: {pasta_resultados}")

        regiao_por_uf = {
            "AC": "Norte", "AP": "Norte", "AM": "Norte", "PA": "Norte", "RO": "Norte", "RR": "Norte", "TO": "Norte",
            "AL": "Nordeste", "BA": "Nordeste", "CE": "Nordeste", "MA": "Nordeste", "PB": "Nordeste", "PE": "Nordeste", "PI": "Nordeste", "RN": "Nordeste", "SE": "Nordeste",
            "DF": "Centro-Oeste", "GO": "Centro-Oeste", "MT": "Centro-Oeste", "MS": "Centro-Oeste",
            "ES": "Sudeste", "MG": "Sudeste", "RJ": "Sudeste", "SP": "Sudeste",
            "PR": "Sul", "RS": "Sul", "SC": "Sul",
        }

        total_movidos = 0

        for arquivo in pasta.glob("*.xlsx"):
            nome_sem_ext = arquivo.stem.strip()

            if nome_sem_ext.upper().startswith("TDAS "):
                uf = nome_sem_ext[5:].strip().upper()
            else:
                uf = nome_sem_ext.upper()

            regiao = regiao_por_uf.get(uf)
            if not regiao:
                print(f"Aviso: UF nao reconhecida no arquivo {arquivo.name}, pulando...")
                continue

            pasta_regiao = pasta / regiao
            pasta_regiao.mkdir(parents=True, exist_ok=True)

            destino = pasta_regiao / arquivo.name
            if destino.exists():
                destino.unlink()

            shutil.move(str(arquivo), str(destino))
            total_movidos += 1

        print(f"Arquivos separados por regiao com sucesso! {total_movidos} arquivo(s) movido(s)")
        return total_movidos

    except FileNotFoundError as e:
        print(f"Erro: {e}")
        return 0
    except Exception as e:
        print(f"Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
        return 0


if __name__ == "__main__":
    main_dir = Path(main_dir)
    comp_dir = Path(main_dir).parent.parent
    # selecionar o arquivo TABELA_INTELIPOST mais recentemente modificado em todo o histórico
    results_dir = main_dir / 'results'
    candidatos = [p for p in results_dir.rglob("TABELA_INTELIPOST_*.xlsx") if p.is_file()] if results_dir.exists() else []
    if candidatos:
        caminho_arquivo = max(candidatos, key=lambda p: p.stat().st_mtime)
    else:
        caminho_arquivo = results_dir / 'TABELA_INTELIPOST.xlsx'
    caminho_arquivo_complementar = comp_dir / "atualizacao_de_fretes_br" / 'separacao_intelipost' / 'TDAS_consolidado.xlsx'
    nome_worksheet = "Sheet"
    nome_worksheet_saida = "Planilha1"
    coluna_filtro = "UF"
    coluna_origem = "CEP DE"
    colunas_complementares = ["CEP DE", "CEP ATÉ"]
    
    pasta_resultados = main_dir / 'separacao_intelipost' / "results_grupo1"
    pasta_resultados.mkdir(parents=True, exist_ok=True)


    print("Etapa 1: Criando arquivos por UF...")
    arquivos = separar_UFs(
        caminho_arquivo=caminho_arquivo,
        nome_worksheet=nome_worksheet,
        coluna_filtro=coluna_filtro,
        coluna_origem=coluna_origem,
        pasta_saida=pasta_resultados
    )
    print(f"Separação de CEPs finalizada! {arquivos} arquivo(s) criado(s)")
    
    print("\nEtapa 2: Adicionando dados complementares...")
    arquivos_complementares = adicionar_TDA(
        caminho_arquivo_complementar=caminho_arquivo_complementar,
        coluna_filtro=coluna_filtro,
        colunas_copia=colunas_complementares,
        pasta_saida=pasta_resultados
    )

    print("\nEtapa 3: Criando tabelas de TDA...")
    tabelas_tda = criar_tabelas_tda(
        caminho_arquivo_complementar=caminho_arquivo_complementar,
        coluna_filtro=coluna_filtro,
        colunas_copia=["CEP DE", "CEP ATÉ", "R$"],
        pasta_saida=pasta_resultados
    )

    print("\nEtapa 4: Separando arquivos por regiao...")
    arquivos_por_regiao = separar_arquivos_por_regiao(
        pasta_resultados=pasta_resultados
    )
