from pathlib import Path
from os import chdir
from datetime import datetime
import pandas as pd
import functions
import time


# Ajusta diretório raiz
root_directory = Path(__file__).parent.parent.absolute()
chdir(str(root_directory))

INPUT_FILE = root_directory / 'cotacao' / "cep_cotacao.xlsx"
PRIORITY_FILE = root_directory / 'cotacao' / "prioridade_estados.xlsx"

# configuração de data
dia_mes = datetime.now().strftime("%d-%m")

# Base do nome (vai gerar um por UF)
OUTPUT_DIR = root_directory.parent / "atualizacao_vtex" / "results" / "quote_results" / dia_mes
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
def process_orders():
    # --- Lê dados principais
    df_ceps = pd.read_excel(INPUT_FILE)

    # --- Lê prioridade de estados (ordem do arquivo)
    df_priority = pd.read_excel(PRIORITY_FILE)
    if "Sigla" not in df_priority.columns:
        raise ValueError("Arquivo precisa ter a coluna 'Sigla'.")

    # Lista de UFs na ordem do arquivo (remove vazios/NaN e normaliza)
    uf_priority = (
        df_priority["Sigla"]
        .dropna()
        .astype(str)
        .str.strip()
        .str.upper()
        .tolist()
    )

    # --- Normaliza CEP e UF no df_ceps
    df_ceps["CEP"] = df_ceps["CEP"].apply(functions.normalize_zip_code)
    df_ceps = df_ceps[df_ceps["CEP"].notna() & (df_ceps["CEP"].str.len() == 9)].copy()
    df_ceps["UF"] = df_ceps["UF"].astype(str).str.strip().str.upper()

    # --- Verifica coluna de transportadora / method_id
    transport_column = None
    for candidate in ["ID", "id", "Id"]:
        if candidate in df_ceps.columns:
            transport_column = candidate
            break

    if transport_column is None:
        raise ValueError("Arquivo precisa ter a coluna 'id'.")

    # --- Processa estado por estado, na ordem definida
    for uf in uf_priority:
        df_ceps_uf = df_ceps[df_ceps["UF"] == uf].copy()

        if df_ceps_uf.empty:
            print(f"{datetime.now()} - UF={uf} sem CEPs no arquivo. Pulando.")
            continue

        results = []
        output_file = OUTPUT_DIR / f"resultado_cotacoes_vtex_{uf}.xlsx"

        print(f"{datetime.now()} - Iniciando cotações UF={uf} | CEPs={len(df_ceps_uf)}")

        for _, row_cep in df_ceps_uf.iterrows():
            state = uf
            region = str(row_cep.get("REGIÃO", "")).strip()
            destination_zip_code = row_cep["CEP"]
            method_id_raw = row_cep.get(transport_column)

            try:
                if pd.isna(method_id_raw):
                    raise ValueError("transportadora ausente")

                method_id = int(method_id_raw)

                print(
                    f"{datetime.now()} - UF={state} | CEP={destination_zip_code} | "
                    f"method_id={method_id}"
                )

                response = functions.quote_request(
                    destination_zip_code=destination_zip_code,
                    envio_id=method_id
                )

                if response:
                    (quote_id, delivery_estimate_business_days, delivery_method_name) = response

                    results.append({
                        "uf": state,
                        "regiao": region,
                        "cep": destination_zip_code,
                        "transportadora id": method_id,
                        "id_cotacao": quote_id,
                        "transportador": delivery_method_name,
                        "prazo_entrega": delivery_estimate_business_days,
                    })
                else:
                    results.append({
                        "uf": state,
                        "regiao": region,
                        "cep": destination_zip_code,
                        "transportadora id": method_id,
                        "erro": "Nenhuma opção de entrega"
                    })

            except Exception as e:
                print(f"{datetime.now()} - Erro UF={state} CEP={destination_zip_code} transportadora={method_id_raw}: {e}")
                results.append({
                    "uf": state,
                    "regiao": region,
                    "cep": destination_zip_code,
                    "transportadora": method_id_raw,
                    "erro": str(e)
                })

        pd.DataFrame(results).to_excel(output_file, index=False)
        print(f"{datetime.now()} - Finalizado UF={uf}: {output_file}")

    print(f"{datetime.now()} - Finalizado geral (por UF).")


if __name__ == "__main__":
    process_orders()