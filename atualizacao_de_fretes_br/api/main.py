from pathlib import Path
from os import chdir
from datetime import datetime
import pandas as pd
import functions
import time


# Ajusta diretório raiz
root_directory = Path(__file__).parent.absolute()
chdir(str(root_directory))

INPUT_FILE = "cep_cotacao_comercial.xlsx"
PRIORITY_FILE = "prioridade_estados.xlsx"

# configuração de data
dia_mes = datetime.now().strftime("%d-%m")

# Base do nome (vai gerar um por UF)
OUTPUT_DIR = root_directory.parent / "create_sheet" / "cotacoes" / dia_mes
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def process_orders():
    # --- Lê dados principais
    df_ceps = pd.read_excel(INPUT_FILE, sheet_name="cep's")
    df_products = pd.read_excel(INPUT_FILE, sheet_name="produtos")

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

    # --- Normaliza SKU (opcional, mas ajuda)
    if "SKU" in df_products.columns:
        df_products["SKU"] = df_products["SKU"].astype(str).str.strip()

    # --- Processa estado por estado, na ordem definida
    for uf in uf_priority:
        df_ceps_uf = df_ceps[df_ceps["UF"] == uf].copy()

        if df_ceps_uf.empty:
            print(f"{datetime.now()} - UF={uf} sem CEPs no arquivo. Pulando.")
            continue

        results = []
        output_file = OUTPUT_DIR / f"resultado_cotacoes_{uf}.xlsx"

        print(f"{datetime.now()} - Iniciando cotações UF={uf} | CEPs={len(df_ceps_uf)}")

        for _, row_cep in df_ceps_uf.iterrows():
            state = uf
            region = str(row_cep.get("REGIÃO", "")).strip()
            destination_zip_code = row_cep["CEP"]

            for _, row_prod in df_products.iterrows():
                try:
                    sku = str(row_prod.get("SKU", "")).strip()
                    weight_range = str(row_prod.get("RANGE FRETE (KG)", "")).strip()

                    weight_raw = row_prod.get("PESO PRODUTO")
                    value_raw = row_prod.get("VALOR PRODUTO")

                    weight = float(weight_raw) if pd.notna(weight_raw) else 0.0
                    cost_of_goods = float(value_raw) if pd.notna(value_raw) else 0.0

                    print(
                        f"{datetime.now()} - UF={state} | CEP={destination_zip_code} | "
                        f"SKU={sku} | weight={weight} | value={cost_of_goods}"
                    )

                    response = functions.quote_request(
                        destination_zip_code=destination_zip_code,
                        products=sku,
                        weight=weight,
                        cost_of_goods=cost_of_goods
                    )

                    if response:
                        (quote_id, delivery_method_id, delivery_estimate_business_days,
                         final_shipping_cost, provider_shipping_cost, delivery_method_name) = response

                        results.append({
                            "uf": state,
                            "regiao": region,
                            "cep": destination_zip_code,
                            "sku": sku,
                            "faixa_peso": weight_range,
                            "peso_produto": weight,
                            "valor_produto": cost_of_goods,
                            "id_cotacao": quote_id,
                            "id_transportador": delivery_method_id,
                            "transportador": delivery_method_name,
                            "prazo_entrega": delivery_estimate_business_days,
                            "valor_comercial": final_shipping_cost,
                            "valor_transportador": provider_shipping_cost
                        })
                    else:
                        results.append({
                            "uf": state,
                            "regiao": region,
                            "cep": destination_zip_code,
                            "sku": sku,
                            "faixa_peso": weight_range,
                            "peso_produto": weight,
                            "valor_produto": cost_of_goods,
                            "erro": "Nenhuma opção de entrega"
                        })

                except Exception as e:
                    print(f"{datetime.now()} - Erro UF={state} CEP={destination_zip_code} SKU={row_prod.get('SKU')}: {e}")
                    results.append({
                        "uf": state,
                        "regiao": region,
                        "cep": destination_zip_code,
                        "sku": row_prod.get("SKU"),
                        "faixa_peso": str(row_prod.get("RANGE FRETE (KG)", "")).strip(),
                        "peso_produto": row_prod.get("PESO PRODUTO"),
                        "valor_produto": row_prod.get("VALOR PRODUTO"),
                        "erro": str(e)
                    })

        pd.DataFrame(results).to_excel(output_file, index=False)
        print(f"{datetime.now()} - Finalizado UF={uf}: {output_file}")

    print(f"{datetime.now()} - Finalizado geral (por UF).")


if __name__ == "__main__":
    process_orders()