import key
from datetime import datetime
import requests

def normalize_zip_code(cep) -> str | None:
    """Normaliza o CEP para o formato 00000-000, garantindo 8 dígitos."""
    # Trata None e NaN (pandas)
    if cep is None:
        return None
    try:
        # Converte sempre para string
        cep = str(cep)
    except Exception:
        return None
    # Remove tudo que não for número
    cep = ''.join(filter(str.isdigit, cep))
    if not cep:
        return None
    # Garante exatamente 8 dígitos
    cep = cep.zfill(8)
    # Formata no padrão 00000-000
    return f"{cep[:5]}-{cep[5:]}"

def quote_request(destination_zip_code: str, envio_id: str):
    """
    Realiza a cotação na Intelipost enviando múltiplos produtos
    """
    origin_zip_code = '95706-350'

    base_url, api_token = key.quote_connection()
    url = f"{base_url}/quote_by_product"

    payload = {
        "origin_zip_code": origin_zip_code,
        "destination_zip_code": destination_zip_code,
        "products": [
            {
                "quantity" : 1,
                "weight": 30,
                "volume_type": "BOX",
                "cost_of_goods": 600,
                "width": 10,
                "height": 10,
                "length": 25
            }
        ],
        "additional_information": {
        "delivery_method_ids": [envio_id]},
        "skip_quote_rules": True,
    }

    headers = {"api-key": api_token,"Content-Type": "application/json","Accept": "application/json"}

    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()

        data = response.json()
        quote_id = data.get("content", {}).get("id")
        delivery_options = data.get("content", {}).get("delivery_options", [])

        if not delivery_options:
            return None

        first_option = delivery_options[0]

        return (
            quote_id,
            first_option.get("delivery_estimate_business_days"),
            first_option.get("delivery_method_name")
        )

    except requests.exceptions.RequestException as e:
        print(f"Erro na requisição Intelipost: {e}")
        return None
    