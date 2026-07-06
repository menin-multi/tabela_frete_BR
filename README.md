# Atualizacao de Fretes BR

Automacao em Python para cotacao, consolidacao, formatacao e envio dos arquivos de frete usados no fluxo BR.

## Como o fluxo funciona

1. `main.py` execucao completa.
2. `api/main.py` executa as cotacoes na Intelipost.
3. `create_sheet/create.py` gera a tabela base sem formulas.
4. `create_sheet2/main.py` executa a versao com formulas e formatacao.
5. `atualizacao_vtex/main.py` cota os prazos VTEX com argumentos adicional de transportadora e gera a saida para tabela VTEX.
6. `separacao_intelipost/atualizar_tabela.py` separa e consolida os arquivos da tabela Intelipost para subir no multi360.
7. `email_sender.py` zipa os anexos e envia o email final.

## Requisitos

- No `.env` alterar o MAIN_DIR para o diretório do main.py.
- Alterar também o WEBMAIL_LOGIN, WEBMAIL_PASSWORD para os seus dados,
- Adicionar WEBMAIL_URL='https://webmail.multimoveis.com/'.
- Configurar a key em `api/env 1` e `atualizacao_vtex/cotacao/env 1`
- confere se tem o firefox no pc, porque vai utilizar o selenium pra fazer o envio do email.
- instala as dependecias.

## Execucao

- Apagar os dados das pastas:
1. `atualizacao_vtex\results\quote_results`
2. `atualizacao_vtex\quote_results\precificacao_br`,
3. `create_sheet\cotacoes`,
4. `\senders`,
5. `\separacao_intelipost\results`.
6. `intelipost_add_value\grupo1\results`,
7. `intelipost_add_value\grupo1\separacao_intelipost`,
8. `intelipost_add_value\grupo2\results`,
9. `intelipost_add_value\grupo2\separacao_intelipost`,

Executar a partir da raiz do projeto:

```bash
python main.py
```

## Dependencias

As dependencias tão em `requirements.txt`.

## Observacoes

- O projeto usa caminhos locais e arquivos Excel como base do fluxo.
- Nao altere os nomes e localizacoes esperadas das planilhas sem revisar as rotinas de leitura.