import time
import os
import datetime
from datetime import datetime
from pathlib import Path
import subprocess
import ctypes
import re
import zipfile
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.keys import Keys
from dotenv import load_dotenv
from logging_utils import configure_logging

load_dotenv()
logger = configure_logging(__name__)
print = logger.info


def resolve_env_path(env_name, default_base=None):
	raw_value = os.getenv(env_name)
	if not raw_value:
		raise FileNotFoundError(f"Variável de ambiente não definida: {env_name}")

	raw_value = raw_value.strip().strip('"').strip("'")
	if raw_value.startswith("MAIN_DIR +"):
		base_dir = Path(os.getenv("MAIN_DIR", default_base or ""))
		suffix = raw_value.split("+", 1)[1].strip().strip('"').strip("'")
		return base_dir / suffix.lstrip("\\/")

	path = Path(raw_value)
	if not path.is_absolute() and default_base:
		return Path(default_base) / path
	return path

#CAMINHO PRINCIPAL#
MAIN_DIR = resolve_env_path('MAIN_DIR')
#ATUALIZACAO_VTEX#
vtex_dir = resolve_env_path('VTEX_DIR', MAIN_DIR)
quote_results_vtex = resolve_env_path('QUOTE_RESULTS_VTEX', MAIN_DIR)
erros_vtex = resolve_env_path('ERROS_VTEX', MAIN_DIR)
#CREATE_SHEET#
quote_results_intelipost = resolve_env_path('QUOTE_RESULTS_INTELIPOST', MAIN_DIR)
base_rodrigo = resolve_env_path('BASE_RODRIGO', MAIN_DIR)
#CREATE_SHEET2#
xlsx_dir = resolve_env_path('XLSX_DIR', MAIN_DIR)
#SEPARACAO_INTELIPOST#
intelipost_dir = resolve_env_path('INTELIPOST_DIR', MAIN_DIR)
grupo1 = resolve_env_path('GRUPO1', MAIN_DIR)
grupo2 = resolve_env_path('GRUPO2', MAIN_DIR)


username = os.getenv("WEBMAIL_LOGIN")
password = os.getenv("WEBMAIL_PASSWORD")
link_email = os.getenv("WEBMAIL_URL")

def log_tamanho_arquivo(arquivo):
    tamanho = arquivo.stat().st_size
    print(f"[TAMANHO ORIGEM] {arquivo.name}: {tamanho} bytes")
    return tamanho

def bring_window_to_front(driver):
	if os.name != "nt":
		return

	try:
		pid = driver.capabilities.get("moz:processID")
		if not pid:
			return

		pid = int(pid)
		user32 = ctypes.windll.user32
		windows = []

		@ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
		def enum_windows_callback(hwnd, lparam):
			process_id = ctypes.c_ulong()
			user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
			if process_id.value == pid and user32.IsWindowVisible(hwnd):
				windows.append(hwnd)
				return False
			return True

		user32.EnumWindows(enum_windows_callback, 0)
		if not windows:
			return

		hwnd = windows[0]
		user32.ShowWindow(hwnd, 9)
		user32.ShowWindow(hwnd, 3)
		user32.SetForegroundWindow(hwnd)
		time.sleep(0.5)
	except Exception:
		pass

#CONFIGURAÇÃO TEMPO#
diahoje = datetime.now().date()
data_formatada = diahoje.strftime("%d-%m-%Y")

#CONFIGURAÇÃO ANEXO#
project_dir = MAIN_DIR

#CONFIGURAÇÕES EMAIL#
remetentes = ["daniel.michelon@multimoveis.com","geverton.hemsing@multimoveis.com", "leonardo.pozzobon@multimoveis.com"]
cc = "giovanni.menin@multimoveis.com"
#remetentes = ["giovanni.menin@multimoveis.com", "giovanni.menin@multimoveis.com"]
assunto = f'Atualização de fretes BR - {data_formatada}'
hora = int(datetime.now().strftime("%H"))
if hora < 12:
	periodo = "Bom dia"
elif 12 <= hora < 18:
	periodo = "Boa tarde"
else:	periodo = "Boa noite"
corpo = f"{periodo}, pessoal! Segue a cotação de frete BR.\n\nArquivos enviados:\n- Precificação Intelipost e VTEX\n- Tabela de frete com fórmulas\n- Cotações VTEX\n- Base_rodrigo.\n- Arquivo com erros de cotação.\n\nOs arquivos devem ser extraídos com 7zip.\n\nEmail enviado automaticamente."

def zip_senders():
	senders_dir = MAIN_DIR / "senders"
	senders_dir.mkdir(parents=True, exist_ok=True)

	def zip_folder(source_dir, zip_name):
		source_dir = Path(source_dir)
		if not source_dir.exists():
			raise FileNotFoundError(f"Pasta não encontrada: {source_dir}")

		zip_path = senders_dir / zip_name
		with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_LZMA) as archive:
			for file_path in source_dir.rglob("*"):
				if file_path.is_file():
					archive.write(file_path, arcname=file_path.relative_to(source_dir.parent))
		return zip_path

	def zip_latest_file(source_dir, zip_name, patterns=("*.xlsx", "*.csv")):
		source_dir = Path(source_dir)
		if not source_dir.exists():
			raise FileNotFoundError(f"Pasta não encontrada: {source_dir}")
		if isinstance(patterns, str):
			patterns = (patterns,)
		arquivos = [
			item
			for pattern in patterns
			for item in source_dir.rglob(pattern)
			if item.is_file()
		]
		if not arquivos:
			raise FileNotFoundError(
				f"Nenhum arquivo correspondente a {patterns} encontrado em: {source_dir}"
			)
		latest_file = max(arquivos, key=lambda item: item.stat().st_mtime)
		zip_path = senders_dir / zip_name
		with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_LZMA) as archive:
			archive.write(latest_file, arcname=latest_file.name)
		return zip_path
	
	

	zip_folder(vtex_dir, f"precificacao_vtex.zip")
	zip_folder(intelipost_dir, f"precificacao_intelipost.zip")
	zip_folder(grupo1, f"grupo1.zip")
	zip_folder(grupo2, f"grupo2.zip")
	zip_latest_file(quote_results_vtex, f"cotacoes_vtex.zip")
	zip_latest_file(xlsx_dir, f"tabela_frete(formulas).zip")
	zip_latest_file(base_rodrigo, f"base_rodrigo.zip", patterns=("BASE_RODRIGO_*.csv", "BASE_RODRIGO_*.xlsx"))
	zip_latest_file(erros_vtex, f'erros cotação vtex e intelipost.zip')

	print(f"Arquivos zipados em {senders_dir}")


def setup_driver():
	options = Options()
	#options.add_argument("-headless")
	driver = webdriver.Firefox(options=options)
	try:
		driver.maximize_window()
	except Exception:
		pass
	try:
		driver.fullscreen_window()
	except Exception:
		pass
	bring_window_to_front(driver)
	return driver


def try_find_element(driver, xpaths, timeout=5):
	for xpath in xpaths:
		try:
			element = WebDriverWait(driver, timeout).until(
				EC.visibility_of_element_located((By.XPATH, xpath))
			)
			text = element.text.strip()
			if text:
				return text
		except Exception:
			continue
	return None


def login_webmail(driver, username, password):
	xpath_username = "//input[@name='username']"
	xpath_password = "//input[@name='password']"

	for tentativa in range(1, 4):
		try:
			time.sleep(2)
			username_input = WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.XPATH, xpath_username))
			)
			username_input.clear()
			username_input.send_keys(username)

			password_input = WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.XPATH, xpath_password))
			)
			password_input.clear()
			password_input.send_keys(password)

			time.sleep(0.8)
			username_input.send_keys(u'\ue007')
			
			time.sleep(1.5)
		
			return True
		except WebDriverException as exc:
			if tentativa < 3:
				time.sleep(1.5)
				continue
			print(f"Erro ao logar: {exc}")
		except Exception as exc:
			if tentativa < 3:
				time.sleep(1.5)
				continue
			print(f"Erro ao logar: {exc}")
	return False


def select_anexo(driver, caminho_arquivo=None):
	file_path = Path(caminho_arquivo) if caminho_arquivo else MAIN_DIR / "senders"
	if not file_path.exists():
		raise FileNotFoundError(f"Arquivo ou pasta para anexo não encontrado: {file_path}")
	if file_path.is_dir():
		arquivos = sorted(item for item in file_path.iterdir() if item.is_file())
		if not arquivos:
			raise FileNotFoundError(f"Nenhum arquivo encontrado para anexo em: {file_path}")
		arquivos_para_anexar = arquivos
	else:
		arquivos_para_anexar = [file_path]
	def find_file_input():
		return WebDriverWait(driver, 10).until(
			lambda current_driver: current_driver.find_elements(By.CSS_SELECTOR, "input[type='file']")[0]
			if current_driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
			else None
		)
	try:
		for indice, arquivo in enumerate(arquivos_para_anexar, start=1):
			file_input = find_file_input()
			file_input.send_keys(str(arquivo))
			time.sleep(4)
		return
	except Exception as exc:
		pass
	clipboard_value = str(file_path).replace("'", "''")
	subprocess.run(
		[
			"powershell",
			"-NoProfile",
			"-Command",
			f"Set-Clipboard -Value '{clipboard_value}'",
		],
		check=True,
	)

	key_event = ctypes.windll.user32.keybd_event
	key_event(0x11, 0, 0, 0)
	key_event(0x56, 0, 0, 0)
	key_event(0x56, 0, 2, 0)
	key_event(0x11, 0, 2, 0)
	time.sleep(0.3)
	key_event(0x0D, 0, 0, 0)
	key_event(0x0D, 0, 2, 0)
	time.sleep(2)


def normalize_recipients(recipients):
	if isinstance(recipients, str):
		return [item.strip() for item in re.split(r"[;,\n]+", recipients) if item.strip()]

	return [str(item).strip() for item in recipients if str(item).strip()]


def email_send(driver, remetentes, cc, assunto, corpo):
	xpath_new_email = "//button[@class='_fce_h _fce_f ms-fwt-r ms-fcl-np o365button']"
	xpath_remetente = "//*[@class='_fp_F']//input[@aria-label='Para']"
	xpath_copia = "//*[@class='_fp_F']//input[@aria-label='Cc']"
	xpath_assunto = "//input[@class='_mcp_Q1 _mcp_R1 ms-font-weight-regular ms-bg-color-white-hover ms-font-color-neutralPrimary allowTextSelection textbox ms-font-s ms-fwt-sl ms-fcl-np ms-bcl-nta ms-bcl-nsa-h']"
	xpath_anexo = "//button[@title='Anexar']"
	xpath_corpo = "//*[@aria-label='Corpo da mensagem']"
	xpath_envio = "//button[@title='Enviar']"


	for tentativa in range(1, 4):
		try:
			new_email_button = WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.XPATH, xpath_new_email))
			)
			new_email_button.click()

			time.sleep(2)
			remetente_input = WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.XPATH, xpath_remetente))
			)
			remetente_input.clear()
			lista_remetentes = normalize_recipients(remetentes)
			for indice, remetente in enumerate(lista_remetentes):
				time.sleep(1.5)
				remetente_input.send_keys(remetente)
				time.sleep(1.5)
				remetente_input.send_keys(Keys.ENTER)
				if indice < len(lista_remetentes) - 1:
					time.sleep(1.5)
			time.sleep(0.5)

			copia_input = WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.XPATH, xpath_copia))
			)
			copia_input.clear()
			copia_input.send_keys(cc)
			time.sleep(0.5)

			assunto_input = WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.XPATH, xpath_assunto))
			)
			assunto_input.clear()
			assunto_input.send_keys(assunto)
			time.sleep(0.5)

			corpo_input = WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.XPATH, xpath_corpo))
			)
			corpo_input.click()
			driver.execute_script(
				"""
				const element = arguments[0];
				element.focus();
				const range = document.createRange();
				range.selectNodeContents(element);
				range.collapse(true);
				const selection = window.getSelection();
				selection.removeAllRanges();
				selection.addRange(range);
				""",
				corpo_input,
			)
			corpo_input.send_keys(corpo)
			time.sleep(0.5)
	
			arquivos_anexos = [
				MAIN_DIR / "senders" / "tabela_frete(formulas).zip",
				MAIN_DIR / "senders" / "base_rodrigo.zip",
				MAIN_DIR / "senders" / "precificacao_vtex.zip",
				MAIN_DIR / "senders" / "precificacao_intelipost.zip",
				MAIN_DIR / "senders" / "grupo1.zip",
				MAIN_DIR / "senders" / "grupo2.zip",
				MAIN_DIR / "senders" / "cotacoes_vtex.zip",
				MAIN_DIR / "senders" / "erros cotação vtex e intelipost.zip"
			]

			total_anexos = len(arquivos_anexos)
			for indice, arquivo_anexo in enumerate(arquivos_anexos, start=1):
				anexo_button = WebDriverWait(driver, 10).until(
					EC.element_to_be_clickable((By.XPATH, xpath_anexo))
				)
				anexo_button.click()
				time.sleep(1)
				bring_window_to_front(driver)
				select_anexo(driver, arquivo_anexo)
				print(f"Anexo {indice}/{total_anexos} adicionado: {arquivo_anexo.name}")
				time.sleep(20)

			remetente_input.send_keys(u'\ue007')
			
			time.sleep(3)
			envio_button = WebDriverWait(driver, 10).until(
				EC.element_to_be_clickable((By.XPATH, xpath_envio))
			)
			envio_button.click()

			print(f"Email enviado para {remetentes}")
			return True
		except WebDriverException as exc:
			if tentativa < 3:
				time.sleep(1.5)
				continue
			print(f"Erro ao enviar email para {remetentes}: {exc}")
		except Exception as exc:
			if tentativa < 3:
				time.sleep(1.5)
				continue
			print(f"Erro ao enviar email para {remetentes}: {exc}")
	return False

def main():
	driver = setup_driver()
	try:
		zip_senders()
		time.sleep(3)
		driver.get(link_email)
		login_webmail(driver,username, password)
		email_send(driver, remetentes, cc, assunto, corpo)
	except Exception as exc:
		print(f"Erro ao executar o script: {exc}")
		if driver:
			driver.quit()
	finally:
		driver.quit()

if __name__ == "__main__":
	main()