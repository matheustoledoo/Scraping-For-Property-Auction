import time
import logging
from typing import List, Tuple

import pandas as pd
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE_URL = (
    "https://www.megaleiloes.com.br/imoveis?tov=igbr&valor_max=5000000&"
    "tipo%5B0%5D=1&pagina={page}"
)


def init_driver() -> webdriver.Chrome:
    """
    Instala e inicializa o ChromeDriver em modo headless.
    """
    chromedriver_autoinstaller.install()
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=options)
    return driver


def collect_links(driver: webdriver.Chrome, pages: int) -> List[Tuple[str, str]]:
    """
    Coleta todos os links de imóveis e seus status. Se pages < 0, varre até não encontrar mais cards.
    Retorna lista de tuplas (link, status).
    """
    all_links = []
    current_page = 1

    while True:
        if pages >= 0 and current_page > pages:
            break

        url = BASE_URL.format(page=current_page)
        logger.info(f"Acessando página {current_page}: {url}")
        driver.get(url)
        time.sleep(3)

        cards = driver.find_elements(By.XPATH,
            '//div[contains(@class, "col-sm-6 col-md-4 col-lg-3")]'
        )
        logger.info(f"Encontrados {len(cards)} imóveis na página {current_page}.")
        if not cards:
            break

        for card in cards:
            try:
                status = card.find_element(
                    By.XPATH, './/div[contains(@class, "card-status")]'
                ).text.strip()
            except Exception:
                status = ""
            try:
                link = card.find_element(
                    By.XPATH, './/a[contains(@class, "card-title")]'
                ).get_attribute('href')
            except Exception:
                continue
            all_links.append((link, status))

        current_page += 1

    return all_links


def process_links(driver: webdriver.Chrome, link_status: List[Tuple[str, str]]) -> List[dict]:
    """
    Visita cada link de imóvel, extrai informações e retorna lista de dicionários.
    """
    results = []

    for idx, (link, status) in enumerate(link_status, start=1):
        logger.info(f"Processando {idx}/{len(link_status)}: {link}")
        driver.execute_script("window.open(arguments[0]);", link)
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(2)

        data = {"link": link, "status": status}

        # Helpers para extrair campos
        def get_text(xpath: str) -> str:
            try:
                return driver.find_element(By.XPATH, xpath).text.strip()
            except Exception:
                return None

        # pega a string completa, ex:
        raw_loc = get_text('/html/body/div[3]/div[3]/div[2]/div[2]/div/div/div[1]/div[1]/div[2]')
        if raw_loc:
            # separa em partes pelo caractere “,”
            parts = [p.strip() for p in raw_loc.split(',')]
            if len(parts) >= 2:
                # o último elemento é sempre o estado
                estado = parts[-1]
                # o penúltimo elemento é a cidade
                cidade = parts[-2]
            else:
                cidade = None
                estado = None
        else:
            cidade = None
            estado = None

        data.update({
            "titulo_leilao": get_text('//h1[contains(@class, "section-header")]'),
            "cidade": cidade,
            "estado": estado,
            "tipo_leilao": get_text('//div[contains(@class, "batch-type")]'),
            "numero_processo": get_text(
                '/html/body/div[3]/div[3]/div[2]/div[2]/div/div/div[2]/div[1]/div[2]/a'
            ),
            "valor_imovel": get_text('//span[contains(@class, "card-instance-value")]') or get_text('/html/body/div[3]/div[3]/div[1]/div[2]/div/div[2]'),
            "edital_leilao": driver.find_element(
                By.XPATH, '/html/body/div[3]/div[3]/div[3]/div[3]/div[2]/a[2]'
            ).get_attribute('href') if driver.find_elements(
                By.XPATH, '/html/body/div[3]/div[3]/div[3]/div[3]/div[2]/a[2]'
            ) else None,
            "laudo_avaliacao": driver.find_element(
                By.XPATH, '/html/body/div[3]/div[3]/div[3]/div[3]/div[2]/a[3]'
            ).get_attribute('href') if driver.find_elements(
                By.XPATH, '/html/body/div[3]/div[3]/div[3]/div[3]/div[2]/a[3]'
            ) else None,
            "matricula": driver.find_element(
                By.XPATH, '/html/body/div[3]/div[3]/div[3]/div[3]/div[2]/a[4]'
            ).get_attribute('href') if driver.find_elements(
                By.XPATH, '/html/body/div[3]/div[3]/div[3]/div[3]/div[2]/a[4]'
            ) else None,
            "descricao_lote": get_text('//div[contains(@class, "description")]')
        })

        results.append(data)

        # Fecha aba e retorna
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(1)

    return results


def run(pages: int) -> pd.DataFrame:
    """
    Executa todo o fluxo de scraping e retorna um DataFrame com os dados.
    pages: número de páginas a raspar; -1 para todas.
    """
    driver = init_driver()
    try:
        links = collect_links(driver, pages)
        raw_data = process_links(driver, links)
    finally:
        driver.quit()

    df = pd.DataFrame(raw_data)
    return df
