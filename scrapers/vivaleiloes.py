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
    "https://www.vivaleiloes.com.br/busca/"
    "?Engine=Start&Pagina={page}&Busca=&Mapa=&ID_Categoria=55&PaginaIndex=3"
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
    return webdriver.Chrome(options=options)


def collect_links(driver: webdriver.Chrome, pages: int) -> List[Tuple[str, str]]:
    """
    Coleta todos os links de lotes e seus status. Se pages < 0, varre até não encontrar mais.
    Retorna lista de tuplas (link, status).
    """
    all_links: List[Tuple[str, str]] = []
    current_page = 1

    while True:
        if pages >= 0 and current_page > pages:
            break

        url = BASE_URL.format(page=current_page)
        logger.info(f"Acessando VivaLeilões página {current_page}: {url}")
        driver.get(url)
        time.sleep(3)

        cards = driver.find_elements(By.XPATH, '//div[contains(@class,"dg-leiloes-item-col")]')
        logger.info(f"Encontrados {len(cards)} lotes na página {current_page}.")
        if not cards:
            break

        for card in cards:
            try:
                status = card.find_element(
                    By.XPATH, './/span[contains(@class,"BoxBtLoteLabel")]'
                ).text.strip()
            except Exception:
                status = ""
            try:
                link = card.find_element(
                    By.XPATH, './/a[contains(@class,"dg-btn-lote-online")]'
                ).get_attribute("href")
            except Exception:
                continue
            all_links.append((link, status))

        current_page += 1

    return all_links


def process_links(driver: webdriver.Chrome, link_status: List[Tuple[str, str]]) -> List[dict]:
    """
    Visita cada link de lote, extrai informações e retorna lista de dicionários.
    """
    results: List[dict] = []

    for idx, (link, status) in enumerate(link_status, start=1):
        logger.info(f"Processando {idx}/{len(link_status)}: {link}")
        driver.execute_script("window.open(arguments[0]);", link)
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(2)

        def get_text(xpath: str) -> str | None:
            try:
                return driver.find_element(By.XPATH, xpath).text.strip()
            except Exception:
                return None

        def get_href(xpath: str) -> str | None:
            elems = driver.find_elements(By.XPATH, xpath)
            return elems[0].get_attribute("href") if elems else None

        # Extrai raw de cidade/estado e separa
        raw_loc = get_text('/html/body/main/div/section/div/div/div/div[1]/div[2]')
        if raw_loc:
            parts = raw_loc.split('/', 1)
            cidade = parts[0].strip()
            estado = parts[1].strip()
        else:
            cidade = None
            estado = None

        data = {
            "link": link,
            "status": status,
            "titulo_leilao": get_text('//div[contains(@class,"dg-titulo")]'),
            "cidade": cidade,
            "estado": estado,
            "tipo_leilao": "Judicial",
            "numero_processo": get_text('/html/body/main/div/section/div/div/div/div[1]/div[1]/a'),
            "valor_imovel": (
                get_text('//span[contains(@class,"ValorMinimoLanceSegundaPraca")]')
                or get_text('//span[contains(@class,"ValorMinimoLancePrimeiraPraca")]')
            ),
            "edital_leilao": get_href('/html/body/section[2]/div/div[2]/div/div/div/div[1]/ul/li[6]/a'),
            "laudo_avaliacao": get_href('/html/body/section[2]/div/div[2]/div/div/div/div[1]/ul/li[1]/a'),
            "matricula": get_href('/html/body/section[2]/div/div[2]/div/div/div/div[1]/ul/li[3]/a'),
            "descricao_lote": get_text('//div[contains(@class,"dg-lote-descricao-txt")]'),
        }

        results.append(data)
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

    return pd.DataFrame(raw_data)
