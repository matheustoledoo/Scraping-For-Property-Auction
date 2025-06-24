import time
import logging
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE_URL = (
    "https://www.vivaleiloes.com.br/busca/"
    "#Engine=Start&Pagina={page}&Busca=&Mapa=&ID_Categoria=55&PaginaIndex=3"
)

def init_driver() -> webdriver.Chrome:
    """
    Inicializa o ChromeDriver usando o Chromium e chromedriver instalados via apt.
    """
    # Usar o binário Chromium e o driver do container
    chrome_bin = os.getenv("CHROME_BIN", "/usr/bin/chromium")
    driver_path = os.getenv("CHROMEDRIVER_PATH", "/usr/bin/chromedriver")

    options = Options()
    options.binary_location = chrome_bin
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    service = Service(executable_path=driver_path)
    return webdriver.Chrome(service=service, options=options)


def collect_links(driver: webdriver.Chrome, pages: int) -> list[tuple[str, str]]:
    all_links = []
    current_page = 1
    while True:
        if pages >= 0 and current_page > pages:
            break
        url = BASE_URL.format(page=current_page)
        logger.info(f"Acessando Viva Leilões página {current_page}: {url}")
        driver.get(url)
        time.sleep(4)
        cards = driver.find_elements(By.XPATH,
            '//div[contains(@class,"dg-leiloes-item-col")]')
        if not cards:
            break
        for card in cards:
            try:
                status = card.find_element(By.XPATH,
                    './/span[contains(@class,"BoxBtLoteLabel")]'
                ).text.strip()
            except:
                status = ""
            try:
                link = card.find_element(By.XPATH,
                    './/a[contains(@class,"dg-btn-lote-online")]'
                ).get_attribute("href")
            except:
                continue
            all_links.append((link, status))
        current_page += 1
    return all_links


def process_links(driver: webdriver.Chrome, link_status: list[tuple[str, str]]) -> list[dict]:
    results = []
    for idx, (link, status) in enumerate(link_status, start=1):
        logger.info(f"VivaLeilões {idx}/{len(link_status)}: {link}")
        driver.execute_script("window.open(arguments[0]);", link)
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(3)
        data = {"link": link, "status": status}
        def get_text(xpath: str):
            try:
                return driver.find_element(By.XPATH, xpath).text.strip()
            except:
                return None
        def get_href(xpath: str):
            elems = driver.find_elements(By.XPATH, xpath)
            return elems[0].get_attribute("href") if elems else None
        data.update({
            "titulo_leilao": get_text('//div[contains(@class,"dg-titulo")]'),
            "tipo_leilao": "Judicial",
            "numero_processo": get_text('/html/body/main/div/section/div/div/div/div[1]/div[1]/a'),
            "valor_imovel": get_text('//span[contains(@class,"ValorMinimoLanceSegundaPraca")]')
                             or get_text('//span[contains(@class,"ValorMinimoLancePrimeiraPraca")]'),
            "edital_leilao": get_href('/html/body/section[2]/div/div[2]/div/div/div/div[1]/ul/li[6]/a'),
            "laudo_avaliacao": get_href('/html/body/section[2]/div/div[2]/div/div/div/div[1]/ul/li[1]/a'),
            "matricula": get_href('/html/body/section[2]/div/div[2]/div/div/div/div[1]/ul/li[3]/a'),
            "descricao_lote": get_text('//div[contains(@class,"dg-lote-descricao-txt")]')
        })
        results.append(data)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(1)
    return results


def run(pages: int) -> pd.DataFrame:
    driver = init_driver()
    try:
        links = collect_links(driver, pages)
        raw = process_links(driver, links)
    finally:
        driver.quit()
    return pd.DataFrame(raw)
