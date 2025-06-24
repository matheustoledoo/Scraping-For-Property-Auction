import time
import logging
import os
import pandas as pd
# chromedriver_autoinstaller removed to use shipped driver
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE_URL = (
    "https://www.vivaleiloes.com.br/busca/"
    "?Engine=Start&Pagina={page}&Busca=&Mapa=&ID_Categoria=55&PaginaIndex=3"
)

def init_driver() -> webdriver.Chrome:
    """
    Usa o Chromium e ChromeDriver instalados no container, com anti-detect.
    """
    chrome_bin = os.environ.get("CHROME_BIN", "/usr/bin/chromium")
    driver_path = os.environ.get("CHROMEDRIVER_PATH", "/usr/bin/chromedriver")

    options = Options()
    options.binary_location = chrome_bin
    # cabeçalho comum de desktop para evitar bot-block
    options.add_argument(
        "--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/137.0.7151.119 Safari/537.36"
    )
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    service = Service(executable_path=driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    # remove flag webdriver
    driver.execute_cdp_cmd(
        'Page.addScriptToEvaluateOnNewDocument',
        {'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined});'}
    )
    return driver


def collect_links(driver: webdriver.Chrome, pages: int) -> list[tuple[str, str]]:
    all_links = []
    for current_page in range(1, pages + 1 if pages >= 0 else 50):
        url = BASE_URL.format(page=current_page)
        logger.info(f"Acessando Viva Leilões página {current_page}: {url}")
        driver.get(url)
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.dg-leiloes-item-col'))
            )
        except Exception:
            logger.info("Nenhum card encontrado ou timeout na página.")
            break
        cards = driver.find_elements(By.CSS_SELECTOR, 'div.dg-leiloes-item-col')
        if not cards:
            break
        for card in cards:
            try:
                status = card.find_element(By.CSS_SELECTOR, 'span.BoxBtLoteLabel').text.strip()
            except:
                status = ""
            try:
                link = card.find_element(By.CSS_SELECTOR, 'a.dg-btn-lote-online').get_attribute("href")
            except:
                continue
            all_links.append((link, status))
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
    return all_links


def process_links(driver: webdriver.Chrome, link_status: list[tuple[str, str]]) -> list[dict]:
    results = []
    for idx, (link, status) in enumerate(link_status, start=1):
        logger.info(f"VivaLeilões {idx}/{len(link_status)}: {link}")
        driver.execute_script("window.open(arguments[0]);", link)
        driver.switch_to.window(driver.window_handles[-1])
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.dg-titulo'))
        )
        time.sleep(1)
        data = {"link": link, "status": status}
        def get_text(selector):
            try:
                return driver.find_element(By.CSS_SELECTOR, selector).text.strip()
            except:
                return None
        def get_href(selector):
            elems = driver.find_elements(By.CSS_SELECTOR, selector)
            return elems[0].get_attribute("href") if elems else None
        data.update({
            "titulo_leilao": get_text('div.dg-titulo'),
            "tipo_leilao": "Judicial",
            "numero_processo": get_text('a[href*="numero_processo"]'),
            "valor_imovel": get_text('span.ValorMinimoLanceSegundaPraca')
                             or get_text('span.ValorMinimoLancePrimeiraPraca'),
            "edital_leilao": get_href('ul li:nth-child(6) a'),
            "laudo_avaliacao": get_href('ul li:nth-child(1) a'),
            "matricula": get_href('ul li:nth-child(3) a'),
            "descricao_lote": get_text('div.dg-lote-descricao-txt')
        })
        results.append(data)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(0.5)
    return results


def run(pages: int) -> pd.DataFrame:
    driver = init_driver()
    try:
        links = collect_links(driver, pages)
        if not links:
            logger.warning("Nenhum link coletado")
        raw = process_links(driver, links)
    finally:
        driver.quit()
    return pd.DataFrame(raw)
