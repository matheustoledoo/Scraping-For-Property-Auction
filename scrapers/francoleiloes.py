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
    "https://www.francoleiloes.com.br/busca/#Engine=Start&Pagina={page}&RangeValores=0&OrientacaoBusca=0&Busca=&Mapa=&ID_Categoria=0&ID_Estado=-1&ID_Cidade=-1&Bairro=&ID_Regiao=0&ValorMinSelecionado=0&ValorMaxSelecionado=0&Ordem=0&QtdPorPagina=24&ID_Leiloes_Status=&SubStatus=&PaginaIndex=3&BuscaProcesso=&NomesPartes=&CodLeilao=&TiposLeiloes=[]&CFGs=[]"
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
            '//div[contains(@class, "col-xs-12 col-sm-6 col-md-4 col-lg-3 dg-leiloes-item-col ")]'
        )
        logger.info(f"Encontrados {len(cards)} imóveis na página {current_page}.")
        if not cards:
            break

        for card in cards:
            try:
                status = card.find_element(
                    By.XPATH, './/span[contains(@class, "BoxBtLoteLabel")]'
                ).text.strip()
            except Exception:
                status = ""
            try:
                # busca a DIV flex-1 dentro de 'card' e, a partir dela, o A
                link_elem = card.find_element(
                    By.XPATH,
                    './/div[contains(@class, "dg-leiloes-lista-img")]/a'
                )
                link = link_elem.get_attribute('href')
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
        raw_loc = get_text('/html/body/section[5]/div/div[2]/div[1]/div/div/div')
        if raw_loc:
            # separa em partes pelo caractere “-”
            parts = [p.strip() for p in raw_loc.split('-')]
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
            "titulo_leilao": get_text('/html/body/main/div/div[1]/div[1]/h1/span[1]') or get_text('//div[contains(@class, "dg-lote-nome")]'),
            "cidade": cidade,
            "estado": estado,
            "tipo_leilao": 'Extrajudicial',
            "numero_processo": get_text(
                '/html/body/section[4]/div/div[2]/div/div[4]/article/a/p/span[1]'
            ) or get_text('//span[contains(@class, "numeroProcessoTxt")]'),
            "valor_imovel": get_text('/html/body/main/div/div[2]/div[2]/div/div[1]/div/div[3]/div/div[3]/span') or get_text('/html/body/main/div/div[2]/div[2]/div/div[1]/div/div[2]/div/div[3]/span') or get_text('/html/body/main/div/div[2]/div[2]/div/div[1]/div/div[1]/div/div[3]/span') or get_text('//div[contains(@class, "data-right")]'),
            "edital_leilao": driver.find_element(
                By.XPATH, '/html/body/section[4]/div/div/div[2]/div/div/div/div/ul/li[2]/a[2]'
            ).get_attribute('href') if driver.find_elements(
                By.XPATH, '/html/body/section[4]/div/div[2]/div/div[2]/div/div[2]/div[2]/ul/li/a'
            ) else None,
            "laudo_avaliacao": driver.find_element(
                By.XPATH, '/html/body/section[4]/div/div[2]/div/div[2]/div/div[2]/div[2]/ul/li[1]/a'
            ).get_attribute('href') if driver.find_elements(
                By.XPATH, '/html/body/div[3]/div[3]/div[3]/div[3]/div[2]/a[3]'
            ) else None,
            "matricula": driver.find_element(
                By.XPATH, '/html/body/section[4]/div/div/div[2]/div/div/div/div/ul/li[3]/a[2]'
            ).get_attribute('href') if driver.find_elements(
                By.XPATH, '/html/body/div[3]/div[3]/div[3]/div[3]/div[2]/a[4]'
            ) else None,
            "descricao_lote": get_text('/html/body/section[3]/div/div/div[2]/div')
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
