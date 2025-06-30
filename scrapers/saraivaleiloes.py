import time
import re
import pandas as pd
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options


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


def collect_links(driver: webdriver.Chrome, pages: int) -> list[tuple[str, str]]:
    """
    Coleta todos os links de imóveis e seus status.
    pages < 0 percorre até não achar mais.
    Retorna lista de tuplas (link, status).
    """
    BASE_URL = "https://www.saraivaleiloes.com.br/buscador?page={page}&categoria=2"
    all_links: list[tuple[str, str]] = []
    current = 1

    while True:
        if pages >= 0 and current > pages:
            break

        url = BASE_URL.format(page=current)
        print(f"[Saraiva] Acessando página {current}: {url}")
        driver.get(url)
        time.sleep(3)

        cards = driver.find_elements(
            By.XPATH,
            '//article[contains(@class, "lote-main lote-main-status-1")]'
        )
        if not cards:
            break

        for card in cards:
            try:
                status = card.find_element(
                    By.XPATH, './/strong[contains(@class,"strong-status")]'
                ).text.strip()
            except:
                status = ""
            try:
                link = card.find_element(
                    By.XPATH, './/a[contains(@class,"link-img")]'
                ).get_attribute("href")
            except:
                continue
            all_links.append((link, status))
        current += 1

    return all_links


def process_links(driver: webdriver.Chrome, link_status: list[tuple[str, str]]) -> list[dict]:
    """
    Visita cada link coletado e extrai dados.
    Retorna lista de dicionários.
    """
    results: list[dict] = []

    for idx, (link, status) in enumerate(link_status, start=1):
        print(f"[Saraiva] Processando {idx}/{len(link_status)}: {link}")
        driver.execute_script("window.open(arguments[0]);", link)
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(2)

        def get_text(xpath: str) -> str | None:
            try:
                return driver.find_element(By.XPATH, xpath).text.strip()
            except:
                return None

        # Obter local (cidade - estado) e separar
        raw_loc = get_text('/html/body/section[4]/div/div[1]/div[1]/div/strong')
        if raw_loc:
            parts = raw_loc.split('-', 1)
            cidade = parts[0].strip()
            estado = parts[1].strip()
        else:
            cidade = None
            estado = None

        data = {
            "link": link,
            "status": status,
            "titulo_leilao": get_text('/html/body/section[4]/div/div[1]/div[1]/h1'),
            "cidade": cidade,
            "estado": estado,
            "tipo_leilao": get_text('/html/body/section[4]/div/div[2]/div/div[5]/ul[2]/li[4]/p'),
            "numero_processo": get_text('//span[contains(@class,"numeroProcessoTxt")]'),
            "valor_imovel": (
                get_text('/html/body/section[4]/div/div[2]/div/div[2]/div/div[1]/ul[2]/li[2]/div[2]/strong')
                or get_text('/html/body/section[4]/div/div[2]/div/div[2]/div/div[1]/ul[2]/li[1]/div[2]/strong')
            ),
            "edital_leilao": None,
            "matricula": None,
            "descricao_lote": get_text('//div[contains(@class,"line-text")]')
        }

        # edital
        try:
            data["edital_leilao"] = driver.find_element(
                By.XPATH,
                '/html/body/section[4]/div/div[2]/div/div[2]/div/div[2]/div[2]/ul/li[1]/a'
            ).get_attribute("href")
        except:
            pass

        # matrícula
        try:
            data["matricula"] = driver.find_element(
                By.XPATH,
                '/html/body/section[4]/div/div[2]/div/div[2]/div/div[2]/div[2]/ul/li[2]/a'
            ).get_attribute("href")
        except:
            pass

        results.append(data)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(1)

    return results


def run(pages: int) -> pd.DataFrame:
    """
    Executa o scraping e retorna DataFrame com colunas:
    link, status, titulo_leilao, cidade, estado, tipo_leilao, numero_processo,
    valor_imovel, edital_leilao, matricula, descricao_lote.
    pages: número de páginas a raspar; -1 para todas.
    """
    driver = init_driver()
    try:
        links = collect_links(driver, pages)
        raw = process_links(driver, links)
    finally:
        driver.quit()

    return pd.DataFrame(raw)
