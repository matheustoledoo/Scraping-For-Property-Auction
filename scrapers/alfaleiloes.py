import time
import logging
import pandas as pd
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

logger = logging.getLogger(__name__)

BASE_URL = (
    "https://www.alfaleiloes.com/leiloes/?&page={page}"
    "&categoria=35&categoria=18&categoria=19"
    "&categoria=24&categoria=23&categoria=26&categoria=27&search="
)

def init_driver():
    chromedriver_autoinstaller.install()
    opts = Options()
    opts.add_argument("--headless")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    return webdriver.Chrome(options=opts)

def collect_links(driver, pages):
    all_links = []
    status = {}
    page = 1
    total = None if pages < 0 else pages
    while True:
        if total and page > total:
            break
        url = BASE_URL.format(page=page)
        logger.info(f"Acessando AlfaLeilões página {page}: {url}")
        driver.get(url)
        time.sleep(4)
        cards = driver.find_elements(
            By.XPATH, '//div[@class="cards-wrapper"]/div[@class="home-leiloes-cards"]'
        )
        if not cards:
            break
        for c in cards:
            try:
                st = c.find_element(By.CLASS_NAME, "card-status") \
                     .find_element(By.TAG_NAME, "p").text.strip()
            except:
                st = ""
            if st.lower() == "vendido":
                continue
            try:
                link = c.find_element(By.XPATH, './/a[@class="btn-card"]') \
                        .get_attribute("href")
                all_links.append(link)
                status[link] = st
            except:
                pass
        page += 1
    return all_links, status

def process_links(driver, links, status):
    data = []
    for idx, link in enumerate(links, start=1):
        logger.info(f"Processando AlfaLeilões {idx}/{len(links)}: {link}")
        driver.execute_script("window.open(arguments[0]);", link)
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(3)
        def gt(xpath):
            try:
                return driver.find_element(By.XPATH, xpath).text.strip()
            except:
                return None

        titulo   = gt('//div[contains(@class,"title-lote-leiloes")]')
        tipo     = gt('//*[@id="lotes"]/div[1]/div/h1')
        proc     = gt('//*[@id="lotes"]/div[1]/div/div[4]/div[1]/a')
        valor    = gt('//span[contains(@class,"line-through")]')
        try:
            edital = driver.find_element(
                By.XPATH,
                '//a[contains(translate(text(),"EDITAL","edital"),"edital")]'
            ).get_attribute("href")
        except:
            edital = None
        try:
            driver.find_element(
                By.XPATH,
                '//a[contains(translate(text(),"DOCUMENTOS","documentos"),"documentos")]'
            ).click()
            time.sleep(1)
            links_docs = [
                a.get_attribute("href")
                for a in driver.find_elements(By.CSS_SELECTOR, ".modal-body-doc a")
            ]
        except:
            links_docs = []

        descricao = gt('//div[contains(@class,"content")]')
        st = status.get(link, "")

        data.append({
            "link": link,
            "status": st,
            "titulo_leilao": titulo,
            "tipo_leilao": tipo,
            "numero_processo": proc,
            "valor_imovel": valor,
            "edital_leilao": edital,
            "documentos": ";".join(links_docs),
            "descricao_lote": descricao
        })

        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(1)

    return data

def run(pages: int) -> pd.DataFrame:
    driver = init_driver()
    try:
        links, status = collect_links(driver, pages)
        rows = process_links(driver, links, status)
    finally:
        driver.quit()
    return pd.DataFrame(rows)
