import requests
from bs4 import BeautifulSoup
import pandas as pd
import logging
import time
from typing import List, Tuple

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BASE_URL = (
    "https://www.megaleiloes.com.br/imoveis?tov=igbr&valor_max=5000000&"
    "tipo%5B0%5D=1&pagina={page}"
)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
}


def collect_links(pages: int) -> List[Tuple[str, str]]:
    all_links = []
    current_page = 1

    while True:
        if pages >= 0 and current_page > pages:
            break

        url = BASE_URL.format(page=current_page)
        logger.info(f"Acessando página {current_page}: {url}")
        response = requests.get(url, headers=HEADERS)
        soup = BeautifulSoup(response.text, "html.parser")

        cards = soup.select('div[class*="col-md-4"]')
        logger.info(f"Encontrados {len(cards)} imóveis na página {current_page}.")
        if not cards:
            break

        for card in cards:
            try:
                link_tag = card.select_one("a.card-title")
                status_tag = card.select_one("div.card-status")

                if link_tag and link_tag.get("href"):
                    link = link_tag["href"]
                    status = status_tag.get_text(strip=True) if status_tag else ""
                    all_links.append((link, status))
            except Exception as e:
                logger.warning(f"Erro ao coletar card: {e}")
                continue

        current_page += 1
        time.sleep(1)

    return all_links


def process_links(link_status: List[Tuple[str, str]]) -> List[dict]:
    results = []

    for idx, (link, status) in enumerate(link_status, start=1):
        logger.info(f"Processando {idx}/{len(link_status)}: {link}")
        try:
            response = requests.get(link, headers=HEADERS)
            soup = BeautifulSoup(response.text, "html.parser")

            def get_text(selector):
                el = soup.select_one(selector)
                return el.get_text(strip=True) if el else None

            def get_href_by_index(index: int):
                try:
                    anchors = soup.select("div.documents a")
                    return anchors[index]["href"] if len(anchors) > index else None
                except Exception:
                    return None

            data = {
                "link": link,
                "status": status,
                "titulo_leilao": get_text("h1.section-header"),
                "tipo_leilao": get_text("div.batch-type"),
                "numero_processo": get_text("div.batch-information a"),
                "valor_imovel": get_text("div.value"),
                "edital_leilao": get_href_by_index(1),
                "laudo_avaliacao": get_href_by_index(2),
                "matricula": get_href_by_index(3),
                "descricao_lote": get_text("div.description")
            }

            results.append(data)
        except Exception as e:
            logger.warning(f"Erro ao processar {link}: {e}")
        time.sleep(1)

    return results


def run(pages: int) -> pd.DataFrame:
    links = collect_links(pages)
    raw_data = process_links(links)
    df = pd.DataFrame(raw_data)
    return df
