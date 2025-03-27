# ============================================================
# 1. Importação de módulos e configuração inicial
# ============================================================
from selenium import webdriver                                  # Para controlar o navegador
from selenium.webdriver.common.by import By                     # Para localizar elementos
from selenium.webdriver.chrome.service import Service           # Para configurar o ChromeDriver
from selenium.webdriver.chrome.options import Options            # Para definir opções do Chrome (ex: modo headless)
import time                                                     # Para utilizar time.sleep()
import pandas as pd                                             # Para criar a planilha (CSV)

def print_header(message):
    print("\n" + "=" * 70)
    print(f"{message}".center(70))
    print("=" * 70 + "\n")

# ============================================================
# 2. Configuração do Chrome (Headless)
# ============================================================
print_header("Configuração do Chrome (Headless)")
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

chrome_driver_path = "C:\\Users\\mathe\\Desktop\\chromedriver-win64\\chromedriver.exe"
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# ============================================================
# 3. Coleta dos Links dos Imóveis de Todas as Páginas Desejadas
# ============================================================
base_url = "https://www.alfaleiloes.com/leiloes/?&page={page}&categoria=35&categoria=18&categoria=19&categoria=24&categoria=23&categoria=26&categoria=27&search="

print_header("Coleta de Links dos Imóveis")
paginas_input = input("Digite o número de páginas a serem raspadas (ou 'todas'): ")
if paginas_input.lower() == "todas":
    total_pages = None
else:
    total_pages = int(paginas_input)
print(f"[INFO] Páginas a raspar: {'Todas' if total_pages is None else total_pages}")

all_links = []
current_page = 1

while True:
    if total_pages is not None and current_page > total_pages:
        break

    current_url = base_url.format(page=current_page)
    print_header(f"Coletando Links - Página {current_page}")
    print(f"[INFO] Acessando: {current_url}")
    driver.get(current_url)
    time.sleep(5)

    leilao_items = driver.find_elements(By.XPATH, '//div[@class="cards-wrapper"]/div[@class="home-leiloes-cards"]')
    print(f"[INFO] Itens encontrados na Página {current_page}: {len(leilao_items)}")
    
    if len(leilao_items) == 0:
        print("[INFO] Nenhum item encontrado nesta página. Encerrando coleta.")
        break

    for index, item in enumerate(leilao_items, start=1):
        try:
            link_element = item.find_element(By.XPATH, './/a[@class="btn-card"]')
            link = link_element.get_attribute('href')
            all_links.append(link)
            print(f"  [OK] Item {index}: {link}")
        except Exception as e:
            print(f"  [ERRO] Item {index}: {e}")
    
    print(f"[INFO] Total de links coletados até a Página {current_page}: {len(all_links)}")
    current_page += 1

print_header("Total de Links Coletados")
print(f"[INFO] Total de imóveis coletados: {len(all_links)}")

# ============================================================
# 4. Processamento dos Imóveis (Extração dos Dados)
# ============================================================
all_imoveis_data = []

for i, link in enumerate(all_links, start=1):
    print_header(f"Processando Imóvel {i}/{len(all_links)}")
    print(f"[INFO] URL: {link}")
    # Abre o imóvel em uma nova aba para não interferir na listagem
    driver.execute_script("window.open(arguments[0]);", link)
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(3)

    try:
        titulo_leilao = driver.find_element(By.CLASS_NAME, "title-lote-leiloes").text
        print(f"  [OK] Título: {titulo_leilao}")
    except Exception as e:
        titulo_leilao = None
        print(f"  [ERRO] Título: {e}")

    try:
        tipo_leilao = driver.find_element(By.XPATH, '//*[@id="lotes"]/div[1]/div/h1').text
        print(f"  [OK] Tipo: {tipo_leilao}")
    except Exception as e:
        tipo_leilao = None
        print(f"  [ERRO] Tipo: {e}")

    try:
        numero_processo = driver.find_element(By.XPATH, '//*[@id="lotes"]/div[1]/div/div[4]/div[1]/a').text
        print(f"  [OK] Nº Processo (via a): {numero_processo}")
    except Exception as e:
        try:
            numero_processo = driver.find_element(By.XPATH, '//*[@id="lotes"]/div[1]/div/div[4]/div[1]/p[2]').text
            print(f"  [OK] Nº Processo (via p[2]): {numero_processo}")
        except Exception as e2:
            numero_processo = None
            print(f"  [ERRO] Nº Processo: {e2}")

    try:
        valor_imovel = driver.find_element(By.CLASS_NAME, "line-through").text
        print(f"  [OK] Valor (via classe): {valor_imovel}")
    except Exception as e:
        try:
            valor_imovel = driver.find_element(By.XPATH, '/html/body/div[2]/section[2]/div[1]/div/div[5]/ul/li[3]/p').text
            print(f"  [OK] Valor (via XPath): {valor_imovel}")
        except Exception as e2:
            valor_imovel = None
            print(f"  [ERRO] Valor: {e2}")

    try:
        edital_leilao = driver.find_element(By.XPATH, '//a[contains(translate(text(),"EDITAL","edital"), "edital")]').get_attribute('href')
        print(f"  [OK] Edital: {edital_leilao}")
    except Exception as e:
        edital_leilao = None
        print(f"  [ERRO] Edital: {e}")

    try:
        link_docs = driver.find_element(By.XPATH, '//a[contains(translate(text(),"DOCUMENTOS","documentos"), "documentos")]')
        link_docs.click()
        time.sleep(2)
        docs_container = driver.find_element(By.CLASS_NAME, "modal-body-doc")
        doc_links = docs_container.find_elements(By.TAG_NAME, "a")
        nomes_docs = [
            "Certidão de Matrícula",
            "Laudo de Avaliação",
            "Débitos Tributários",
            "Débito Exequendo/Condominial",
            "Manual de Participação"
        ]
        documentos_dict = {}
        for j, doc_link in enumerate(doc_links):
            href = doc_link.get_attribute('href')
            if j < len(nomes_docs):
                documentos_dict[nomes_docs[j]] = href
            else:
                documentos_dict[f"Documento {j+1}"] = href
        print("  [OK] Documentos:")
        for nome, link_doc in documentos_dict.items():
            print(f"       {nome:30}: {link_doc}")
        documentos = documentos_dict
    except Exception as e:
        documentos = None
        print(f"  [ERRO] Documentos: {e}")

    try:
        descricao_lote = driver.find_element(By.CLASS_NAME, "content").text
        print("  [OK] Descrição do Lote extraída.")
    except Exception as e:
        descricao_lote = None
        print(f"  [ERRO] Descrição do Lote: {e}")

    imovel_info = {
        "link": link,
        "titulo_leilao": titulo_leilao,
        "tipo_leilao": tipo_leilao,
        "numero_processo": numero_processo,
        "valor_imovel": valor_imovel,
        "edital_leilao": edital_leilao,
        "documentos": documentos,
        "descricao_lote": descricao_lote
    }
    all_imoveis_data.append(imovel_info)
    print_header(f"Dados Extraídos do Imóvel {i}")
    for chave, valor in imovel_info.items():
        print(f"{chave:20}: {valor}")
    
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(1)

# ============================================================
# 5. Resumo Final dos Dados Coletados e Exportação para Excel
# ============================================================
print_header("Resumo Final dos Dados Coletados")
for idx, data in enumerate(all_imoveis_data, start=1):
    print(f"Imóvel {idx}:")
    for key, value in data.items():
        print(f"    {key:20}: {value}")
    print("-" * 70)

# Converte a lista de dicionários para DataFrame
def format_documentos(doc_dict):
    if not doc_dict:
        return ""
    return "\n".join([f"{nome}: {link}" for nome, link in doc_dict.items()])

# Prepara os dados para o Excel
dados_excel = []
for item in all_imoveis_data:
    item_excel = item.copy()
    item_excel["documentos"] = format_documentos(item.get("documentos"))
    dados_excel.append(item_excel)

df = pd.DataFrame(dados_excel)

# Caminho do arquivo de saída (.xlsx)
output_file = "imoveis_formatado.xlsx"

# Criação do Excel com ajustes visuais
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Imoveis')
    workbook  = writer.book
    worksheet = writer.sheets['Imoveis']
    
    # Ajuste automático da largura das colunas
    for i, column in enumerate(df.columns):
        max_len = df[column].astype(str).map(len).max()
        worksheet.set_column(i, i, min(max_len + 2, 80))  # limite de largura

print(f"[INFO] Dados exportados com sucesso para {output_file}")


# ============================================================
# 6. Finaliza o Navegador
# ============================================================
print_header("Finalizando o Scraping e Fechando o Navegador")
driver.quit()