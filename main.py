import os
import subprocess
import tempfile
import pandas as pd
import streamlit as st

def run_scraper(script: str, pages: str) -> pd.DataFrame:
    """Executa o script de scraping em um subprocesso e retorna o DataFrame."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        env = os.environ.copy()
        env["PAGINAS"] = pages
        env["OUTPUT_FILE"] = tmp.name
        subprocess.run(["python", script], check=True, env=env)
        df = pd.read_excel(tmp.name)
    return df


st.title("Interface de Scraping de Leilões")

# Seletor de sites para scraping
st.subheader("Selecione os sites para realizar o scraping:")
mega_selected = st.checkbox("Mega Leilões")
alfa_selected = st.checkbox("Alfa Leilões")
viva_selected = st.checkbox("Viva Leilões")

# Campo para entrada do número de páginas
paginas = st.text_input("Digite o número de páginas a serem raspadas (ou 'todas'):", "1")

# Botão para iniciar o scraping
if st.button("Iniciar Scraping"):
    if not mega_selected and not alfa_selected:
        st.warning("Por favor, selecione pelo menos um site para realizar o scraping.")
    else:
        dados_coletados = []

        if mega_selected:
            with st.spinner("Raspando dados da Mega Leilões..."):
                dados_mega = run_scraper("megaleiloes.py", paginas)
                dados_coletados.append(dados_mega)
            st.success("Dados da Mega Leilões coletados com sucesso!")

        if alfa_selected:
            with st.spinner("Raspando dados da Alfa Leilões..."):
                dados_alfa = run_scraper("alfaleiloes.py", paginas)
                dados_coletados.append(dados_alfa)
            st.success("Dados da Alfa Leilões coletados com sucesso!")

        if viva_selected:
            with st.spinner("Raspando dados da Viva Leilões..."):
                dados_viva = run_scraper("vivaleiloes.py", paginas)
                dados_coletados.append(dados_viva)
            st.success("Dados da Viva Leilões coletados com sucesso!")

        if dados_coletados:
            df = pd.concat(dados_coletados, ignore_index=True)
            st.dataframe(df)

            # Botão para download dos dados em Excel
            if st.button("Gerar arquivo Excel"):
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_out:
                    df.to_excel(tmp_out.name, index=False)
                with open(tmp_out.name, "rb") as f:
                    st.download_button(
                        label="Download dos dados em Excel",
                        data=f,
                        file_name="dados_leiloes.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
