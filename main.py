import streamlit as st
import pandas as pd
from megaleiloes import scrape_megaleiloes
from alfaleiloes import scrape_alfaleiloes

st.title("Interface de Scraping de Leilões")

# Seletor de sites para scraping
st.subheader("Selecione os sites para realizar o scraping:")
mega_selected = st.checkbox("Mega Leilões")
alfa_selected = st.checkbox("Alfa Leilões")

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
                dados_mega = scrape_megaleiloes(paginas)
                dados_coletados.extend(dados_mega)
            st.success("Dados da Mega Leilões coletados com sucesso!")

        if alfa_selected:
            with st.spinner("Raspando dados da Alfa Leilões..."):
                dados_alfa = scrape_alfaleiloes(paginas)
                dados_coletados.extend(dados_alfa)
            st.success("Dados da Alfa Leilões coletados com sucesso!")

        if dados_coletados:
            df = pd.DataFrame(dados_coletados)
            st.dataframe(df)

            # Botão para download dos dados em Excel
            if st.button("Download dos dados em Excel"):
                df.to_excel("dados_leiloes.xlsx", index=False)
                st.success("Arquivo Excel gerado com sucesso!")
