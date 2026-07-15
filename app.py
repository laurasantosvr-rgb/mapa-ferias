import streamlit as st
from gerar_mapa import gerar_mapa

st.set_page_config(
    page_title="Mapa de Férias",
    layout="wide"
)

st.title("📅 Mapa Anual de Férias")

ficheiro = st.file_uploader(
    "Selecione o ficheiro Excel",
    type=["xlsx"]
)

if ficheiro:

    if st.button("Gerar mapa"):

        caminho = gerar_mapa(ficheiro)

        with open(caminho, "rb") as f:

            st.download_button(
                "📥 Download",
                f,
                file_name="Mapa_Ferias.xlsx"
            )