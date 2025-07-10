# app.py
import streamlit as st
import os
from pathlib import Path
import pandas as pd
import numpy as np
import shutil
from zipfile import ZipFile
from datetime import datetime
from io import BytesIO
import camelot
import PyPDF2

# Configuração do app
st.set_page_config(page_title="Gerador de Consolidado", layout="centered")
st.title("📄 Gerador de Consolidado ZIP a partir de PDF")

st.markdown("Envie o PDF abaixo para gerar automaticamente as tabelas e o pacote consolidado.")

# Diretórios temporários
BASE_DIR = Path("temp_data")
CONVERTED_DIR = BASE_DIR / "convertidos"
BASE_DIR.mkdir(exist_ok=True)
CONVERTED_DIR.mkdir(exist_ok=True)

# Funções
def salvar_pdf(uploaded_file):
    pdf_path = BASE_DIR / uploaded_file.name
    with open(pdf_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return pdf_path

def processar_pdf(pdf_path, output_dir):
    tables = camelot.read_pdf(str(pdf_path), pages="all", flavor="lattice")
    if not tables:
        tables = camelot.read_pdf(str(pdf_path), pages="all", flavor="stream")
    
    arquivos = []
    for i, tabela in enumerate(tables, 1):
        df = tabela.df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        df.replace(["", "nan", "NaN", "NULL"], np.nan, inplace=True)
        df.dropna(how="all", inplace=True)
        df.fillna("", inplace=True)
        df.reset_index(drop=True, inplace=True)

        xlsx_path = output_dir / f"{pdf_path.stem}_tabela_{i}.xlsx"
        df.to_excel(xlsx_path, index=False)
        arquivos.append(xlsx_path)
    return arquivos

def consolidar_xlsx(arquivos, caminho_saida):
    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        for i, arq in enumerate(arquivos, 1):
            df = pd.read_excel(arq)
            df.to_excel(writer, sheet_name=f"Tabela_{i}", index=False)

def extrair_nome_amigavel(pdf_path):
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        texto = "\n".join(p.extract_text() for p in reader.pages if p.extract_text())
    idx = texto.find("Nome Amigável")
    if idx != -1:
        palavras = texto[idx:].split()
        try:
            pos = palavras.index("Amigável")
            return palavras[pos + 1]
        except:
            return "sem_nome"
    return "sem_nome"

def criar_txt(destino):
    conteudo = "Olá! Este é um arquivo de texto criado com Python."
    with open(destino, "w", encoding="utf-8") as f:
        f.write(conteudo)
    return destino

def zipar_conteudo(nome_zip, arquivos):
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "w") as zipf:
        for path in arquivos:
            zipf.write(path, arcname=Path(path).name)
    zip_buffer.seek(0)
    return zip_buffer

# Upload principal
uploaded_pdf = st.file_uploader("📤 Envie o arquivo PDF", type=["pdf"])
arquivos_extras = st.file_uploader("📎 Anexar outros arquivos para incluir no .zip (opcional)", type=None, accept_multiple_files=True)

if uploaded_pdf and st.button("🚀 Gerar ZIP"):
    with st.spinner("⏳ Processando o PDF..."):

        pdf_path = salvar_pdf(uploaded_pdf)
        nome_amigavel = extrair_nome_amigavel(pdf_path)
        data_hoje = datetime.today().strftime("%Y%m%d")
        nome_zip = f"{nome_amigavel}_{data_hoje}.zip"

        arquivos_excel = processar_pdf(pdf_path, CONVERTED_DIR)

        consolidado_path = CONVERTED_DIR / "consolidado_final.xlsx"
        consolidar_xlsx(arquivos_excel, consolidado_path)

        txt_path = CONVERTED_DIR / "arquivo_info.txt"
        criar_txt(txt_path)

        paths_para_zip = [pdf_path, consolidado_path, txt_path]

        if arquivos_extras:
            for arq in arquivos_extras:
                extra_path = CONVERTED_DIR / arq.name
                with open(extra_path, "wb") as f:
                    f.write(arq.getbuffer())
                paths_para_zip.append(extra_path)

        buffer_zip = zipar_conteudo(nome_zip, paths_para_zip)

    st.success("✅ ZIP gerado com sucesso!")
    st.download_button("📦 Baixar ZIP", data=buffer_zip, file_name=nome_zip, mime="application/zip")
