# ✅ SCRIPT COMPLETO - PDF → Excel → Zip com Nome Personalizado (com PDFPlumber para Deploy Online)

# Ocultar/Mostrar Código com botão (Javascript)
from IPython.display import HTML
HTML('''
<script>
  code_show=true; 
  function code_toggle() {
    if (code_show){
      $('div.input').hide();
    } else {
      $('div.input').show();
    }
    code_show = !code_show
  } 
  $( document ).ready(code_toggle);
</script>
<form action="javascript:code_toggle()">
  <input type="submit" value="🔧 Mostrar/Ocultar Código">
</form>
''')

# Instalações silenciosas (Colab ou Local)
!pip install -q pdfplumber tabulate openpyxl xlsxwriter PyPDF2 tqdm streamlit > /dev/null

import os
import shutil
import logging
from pathlib import Path
import pandas as pd
import numpy as np
import pdfplumber
from tqdm.notebook import tqdm
from datetime import datetime
from zipfile import ZipFile
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

import PyPDF2
from openpyxl import load_workbook
from google.colab import files

# Diretórios
BASE_DIR = Path("/content/sample_pdfs")
CONVERTED_DIR = BASE_DIR / "convertidos"
BASE_DIR.mkdir(parents=True, exist_ok=True)
CONVERTED_DIR.mkdir(parents=True, exist_ok=True)

# Upload de PDFs
def upload_pdfs(dest_dir: Path) -> list[Path]:
    uploaded = files.upload()
    pdf_paths = []
    for fname in uploaded.keys():
        if fname.lower().endswith(".pdf"):
            src = Path("/content") / fname
            dst = dest_dir / fname
            shutil.move(str(src), str(dst))
            pdf_paths.append(dst)
    return pdf_paths

# Processamento com pdfplumber
def processar_pdf(pdf_path: Path, output_dir: Path) -> list[Path]:
    out_dir = output_dir / pdf_path.stem
    out_dir.mkdir(parents=True, exist_ok=True)
    arquivos = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for j, table in enumerate(tables):
                df = pd.DataFrame(table)
                df.replace(["", "nan", "NaN", "NULL"], np.nan, inplace=True)
                df.dropna(how="all", inplace=True)
                df.fillna("", inplace=True)
                df.reset_index(drop=True, inplace=True)
                out_file = out_dir / f"{pdf_path.stem}_tabela_{i+1}_{j+1}.xlsx"
                df.to_excel(out_file, index=False)
                arquivos.append(out_file)
    return arquivos

# Consolidação
def consolidar_xlsx(arquivos: list[Path], arquivo_saida: Path):
    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        for i, arquivo in enumerate(sorted(arquivos), 1):
            try:
                df = pd.read_excel(arquivo)
                if df.empty or df.shape[1] < 2:
                    continue
                nome = f"Tabela_{i}"
                df.to_excel(writer, sheet_name=nome[:31], index=False)
            except:
                continue

# Nome amigável
def extrair_nome_amigavel(pdf_path: Path):
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
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

# Criar TXT
def criar_txt(destino: Path):
    conteudo = "Olá! Este é um arquivo de texto criado com Python."
    with open(destino, "w", encoding="utf-8") as f:
        f.write(conteudo)
    return destino

# Upload de extras
def upload_arquivos_extras(destino_dir: Path) -> list[Path]:
    uploaded = files.upload()
    paths = []
    for fname in uploaded.keys():
        src = Path("/content") / fname
        dst = destino_dir / fname
        shutil.move(str(src), str(dst))
        paths.append(dst)
    return paths

# Função principal
def main():
    pdfs = upload_pdfs(BASE_DIR)
    if not pdfs:
        return

    pdf_origem = pdfs[0]
    arquivos_gerados = processar_pdf(pdf_origem, CONVERTED_DIR)
    consolidado_path = CONVERTED_DIR / "consolidado_final.xlsx"
    consolidar_xlsx(arquivos_gerados, consolidado_path)

    txt_path = CONVERTED_DIR / "arquivo_info.txt"
    criar_txt(txt_path)

    arquivos_extras = upload_arquivos_extras(CONVERTED_DIR)

    nome_amigavel = extrair_nome_amigavel(pdf_origem)
    data_str = datetime.today().strftime("%Y%m%d")
    nome_zip = f"{nome_amigavel}_{data_str}.zip"
    caminho_zip = BASE_DIR / nome_zip

    with ZipFile(caminho_zip, "w") as zipf:
        zipf.write(pdf_origem, arcname=pdf_origem.name)
        zipf.write(consolidado_path, arcname=consolidado_path.name)
        zipf.write(txt_path, arcname=txt_path.name)
        for arq in arquivos_extras:
            zipf.write(arq, arcname=arq.name)

    files.download(str(caminho_zip))

if __name__ == "__main__":
    main()