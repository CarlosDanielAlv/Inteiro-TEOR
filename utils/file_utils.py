import streamlit as st
import zipfile
import os
import shutil

def download_files(modelo_selecionado):
    # Define o nome do arquivo ZIP
    zip_filename = f"{modelo_selecionado}.zip"
    data_directory = 'data'

    # Verifica se a pasta 'data' contém arquivos PDF
    pdf_files = [f for f in os.listdir(data_directory) if f.endswith('.pdf')]

    if not pdf_files:
        st.warning("Nenhum arquivo PDF encontrado na pasta data.")
        return

    # Cria o arquivo ZIP e adiciona os arquivos PDF da pasta 'data'
    with zipfile.ZipFile(zip_filename, 'w') as zf:
        for pdf_file in pdf_files:
            file_path = os.path.join(data_directory, pdf_file)
            zf.write(file_path, os.path.basename(file_path))

    # Oferece o download do arquivo ZIP
    with open(zip_filename, "rb") as f:
        st.download_button(
            label="Baixar documentos",
            data=f,
            file_name=zip_filename,
            mime="application/zip"
        )

    # Remove os arquivos da pasta 'data' após o download
    clear_data_directory(data_directory)

    # Remove o arquivo zip após o download
    try:
        if os.path.exists(zip_filename):
            os.remove(zip_filename)
    except Exception as e:
        print(f"Erro ao remover o arquivo {zip_filename}: {e}")

def clear_data_directory(directory):
    """
    Remove todos os arquivos dentro da pasta especificada.
    """
    try:
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
    except Exception as e:
        print(f"Erro ao limpar o diretório {directory}: {e}")