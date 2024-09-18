import streamlit as st
import pandas as pd
from utils.document_generator import generate_documents
from utils.file_utils import download_files
from PIL import Image

# Função para iniciar a geração dos documentos
def start_generator():
    if uploaded_file and modelos_selecionados:
        # Exibe um spinner enquanto os documentos estão sendo gerados
        with st.spinner('Gerando documentos...'):
            df = pd.read_excel(uploaded_file)
            documentos_gerados = []

            # Criar lista de caminhos de templates dos modelos selecionados
            template_paths = [modelos[modelo]
                              for modelo in modelos_selecionados]

            # Gerar documentos para os modelos selecionados
            documentos_gerados += generate_documents(df, template_paths)

            st.success("Documentos gerados com sucesso!")
        # Botão para baixar os documentos
        download_files("Inteiro TEOR PDFs")  # Nome do zip
    else:
        st.warning("Faça o upload e selecione ao menos um modelo")



# --------------------------------------------------------------- Configurações da página
# Defina o favicon
favicon = Image.open("favicon.png")  # Pode ser .ico, .png, .jpg, etc.

st.set_page_config(
    page_title="Gerador de Inteiro TEOR",
    page_icon=favicon,  # Favicon personalizado
    layout="centered"
)

st.markdown('<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">', unsafe_allow_html=True)

st.markdown("""
<nav class="navbar fixed-top navbar-expand-lg navbar-dark" style="background-color: #3498DB; color: #fff;">
  <a class="navbar-brand" href="https://youtube.com/dataprofessor" target="_blank">Data Professor</a>
  <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNav" aria-controls="navbarNav" aria-expanded="false" aria-label="Toggle navigation">
    <span class="navbar-toggler-icon"></span>
  </button>
  <div class="collapse navbar-collapse" id="navbarNav">
    <ul class="navbar-nav">
      <li class="nav-item active">
        <a class="nav-link disabled" href="#">Home <span class="sr-only">(current)</span></a>
      </li>
      <li class="nav-item">
        <a class="nav-link" href="https://youtube.com/dataprofessor" target="_blank">YouTube</a>
      </li>
      <li class="nav-item">
        <a class="nav-link" href="https://twitter.com/thedataprof" target="_blank">Twitter</a>
      </li>
    </ul>
  </div>
</nav>
""", unsafe_allow_html=True)

st.markdown('''# **Gerador de Inteiro TEOR**
Gere Inteiro TEOR DA R1 e R2.
''')

# Organizar todos os componentes dentro de um formulário
with st.form(key='document_form'):
    # Upload do arquivo Excel
    uploaded_file = st.file_uploader(
        "Envie a planilha Excel com os dados", type=["xlsx"])

    # Multiselect para escolher mais de um modelo
    modelos = {
        "Defesa de Autuação": "modelo_da.docx",
        "Recurso Primeira instância": "modelo_r1.docx",
        "Recurso Segunda instância": "modelo_r2.docx"
    }
    modelos_selecionados = st.multiselect("Selecione os modelos de documento", list(
        modelos.keys()), placeholder="Selecione um ou mais modelos")

    # Botão para iniciar a geração dos documentos
    submit_button = st.form_submit_button(label="Iniciar")

# Adicionar a logo com centralização e redimensionamento
image = Image.open("logo.png")  # Certifique-se de ter o arquivo da logo
st.image(image, caption="", use_column_width=False, width=200)

# Lógica para o botão de envio
if submit_button:
    start_generator()


# Adicionar o footer com HTML e CSS embutido
footer = """
    <style>
        .footer {
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            background-color: #262730;
            color: #fff;
            text-align: center;
            padding: 10px;
            font-size: 14px;
        }
    </style>
    <div class="footer">
        <p>© 2024 - Todos os direitos reservados | Desenvolvido por Carlos Daniel da Silva Alves</p>
    </div>
"""
st.markdown(footer, unsafe_allow_html=True)

st.markdown("""
<script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
""", unsafe_allow_html=True)
