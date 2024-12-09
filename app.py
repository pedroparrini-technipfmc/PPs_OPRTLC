import os
import shutil
#import PyPDF2
import pikepdf
import pythoncom
import comtypes.client
import streamlit as st
#from io import BytesIO

# Importar funções

def convert_docx_2_pdf(input_file):

    auxiliar_file = 'uploads\\auxiliar.pdf'
    output_file = input_file[:-5] + ".pdf"
    word = comtypes.client.CreateObject("Word.Application")
    docx_path = os.path.abspath(input_file)
    pdf_path = os.path.abspath(auxiliar_file)
    pdf_format = 17
    word.Visible = False
    in_file = word.Documents.Open(docx_path)
    in_file.Saveas(pdf_path, FileFormat=pdf_format)
    in_file.Close()
    word.Quit()
    os.remove(input_file)

    return [output_file, auxiliar_file]

def merge_pdfs(pdf_pp_path, pdf_desenho_path, pagina_inicial, pagina_final, pdf_final):

    # Abrir os PDFs
    with pikepdf.Pdf.open(pdf_pp_path) as pdf_base, pikepdf.Pdf.open(pdf_desenho_path) as pdf_substituto:

        # Criar um novo PDF para o resultado
        pdf_resultante = pikepdf.Pdf.new()

        # Copiar páginas do PDF base, substituindo as páginas em branco
        for i, page in enumerate(pdf_base.pages):
            if pagina_inicial <= i <= pagina_final:
                # Substituir páginas em branco pelas do PDF substituto
                pdf_resultante.pages.append(pdf_substituto.pages[i - pagina_inicial])
            else:
                # Adicionar a página original
                pdf_resultante.pages.append(page)

        # Salvar o PDF resultante
        pdf_resultante.save(pdf_final)

def qtd_pags(pdf_path):

    with pikepdf.Pdf.open(pdf_path) as pdf:

        return len(pdf.pages)
    
def clean_folder(pasta):
    
    # Limpar todos os arquivos da pasta uploads
    for filename in os.listdir(pasta):
        file_path = os.path.join(pasta, filename)
        if os.path.isfile(file_path) or os.path.islink(file_path):
            os.unlink(file_path)  # Remove arquivo ou link
        elif os.path.isdir(file_path):
            shutil.rmtree(file_path)  # Remove diretório

# Configuração inicial da página

st.set_page_config(page_title="Rev. PPs", page_icon="✅", layout="centered", initial_sidebar_state="expanded")

st.title("Revisão de Procedimentos Padrão")

# Barra lateral com duas opções

st.sidebar.title("Menu")
option = st.sidebar.selectbox("Escolha uma opção:", ("Início da Revisão", "Fim da Revisão"))

clean_folder("uploads")

if option == "Início da Revisão":

    pass

elif option == "Fim da Revisão":

    # Upload de arquivos
    st.write("Faça o upload dos arquivos para serem modificados.")
    arquivo_docx = st.file_uploader("Carregar arquivo .docx", type="docx")

    if arquivo_docx:

        arquivo_docx_path = os.path.join("uploads", arquivo_docx.name)
        with open(arquivo_docx_path, "wb") as f:
            f.write(arquivo_docx.getbuffer())

    pdf_desenho = st.file_uploader("Carregar arquivo .pdf", type="pdf")

    if pdf_desenho:

        pdf_desenho_path = os.path.join("uploads", pdf_desenho.name)
        with open(pdf_desenho_path, "wb") as f:
            f.write(pdf_desenho.getbuffer())

        pagina_inicial = int(st.number_input("Página inicial dos desenhos", value=3))
        pagina_final = int(st.number_input("Página final dos desenhos", value=qtd_pags(pdf_desenho) + pagina_inicial -1))

    if arquivo_docx and pdf_desenho and pagina_inicial and pagina_final:

        pythoncom.CoInitialize()

        pdf_pp_path = convert_docx_2_pdf(arquivo_docx_path)

        merge_pdfs(pdf_pp_path[1], pdf_desenho_path, pagina_inicial - 1, pagina_final - 1, pdf_pp_path[0])

        with open(pdf_pp_path[0], "rb") as file:
            st.download_button(
                label="Baixar PDF Modificado",
                data=file,
                file_name=os.path.basename(pdf_pp_path[0]),
                mime="application/pdf"
            )

