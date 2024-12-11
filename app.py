import streamlit as st
import pdfplumber
import pandas as pd

def pdf_para_dataframe(ruta_arquivo):
    """
    Extrai tabelas de um arquivo PDF e combina em um único DataFrame.
    """
    try:
        with pdfplumber.open(ruta_arquivo) as pdf:
            tabelas = []
            for pagina in pdf.pages:
                # Extrai tabelas da página
                page_tables = pagina.extract_tables()
                for tabela in page_tables:
                    # Converte tabela para DataFrame
                    df = pd.DataFrame(tabela[1:], columns=tabela[0])
                    tabelas.append(df)

        # Combina todas as tabelas em um único DataFrame
        dataframe_combinado = pd.concat(tabelas, ignore_index=True)
        return dataframe_combinado
    except Exception as e:
        st.error(f"Erro ao processar o PDF: {e}")
        return None


# Configurar Streamlit
st.title("Bot para Converter PDF em Excel")

# Upload do arquivo PDF
uploaded_file = st.file_uploader("Envie o arquivo PDF", type=["pdf"])

if uploaded_file:
    st.write("Arquivo carregado com sucesso!")
    # Converter PDF para DataFrame
    df = pdf_para_dataframe(uploaded_file)

    if df is not None:
        st.write("Tabela extraída do PDF:")
        st.dataframe(df)

        # Limpeza dos dados
        st.write("Processando e limpando dados...")
        try:
            # Renomear colunas
            df.columns = ['Clase', 'Nome', 'NumInsc', 'NroFilho', 'DataNac', 
                          'GEO', 'PORT', 'EDU', 'Objetivas', 'Discursiva', 
                          'Titulos', 'Total']

            # Remover colunas desnecessárias
            colunas_remover = ['Clase', 'Discursiva', 'Titulos']
            df = df.drop(columns=colunas_remover)

            # Substituir "Elimin." por 0 na coluna 'Total'
            df['Total'] = df['Total'].replace('Elimin.', 0)

            # Conversões de tipos de dados
            df['NroFilho'] = pd.to_numeric(df['NroFilho'], errors='coerce').fillna(0).astype(int)
            df['DataNac'] = pd.to_datetime(df['DataNac'], errors='coerce')
            colunas_numericas = ['GEO', 'PORT', 'EDU', 'Objetivas', 'Total']
            df[colunas_numericas] = df[colunas_numericas].apply(pd.to_numeric, errors='coerce')

            # Remover linhas duplicadas
            df = df.drop_duplicates()

            # Remover linhas com valores ausentes e resetar índice
            df = df.dropna().reset_index(drop=True)

            st.write("Dados limpos:")
            st.dataframe(df)

            # Botão para download do Excel
            excel_file = "dados_limpos_pdf.xlsx"
            df.to_excel(excel_file, index=False)
            with open(excel_file, "rb") as file:
                st.download_button("Baixar Excel Limpo", data=file, file_name="dados_limpos_pdf.xlsx")
        except Exception as e:
            st.error(f"Erro durante a limpeza dos dados: {e}")
    else:
        st.error("Não foi possível extrair as tabelas do PDF.")

