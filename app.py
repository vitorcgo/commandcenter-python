import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Separar Atendimentos por Especialidade e Convênio", layout="wide")
st.title("Separar Atendimentos por Especialidade e Convênio")

st.markdown("_Envie sua planilha de atendimentos para organizar automaticamente por especialidade, tipo de convênio e data._")

uploaded_file = st.file_uploader("Escolha a planilha de atendimentos (.xls ou .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    # Verifica e converte .xls para .xlsx, se necessário
    file_name = uploaded_file.name
    if file_name.endswith(".xls"):
        df_temp = pd.read_excel(uploaded_file, sheet_name="Report", engine="xlrd")
        buffer_xlsx = BytesIO()
        with pd.ExcelWriter(buffer_xlsx, engine="openpyxl") as writer:
            df_temp.to_excel(writer, sheet_name="Report", index=False)
        buffer_xlsx.seek(0)
        df = pd.read_excel(buffer_xlsx, sheet_name="Report")
    else:
        df = pd.read_excel(uploaded_file, sheet_name="Report")

    df.columns = [str(col).strip().lower() for col in df.columns]

    # Tentar detectar automaticamente
    col_especialidade = next((col for col in df.columns if "especialidade" in col), None)
    col_convenio = next((col for col in df.columns if "convenio" in col), None)
    col_data = next((col for col in df.columns if "data" in col), None)

    # Se não encontrar, permitir seleção manual
    if not all([col_especialidade, col_convenio, col_data]):
        st.warning("Não foi possível detectar automaticamente. Por favor, selecione manualmente as colunas corretas abaixo:")
        col_especialidade = st.selectbox("Coluna de Especialidade:", df.columns)
        col_convenio = st.selectbox("Coluna de Convênio:", df.columns)
        col_data = st.selectbox("Coluna de Data:", df.columns)

    # Selecionar e renomear colunas
    df = df[[col_especialidade, col_convenio, col_data]]
    df.columns = ["Especialidade", "Convenio", "Data"]

    especialidades_desejadas = ["CLI", "PED", "ORT"]
    df = df[df["Especialidade"].isin(especialidades_desejadas)]

    # Tipo de convênio
    df["TipoConvenio"] = df["Convenio"].apply(lambda x: "GRUPO" if "AMIL" in str(x).upper() else "EXTRA GRUPO")

    # Padronizar data
    df["Data"] = pd.to_datetime(df["Data"]).dt.date

    # Agrupamento
    resumo = df.groupby(["Especialidade", "TipoConvenio", "Data"]).size().reset_index(name="Total")

    # Pivot para formato de planilha final
    tabela_formatada = resumo.pivot_table(
        index=["Especialidade", "TipoConvenio"],
        columns="Data",
        values="Total",
        fill_value=0
    )

    st.subheader("Visualização dos Dados Formatados")
    st.dataframe(tabela_formatada)

    # Geração do Excel para download
    buffer = BytesIO()
    tabela_formatada.to_excel(buffer)
    buffer.seek(0)

    st.download_button(
        label="📄 Baixar Excel Formatado",
        data=buffer,
        file_name="planilha_formatada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Rodapé com crédito
st.markdown("---")
st.markdown("**Criado por Vitor Cavalcante Gomes**")
