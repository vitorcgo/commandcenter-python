import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Separar Atendimentos por Especialidade e Convênio", layout="wide")
st.title("Separar Atendimentos por Especialidade e Convênio")
st.markdown("_Envie sua planilha de atendimentos para organizar automaticamente por especialidade, tipo de convênio e data._")

uploaded_file = st.file_uploader("Escolha a planilha de atendimentos (.xls ou .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    file_name = uploaded_file.name.lower()

    try:
        if file_name.endswith(".xls"):
            df = pd.read_excel(uploaded_file, sheet_name="Report", engine="xlrd", header=None)
        else:
            df = pd.read_excel(uploaded_file, sheet_name="Report", engine="openpyxl", header=None)
    except Exception as e:
        st.error(f"Erro ao abrir a planilha: {e}")
        st.stop()

    # Seleciona colunas A (0), D (3), N (13)
    df = df.iloc[:, [0, 3, 13]]
    df.columns = ["Especialidade", "Convenio", "Data"]

    # Remove linhas com qualquer valor faltando
    df = df.dropna(subset=["Especialidade", "Convenio", "Data"])

    # Filtros das especialidades desejadas
    especialidades_desejadas = ["CLI", "PED", "ORT"]
    df = df[df["Especialidade"].isin(especialidades_desejadas)]

    # Classificação do tipo de convênio
    df["TipoConvenio"] = df["Convenio"].apply(lambda x: "GRUPO" if "AMIL" in str(x).upper() else "EXTRA GRUPO")

    # Padronização da data
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.date
    df = df.dropna(subset=["Data"])  # remove linhas onde a data virou NaT

    # Agrupamento
    resumo = df.groupby(["Especialidade", "TipoConvenio", "Data"]).size().reset_index(name="Total")

    # Pivot para formato final
    tabela_formatada = resumo.pivot_table(
        index=["Especialidade", "TipoConvenio"],
        columns="Data",
        values="Total",
        fill_value=0
    )

    st.subheader("Visualização dos Dados Formatados")
    st.dataframe(tabela_formatada)

    # Download
    buffer = BytesIO()
    tabela_formatada.to_excel(buffer)
    buffer.seek(0)

    st.download_button(
        label="📄 Baixar Excel Formatado",
        data=buffer,
        file_name="planilha_formatada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # TOTAL DE PACIENTES POR DIA
    total_por_dia = df.groupby("Data").size().reset_index(name="TotalPacientes")

    if not total_por_dia.empty:
        dia_mais = total_por_dia.loc[total_por_dia["TotalPacientes"].idxmax()]
        dia_menos = total_por_dia.loc[total_por_dia["TotalPacientes"].idxmin()]

        st.markdown("### 📊 Análise de Atendimentos")
        st.markdown(f"🔝 **Maior movimento:** {dia_mais['Data'].strftime('%d/%m/%Y')} com **{dia_mais['TotalPacientes']} pacientes**")
        st.markdown(f"🔻 **Menor movimento:** {dia_menos['Data'].strftime('%d/%m/%Y')} com **{dia_menos['TotalPacientes']} pacientes**")

# Rodapé
st.markdown("---")
st.markdown("**Criado por Vitor Cavalcante Gomes - vitor.cavalcante@amil.com.br - www.vitorgomes.tech**")
