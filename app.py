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
        df_temp = pd.read_excel(uploaded_file, sheet_name="Report", engine="xlrd", header=None)
        buffer_xlsx = BytesIO()
        with pd.ExcelWriter(buffer_xlsx, engine="openpyxl") as writer:
            df_temp.to_excel(writer, sheet_name="Report", index=False, header=False)
        buffer_xlsx.seek(0)
        df = pd.read_excel(buffer_xlsx, sheet_name="Report", header=None)
    else:
        df = pd.read_excel(uploaded_file, sheet_name="Report", header=None)

    # Seleciona colunas específicas: A (0), D (3), N (13)
    df = df.iloc[:, [0, 3, 13]]
    df.columns = ["Especialidade", "Convenio", "Data"]

    # Limpa campos de texto
    df["Especialidade"] = df["Especialidade"].astype(str).str.strip()
    df["Convenio"] = df["Convenio"].astype(str).str.strip()

    # Remove linhas com campos em branco ou com "NaN" disfarçado
    df = df[
        (df["Especialidade"].str.upper() != "NAN") & (df["Especialidade"] != "") &
        (df["Convenio"].str.upper() != "NAN") & (df["Convenio"] != "")
    ]

    # Converte datas e remove inválidas
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.date
    df = df.dropna(subset=["Data"])

    # Filtros das especialidades desejadas
    especialidades_desejadas = ["CLI", "PED", "ORT"]
    df = df[df["Especialidade"].isin(especialidades_desejadas)]

    # Classificação do tipo de convênio
    df["TipoConvenio"] = df["Convenio"].apply(lambda x: "GRUPO" if "AMIL" in str(x).upper() else "EXTRA GRUPO")

    # Agrupamento
    resumo = df.groupby(["Especialidade", "TipoConvenio", "Data"]).size().reset_index(name="Total")

    # Pivot final para tabela formatada
    tabela_formatada = resumo.pivot_table(
        index=["Especialidade", "TipoConvenio"],
        columns="Data",
        values="Total",
        fill_value=0
    )

    st.subheader("📊 Visualização dos Dados Formatados")
    st.dataframe(tabela_formatada)

    # Geração do arquivo para download
    buffer = BytesIO()
    tabela_formatada.to_excel(buffer)
    buffer.seek(0)

    st.download_button(
        label="📄 Baixar Excel Formatado",
        data=buffer,
        file_name="planilha_formatada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Total de pacientes por dia
    total_por_dia = df.groupby("Data").size().reset_index(name="TotalPacientes")

    # Maior e menor volume
    if not total_por_dia.empty:
        dia_mais = total_por_dia.sort_values("TotalPacientes", ascending=False).iloc[0]
        dia_menos = total_por_dia.sort_values("TotalPacientes", ascending=True).iloc[0]

        st.markdown("### 🔍 Análise de Atendimentos")
        st.markdown(f"📈 **Maior movimento:** {dia_mais['Data'].strftime('%d/%m/%Y')} com **{dia_mais['TotalPacientes']} pacientes**")
        st.markdown(f"📉 **Menor movimento:** {dia_menos['Data'].strftime('%d/%m/%Y')} com **{dia_menos['TotalPacientes']} pacientes**")

# Rodapé
st.markdown("---")
st.markdown("**Criado por Vitor Cavalcante Gomes - vitor.cavalcante@amil.com.br - www.vitorgomes.tech**")
