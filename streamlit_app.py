import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime  # ✅ Importado para gerar o nome com a data

# Lista de prefixos a excluir no campo "Nº do Item"
prefixos_excluir = [
    "AE5", "ARM", "COB", "DLM", "FDC", "HTS", "HS2", "LD6", "EL8", "MIC",
    "VFD", "RHT", "RMO", "RD3", "RD5", "RD2", "R30", "TRE", "TUB", "TIV",
    "SIP", "SPS", "ACN", "AEC", "AGA", "BRN", "CAL", "CBD", "CHT", "CMI",
    "PWD", "DF2", "PSK", "T30", "SC5", "MGL", "ME1", "ETL", "EQ1", "ATH",
    "AMD", "AKT", "AHN"
]

# Descrições que devem sempre ser mantidas
descricoes_preservadas = [
    "fortbio 1008", "fortbio 1009", "fortbio 1007", "fortbio 1010",
    "fortdoss 70"
]

# Códigos válidos com 6 dígitos + ponto
codigos_6_digitos_permitidos = [
    "973473.L1", "973514.L1", "977259.L1", "973515.L1", "148326.K6",
    "148478.L1", "222654.L1"
]

# Função que valida se o código é permitido
def item_valido(item):
    if not isinstance(item, str):
        return False
    if len(item) > 4 and item[3] == '.':
        return True
    if item in codigos_6_digitos_permitidos:
        return True
    return False

# Função principal de processamento
def processar_planilha(df):
    df = df.rename(columns=lambda x: x.strip())
    df["desc_lower"] = df["Descrição"].astype(str).str.lower()

    df_filtrado = df[
        df["Nº do Item"].apply(item_valido) &
        ~df["Nº do Item"].str[:3].isin(prefixos_excluir) &
        ~(
            df["desc_lower"].str.startswith("coladur") &
            ~df["desc_lower"].isin(descricoes_preservadas)
        ) &
        (df["Quantidade Disponível"] < 0)
    ].copy()

    # Limpa coluna descrição e inverte valor negativo
    df_filtrado.drop(columns=["desc_lower"], inplace=True)
    df_filtrado["Descrição"] = ""
    df_filtrado["Quantidade Disponível"] = df_filtrado["Quantidade Disponível"].abs()
    df_filtrado = df_filtrado.sort_values(by="Nº do Item")
    return df_filtrado

# Função auxiliar para converter o DataFrame para bytes
def converter_para_excel(df):
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)
    return buffer

# Interface do Streamlit
st.set_page_config(page_title="Filtro Automático de Planilhas", page_icon="📄")
st.title("📊 Filtro automático de planilhas")

arquivo = st.file_uploader("📂 Selecione a planilha Excel", type=["xlsx"])

if arquivo:
    try:
        df = pd.read_excel(arquivo)
        df_resultado = processar_planilha(df)
        st.success("✅ Planilha processada com sucesso!")

        excel_bytes = converter_para_excel(df_resultado)

        # ✅ Gera nome do arquivo com a data atual
        nome_arquivo = datetime.today().strftime('%Y-%m-%d') + ".xlsx"

        st.download_button("📥 Baixar planilha filtrada",
                           data=excel_bytes,
                           file_name=nome_arquivo,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"❌ Erro ao processar: {e}")
