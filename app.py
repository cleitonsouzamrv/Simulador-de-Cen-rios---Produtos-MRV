import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
from PIL import Image

# Configuração inicial da página
st.set_page_config(
    page_title="Produtos MRV",
    page_icon="logo_mrv_light.png",
    layout="wide"
)

# Exibir logo no topo
logo = Image.open("logo_mrv_light.png")
st.image(logo, width=240)
st.title("Simulador de Cenários - Produtos MRV")

SAVE_DIR = "cenarios_salvos"
os.makedirs(SAVE_DIR, exist_ok=True)

uploaded_file = st.file_uploader("Faça upload da planilha 'Base_Valores'", type=["xlsx"])
if uploaded_file:
    df_atributos = pd.read_excel(uploaded_file, sheet_name=None)
    df_opcoes = df_atributos[list(df_atributos.keys())[0]]
    df_classificacao = df_atributos.get("Classificacao_Produto", pd.DataFrame())
    st.success("Planilha carregada com sucesso!")

    nome_cenario = st.text_input("Digite um nome para o cenário")

    custos_selecionados = []
    subtotais_dict = {}
    st.header("Selecione os atributos")

    for dimensao in df_opcoes["Dimensão"].unique():
        df_dim = df_opcoes[df_opcoes["Dimensão"] == dimensao]
        st.subheader(dimensao)
        subtotal = 0.0
        colunas = st.columns(4)
        idx_col = 0

        for atributo in df_dim["Atributos"].unique():
            df_attr = df_dim[df_dim["Atributos"] == atributo]
            modo = df_attr["Streamlit"].iloc[0]
            opcoes = df_attr.apply(lambda row: f"{row['Tipo']} (R$ {row['Custo/UH']})", axis=1).tolist()

            with colunas[idx_col]:
                if modo == "Lista Suspensa":
                    escolha_formatada = st.selectbox(f"{atributo}", opcoes, key=f"select_{atributo}")
                    idx = opcoes.index(escolha_formatada)
                    tipo = df_attr.iloc[idx]["Tipo"]
                    custo = df_attr.iloc[idx]["Custo/UH"]
                    custos_selecionados.append({"Dimensão": dimensao, "Atributo": atributo, "Tipo": tipo, "Custo/UH": custo})
                    custo_float = float(str(custo).replace("R$", "").replace(".", "").replace(",", "."))
                    subtotal += custo_float

                elif modo == "Checkbox":
                    label = f"{atributo} (R$ {df_attr.iloc[0]['Custo/UH']})"
                    if st.checkbox(label, key=f"check_{atributo}"):
                        custo = df_attr.iloc[0]["Custo/UH"]
                        custos_selecionados.append({"Dimensão": dimensao, "Atributo": atributo, "Tipo": "Sim", "Custo/UH": custo})
                        custo_float = float(str(custo).replace("R$", "").replace(".", "").replace(",", "."))
                        subtotal += custo_float

            idx_col = (idx_col + 1) % 4

        subtotais_dict[dimensao] = subtotal
        st.markdown(f"**Subtotal {dimensao}: R$ {subtotal:,.2f}**")

    if custos_selecionados:
        df_resultado = pd.DataFrame(custos_selecionados)
        st.subheader("Resumo do Cenário Selecionado")
        st.dataframe(df_resultado)

        df_resultado["Custo/UH"] = df_resultado["Custo/UH"].replace({"R\$": "", ".": "", ",": "."}, regex=True).astype(float)
        total = df_resultado["Custo/UH"].sum()
        st.metric("Custo Total do Cenário", f"R$ {total:,.2f}")

        produto = "NÃO CLASSIFICADO"
        produto_cor = "gray"
        if not df_classificacao.empty:
            for _, row in df_classificacao.iterrows():
                minimo = float(str(row[0]).replace("R$", "").replace(".", "").replace(",", "."))
                maximo = float(str(row[1]).replace("R$", "").replace(".", "").replace(",", "."))
                if minimo <= total <= maximo:
                    produto = row[2]
                    break

        if produto == "ESSENCIAL":
            produto_cor = "#FF8D03"
        elif produto == "ECO":
            produto_cor = "#00B050"
        elif produto == "BIO":
            produto_cor = "#B14FDA"

        st.markdown(f"<h4>Classificação do Produto: <span style='color:{produto_cor}'>{produto}</span></h4>", unsafe_allow_html=True)

        st.markdown("**Tabela de Classificação**")
        st.dataframe(df_classificacao)

        st.subheader("Distribuição de Custos por Dimensão")
        df_subtotais = pd.DataFrame(list(subtotais_dict.items()), columns=["Dimensão", "Subtotal"])
        df_subtotais = df_subtotais.sort_values("Subtotal", ascending=False)
        fig = px.bar(df_subtotais, x="Subtotal", y="Dimensão", orientation="h",
                     title="Custo por Dimensão", labels={"Subtotal": "Custo (R$)", "Dimensão": "Dimensão"},
                     text="Subtotal", color_discrete_sequence=["#006B3F"])
        fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
        st.plotly_chart(fig)

        if nome_cenario.strip() == "":
            st.warning("Você deve inserir um nome para o cenário antes de salvar.")
        else:
            if st.button("Salvar cenário"):
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                file_name = f"cenario_{nome_cenario}_{ts}.xlsx"
                file_path = os.path.join(SAVE_DIR, file_name)

                df_resultado["Classificação"] = produto
                df_resultado["Cenário"] = file_name.replace(".xlsx", "")

                with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name="Base Consolidada")
                    pd.DataFrame({
                        "Nome do Cenário": [nome_cenario],
                        "Data": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                        "Classificação": [produto],
                        "Custo Total": [total]
                    }).to_excel(writer, sheet_name="Metadados", index=False)
                    df_classificacao.to_excel(writer, index=False, sheet_name="Classificacao_Produto")

                st.success("Cenário salvo com sucesso!")
                with open(file_path, "rb") as f:
                    st.download_button("📥 Baixar planilha do cenário", f,
                                       file_name=file_name,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("Selecionar cenários salvos para exportação")

    arquivos_validos = []
    for file in os.listdir(SAVE_DIR):
        if file.endswith(".xlsx"):
            path = os.path.join(SAVE_DIR, file)
            try:
                xl = pd.ExcelFile(path)
                if "Base Consolidada" in xl.sheet_names:
                    arquivos_validos.append(file)
            except:
                continue

    if arquivos_validos:
        selected_files = st.multiselect(
            "Escolha os cenários que deseja exportar:",
            arquivos_validos,
            default=[]
        )

        if selected_files:
            df_all = []
            for file in selected_files:
                path = os.path.join(SAVE_DIR, file)
                try:
                    df = pd.read_excel(path, sheet_name="Base Consolidada")
                    df_all.append(df)
                except:
                    continue

            if df_all:
                df_final = pd.concat(df_all, ignore_index=True)
                with pd.ExcelWriter("cenarios_exportados.xlsx", engine="xlsxwriter") as writer:
                    df_final.to_excel(writer, index=False, sheet_name="Base Consolidada")

                with open("cenarios_exportados.xlsx", "rb") as f:
                    st.download_button("📦 Baixar base consolidada selecionada", f,
                                       file_name="cenarios_exportados.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("Selecione pelo menos um cenário para exportar.")
    else:
        st.info("Nenhum cenário com estrutura válida foi encontrado.")
