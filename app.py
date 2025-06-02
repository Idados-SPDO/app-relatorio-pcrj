import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
import datetime
import zipfile

from utils.data_utils import mask_code, prepare_df, split_quartil_decreto
from utils.excel_utils import make_excel_with_headers
from utils.doc_utils import generate_full_doc, generate_price_only_doc

# ─────────── Configurações iniciais de Streamlit ───────────
st.set_page_config(page_title="SPDO - Automatizador de Relatórios PCRJ",page_icon="fgv_logo.png")
st.title("Automatizador de Relatórios PCRJ")
st.logo("logo_ibre.png")

# ─────────── Cálculo dinâmico de validade e nome dos arquivos ───────────
today = datetime.date.today()
end_date = today + datetime.timedelta(days=14)
today_str = today.strftime("%d/%m/%Y")
end_str = end_date.strftime("%d/%m/%Y")
validade = f"{today_str} a {end_str}"

current_year = today.year
current_quartil = pd.to_datetime(today).quarter
document_name = f"{current_year}Q{current_quartil}"

uploaded = st.file_uploader("Coloque o arquivo aqui", type="txt")
if uploaded is not None:
    data = uploaded.read().decode("latin-1")
    df = pd.read_csv(
        StringIO(data),
        sep="@",
        header=None,
        decimal=",",
        engine="python",
    )

    df = df.drop(df.columns[[7, 9]], axis=1)
    df.columns = [
        "Código do Item",
        "Dado1",
        "Dado2",
        "Dado3",
        "Ano",
        "Unidade",
        "Preço Atacado",
        "Preço Varejo",
        "Preço Praticado",
        "Produto",
        "Descrição",
    ]

    df["Código do Item"] = df["Código do Item"].apply(mask_code)

    quartil_df, decreto_df = split_quartil_decreto(df)

    tab_quartil, tab_decreto = st.tabs(["Quartil", "Decreto (Média)"])
    with tab_quartil:
        st.header("Quartil")
        st.dataframe(quartil_df)
    with tab_decreto:
        st.header("Decreto (Média)")
        st.dataframe(decreto_df)

    # ─────────── Preparar outputs formatados ───────────
    quartil_out = prepare_df(quartil_df)
    decreto_out = prepare_df(decreto_df)

    texto_cabecalho = (
        "Prefeitura da Cidade do Rio de Janeiro\n"
        "Tabela de Preços de Mercado de Gêneros Alimentícios\n"
        f"Validade: {validade}"
    )
    texto_subcabecalho = (
        "A tabela é referência para as aquisições realizadas pelos diversos órgãos do município "
        "e tem o preço dos itens apurado conforme estabelecido no Art. 1º do Decreto nº 51.017/2022 "
        "e alterações, que estabelece que o preço praticado pelo município e divulgado nesta tabela "
        "seja um preço intermediário entre os preços no mercado de atacado e de varejo."
    )

    # ─────────── Gerar bytes de cada arquivo ───────────
    bytes_excel_quartil = make_excel_with_headers(
        quartil_out,
        sheet="Quartil",
        text1=texto_cabecalho,
        text2=texto_subcabecalho,
        name="",
    )
    bytes_excel_quartil_praticado = make_excel_with_headers(
        quartil_out,
        sheet="Quartil - Praticado",
        text1=texto_cabecalho,
        text2=texto_subcabecalho,
        name="preço_praticado",
    )

    bytes_excel_decreto = make_excel_with_headers(
        decreto_out,
        sheet="Decreto",
        text1=texto_cabecalho,
        text2=texto_subcabecalho,
        name="",
    )
    bytes_excel_decreto_praticado = make_excel_with_headers(
        decreto_out,
        sheet="Decreto - Praticado",
        text1=texto_cabecalho,
        text2=texto_subcabecalho,
        name="preço_praticado",
    )
    bytes_docx_quartil = generate_full_doc(quartil_df, validade)
    bytes_docx_quartil_praticado = generate_price_only_doc(quartil_df, validade)

    bytes_docx_decreto = generate_full_doc(decreto_df, validade)
    bytes_docx_decreto_praticado = generate_price_only_doc(decreto_df, validade)

    # ─────────── Criar ZIP em memória ───────────
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"Quartil - GENALIM{document_name}.xlsx", bytes_excel_quartil)
        zf.writestr(
            f"Quartil - PRE_TAB_{document_name}.xlsx", bytes_excel_quartil_praticado
        )
        zf.writestr(f"Contrato - GENALIM{document_name}.xlsx", bytes_excel_decreto)
        zf.writestr(
            f"Contrato - PRE_TAB_{document_name}.xlsx", bytes_excel_decreto_praticado
        )
        zf.writestr(f"Quartil - GENALIM{document_name}.doc", bytes_docx_quartil)
        zf.writestr(
            f"Quartil - PRE_TAB_{document_name}.doc", bytes_docx_quartil_praticado
        )
        zf.writestr(f"Decreto - GENALIM{document_name}.doc", bytes_docx_decreto)
        zf.writestr(
            f"Decreto - PRE_TAB_{document_name}.doc", bytes_docx_decreto_praticado
        )

    zip_buffer.seek(0)

    # ─────────── Botão de download único para o ZIP ───────────
    st.download_button(
        label="📥 Baixar todos os Relatórios (Zip)",
        data=zip_buffer,
        file_name=f"Relatorios_{document_name}.zip",
        mime="application/zip",
    )
