# relatorios_pcrj/app.py

import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
import datetime
import zipfile

from utils.data_utils import mask_code, prepare_df, split_quartil_contrato
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

# ─────────── Upload do arquivo TXT ───────────
uploaded = st.file_uploader("Coloque o arquivo aqui", type="txt")
if uploaded is not None:
    # ───────────── Leitura e tratamento inicial ─────────────
    data = uploaded.read().decode("latin-1")
    df = pd.read_csv(
        StringIO(data),
        sep="@",
        header=None,
        decimal=",",
        engine="python",
    )

    # Remover colunas indesejadas (índices 7 e 9)
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

    # Formatar "Código do Item"
    df["Código do Item"] = df["Código do Item"].apply(mask_code)

    # Separar em quartil e contrato
    quartil_df, contrato_df = split_quartil_contrato(df)

    # Exibir tabelas no Streamlit
    tab_quartil, tab_contrato = st.tabs(["Quartil", "Contrato Média"])
    with tab_quartil:
        st.header("Quartil")
        st.dataframe(quartil_df)
    with tab_contrato:
        st.header("Contrato Média")
        st.dataframe(contrato_df)

    # ─────────── Preparar outputs formatados ───────────
    quartil_out = prepare_df(quartil_df)
    contrato_out = prepare_df(contrato_df)

    # Cabeçalhos para Excel/Word
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
    # Excel Quartil (6 colunas)
    bytes_excel_quartil = make_excel_with_headers(
        quartil_out,
        sheet="Quartil",
        text1=texto_cabecalho,
        text2=texto_subcabecalho,
        name="",
    )
    # Excel Quartil Praticado (5 colunas + Nº)
    bytes_excel_quartil_praticado = make_excel_with_headers(
        quartil_out,
        sheet="Quartil - praticado",
        text1=texto_cabecalho,
        text2=texto_subcabecalho,
        name="preço_praticado",
    )

    # Excel Contrato (6 colunas)
    bytes_excel_contrato = make_excel_with_headers(
        contrato_out,
        sheet="Contrato",
        text1=texto_cabecalho,
        text2=texto_subcabecalho,
        name="",
    )
    # Excel Contrato Praticado (5 colunas + Nº)
    bytes_excel_contrato_praticado = make_excel_with_headers(
        contrato_out,
        sheet="Contrato - praticado",
        text1=texto_cabecalho,
        text2=texto_subcabecalho,
        name="preço_praticado",
    )

    # Word Quartil (todos os campos)
    bytes_docx_quartil = generate_full_doc(quartil_df, validade)
    # Word Quartil Preço Praticado (só 4 colunas + Nº)
    bytes_docx_quartil_praticado = generate_price_only_doc(quartil_df, validade)

    # Word Contrato (todos os campos)
    bytes_docx_contrato = generate_full_doc(contrato_df, validade)
    # Word Contrato Preço Praticado (só 4 colunas + Nº)
    bytes_docx_contrato_praticado = generate_price_only_doc(contrato_df, validade)

    # ─────────── Criar ZIP em memória ───────────
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        # Excel
        zf.writestr(f"Quartil - GENALIM{document_name}.xlsx", bytes_excel_quartil)
        zf.writestr(
            f"Quartil - PRE_TAB_{document_name}.xlsx", bytes_excel_quartil_praticado
        )
        zf.writestr(f"Contrato - GENALIM{document_name}.xlsx", bytes_excel_contrato)
        zf.writestr(
            f"Contrato - PRE_TAB_{document_name}.xlsx", bytes_excel_contrato_praticado
        )
        # Word
        zf.writestr(f"Quartil - GENALIM{document_name}.doc", bytes_docx_quartil)
        zf.writestr(
            f"Quartil - PRE_TAB_{document_name}.doc", bytes_docx_quartil_praticado
        )
        zf.writestr(f"Contrato - GENALIM{document_name}.doc", bytes_docx_contrato)
        zf.writestr(
            f"Contrato - PRE_TAB_{document_name}.doc", bytes_docx_contrato_praticado
        )

    zip_buffer.seek(0)

    # ─────────── Botão de download único para o ZIP ───────────
    st.download_button(
        label="📥 Baixar todos os Relatórios (Zip)",
        data=zip_buffer,
        file_name=f"Relatorios_{document_name}.zip",
        mime="application/zip",
    )
