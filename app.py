import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
import datetime
import zipfile
import calendar

from utils.data_utils import mask_code, prepare_df, split_quartil_decreto
from utils.excel_utils import make_excel_with_headers
from utils.doc_utils import generate_full_doc, generate_price_only_doc
from utils.pdf_utils import convert_doc_to_pdf

# ─────────── Configurações iniciais de Streamlit ───────────
st.set_page_config(
    page_title="SPDO - Automatizador de Relatórios PCRJ",
    page_icon="fgv_logo.png"
)
st.title("Automatizador de Relatórios PCRJ")
st.logo("logo_ibre.png")

# ─────────── Cálculo dinâmico de validade e nome dos arquivos ───────────
today = datetime.date.today()
day = today.day
month = today.month
year = today.year

if day <= 15:
    # Segunda quinzena do mês atual: 16 até último dia do mês
    start_date = datetime.date(year, month, 16)
    last_day = calendar.monthrange(year, month)[1]
    end_date = datetime.date(year, month, last_day)
else:
    # Primeira quinzena do próximo mês: 1 até 15
    if month < 12:
        next_month = month + 1
        next_year = year
    else:
        next_month = 1
        next_year = year + 1
    start_date = datetime.date(next_year, next_month, 1)
    end_date = datetime.date(next_year, next_month, 15)

# Formatar “dd/mm/YYYY a dd/mm/YYYY”
today_str = start_date.strftime("%d/%m/%Y")
end_str = end_date.strftime("%d/%m/%Y")
validade = f"{today_str} a {end_str}"

current_year = today.year
current_quartil = pd.to_datetime(today).quarter
document_name = f"{current_year}Q{current_quartil}"

uploaded = st.file_uploader("Coloque o arquivo aqui", type="txt")
if uploaded is not None:
    # ─────────── Ler e tratar o TXT enviado ───────────
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

    # ─────────── Preparar outputs para Excel e DOCX ───────────
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

    # ─────────── Gerar bytes de cada arquivo Excel (.xlsx) ───────────
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

    # ─────────── Gerar bytes de cada arquivo DOCX (.docx) ───────────
    bytes_docx_quartil = generate_full_doc(quartil_df, validade)
    bytes_docx_quartil_praticado = generate_price_only_doc(quartil_df, validade)

    bytes_docx_decreto = generate_full_doc(decreto_df, validade)
    bytes_docx_decreto_praticado = generate_price_only_doc(decreto_df, validade)

    # ─────────── Converter cada DOCX em PDF ───────────
    pdf_bytes_quartil = convert_doc_to_pdf(bytes_docx_quartil, f"Quartil_{document_name}")
    pdf_bytes_quartil_praticado = convert_doc_to_pdf(bytes_docx_quartil_praticado, f"QuartilPraticado_{document_name}")

    pdf_bytes_contrato = convert_doc_to_pdf(bytes_docx_decreto, f"Contrato_{document_name}")
    pdf_bytes_contrato_praticado = convert_doc_to_pdf(bytes_docx_decreto_praticado, f"ContratoPraticado_{document_name}")

    # Se qualquer PDF falhar, mostra aviso (mas continua gerando o ZIP)
    if None in (pdf_bytes_quartil, pdf_bytes_quartil_praticado, pdf_bytes_contrato, pdf_bytes_contrato_praticado):
        st.error("Pelo menos um PDF não pôde ser gerado. Verifique se o Word está instalado corretamente.")

    # ─────────── Criar ZIP em memória (BytesIO) e incluir todos os arquivos ───────────
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        # 1) Excel (.xlsx)
        zf.writestr(f"Quartil - GENALIM_{document_name}.xlsx", bytes_excel_quartil)
        zf.writestr(
            f"Quartil - PRE_TAB_{document_name}.xlsx",
            bytes_excel_quartil_praticado,
        )
        zf.writestr(f"Contrato - GENALIM_{document_name}.xlsx", bytes_excel_decreto)
        zf.writestr(
            f"Contrato - PRE_TAB_{document_name}.xlsx",
            bytes_excel_decreto_praticado,
        )

        # 2) DOCX (.docx)
        zf.writestr(f"Quartil - GENALIM_{document_name}.docx", bytes_docx_quartil)
        zf.writestr(
            f"Quartil - PRE_TAB_{document_name}.docx",
            bytes_docx_quartil_praticado,
        )
        zf.writestr(f"Contrato - GENALIM_{document_name}.docx", bytes_docx_decreto)
        zf.writestr(
            f"Contrato - PRE_TAB_{document_name}.docx",
            bytes_docx_decreto_praticado,
        )

        # 3) PDF (se gerados com sucesso)
        if pdf_bytes_quartil is not None:
            zf.writestr(f"Quartil - GENALIM_{document_name}.pdf", pdf_bytes_quartil)
        else:
            zf.writestr(
                f"Aviso_Conversao_PDF_Quartil_{document_name}.txt",
                "Não foi possível gerar o PDF a partir do Quartil DOCX.",
            )

        if pdf_bytes_quartil_praticado is not None:
            zf.writestr(f"Quartil - PRE_TAB_{document_name}.pdf", pdf_bytes_quartil_praticado)
        else:
            zf.writestr(
                f"Aviso_Conversao_PDF_QuartilPraticado_{document_name}.txt",
                "Não foi possível gerar o PDF a partir do Quartil (Praticado) DOCX.",
            )

        if pdf_bytes_contrato is not None:
            zf.writestr(f"Contrato - GENALIM_{document_name}.pdf", pdf_bytes_contrato)
        else:
            zf.writestr(
                f"Aviso_Conversao_PDF_Contrato_{document_name}.txt",
                "Não foi possível gerar o PDF a partir do Contrato DOCX.",
            )

        if pdf_bytes_contrato_praticado is not None:
            zf.writestr(f"Contrato - PRE_TAB_{document_name}.pdf", pdf_bytes_contrato_praticado)
        else:
            zf.writestr(
                f"Aviso_Conversao_PDF_ContratoPraticado_{document_name}.txt",
                "Não foi possível gerar o PDF a partir do Contrato (Praticado) DOCX.",
            )

    # Voltar o ponteiro para o início do buffer
    zip_buffer.seek(0)

    # ─────────── Botão de download único para o ZIP com tudo dentro ───────────
    st.download_button(
        label="📥 Baixar todos os Relatórios (Zip)",
        data=zip_buffer,
        file_name=f"Relatorios_{document_name}.zip",
        mime="application/zip",
    )
