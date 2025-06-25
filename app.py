import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
import datetime
import zipfile
import calendar
import snowflake.connector
from snowflake.snowpark import Session


from utils.data_utils import mask_code, prepare_df, split_quartil_decreto
from utils.excel_utils import make_excel_with_headers
from utils.doc_utils import generate_full_doc, generate_price_only_doc

# ─────────── Configurações iniciais de Streamlit ───────────
st.set_page_config(
    page_title="SPDO - Automatizador de Relatórios PCRJ",
    page_icon="fgv_logo.png",
    initial_sidebar_state="expanded",
    layout="wide"
)
st.title("Automatizador de Relatórios PCRJ")
st.logo("logo_ibre.png")

@st.cache_resource
def get_session():
    return Session.builder.configs(st.secrets["snowflake"]).create()

session = get_session()

def load_sazonalidade():
    sql = "SELECT * FROM BASES_SPDO.DB_APP_RELATORIO_PCRJ.TB_SAZONALIDADE"
    return session.sql(sql).to_pandas()


uploaded_sazonalidade = st.sidebar.file_uploader(
    "Atualizar tabela de sazonalidade:", 
    type=["xlsx"]
)

if uploaded_sazonalidade is not None:
    # Lê a primeira planilha, mantendo todos os campos como string
    # Esse dado vai para a snow:
    df = pd.read_excel(uploaded_sazonalidade, sheet_name=0, dtype=str)
    
    df.columns = [
         'COD_EXT', 'COD_FGV',
         'ESPEC_CLIENTE', 'UNIDADE',
         'ALTA_OFERTA', 'REGULAR', 'BAIXA_OFERTA'
    ]
    # Até aqui

    df_long = df.melt(
        id_vars=['COD_EXT','COD_FGV','ESPEC_CLIENTE','UNIDADE'],
        value_vars=['ALTA_OFERTA','REGULAR','BAIXA_OFERTA'],
        var_name='OFERTA',
        value_name='MESES'
    )

    df_long['MESES'] = (
        df_long['MESES']
        .str.replace('/', ', ')
        .str.split(r',\s*')
    )
    
    df_long = df_long.explode('MESES')
    df_long['MESES'] = df_long['MESES'].str.strip()
    df_pivot = (
        df_long
        .pivot_table(
            index=['COD_EXT','COD_FGV','ESPEC_CLIENTE','UNIDADE'],
            columns='MESES',
            values='OFERTA',
            aggfunc='first'     
        )
        .reset_index()
    )
    meses_ordem = [
        'Janeiro','Fevereiro','Março','Abril','Maio','Junho',
        'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'
    ]
    df_pivot = df_pivot[
        ['COD_EXT','COD_FGV','ESPEC_CLIENTE','UNIDADE']
        + [m for m in meses_ordem if m in df_pivot.columns]
    ]
    df_pivot = df_pivot.rename(columns=lambda c: c.upper() if c in meses_ordem else c)
    df_pivot = df_pivot.rename(columns={'MARÇO': 'MARCO'})

    snow_df = session.create_dataframe(df_pivot)
    snow_df.write.mode("overwrite").save_as_table("BASES_SPDO.DB_APP_RELATORIO_PCRJ.TB_SAZONALIDADE")
    st.success("Tabela de Sazonalidade atualizada com sucesso!")

df_snow = load_sazonalidade()
month_map = {
    1: 'JANEIRO',   2: 'FEVEREIRO', 3: 'MARCO',    4: 'ABRIL',
    5: 'MAIO',      6: 'JUNHO',     7: 'JULHO',    8: 'AGOSTO',
    9: 'SETEMBRO', 10: 'OUTUBRO',  11: 'NOVEMBRO',12: 'DEZEMBRO'
}
current_month_col = month_map[datetime.date.today().month]

# 2) seleciona apenas as colunas desejadas
cols_to_show = ['COD_EXT','COD_FGV','ESPEC_CLIENTE','UNIDADE', current_month_col]
df_mes_atual = df_snow[cols_to_show]

# 3) exibe no Streamlit
def show_item_warning(df):
    df = df[['COD_EXT', 'COD_FGV', 'ESPEC_CLIENTE']].astype(str)
    if df.empty:
        return "Sem registros."
    linhas = df.apply(lambda row: f"{row['ESPEC_CLIENTE'].capitalize()}", axis=1)
    # uma quebra simples entre linhas:
    return " - ".join(linhas.tolist())

st.write(f"### Sazonalidade para {current_month_col}")
def alert_custom(msg: str, bg: str = "#FFA500", text: str = "#000"):
    html = f"""
    <div style="
        background-color: {bg};
        color: {text};
        padding: 0.75em 1em;
        border-radius: 0.25em;
        margin-bottom: 1em;
    ">
      <strong>Atenção:</strong> {msg}
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
with col1.expander(f"Alta Oferta:"):
    df_alt = df_mes_atual[df_mes_atual[current_month_col] == "ALTA_OFERTA"]
    texto = show_item_warning(df_alt)
    if texto:
        alert_custom(texto, bg="#E89854", text="white")
    else:
        st.info("Sem registros.")

with col2.expander("Média Oferta:"):
    df_med = df_mes_atual[df_mes_atual[current_month_col] == "REGULAR"]
    texto = show_item_warning(df_med)
    if texto:
        alert_custom(texto, bg="#E8CF54", text="black")
    else:
        st.info("Sem registros.")
with col3.expander("Baixa Oferta:"):
    df_baix = df_mes_atual[df_mes_atual[current_month_col] == "BAIXA_OFERTA"]
    texto = show_item_warning(df_baix)
    if texto:
       alert_custom(texto, bg="#8CE854", text="black")
    else:
        st.info("Sem registros.")
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

uploaded = st.sidebar.file_uploader("Coloque o arquivo GENEROSCGM:", type="txt")
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

        
    # Voltar o ponteiro para o início do buffer
    zip_buffer.seek(0)

    # ─────────── Botão de download único para o ZIP com tudo dentro ───────────
    st.download_button(
        label="📥 Baixar todos os Relatórios (Zip)",
        data=zip_buffer,
        file_name=f"Relatorios_{document_name}.zip",
        mime="application/zip",
    )
