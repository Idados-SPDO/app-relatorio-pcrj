import pandas as pd

def mask_code(val):
    """
    Formata o código numérico como 'AAAA.BB.CCC-DD'.
    """
    s = str(val).zfill(11)
    return f"{s[:4]}.{s[4:6]}.{s[6:9]}-{s[9:]}"


def prepare_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Recebe o DataFrame bruto e retorna aquele com colunas:
      ["Código do Item", "Descrição do Item", "Unidade",
       "Preço Atacado", "Preço Varejo", "Preço Praticado"],
    formatando preços e criando a coluna "Descrição do Item".
    """
    df = df.copy()

    def combine(r):
        prod = str(r["Produto"])
        desc = r["Descrição"]
        if pd.isna(desc) or str(desc).strip() in ("", "-"):
            return prod
        return f"{prod}\n{str(desc).strip()}"

    df["Descrição do Item"] = df.apply(combine, axis=1)

    # Preço Atacado
    df["Preço Atacado"] = pd.to_numeric(df["Preço Atacado"], errors="coerce")
    df["Preço Atacado"] = df["Preço Atacado"].apply(
        lambda x: f"{x:.2f}".replace(".", ",") if pd.notna(x) else ""
    )

    # Preço Varejo
    df["Preço Varejo"] = pd.to_numeric(df["Preço Varejo"], errors="coerce")
    df["Preço Varejo"] = df["Preço Varejo"].apply(
        lambda x: f"{x:.2f}".replace(".", ",") if pd.notna(x) else ""
    )

    # Preço Praticado
    df["Preço Praticado"] = pd.to_numeric(df["Preço Praticado"], errors="coerce")
    df["Preço Praticado"] = df["Preço Praticado"].apply(
        lambda x: f"{x:.2f}".replace(".", ",") if pd.notna(x) else ""
    )

    return df[
        [
            "Código do Item",
            "Descrição do Item",
            "Unidade",
            "Preço Atacado",
            "Preço Varejo",
            "Preço Praticado",
        ]
    ]


def split_quartil_decreto(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Recebe o DataFrame original (já com colunas renomeadas), aplica mask_code  
    e devolve dois DataFrames: um com itens que começam por "89" (quartil) e outro "90" (decreto).
    """
    quartil_df = df[df["Código do Item"].astype(str).str.startswith("89")].reset_index(drop=True)
    decreto_df = df[df["Código do Item"].astype(str).str.startswith("90")].reset_index(drop=True)
    return quartil_df, decreto_df
