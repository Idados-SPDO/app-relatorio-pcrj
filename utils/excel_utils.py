import pandas as pd
from io import BytesIO

def make_excel_with_headers(
    df_export: pd.DataFrame,
    sheet: str,
    text1: str,
    text2: str,
    name: str = "",
) -> bytes:
    """
    Gera um arquivo Excel em bytes, mescla dois cabeçalhos (text1 e text2) e:
    - se name == "preço_praticado", usa apenas 4 colunas + insere coluna "Nº"
    - caso contrário, escreve as 6 colunas originais.
    Retorna os bytes prontos para escrita.
    """
    buf = BytesIO()

    if name == "preço_praticado":
        df4 = df_export[
            [
                "Código do Item",
                "Descrição do Item",
                "Unidade",
                "Preço Praticado",
            ]
        ].copy()
        df4.insert(0, "Nº", range(1, len(df4) + 1))
        df4.rename(columns={"Preço Praticado": "Preço (em R$)"}, inplace=True)

        merge_range_header1 = "A1:E1"
        merge_range_header2 = "A2:E2"
        widths = [5, 15, 60, 12, 12]
        df_to_write = df4
    else:
        df6 = df_export.copy()
        merge_range_header1 = "A1:F1"
        merge_range_header2 = "A2:F2"
        widths = [15, 60, 10, 12, 12, 12]
        df_to_write = df6

    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df_to_write.to_excel(writer, sheet_name=sheet, index=False, startrow=2)
        wb = writer.book
        ws = writer.sheets[sheet]

        fmt1 = wb.add_format(
            {
                "align": "center",
                "valign": "vcenter",
                "bold": True,
                "text_wrap": True,
            }
        )
        fmt2 = wb.add_format(
            {
                "align": "center",
                "valign": "vcenter",
                "bold": False,
                "text_wrap": True,
            }
        )
        center_fmt = wb.add_format(
            {"align": "center", "valign": "vcenter", "text_wrap": True}
        )
        left_fmt = wb.add_format(
            {"align": "left", "valign": "vcenter", "text_wrap": True}
        )

        num_fmt = wb.add_format({
            "num_format": "0.00", 
            "align": "center", 
            "valign": "vcenter", 
            "text_wrap": True
        })

        ws.merge_range(merge_range_header1, text1, fmt1)
        ws.merge_range(merge_range_header2, text2, fmt2)
        ws.set_row(0, 50)
        ws.set_row(1, 80)

        for idx, w in enumerate(widths):
            if name == "preço_praticado":
                # df4 tem colunas: Nº(0), Código(1), Desc(2), Unidade(3), Preço(4)
                if idx == 2:
                    fmt = left_fmt
                elif idx == 4:
                    fmt = num_fmt
                else:
                    fmt = center_fmt
            else:
                # df6 tem colunas: Código(0), Desc(1), Unidade(2), Atacado(3), Varejo(4), Praticado(5)
                if idx == 1:
                    fmt = left_fmt
                elif idx in (3, 4, 5):
                    fmt = num_fmt
                else:
                    fmt = center_fmt

            ws.set_column(idx, idx, w, fmt)

        ws.set_default_row(60)

    buf.seek(0)
    return buf.getvalue()
