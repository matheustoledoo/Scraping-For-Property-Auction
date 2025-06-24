from fastapi import FastAPI, Form
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.responses import FileResponse
import io
import pandas as pd
from scrapers.megaleiloes import run as scrape_mega
from scrapers.vivaleiloes import run as scrape_viva
from scrapers.saraivaleiloes import run as scrape_saraiva

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.table import Table, TableStyleInfo

app = FastAPI()

@app.get("/", response_class=FileResponse)
async def index():
    return FileResponse("frontend/index.html")

@app.post("/scrape")
async def scrape(
    sites: list[str] = Form(...),
    pages: str      = Form(...),
    min_valor: str  = Form(None),
    max_valor: str  = Form(None),
    leilao_types: list[str] = Form(None),
    include_summary: bool    = Form(False)
):
    pages_int = -1 if pages.lower() == "todas" else int(pages)
    min_v = float(min_valor) if min_valor else None
    max_v = float(max_valor) if max_valor else None

    # coleta
    dfs: dict[str, pd.DataFrame] = {}
    if "mega_leiloes"  in sites:
        dfs["Mega Leilões"]  = scrape_mega(pages_int)
    if "viva_leiloes"  in sites:
        dfs["Viva Leilões"]  = scrape_viva(pages_int)
    if "saraiva_leiloes" in sites:
        dfs["Saraiva Leilões"] = scrape_saraiva(pages_int)

    # filtros
    for name, df in dfs.items():
        if not df.empty:
            vals = (
                df["valor_imovel"]
                .str.replace(r"[^\d,]", "", regex=True)
                .str.replace(",", ".")
                .astype(float)
            )
            df["_valor_num"] = vals
            if min_v is not None:
                df = df[vals >= min_v]
            if max_v is not None:
                df = df[vals <= max_v]

        if leilao_types and "tipo_leilao" in df.columns:
            df = df[df["tipo_leilao"].isin(leilao_types)]

        dfs[name] = df

    # gera Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wb = writer.book

        for sheet, df in dfs.items():
            df.drop(columns=["_valor_num"], errors="ignore") \
              .to_excel(writer, index=False, sheet_name=sheet)
            ws = writer.sheets[sheet]
            last_col, last_row = ws.max_column, ws.max_row

            # tabela
            tbl = Table(
                displayName=sheet.replace(" ","") + "_Tbl",
                ref=f"A1:{chr(64+last_col)}{last_row}"
            )
            tbl.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium9", showRowStripes=True
            )
            ws.add_table(tbl)

            # hyperlinks
            link_cols = {"link","edital_leilao","laudo_avaliacao","matricula"}
            for idx, col_cells in enumerate(ws.iter_cols(min_row=2, max_row=last_row), start=1):
                hdr = ws.cell(row=1, column=idx).value
                if hdr in link_cols:
                    for c in col_cells:
                        if isinstance(c.value, str) and c.value.startswith("http"):
                            c.hyperlink = c.value
                            c.font = Font(color="0000FF", underline="single")

            # cond. formatting
            if "_valor_num" in df.columns and last_row >= 2:
                avg = df["_valor_num"].mean()
                ci  = df.columns.get_loc("_valor_num") + 1
                rng = f"{chr(64+ci)}2:{chr(64+ci)}{last_row}"
                ws.conditional_formatting.add(
                    rng,
                    CellIsRule(
                        operator="greaterThan",
                        formula=[str(avg)],
                        fill=PatternFill("solid", fgColor="FFC7CE")
                    )
                )

            ws.freeze_panes = "A2"
            for col in ws.columns:
                w = min(max(len(str(c.value)) for c in col if c.value) + 5, 30)
                ws.column_dimensions[col[0].column_letter].width = w

        # aba Resumo
        if include_summary:
            s = wb.create_sheet("Resumo")
            s.sheet_view.showGridLines = False

            combined = pd.concat(dfs.values(), ignore_index=True)
            vals = combined["_valor_num"]

            # ---- Título ----
            s.merge_cells("A1:C1")
            s["A1"] = "Resumo da Raspagem"
            s["A1"].font = Font(bold=True, size=16)
            s["A1"].alignment = Alignment(horizontal="center", vertical="center")
            s.row_dimensions[1].height = 24

            # ---- Largura das colunas ----
            s.column_dimensions["A"].width = 30
            s.column_dimensions["B"].width = 20
            s.column_dimensions["C"].width = 15

            # ---- Estatísticas ----
            stats = [
                ("Total de imóveis", len(combined)),
                ("Valor médio (R$)", f"{vals.mean():,.2f}"),
                ("Valor mínimo (R$)", f"{vals.min():,.2f}"),
                ("Valor máximo (R$)", f"{vals.max():,.2f}"),
                ("Mediana (R$)", f"{vals.median():,.2f}")
            ]
            for i, (lbl, val) in enumerate(stats, start=2):
                s[f"A{i}"] = lbl
                s[f"B{i}"] = val
                s[f"A{i}"].font = Font(bold=True)
                s[f"A{i}"].alignment = Alignment(horizontal="right")
                s[f"B{i}"].alignment = Alignment(horizontal="left")

            # ---- Cabeçalho da tabela de contagem por tipo ----
            start = len(stats) + 4
            s[f"A{start}"] = "Tipo"
            s[f"B{start}"] = "Quantidade"
            for col_cell in (f"A{start}", f"B{start}"):
                cell = s[col_cell]
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="4F81BD")
                cell.alignment = Alignment(horizontal="center")

            # ---- Linhas de contagem por tipo ----
            counts = combined["tipo_leilao"].value_counts()
            for j, (t, cnt) in enumerate(counts.items(), start=start + 1):
                s[f"A{j}"] = t
                s[f"B{j}"] = cnt
                s[f"A{j}"].alignment = Alignment(horizontal="left")
                s[f"B{j}"].alignment = Alignment(horizontal="center")

            # ---- Gráfico de barras ----
            chart = BarChart()
            chart.title = "Total por Tipo"
            chart.style = 10
            chart.width = 12
            chart.height = 6
            data = Reference(s, min_col=2, min_row=start, max_row=start + len(counts))
            cats = Reference(s, min_col=1, min_row=start + 1, max_row=start + len(counts))
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            for k, ser in enumerate(chart.series):
                ser.graphicalProperties.solidFill = ["4F81BD", "C0504D"][k % 2]
                ser.dLbls = DataLabelList()
                ser.dLbls.showVal = True
            s.add_chart(chart, "D2")

    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="leiloes.xlsx"'}
    )

# uvicorn main:app --reload --host 127.0.0.1 --port 8080
