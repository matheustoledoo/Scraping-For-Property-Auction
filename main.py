from fastapi import FastAPI, Form
from fastapi.responses import StreamingResponse, HTMLResponse
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

@app.get("/", response_class=HTMLResponse)
async def index():
    return """
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8"/>
      <meta name="viewport" content="width=device-width, initial-scale=1"/>
      <title>Leilões Scraper</title>
      <!-- Bootstrap 5 -->
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet"/>
      <!-- Google Fonts -->
      <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet"/>
      <!-- Material Icons -->
      <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet"/>
      <style>
        :root {
          --c-primary: #1f4068;
          --c-secondary: #3867d6;
          --c-accent: #34ace0;
          --c-bg: rgba(255,255,255,0.75);
          --c-blur: 16px;
          --radius: 1rem;
          --transition: 0.3s;
        }
        * { box-sizing: border-box; }
        body {
          margin: 0;
          font-family: 'Inter', sans-serif;
          background: linear-gradient(135deg, #182848 0%, #4b6cb7 100%);
          color: #333;
        }
        /* Navbar */
        .navbar {
          background: rgba(31,64,104,0.85);
          backdrop-filter: blur(var(--c-blur));
        }
        .navbar-brand, .nav-link {
          color: #fff !important;
          font-weight: 600;
        }
        .nav-link:hover, .nav-link.active {
          color: var(--c-accent) !important;
          text-shadow: 0 0 5px rgba(52,172,224,0.7);
        }
        /* Hero */
        .hero {
          text-align: center;
          color: #fff;
          padding: 4rem 1rem;
        }
        .hero h1 {
          font-size: 3rem;
          font-weight: 700;
          text-shadow: 2px 2px 10px rgba(0,0,0,0.3);
        }
        .hero p {
          font-size: 1.2rem;
          opacity: 0.9;
        }
        /* Glass card */
        .card-scraper {
          background: var(--c-bg);
          backdrop-filter: blur(var(--c-blur));
          border: none;
          border-radius: var(--radius);
          box-shadow: 0 8px 32px rgba(0,0,0,0.2);
          padding: 2rem;
          margin-top: -2rem;
          transition: transform var(--transition);
        }
        .card-scraper:hover {
          transform: translateY(-5px);
        }
        /* Form elements */
        .form-label { font-weight: 600; color: var(--c-primary); }
        .btn-group .btn-outline-secondary {
          transition: background var(--transition), color var(--transition);
        }
        .btn-group .btn-outline-secondary.active,
        .btn-group .btn-outline-secondary:hover {
          background: var(--c-secondary);
          color: #fff;
        }
        .btn-primary {
          background: var(--c-accent);
          border: none;
          font-weight: 600;
          padding: 0.75rem;
          transition: background var(--transition), box-shadow var(--transition);
        }
        .btn-primary:hover {
          background: #227093;
          box-shadow: 0 4px 15px rgba(34,112,147,0.6);
        }
        /* Inputs with icon */
        .input-icon {
          position: relative;
        }
        .input-icon .material-icons {
          position: absolute;
          top: 50%; left: 12px;
          transform: translateY(-50%);
          color: var(--c-secondary);
          font-size: 1.2rem;
        }
        .input-icon input, .input-icon select {
          padding-left: 2.5rem !important;
        }
        /* Advanced panel */
        #advancedFilters {
          background: #eef5ff;
          border-radius: 0.5rem;
          padding: 1rem;
          margin-bottom: 1.5rem;
        }
        /* Footer */
        footer {
          text-align: center;
          padding: 1rem 0;
          color: #eee;
          font-size: 0.9rem;
          margin-top: 3rem;
        }
      </style>
    </head>
    <body>

      <!-- Navbar -->
      <nav class="navbar navbar-expand-lg sticky-top">
        <div class="container">
          <a class="navbar-brand" href="/">Leilões Scraper</a>
          <button class="navbar-toggler" type="button"
                  data-bs-toggle="collapse" data-bs-target="#navMenu">
            <span class="navbar-toggler-icon"></span>
          </button>
          <div class="collapse navbar-collapse" id="navMenu">
            <ul class="navbar-nav ms-auto">
              <li class="nav-item"><a class="nav-link active" href="#">Dashboard</a></li>
              <li class="nav-item"><a class="nav-link" href="#">Leilões</a></li>
              <li class="nav-item"><a class="nav-link" href="#">Relatórios</a></li>
              <li class="nav-item"><a class="nav-link" href="#">Configurações</a></li>
              <li class="nav-item"><a class="nav-link" href="#">Sobre</a></li>
              <li class="nav-item"><a class="nav-link" href="#">Contato</a></li>
            </ul>
          </div>
        </div>
      </nav>

      <!-- Hero -->
      <section class="hero">
        <div class="container">
          <h1>Automatize sua Raspagem</h1>
          <p>Unifique todos <strong>os maiores sites de leilão de imóveis do Brasil</strong> em um só lugar.</p>
        </div>
      </section>

      <!-- Form -->
      <main class="container">
        <div class="row justify-content-center">
          <div class="col-lg-8">
            <div class="card card-scraper">
              <h2 class="text-center mb-4">Configurações de Raspagem</h2>
              <form method="post" action="/scrape" target="_blank"
                    onsubmit="this.querySelector('button').textContent='Carregando…';">

                <!-- Sites -->
                <div class="mb-4">
                  <label class="form-label">Sites a consultar</label>
                  <div class="btn-group w-100" role="group">
                    <input type="checkbox" class="btn-check" name="sites" id="mega" value="mega_leiloes" autocomplete="off" checked>
                    <label class="btn btn-outline-secondary" for="mega">Mega Leilões</label>

                    <input type="checkbox" class="btn-check" name="sites" id="viva" value="viva_leiloes" autocomplete="off">
                    <label class="btn btn-outline-secondary" for="viva">Viva Leilões</label>

                    <input type="checkbox" class="btn-check" name="sites" id="saraiva" value="saraiva_leiloes" autocomplete="off">
                    <label class="btn btn-outline-secondary" for="saraiva">Saraiva Leilões</label>
                  </div>
                  <div class="mt-2 text-end">
                    <button type="button" id="selectAllSites" class="btn btn-sm btn-outline-primary">
                      Adicionar Todos
                    </button>
                  </div>
                </div>

                <!-- Price & Filters -->
                <div id="advancedFilters">
                  <h5>Filtros Avançados</h5>
                  <div class="row gx-3 mb-3">
                    <div class="col-md-6 input-icon">
                      <label class="form-label">Valor Mínimo</label>
                      <input type="number" step="0.01" class="form-control" name="min_valor" id="min_valor" placeholder="Ex: 100.000"/>
                    </div>
                    <div class="col-md-6 input-icon">
                      <label class="form-label">Valor Máximo</label>
                      <input type="number" step="0.01" class="form-control" name="max_valor" id="max_valor" placeholder="Ex: 500.000"/>
                    </div>
                  </div>
                  <div class="row gx-3 mb-3">
                    <div class="col-md-6">
                      <label class="form-label">Status do Leilão</label>
                      <select class="form-select" name="status_filter">
                        <option value="">Todos</option>
                        <option value="aberto">Aberto</option>
                        <option value="futuro">Futuro</option>
                        <option value="encerrado">Encerrado</option>
                      </select>
                    </div>
                    <div class="col-md-6">
                      <label class="form-label">Formato de Saída</label>
                      <select class="form-select" name="output_format">
                        <option value="xlsx">Excel (.xlsx)</option>
                        <option value="csv">CSV (.csv)</option>
                      </select>
                    </div>
                  </div>
                </div>

                <!-- Tipo de Leilão -->
                <fieldset class="mb-4">
                  <legend class="form-label">Tipo de Leilão</legend>
                  <div class="form-check form-check-inline">
                    <input class="form-check-input" type="checkbox" name="leilao_types" id="jud" value="Judicial" checked>
                    <label class="form-check-label" for="jud">Judicial</label>
                  </div>
                  <div class="form-check form-check-inline">
                    <input class="form-check-input" type="checkbox" name="leilao_types" id="extr" value="Extrajudicial" checked>
                    <label class="form-check-label" for="extr">Extrajudicial</label>
                  </div>
                </fieldset>

                <!-- Resumo -->
                <div class="form-check form-switch mb-4">
                  <input class="form-check-input" type="checkbox" name="include_summary" id="sum">
                  <label class="form-check-label" for="sum">Incluir aba de Resumo</label>
                </div>

                <!-- Páginas -->
                <div class="mb-4 input-icon">
                  <label class="form-label">Páginas (ou 'todas')</label>
                  <input type="text" class="form-control" name="pages" id="pages" value="todas" required/>
                </div>

                <!-- Submit -->
                <div class="d-grid">
                  <button type="submit" class="btn btn-primary btn-lg shadow-lg">
                    <i class="material-icons align-middle">autorenew</i>
                    Iniciar Raspagem
                  </button>
                </div>
              </form>
            </div>
          </div>
        </div>
      </main>

      <!-- Footer -->
      <footer>
        &copy; 2025 Geourbe • Todos os direitos reservados
      </footer>

      <!-- Scripts -->
      <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
      <script>
        document.getElementById('selectAllSites').addEventListener('click', () => {
          document.querySelectorAll('input[name="sites"]').forEach(chk => chk.checked = true);
        });
      </script>
    </body>
    </html>
    """



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
        if leilao_types:
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

# uvicorn main:app --reload --host 127.0.0.1 --port 8000
