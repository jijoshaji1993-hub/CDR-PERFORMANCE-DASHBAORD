import os
import pandas as pd
from dash import Dash, dcc, html, dash_table, Input, Output, State, ctx
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import io
import base64
import zipfile
from tempfile import TemporaryDirectory

# === Load & preprocess data ===

FILE_PATH = "Result.xlsx"
xls = pd.ExcelFile(FILE_PATH)
sheets = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

# Clean column names
for sn in sheets:
    sheets[sn].columns = sheets[sn].columns.str.strip()

# Helper: format date columns
def format_date_columns(df):
    df2 = df.copy()
    for col in df2.columns:
        if 'date' in col.lower():
            try:
                df2[col] = pd.to_datetime(df2[col], errors='coerce')
                df2[col] = df2[col].dt.strftime('%d-%m-%Y')
                df2[col] = df2[col].fillna(df[col].astype(str))
            except Exception as e:
                print(f"Error formatting date column {col}: {e}")
    return df2

# Helper: format display values (percentages, rounding)
def format_display_values(df):
    df2 = df.copy()
    percent_cols = [
        "SHORT CALL%", "IVRS ANS%", "ANS%", "CMS Aband%",
        "SL%", "SL %", "Entry Level %", "Second Level %", "Third Level %"
    ]
    round_cols = ["AHT", "IVRS AHT", "AVG WAIT TIME", "Answered", "IVRS_OFFERED", "NET OFFERED"]

    for col in df2.columns:
        col_clean = col.strip()
        if col_clean in percent_cols:
            try:
                df2[col] = pd.to_numeric(df2[col], errors='coerce') * 100
                df2[col] = df2[col].map(lambda x: f"{x:.2f}%" if pd.notna(x) else "N/A")
            except Exception as e:
                print(f"Error formatting percent column {col}: {e}")
        elif col_clean in round_cols:
            try:
                df2[col] = pd.to_numeric(df2[col], errors='coerce').round(0).astype("Int64").astype(str)
            except Exception as e:
                print(f"Error rounding column {col}: {e}")
    return df2

# Apply formatting
for sn in list(sheets.keys()):
    sheets[sn] = format_date_columns(sheets[sn])
    sheets[sn] = format_display_values(sheets[sn])

# === Helpers to get required rows / date filtering ===

def get_mtd_row(df):
    if df.empty or "Date" not in df.columns:
        return None
    df2 = df.copy()
    df2["Date_str"] = df2["Date"].astype(str).str.strip().str.upper()
    mtd = df2[df2["Date_str"] == "MTD"]
    if mtd.empty:
        return None
    return mtd.iloc[0].drop(labels=["Date_str"], errors="ignore")

def get_date_rows(df, n_days=7):
    """Return last n_days rows excluding MTD, sorted descending by date."""
    if df.empty or "Date" not in df.columns:
        return pd.DataFrame()
    df2 = df.copy()
    df2["Date_str"] = df2["Date"].astype(str).str.strip().str.upper()
    df2 = df2[df2["Date_str"] != "MTD"]
    df2["Date_dt"] = pd.to_datetime(df2["Date_str"], dayfirst=True, errors='coerce')
    df2 = df2.dropna(subset=["Date_dt"])
    df2 = df2.sort_values("Date_dt", ascending=False)
    return df2.head(n_days)

def get_today_and_day1(df):
    recent = get_date_rows(df, n_days=2)
    if recent.empty:
        return None, None
    recent = recent.reset_index(drop=True)
    row0 = recent.iloc[0].drop(labels=["Date_str", "Date_dt"], errors="ignore")
    if len(recent) >= 2:
        row1 = recent.iloc[1].drop(labels=["Date_str", "Date_dt"], errors="ignore")
    else:
        row1 = None
    return row0, row1

# Precompute Dashboard metrics
dashboard_df = sheets.get("Dashboard", pd.DataFrame())
dashboard_mtd = get_mtd_row(dashboard_df)
dashboard_last7 = get_date_rows(dashboard_df, n_days=7)
dashboard_today, dashboard_day1 = get_today_and_day1(dashboard_df)

kerala_mtd = get_mtd_row(sheets.get("Kerala", pd.DataFrame()))
tamilnadu_mtd = get_mtd_row(sheets.get("Tamilnadu", pd.DataFrame()))
chennai_mtd = get_mtd_row(sheets.get("Chennai", pd.DataFrame()))

hourly_df = sheets.get("Hourly Performance", pd.DataFrame())

# === KPI Card generation ===

def kpi_card(label, value, is_percent=False, target=None, inverse=False):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        display_val = "N/A"
        color = "black"
    else:
        raw = value
        val_num = None
        if isinstance(raw, str) and raw.endswith("%"):
            try:
                val_num = float(raw.rstrip("%").strip())
            except:
                pass
        else:
            try:
                val_num = float(str(raw).replace(",", "").strip())
            except:
                pass

        if is_percent:
            if val_num is not None:
                display_val = f"{val_num:.2f}%"
            else:
                display_val = str(raw)
        else:
            if label.strip().upper() in ["IVRS_OFFERED", "NET OFFERED", "ANSWERED", "AHT"]:
                try:
                    display_val = f"{int(round(float(str(raw).replace('%','').replace(',','').strip()))):,}"
                except:
                    display_val = str(raw)
            else:
                display_val = str(raw)

        color = "black"
        if target is not None and val_num is not None:
            if inverse:
                color = "green" if val_num >= target else "red"
            else:
                color = "green" if val_num <= target else "red"

    return dbc.Card(
        dbc.CardBody([
            html.H6(label, className="text-muted", style={"fontWeight":"600"}),
            html.H4(display_val, style={"color": color, "fontWeight":"700"})
        ]),
        style={"width":"12rem","margin":"5px","textAlign":"center","boxShadow":"0 0 5px rgba(0,0,0,0.1)"}
    )

# === Analytics / Graph/Table Helpers ===

def trend_graph_last7():
    if dashboard_last7.empty:
        return html.Div("No sufficient historical data for trend graph.")
    df = dashboard_last7.copy()
    metrics = {}
    for m in ["SL%", "ANS%", "AHT"]:
        if m in df.columns:
            metrics[m] = df[m].map(
                lambda v: float(str(v).rstrip('%').replace(',','').strip())
                if isinstance(v, str) else None
            )
    df["Date_dt"] = pd.to_datetime(df["Date"].astype(str).str.strip(), dayfirst=True, errors='coerce')
    fig = go.Figure()
    for m, series in metrics.items():
        fig.add_trace(go.Scatter(
            x=df["Date_dt"],
            y=series,
            mode="lines+markers",
            name=m
        ))
    fig.update_layout(
        title="Trends (Last 7 Days): SL%, ANS%, AHT",
        xaxis_title="Date",
        yaxis_title="Value",
        template="plotly_white",
        hovermode="x unified",
        margin=dict(l=40, r=40, t=60, b=40)
    )
    return dcc.Graph(figure=fig, style={"marginBottom":"30px"})

def daily_summary_table(n_days=5):
    recent = get_date_rows(dashboard_df, n_days=n_days)
    if recent.empty:
        return html.Div("No recent days data.")
    recent = recent.copy()  # avoid slice warning
    fields = ["Date", "SL%", "ANS%", "AHT"]
    data = []
    recent = recent.reset_index(drop=True)
    recent["Date_dt"] = pd.to_datetime(recent["Date"].astype(str).str.strip(), dayfirst=True, errors='coerce')
    for idx, row in recent.iterrows():
        rec = {f: row.get(f, "N/A") for f in fields}
        data.append(rec)

    table = dash_table.DataTable(
        data=data,
        columns=[{"name": f, "id": f} for f in fields],
        style_cell={'textAlign':'center'},
        style_header={'backgroundColor':'#0b4f6c','color':'white','fontWeight':'bold'},
        style_data={'backgroundColor':'#fde2d1','color':'black'},
        page_action='none',
        style_table={'overflowX':'auto'},
    )
    spark_graphs = []
    for metric in ["SL%", "ANS%", "AHT"]:
        if metric in recent.columns:
            y = recent[metric].map(
                lambda v: float(str(v).rstrip('%').replace(',','').strip())
                if isinstance(v, str) else None
            )
            x = recent["Date_dt"]
            fig = go.Figure(go.Scatter(x=x, y=y, mode='lines+markers', marker=dict(size=6)))
            fig.update_layout(
                margin=dict(l=20, r=20, t=20, b=20),
                height=150,
                title=f"{metric} Trend (last {n_days} days)",
                xaxis_title="",
                yaxis_title=metric,
                template="plotly_white"
            )
            spark_graphs.append(dcc.Graph(figure=fig, config={'displayModeBar':False}, style={"marginBottom":"20px"}))
    return html.Div([table] + spark_graphs, style={"marginTop":"20px", "marginBottom":"30px"})

def sla_breach_kpis():
    recent = get_date_rows(dashboard_df, n_days=7)
    cards = []
    if recent.empty:
        return cards
    if "SL%" in recent.columns:
        sl_vals = recent["SL%"].map(
            lambda v: float(str(v).rstrip('%').replace(',','').strip()) if isinstance(v, str) else None
        )
        count_sl_breach = sl_vals.apply(lambda x: (x < 95) if x is not None else False).sum()
    else:
        count_sl_breach = None
    if "AHT" in recent.columns:
        aht_vals = recent["AHT"].map(
            lambda v: float(str(v).replace(',','').strip()) if isinstance(v, str) else None
        )
        count_aht_breach = aht_vals.apply(lambda x: (x > 130) if x is not None else False).sum()
    else:
        count_aht_breach = None

    if count_sl_breach is not None:
        cards.append(kpi_card("SL% Breach Days (Last 7)", str(int(count_sl_breach)), target=None, inverse=False))
    if count_aht_breach is not None:
        cards.append(kpi_card("AHT Breach Days (Last 7)", str(int(count_aht_breach)), target=None, inverse=False))
    return cards

def generate_chart(df, sheet_name):
    try:
        if sheet_name == "Dashboard":
            if "Date" in df.columns and "ANSWERED" in df.columns:
                df2 = df[df["Date"].astype(str).str.upper() != "MTD"].copy()  # use .copy()
                df2["Date_dt"] = pd.to_datetime(df2["Date"].astype(str).str.strip(), dayfirst=True, errors='coerce')
                fig = px.line(df2, x="Date_dt", y="ANSWERED", title="Answered Calls Over Time", markers=True)
                fig.update_layout(margin=dict(l=40, r=40, t=60, b=40))
                return fig
        elif sheet_name == "Hourly Performance":
            required_cols = ["Hour", "Date", "SL% For Kerala", "SL% For Tamilnadu", "SL% For Chennai"]
            if all(col in df.columns for col in required_cols):
                df2 = df[df["Date"].astype(str).str.upper() != "MTD"].copy()
                df2["Date_dt"] = pd.to_datetime(df2["Date"].astype(str).str.strip(), dayfirst=True, errors='coerce')
                fig = px.line(df2, x="Hour",
                              y=["SL% For Kerala", "SL% For Tamilnadu", "SL% For Chennai"],
                              title="SL% by Location", markers=True)
                fig.update_layout(margin=dict(l=40, r=40, t=60, b=40))
                return fig
    except Exception as e:
        print(f"[{sheet_name}] Chart error: {e}")
    # fallback
    fig = go.Figure()
    fig.update_layout(title="No chart available for this sheet", margin=dict(l=40, r=40, t=60, b=40))
    return fig

# === Layout & Navigation ===

app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)
app.title = "CDR Performance Dashboard"

def nav_buttons():
    return dbc.ButtonGroup([
        dbc.Button("🏠 Home", id="btn_home", href="/", color="info"),
        dbc.Button("🔙 Back", id="btn_back", color="secondary", n_clicks=0)
    ], className="mb-3")

def layout_home():
    # KPI cards
    kpi_cards = []
    kpis = [
        ("IVRS_OFFERED", None, False),
        ("NET OFFERED", None, False),
        ("ANSWERED", None, False),
        ("AHT", 130, False),
        ("ANS%", 95, True),
        ("SL%", 95, True),
        ("CMS Aband%", 5, False),
        ("Entry Level %", 95, True),
        ("Second Level %", 95, True),
        ("Third Level %", 95, True),
    ]

    if dashboard_mtd is not None:
        for label, target, inverse in kpis:
            val = dashboard_mtd.get(label, "N/A")
            kpi_cards.append(kpi_card(label, val, is_percent=("%" in label), target=target, inverse=inverse))
        if kerala_mtd is not None:
            kpi_cards.append(kpi_card("KERALA SL%", kerala_mtd.get("SL%", "N/A"), is_percent=True, target=95, inverse=True))
        if tamilnadu_mtd is not None:
            kpi_cards.append(kpi_card("TAMILNADU SL%", tamilnadu_mtd.get("SL%", "N/A"), is_percent=True, target=95, inverse=True))
        if chennai_mtd is not None:
            kpi_cards.append(kpi_card("CHENNAI SL%", chennai_mtd.get("SL%", "N/A"), is_percent=True, target=95, inverse=True))

    # Performance report table
    perf_data, perf_cols = prepare_performance_report()
    perf_styles = get_conditional_styles_perf_report(perf_data, perf_cols)

    # Reports / Download All card
    download_reports_card = dbc.Card(
        dbc.CardBody([
            html.H5("📂 Reports & Downloads", className="text-primary fw-bold mb-3"),
            html.Ul([
                html.Li(html.A(sheet, href=f"/{sheet.replace(' ', '_')}", className="link-primary"))
                for sheet in sheets.keys()
            ]),
            html.Hr(),
            dbc.Button("⬇️ Download All as Excel", id="btn_download_all_excel", color="primary", className="me-2 mb-2", n_clicks=0),
            dbc.Button("⬇️ Download All as CSV (.zip)", id="btn_download_all_csv", color="success", className="mb-2", n_clicks=0),
            dcc.Download(id="download_all_data")
        ]),
        style={"maxWidth": "300px"},
        className="mb-4 shadow-sm p-3"
    )

    return dbc.Container([
        html.H1("📊 CDR Performance Dashboard", className="text-center my-4 text-primary fw-bold"),
        nav_buttons(),
        html.Hr(),

        # Row: KPIs (big) + Reports/Download (side)
        dbc.Row([
            dbc.Col(
                dbc.Card(
                    dbc.CardBody(
                        dbc.Row(kpi_cards, justify="start", className="g-2 flex-wrap")
                    ),
                    className="mb-4 shadow-sm p-3"
                ),
                md=9
            ),
            dbc.Col(
                download_reports_card,
                md=3
            )
        ]),

        # Other sections
        dbc.Card(
            dbc.CardBody([
                html.H4("📍 SLA / AHT Breaches (Last 7 Days)", className="mb-3 text-secondary"),
                dbc.Row(sla_breach_kpis(), className="g-3")
            ]),
            className="mb-4 shadow-sm p-3"
        ),

        dbc.Card(
            dbc.CardBody([
                html.H4("📈 Trends: SL%, ANS%, AHT (Last 7 Days)", className="mb-3 text-secondary"),
                html.Div(trend_graph_last7(), style={"gap": "30px", "display": "flex", "flexDirection": "column"})
            ]),
            className="mb-4 shadow-sm p-3"
        ),

        dbc.Card(
            dbc.CardBody([
                html.H4("🗓️ Daily Summary (Last 5 Days)", className="mb-3 text-secondary"),
                daily_summary_table(n_days=5)
            ]),
            className="mb-4 shadow-sm p-3"
        ),

        html.H3("📋 Performance Report: MTD / Today / Day‑1", className="my-4"),
        dash_table.DataTable(
            id="perf-report-table",
            columns=perf_cols,
            data=perf_data,
            style_cell={'textAlign':'center', 'minWidth':'80px', 'whiteSpace':'normal'},
            style_header={'backgroundColor':'#0b4f6c', 'color':'white', 'fontWeight':'bold'},
            style_data={'backgroundColor':'#fde2d1', 'color':'black'},
            style_data_conditional=perf_styles,
            page_action='none',
            style_table={'overflowX':'auto'},
        ),

        html.Br()
    ], fluid=True)

def layout_sheet(sheet_name):
    df = sheets.get(sheet_name, pd.DataFrame())
    # Ensure a fresh copy for display
    df_display = df.copy()
    return dbc.Container([
        nav_buttons(),
        html.H2(sheet_name, className="text-primary"),
        dbc.Button("⬇️ Download CSV", id="download_csv_btn", color="success", className="me-2", n_clicks=0),
        dbc.Button("⬇️ Download Excel", id="download_excel_btn", color="primary", n_clicks=0),
        dcc.Download(id="download_data"),
        html.Br(), html.Br(),

        dash_table.DataTable(
            id="sheet-table",
            data=df_display.to_dict('records'),
            columns=[{"name": i, "id": i} for i in df_display.columns if i != "Date_str"],
            page_size=15,
            filter_action="native",
            sort_action="native",
            style_table={'overflowX': 'auto'},
            style_header={
                'backgroundColor': '#0b4f6c',
                'color': 'white',
                'fontWeight': 'bold',
                'textAlign': 'center'
            },
            style_data={
                'backgroundColor': '#fde2d1',
                'textAlign': 'center'
            },
        ),
        html.Br(),
        dcc.Graph(id="sheet-graph", figure=generate_chart(df_display, sheet_name))
    ], fluid=True)

# === Performance Report Table helpers ===

def prepare_performance_report():
    if dashboard_mtd is None:
        return [], []
    fields = [col for col in dashboard_mtd.index if col != "Date"]
    def row_dict(r, period):
        if r is None or (hasattr(r, 'empty') and r.empty):
            return {f: "N/A" for f in fields} | {"Period": period}
        d = {f: r.get(f, "N/A") for f in fields}
        d["Period"] = period
        return d

    rows = []
    rows.append(row_dict(dashboard_mtd, "MTD"))
    rows.append(row_dict(dashboard_today, "Today"))
    rows.append(row_dict(dashboard_day1, "Day‑1"))

    # Clean numeric values for table styling comparisons
    for row in rows:
        for f in fields:
            v = row.get(f, "N/A")
            if isinstance(v, str) and v.endswith("%"):
                try:
                    row[f] = float(v.rstrip("%").strip())
                except:
                    pass
            else:
                try:
                    row[f] = float(str(v).replace(",", "").strip())
                except:
                    pass

    cols = [{"name":"Period", "id":"Period"}] + [{"name":f, "id":f} for f in fields]
    return rows, cols

def get_conditional_styles_perf_report(data, columns):
    styles = []
    targets = {
        "ANS%": (95, True),
        "SL%": (95, True),
        "Entry Level %": (95, True),
        "Second Level %": (95, True),
        "Third Level %": (95, True),
        "CMS Aband%": (5, False),
        "AHT": (130, False),
    }
    col_map = {col['name'].strip().upper(): col['id'] for col in columns}

    for kpi, (target_val, higher_is_better) in targets.items():
        key = col_map.get(kpi.strip().upper())
        if not key:
            continue
        for i, row in enumerate(data):
            if key not in row:
                continue
            val = row.get(key)
            if val is None or (not isinstance(val, (int, float))):
                continue
            color = None
            if kpi == "CMS Aband%":
                color = "green" if val < target_val else "red"
            elif kpi == "AHT":
                color = "green" if val <= target_val else "red"
            else:
                if higher_is_better:
                    color = "green" if val >= target_val else "red"
                else:
                    color = "green" if val <= target_val else "red"
            if color:
                styles.append({
                    'if': {'row_index': i, 'column_id': key},
                    'color': color,
                    'fontWeight': '700'
                })
    return styles

# === Callbacks ===

@app.callback(
    Output("page-content", "children"),
    Input("url", "pathname")
)
def display_page(pathname):
    sheet_key = pathname.strip("/").replace("_", " ")
    if pathname == "/" or pathname is None:
        return layout_home()
    elif sheet_key in sheets:
        return layout_sheet(sheet_key)
    else:
        return dbc.Container([
            nav_buttons(),
            html.H2("404 - Page Not Found", className="text-danger"),
            html.P(f"The page '{pathname}' does not exist.")
        ], fluid=True)

@app.callback(
    Output("download_data", "data"),
    Input("download_csv_btn", "n_clicks"),
    Input("download_excel_btn", "n_clicks"),
    State("url", "pathname"),
    prevent_initial_call=True
)
def download_sheet_file(n_csv, n_excel, pathname):
    triggered_id = ctx.triggered_id
    sheet_name = pathname.strip("/").replace("_", " ")
    df = sheets.get(sheet_name, pd.DataFrame())

    if triggered_id == "download_csv_btn":
        return dcc.send_data_frame(df.to_csv, f"{sheet_name}.csv", index=False)
    elif triggered_id == "download_excel_btn":
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        buffer.seek(0)
        b64 = base64.b64encode(buffer.read()).decode()
        return dict(content=b64, filename=f"{sheet_name}.xlsx", base64=True)

@app.callback(
    Output("download_all_data", "data"),
    Input("btn_download_all_excel", "n_clicks"),
    Input("btn_download_all_csv", "n_clicks"),
    prevent_initial_call=True
)
def download_all_reports(n_excel, n_csv):
    triggered_id = ctx.triggered_id
    if triggered_id == "btn_download_all_excel":
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            for sheet, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet[:31], index=False)
        buffer.seek(0)
        b64 = base64.b64encode(buffer.read()).decode()
        return dict(content=b64, filename="All_Reports.xlsx", base64=True)
    elif triggered_id == "btn_download_all_csv":
        with TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "all_reports.zip")
            with zipfile.ZipFile(zip_path, "w") as zipf:
                for sheet, df in sheets.items():
                    csv_file = os.path.join(tmpdir, f"{sheet}.csv")
                    df.to_csv(csv_file, index=False)
                    zipf.write(csv_file, arcname=f"{sheet}.csv")
            with open(zip_path, "rb") as f:
                content = base64.b64encode(f.read()).decode()
            return dict(content=content, filename="All_Reports.zip", base64=True)

# === App layout and server start ===

app.layout = html.Div([
    dcc.Location(id="url"),a
    html.Div(id="page-content")
])

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run(host="0.0.0.0", port=port, debug=True)
