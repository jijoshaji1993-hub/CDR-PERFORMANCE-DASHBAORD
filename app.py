import os
import pandas as pd
from dash import Dash, dcc, html, Input, Output, State, ctx
import dash_bootstrap_components as dbc
from dash import dash_table
import plotly.express as px
import io
import base64

# === Load Excel File ===
FILE_PATH = "Result.xlsx"
xls = pd.ExcelFile(FILE_PATH)
sheets = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

# Strip column name spaces for all sheets
for sheet_name in sheets:
    sheets[sheet_name].columns = sheets[sheet_name].columns.str.strip()

# === Format date columns ===
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

for sheet_name in sheets:
    sheets[sheet_name] = format_date_columns(sheets[sheet_name])

# === Dash App Setup ===
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)
app.title = "CDR Performance Dashboard"

# === Get MTD row ===
def get_mtd_row(df, sheet_name):
    if df.empty:
        print(f"[{sheet_name}] DataFrame is empty.")
        return pd.Series()
    if "Date" not in df.columns:
        print(f"[{sheet_name}] No 'Date' column. Columns: {df.columns.tolist()}")
        return pd.Series()
    df2 = df.copy()
    df2["Date_str"] = df2["Date"].astype(str).str.strip().str.upper()
    mtd_rows = df2[df2["Date_str"] == "MTD"]
    if mtd_rows.empty:
        print(f"[{sheet_name}] No MTD row found.")
        return pd.Series()
    row = mtd_rows.iloc[0].drop(labels="Date_str", errors="ignore")
    return row

# Precompute MTD rows
dashboard_mtd = get_mtd_row(sheets.get("Dashboard", pd.DataFrame()), "Dashboard")
kerala_mtd = get_mtd_row(sheets.get("Kerala", pd.DataFrame()), "Kerala")
tamilnadu_mtd = get_mtd_row(sheets.get("Tamilnadu", pd.DataFrame()), "Tamilnadu")
chennai_mtd = get_mtd_row(sheets.get("Chennai", pd.DataFrame()), "Chennai")

# === KPI Card ===
def kpi_card(label, value, is_percent=False, target=None, inverse=False):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        display_val = "N/A"
    else:
        try:
            if is_percent:
                display_val = f"{float(value):.2%}"
            else:
                if isinstance(value, (int, float)) and float(value).is_integer():
                    display_val = f"{int(value):,}"
                else:
                    display_val = f"{float(value):,.2f}"
        except:
            display_val = str(value)

    color = "black"
    try:
        val_f = float(value)
    except:
        val_f = None

    if target is not None and val_f is not None:
        if inverse:
            color = "green" if val_f >= target else "red"
        else:
            color = "red" if val_f > target else "green"

    card_style = {
        "textAlign": "center",
        "width": "10rem",
        "margin": "4px",
        "backgroundColor": "#f8f9fa",
        "borderRadius": "0.375rem",
        "boxShadow": "0 0.125rem 0.25rem rgba(0,0,0,0.075)"
    }

    return dbc.Card([
        dbc.CardBody([
            html.H6(label, className="text-muted", style={"fontWeight": "600"}),
            html.H4(display_val, style={"color": color, "fontWeight": "700"})
        ])
    ], style=card_style)

# === Home Layout ===
def layout_home():
    kpi_cards = []

    if not dashboard_mtd.empty:
        exclude = {
            "Date", "Date_str", "MTD",
            "DISPOSED_AT_IVR", "DISP < 10SEC", "Disp W/O<10SEC", "SHORT CALL%",
            "IVRS AHT", "Req 4 AGENT", "IVRS ANS%", "OFFERED", "NET ANSWERED", "Short Ans",
            "ABAN CALLS", "SHORT ABAN<10SEC", "Ans Within 90 Sec", "90 SEC ABOVE ABAND CALLS",
            "Ans exceeds 90 Sec", "AVG WAIT TIME", "Entry Level", "Second Level", "Third Level",
            "Entry Level Fixed", "Second Level Fixed", "Third Level Fixed",
            "Entry Level Met", "Second Level Met", "Third Level Met"
        }

        for col in dashboard_mtd.index:
            if col in exclude:
                continue

            val = dashboard_mtd.get(col, None)
            is_percent = False
            col_lower = col.strip().lower()
            if "%" in col or col_lower.endswith("%") or "sl%" in col_lower or "ans%" in col_lower or "aband%" in col_lower:
                is_percent = True

            target = None
            inverse = False
            if col_lower in ["ans%", "sl%", "entry level %", "second level %", "third level %"]:
                target = 0.95
                inverse = True
            elif col_lower == "cms aband%":
                target = 0.05
            elif col_lower == "aht":
                target = 130

            label = col
            if col_lower == "entry level %":
                label = "ENTRY TCBH"
            elif col_lower == "second level %":
                label = "SECOND TCBH"
            elif col_lower == "third level %":
                label = "THIRD TCBH"

            kpi_cards.append(kpi_card(label, val, is_percent=is_percent, target=target, inverse=inverse))

    for region, mtd_row, sheet_name in [
        ("KERALA SL%", kerala_mtd, "Kerala"),
        ("TAMILNADU SL%", tamilnadu_mtd, "Tamilnadu"),
        ("CHENNAI SL%", chennai_mtd, "Chennai")
    ]:
        if not mtd_row.empty:
            possible_cols = ["SL%", "SL% For " + sheet_name, "SL %", "SL % " + sheet_name]
            for pcol in possible_cols:
                if pcol in mtd_row.index:
                    val = mtd_row.get(pcol, None)
                    kpi_cards.append(kpi_card(region, val, is_percent=True, target=0.95, inverse=True))
                    break

    return dbc.Container([
        html.H1("CDR Performance Dashboard", className="text-center text-primary my-4"),
        dbc.Row([dbc.Col(card, width="auto") for card in kpi_cards], justify="start", className="g-2 flex-wrap"),
        html.Hr(),
        html.H3("Reports"),
        html.Ul([html.Li(html.A(sheet, href=f"/{sheet.replace(' ', '_')}")) for sheet in sheets.keys()])
    ], fluid=True)

# === Navigation Buttons ===
def nav_buttons():
    return dbc.ButtonGroup([
        dbc.Button("🏠 Home", id="btn_home", href="/", color="info"),
        dbc.Button("🔙 Back", id="btn_back", color="secondary", n_clicks=0)
    ], className="mb-3")

# === Chart Generator ===
def generate_chart(df, sheet_name):
    try:
        if sheet_name == "Dashboard":
            if "Date" in df.columns and "ANSWERED" in df.columns:
                df2 = df[df["Date"].astype(str).str.upper() != "MTD"]
                return px.line(df2, x="Date", y="ANSWERED", title="Answered Calls Over Time", markers=True)
        elif sheet_name == "Hourly Performance":
            required_cols = ["Hour", "Date", "SL% For Kerala", "SL% For Tamilnadu", "SL% For Chennai"]
            if all(col in df.columns for col in required_cols):
                df2 = df[df["Date"].astype(str).str.upper() != "MTD"]
                return px.line(df2, x="Hour",
                               y=["SL% For Kerala", "SL% For Tamilnadu", "SL% For Chennai"],
                               title="SL% by Location", markers=True)
    except Exception as e:
        print(f"[{sheet_name}] Chart error: {e}")
    return px.scatter(title="No chart available for this sheet")

# === Sheet Layout ===
def layout_sheet(sheet_name):
    df = sheets.get(sheet_name, pd.DataFrame())

    return dbc.Container([
        nav_buttons(),
        html.H2(sheet_name, className="text-primary"),
        dbc.Button("⬇️ Download CSV", id="download_csv_btn", color="success", className="me-2"),
        dbc.Button("⬇️ Download Excel", id="download_excel_btn", color="primary"),
        dcc.Download(id="download_data"),
        html.Br(), html.Br(),

        dash_table.DataTable(
            id="sheet-table",
            data=df.to_dict('records'),
            columns=[{"name": i, "id": i} for i in df.columns if i != "Date_str"],
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
        dcc.Graph(id="sheet-graph", figure=generate_chart(df, sheet_name))
    ], fluid=True)

# === Layout Manager ===
app.layout = html.Div([
    dcc.Location(id="url"),
    html.Div(id="page-content")
])

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

# === Download Callback ===
@app.callback(
    Output("download_data", "data"),
    Input("download_csv_btn", "n_clicks"),
    Input("download_excel_btn", "n_clicks"),
    State("url", "pathname"),
    prevent_initial_call=True
)
def download_file(n_csv, n_excel, pathname):
    triggered = ctx.triggered_id
    sheet_name = pathname.strip("/").replace("_", " ")
    df = sheets.get(sheet_name, pd.DataFrame())

    if triggered == "download_csv_btn":
        return dcc.send_data_frame(df.to_csv, f"{sheet_name}.csv", index=False)
    elif triggered == "download_excel_btn":
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        buffer.seek(0)
        b64 = base64.b64encode(buffer.read()).decode()
        return dict(content=b64, filename=f"{sheet_name}.xlsx", base64=True)

# === Run Server ===
import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run(host="0.0.0.0", port=port)

