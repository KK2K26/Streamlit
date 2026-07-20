import re
from io import BytesIO
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st
from st_aggrid import (
    AgGrid,
    GridOptionsBuilder,
    GridUpdateMode,
    DataReturnMode,
    JsCode,
)


EXCEL_FILE = "Open Tickets in GLPI 8 Apr 1.xlsx"

REQUIRED_COLUMNS = [
    "ID",
    "Title",
    "Status",
    "Priority",
    "Type",
    "Assigned to - Assigned To",
    "Assigned to - Assignment Group",
    "Requester - Requester",
    "Opening date",
    "Description",
    "SLA Bucket",
]

AGING_COLUMNS = [
    "00_24 Hours",
    "01_03 Days",
    "04_05 Days",
    "06_10 Days",
    "11_15 Days",
    "16_31 Days",
    "31_60 Days",
    ">60 Days",
]

STATUS_VALUES = ["New", "Pending", "Assigned"]
TYPE_COLUMNS = ["Incident"]
PRIORITY_COLUMNS = ["P1", "P2", "P3", "P4"]

DETAIL_COLUMNS = [
    "ID",
    "Title",
    "Assigned to - Assigned To",
    "Assigned to - Assignment Group",
    "Requester - Requester",
    "Status",
    "Priority",
    "Type",
    "SLA Bucket",
    "SLA Days",
    "Opening date",
    "Description",
]


st.set_page_config(
    page_title="GLPI Operations Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)


@st.cache_data(show_spinner=False)
def load_data(file_path: str = EXCEL_FILE) -> pd.DataFrame:
    df = pd.read_excel(file_path, engine="openpyxl")
    df.columns = df.columns.astype(str).str.strip()

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[REQUIRED_COLUMNS].copy()

    text_columns = [
        "Title",
        "Status",
        "Priority",
        "Type",
        "Assigned to - Assigned To",
        "Assigned to - Assignment Group",
        "Requester - Requester",
        "Description",
        "SLA Bucket",
    ]

    for col in text_columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["ID"] = df["ID"].fillna("").astype(str).str.strip()
    df["Opening date"] = pd.to_datetime(df["Opening date"], errors="coerce")

    now = pd.Timestamp.now()
    elapsed_hours = (now - df["Opening date"]).dt.total_seconds().div(3600)
    elapsed_days = elapsed_hours.div(24)

    df["SLA Days"] = elapsed_days.fillna(0).clip(lower=0).astype(int)

    df["SLA Aging Bucket"] = pd.cut(
        elapsed_days.fillna(0),
        bins=[-1, 1, 3, 5, 10, 15, 31, 60, float("inf")],
        labels=AGING_COLUMNS,
        right=True,
    ).astype(str)

    df.loc[elapsed_hours <= 24, "SLA Aging Bucket"] = "00_24 Hours"

    df["Assigned Group"] = df["Assigned to - Assignment Group"].replace("", "Unassigned")
    df["SLA Bucket Clean"] = df["SLA Bucket"].replace("", "Blank")

    status_lower = df["Status"].str.lower()

    df["Status Group"] = "Other"
    df.loc[status_lower.str.contains("new", na=False), "Status Group"] = "New"
    df.loc[status_lower.str.contains("pending", na=False), "Status Group"] = "Pending"
    df.loc[
        status_lower.str.contains("assigned|processing", regex=True, na=False),
        "Status Group",
    ] = "Assigned"

    df["Status Clean"] = df["Status"].str.strip().str.title()
    df["Type Clean"] = df["Type"].str.strip().str.title()

    priority = df["Priority"].fillna("").astype(str).str.strip().str.upper()

    df["Priority Clean"] = "Other"

    df.loc[
        priority.isin(["VERY HIGH", "CRITICAL", "P1"]),
        "Priority Clean",
    ] = "P1"

    df.loc[
        priority.isin(["HIGH", "MEDIUM", "P2"]),
        "Priority Clean",
    ] = "P2"

    df.loc[
        priority.isin(["LOW", "P3"]),
        "Priority Clean",
    ] = "P3"

    df.loc[
        priority.isin(["VERY LOW", "P4"]),
        "Priority Clean",
    ] = "P4"

    closed_pattern = r"\b(closed|resolved|solved|cancelled|canceled|done|completed)\b"

    df["Is Open"] = ~status_lower.str.contains(
        closed_pattern,
        regex=True,
        na=False,
    )

    df["Is New"] = df["Status Group"].eq("New")
    df["Is Pending"] = df["Status Group"].eq("Pending")
    df["Is Assigned"] = df["Status Group"].eq("Assigned")

    df["Is Critical"] = df["Priority Clean"].eq("P1")

    df["Is SLA Risk"] = df["SLA Aging Bucket"].isin(
        ["06_10 Days", "11_15 Days", "16_31 Days", "31_60 Days", ">60 Days"]
    )

    requester = df["Requester - Requester"].fillna("").astype(str).str.upper()

    df["Is Dynatrace"] = requester.str.contains(
        r"DYNATRACE|PLTOOLS|PL\s*TOOLS|NETCOOL",
        regex=True,
        na=False,
    )

    return df


def create_css() -> None:
    st.markdown(
        """
        <style>
            html, body, [class*="css"] {
                font-family: "Segoe UI", Arial, sans-serif;
                font-size: 12px;
            }

            .stApp {
                background-color: #ffffff;
            }

            header[data-testid="stHeader"] {
                height: 0rem;
                background: transparent;
            }

            div[data-testid="stToolbar"],
            footer {
                display: none;
            }

            .block-container {
                padding-top: 0.45rem;
                padding-left: 0.50rem;
                padding-right: 0.50rem;
                padding-bottom: 0.45rem;
                max-width: 100%;
            }

            .dashboard-title {
                border: 2px solid #000000;
                background: #f2f2f2;
                color: #111111;
                padding: 8px 12px;
                margin-bottom: 5px;
                font-size: 24px;
                font-weight: 800;
                line-height: 1.15;
            }

            .section-panel {
                border: 2px solid #000000;
                background: #f2f2f2;
                padding: 6px;
                margin-bottom: 5px;
            }

            .section-title {
                font-size: 15px;
                font-weight: 800;
                color: #111111;
                margin-bottom: 5px;
                line-height: 1.1;
            }

            .kpi-panel {
                border: 2px solid #000000;
                background: #f2f2f2;
                padding: 6px;
                margin-bottom: 5px;
            }

            div[data-testid="column"] {
                padding-left: 3px !important;
                padding-right: 3px !important;
            }

            .stButton > button {
                width: 100%;
                min-height: 74px;
                background: #ffffff;
                color: #111111;
                border: 2px solid #000000;
                border-radius: 0px;
                padding: 5px 8px;
                font-weight: 800;
                font-size: 13px;
                line-height: 1.15;
                white-space: pre-line;
                box-shadow: none;
            }

            .stButton > button:hover {
                background: #e9f2ff;
                color: #000000;
                border: 2px solid #000000;
            }

            .stButton > button:focus {
                border: 3px solid #1f4e79;
                box-shadow: none;
                outline: none;
            }

            div[data-testid="stDownloadButton"] > button {
                width: 100%;
                min-height: 34px;
                background: #111111;
                color: #ffffff;
                border: 2px solid #111111;
                border-radius: 0px;
                font-size: 12px;
                font-weight: 800;
                padding: 5px 10px;
            }

            div[data-testid="stDownloadButton"] > button:hover {
                background: #2b2b2b;
                color: #ffffff;
            }

            div[data-testid="stTextInput"] {
                margin-bottom: 4px !important;
            }

            div[data-testid="stTextInput"] input {
                border-radius: 0px;
                border: 1px solid #000000;
                font-size: 12px;
                height: 34px;
            }

            .filter-chip {
                display: inline-block;
                border: 1px solid #000000;
                background: #ffffff;
                color: #111111;
                font-size: 12px;
                font-weight: 800;
                padding: 3px 8px;
                margin-bottom: 5px;
            }

            .ag-theme-balham {
                --ag-font-family: "Segoe UI", Arial, sans-serif;
                --ag-font-size: 12px;
                --ag-header-background-color: #d9d9d9;
                --ag-header-foreground-color: #000000;
                --ag-border-color: #000000;
                --ag-row-border-color: #bfbfbf;
                --ag-odd-row-background-color: #ffffff;
                --ag-row-hover-color: #e9f2ff;
                --ag-selected-row-background-color: #cfe2f3;
            }

            .ag-theme-balham .ag-root-wrapper {
                border: 1px solid #000000;
            }

            .ag-theme-balham .ag-header-cell {
                font-weight: 800;
                border-right: 1px solid #000000;
            }

            .ag-theme-balham .ag-cell {
                border-right: 1px solid #d0d0d0;
            }

            .js-plotly-plot {
                border: 1px solid #000000;
                background: #ffffff;
            }

            .element-container {
                margin-bottom: 0px !important;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def init_state() -> None:
    defaults = {
        "selected_filter_type": "all_open",
        "selected_filter_payload": {},
        "selected_filter_label": "Total Open Tickets",
        "last_matrix_click": "",
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def set_filter(filter_type: str, payload: dict | None = None, label: str | None = None) -> None:
    st.session_state["selected_filter_type"] = filter_type
    st.session_state["selected_filter_payload"] = payload or {}
    st.session_state["selected_filter_label"] = label or filter_type
    st.session_state["last_matrix_click"] = f"{filter_type}_{payload}_{datetime.now().timestamp()}"


def apply_selected_filter(df: pd.DataFrame) -> pd.DataFrame:
    filter_type = st.session_state.get("selected_filter_type", "all_open")
    payload = st.session_state.get("selected_filter_payload", {})

    if filter_type == "all_open":
        return df[df["Is Open"]].copy()

    if filter_type == "new":
        return df[df["Is New"]].copy()

    if filter_type == "pending":
        return df[df["Is Pending"]].copy()

    if filter_type == "critical":
        return df[df["Is Critical"]].copy()

    if filter_type == "sla_risk":
        return df[df["Is SLA Risk"]].copy()

    if filter_type == "status":
        status = payload.get("status", "")
        return df[df["Status Group"].eq(status)].copy()

    if filter_type == "aging_matrix":
        group = payload.get("group")
        bucket = payload.get("bucket")
        filtered = df.copy()

        if group and group != "Grand Total":
            filtered = filtered[filtered["Assigned Group"].eq(group)]

        if bucket and bucket != "Total":
            filtered = filtered[filtered["SLA Aging Bucket"].eq(bucket)]

        return filtered.copy()

    if filter_type == "dynatrace":
        return df[df["Is Dynatrace"]].copy()

    if filter_type == "dynatrace_matrix":
        group = payload.get("group")
        bucket = payload.get("bucket")
        filtered = df[df["Is Dynatrace"]].copy()

        if group and group != "Grand Total":
            filtered = filtered[filtered["Assigned Group"].eq(group)]

        if bucket and bucket != "Total":
            filtered = filtered[filtered["SLA Bucket Clean"].eq(bucket)]

        return filtered.copy()

    if filter_type == "ticket_type_matrix":
        group = payload.get("group")
        ticket_type = payload.get("ticket_type")
        filtered = df.copy()

        if group and group != "Grand Total":
            filtered = filtered[filtered["Assigned Group"].eq(group)]

        if ticket_type and ticket_type != "Total":
            filtered = filtered[filtered["Type Clean"].eq(ticket_type)]

        return filtered.copy()

    if filter_type == "priority_matrix":
        group = payload.get("group")
        priority = payload.get("priority")
        filtered = df.copy()

        if group and group != "Grand Total":
            filtered = filtered[filtered["Assigned Group"].eq(group)]

        if priority and priority != "Total":
            filtered = filtered[filtered["Priority Clean"].eq(priority)]

        return filtered.copy()

    return df.copy()


def export_to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()

    export_df = df[DETAIL_COLUMNS].copy()
    export_df["Opening date"] = export_df["Opening date"].dt.strftime("%Y-%m-%d %H:%M:%S")
    export_df["Opening date"] = export_df["Opening date"].fillna("")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Ticket Details")
        sheet = writer.sheets["Ticket Details"]

        for column_cells in sheet.columns:
            max_len = 0
            col_letter = column_cells[0].column_letter

            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(value))

            sheet.column_dimensions[col_letter].width = min(max_len + 2, 60)

    output.seek(0)
    return output.getvalue()


def build_matrix(
    df: pd.DataFrame,
    row_col: str,
    col_col: str,
    value_columns: list[str],
    row_name: str = "Assigned Group",
) -> pd.DataFrame:
    if df.empty:
        empty = pd.DataFrame(columns=[row_name] + value_columns + ["Total"])
        empty.loc[0] = ["Grand Total"] + ["_" for _ in value_columns] + ["_"]
        empty["_clicked_col"] = ""
        return empty

    matrix_numeric = (
        pd.pivot_table(
            df,
            index=row_col,
            columns=col_col,
            values="ID",
            aggfunc="count",
            fill_value=0,
        )
        .reindex(columns=value_columns, fill_value=0)
        .reset_index()
    )

    matrix_numeric = matrix_numeric.rename(columns={row_col: row_name})
    matrix_numeric["Total"] = matrix_numeric[value_columns].sum(axis=1)

    total_row = {row_name: "Grand Total"}

    for col in value_columns:
        total_row[col] = int(matrix_numeric[col].sum())

    total_row["Total"] = int(matrix_numeric["Total"].sum())

    matrix_numeric = pd.concat(
        [matrix_numeric, pd.DataFrame([total_row])],
        ignore_index=True,
    )

    matrix_display = matrix_numeric.copy()

    for col in value_columns + ["Total"]:
        matrix_display[col] = matrix_display[col].apply(
            lambda x: "_" if pd.isna(x) or int(x) == 0 else int(x)
        )

    matrix_display["_clicked_col"] = ""

    return matrix_display


def extract_aggrid_cell(response: dict, group_col: str = "Assigned Group") -> tuple[str | None, str | None]:
    rows = response.get("selected_rows")

    if rows is None:
        return None, None

    if isinstance(rows, pd.DataFrame):
        if rows.empty:
            return None, None
        row = rows.iloc[0].to_dict()

    elif isinstance(rows, list):
        if len(rows) == 0:
            return None, None
        row = rows[0]

    else:
        return None, None

    group = row.get(group_col)
    clicked_col = row.get("_clicked_col")

    if not clicked_col or clicked_col == group_col:
        clicked_col = "Total"

    return group, clicked_col


def render_matrix_grid(
    matrix_df: pd.DataFrame,
    height: int,
    key: str,
    numeric_columns: list[str],
    group_col: str = "Assigned Group",
    first_col_width: int = 230,
    numeric_width: int = 95,
) -> dict:
    gb = GridOptionsBuilder.from_dataframe(matrix_df)

    gb.configure_default_column(
        sortable=True,
        resizable=True,
        filter=False,
        editable=False,
        suppressMenu=True,
        wrapText=False,
        autoHeight=False,
    )

    gb.configure_column(
        group_col,
        pinned="left",
        width=first_col_width,
        cellStyle={"fontWeight": "700"},
    )

    gb.configure_column("_clicked_col", hide=True)

    for col in numeric_columns:
        gb.configure_column(
            col,
            width=numeric_width,
            cellStyle={
                "textAlign": "center",
                "fontWeight": "800",
                "cursor": "pointer",
            },
        )

    gb.configure_selection(
        selection_mode="single",
        use_checkbox=False,
        suppressRowDeselection=True,
    )

    gb.configure_grid_options(
        rowHeight=29,
        headerHeight=32,
        suppressRowClickSelection=False,
        enableCellTextSelection=True,
        ensureDomOrder=True,
        onCellClicked=JsCode(
            """
            function(params) {
                params.api.deselectAll();
                params.data._clicked_col = params.column.getColId();
                params.node.setSelected(true);
                params.api.applyTransaction({ update: [params.data] });
            }
            """
        ),
        getRowStyle=JsCode(
            """
            function(params) {
                const keys = Object.keys(params.data);
                const firstCol = keys[0];

                if (params.data && params.data[firstCol] === 'Grand Total') {
                    return {
                        'font-weight': '800',
                        'background-color': '#d9d9d9',
                        'border-top': '2px solid #000000'
                    };
                }

                return {};
            }
            """
        ),
    )

    return AgGrid(
        matrix_df,
        gridOptions=gb.build(),
        height=height,
        theme="balham",
        fit_columns_on_grid_load=False,
        allow_unsafe_jscode=True,
        update_mode=GridUpdateMode.SELECTION_CHANGED | GridUpdateMode.VALUE_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        key=key,
    )


def create_kpis(df: pd.DataFrame) -> None:
    st.markdown('<div class="kpi-panel">', unsafe_allow_html=True)

    kpis = [
        {
            "title": "Total Open Tickets",
            "value": int(df["Is Open"].sum()),
            "filter_type": "all_open",
            "payload": {},
            "label": "Total Open Tickets",
        },
        {
            "title": "Tickets in New State",
            "value": int(df["Is New"].sum()),
            "filter_type": "new",
            "payload": {},
            "label": "Tickets in New State",
        },
        {
            "title": "Pending Tickets",
            "value": int(df["Is Pending"].sum()),
            "filter_type": "pending",
            "payload": {},
            "label": "Pending Tickets",
        },
        {
            "title": "Critical Tickets",
            "value": int(df["Is Critical"].sum()),
            "filter_type": "critical",
            "payload": {},
            "label": "Critical Tickets",
        },
        {
            "title": "SLA Risk Tickets",
            "value": int(df["Is SLA Risk"].sum()),
            "filter_type": "sla_risk",
            "payload": {},
            "label": "SLA Risk Tickets",
        },
    ]

    cols = st.columns(5, gap="small")

    for idx, (col, item) in enumerate(zip(cols, kpis)):
        with col:
            clicked = st.button(
                f"{item['title']}\n{item['value']:,}",
                key=f"kpi_{idx}",
                use_container_width=True,
            )

            if clicked:
                set_filter(
                    item["filter_type"],
                    item["payload"],
                    item["label"],
                )
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def create_aging_matrix(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Aging Ticket Status</div>', unsafe_allow_html=True)

    matrix_df = build_matrix(
        df=df,
        row_col="Assigned Group",
        col_col="SLA Aging Bucket",
        value_columns=AGING_COLUMNS,
    )

    response = render_matrix_grid(
        matrix_df=matrix_df,
        height=280,
        key="aging_matrix_grid",
        numeric_columns=AGING_COLUMNS + ["Total"],
        first_col_width=230,
        numeric_width=95,
    )

    group, bucket = extract_aggrid_cell(response)

    if group and bucket and bucket in AGING_COLUMNS + ["Total"]:
        new_filter = {
            "group": group,
            "bucket": bucket,
        }

        if (
            st.session_state.get("selected_filter_type") != "aging_matrix"
            or st.session_state.get("selected_filter_payload") != new_filter
        ):
            set_filter(
                "aging_matrix",
                new_filter,
                f"Aging Ticket Status | {group} | {bucket}",
            )
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def create_status_pie(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Count of Status by Status</div>', unsafe_allow_html=True)

    status_df = (
        df[df["Status Group"].isin(STATUS_VALUES)]
        .groupby("Status Group", as_index=False)
        .agg(Ticket_Count=("ID", "count"))
    )

    status_df = pd.DataFrame({"Status Group": STATUS_VALUES}).merge(
        status_df,
        on="Status Group",
        how="left",
    )

    status_df["Ticket_Count"] = status_df["Ticket_Count"].fillna(0).astype(int)

    fig = px.pie(
        status_df,
        names="Status Group",
        values="Ticket_Count",
        hole=0.35,
        color="Status Group",
        color_discrete_map={
            "New": "#4472C4",
            "Pending": "#ED7D31",
            "Assigned": "#A5A5A5",
        },
    )

    fig.update_traces(
        textposition="inside",
        texttemplate="%{label}<br>%{value}<br>%{percent}",
        hovertemplate="<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>",
        marker=dict(line=dict(color="#000000", width=1)),
    )

    fig.update_layout(
        height=300,
        margin=dict(l=5, r=5, t=5, b=5),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.08,
            xanchor="center",
            x=0.5,
            font=dict(size=11),
        ),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(family="Segoe UI", size=12, color="#111111"),
    )

    selected = st.plotly_chart(
        fig,
        use_container_width=True,
        key="status_pie_chart",
        on_select="rerun",
        selection_mode="points",
    )

    try:
        points = selected.get("selection", {}).get("points", [])

        if points:
            status_label = points[0].get("label")

            if status_label in STATUS_VALUES:
                new_filter = {"status": status_label}

                if (
                    st.session_state.get("selected_filter_type") != "status"
                    or st.session_state.get("selected_filter_payload") != new_filter
                ):
                    set_filter(
                        "status",
                        new_filter,
                        f"Status | {status_label}",
                    )
                    st.rerun()

    except Exception:
        pass

    st.markdown("</div>", unsafe_allow_html=True)


def create_dynatrace_section(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Dynatrace Tickets</div>', unsafe_allow_html=True)

    dyn_df = df[df["Is Dynatrace"]].copy()
    dyn_count = int(len(dyn_df))

    clicked = st.button(
        f"Dynatrace Tickets\n{dyn_count:,}",
        key="dynatrace_kpi",
        use_container_width=True,
    )

    if clicked:
        set_filter("dynatrace", {}, "Dynatrace Tickets")
        st.rerun()

    sla_bucket_columns = sorted(
        [x for x in dyn_df["SLA Bucket Clean"].dropna().unique().tolist() if x]
    )

    if not sla_bucket_columns:
        sla_bucket_columns = ["Blank"]

    matrix_df = build_matrix(
        df=dyn_df,
        row_col="Assigned Group",
        col_col="SLA Bucket Clean",
        value_columns=sla_bucket_columns,
    )

    response = render_matrix_grid(
        matrix_df=matrix_df,
        height=205,
        key="dynatrace_matrix_grid",
        numeric_columns=sla_bucket_columns + ["Total"],
        first_col_width=230,
        numeric_width=95,
    )

    group, bucket = extract_aggrid_cell(response)

    if group and bucket and bucket in sla_bucket_columns + ["Total"]:
        new_filter = {
            "group": group,
            "bucket": bucket,
        }

        if (
            st.session_state.get("selected_filter_type") != "dynatrace_matrix"
            or st.session_state.get("selected_filter_payload") != new_filter
        ):
            set_filter(
                "dynatrace_matrix",
                new_filter,
                f"Dynatrace Tickets | {group} | {bucket}",
            )
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def create_ticket_count_matrix(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Total Ticket Count</div>', unsafe_allow_html=True)

    matrix_df = build_matrix(
        df=df[df["Type Clean"].isin(TYPE_COLUMNS)],
        row_col="Assigned Group",
        col_col="Type Clean",
        value_columns=TYPE_COLUMNS,
    )

    response = render_matrix_grid(
        matrix_df=matrix_df,
        height=250,
        key="ticket_count_matrix_grid",
        numeric_columns=TYPE_COLUMNS + ["Total"],
        first_col_width=230,
        numeric_width=95,
    )

    group, ticket_type = extract_aggrid_cell(response)

    if group and ticket_type and ticket_type in TYPE_COLUMNS + ["Total"]:
        new_filter = {
            "group": group,
            "ticket_type": ticket_type,
        }

        if (
            st.session_state.get("selected_filter_type") != "ticket_type_matrix"
            or st.session_state.get("selected_filter_payload") != new_filter
        ):
            set_filter(
                "ticket_type_matrix",
                new_filter,
                f"Total Ticket Count | {group} | {ticket_type}",
            )
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def create_priority_matrix(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Current Priority</div>', unsafe_allow_html=True)

    matrix_df = build_matrix(
        df=df[df["Priority Clean"].isin(PRIORITY_COLUMNS)],
        row_col="Assigned Group",
        col_col="Priority Clean",
        value_columns=PRIORITY_COLUMNS,
    )

    response = render_matrix_grid(
        matrix_df=matrix_df,
        height=250,
        key="priority_matrix_grid",
        numeric_columns=PRIORITY_COLUMNS + ["Total"],
        first_col_width=230,
        numeric_width=85,
    )

    group, priority = extract_aggrid_cell(response)

    if group and priority and priority in PRIORITY_COLUMNS + ["Total"]:
        new_filter = {
            "group": group,
            "priority": priority,
        }

        if (
            st.session_state.get("selected_filter_type") != "priority_matrix"
            or st.session_state.get("selected_filter_payload") != new_filter
        ):
            set_filter(
                "priority_matrix",
                new_filter,
                f"Current Priority | {group} | {priority}",
            )
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def create_ticket_grid(df: pd.DataFrame) -> None:
    st.markdown('<div class="section-panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Ticket Details</div>', unsafe_allow_html=True)

    selected_label = st.session_state.get("selected_filter_label", "Total Open Tickets")

    st.markdown(
        f'<span class="filter-chip">Selected: {selected_label} | Records: {len(df):,}</span>',
        unsafe_allow_html=True,
    )

    search_col, export_col = st.columns([5, 1.2], gap="small")

    with search_col:
        ticket_search = st.text_input(
            "Search by Ticket Number",
            value="",
            placeholder="Search by Ticket Number",
            key="ticket_number_search",
            label_visibility="collapsed",
        )

    detail_df = df[DETAIL_COLUMNS].copy()

    if ticket_search:
        search_value = str(ticket_search).strip()
        detail_df = detail_df[
            detail_df["ID"].astype(str).str.contains(
                re.escape(search_value),
                case=False,
                na=False,
            )
        ]

    with export_col:
        st.download_button(
            label="Export to Excel",
            data=export_to_excel(detail_df),
            file_name=f"GLPI_Ticket_Details_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    display_df = detail_df.copy()
    display_df["Opening date"] = display_df["Opening date"].dt.strftime("%Y-%m-%d %H:%M:%S")
    display_df["Opening date"] = display_df["Opening date"].fillna("")

    gb = GridOptionsBuilder.from_dataframe(display_df)

    gb.configure_default_column(
        sortable=True,
        filter=True,
        resizable=True,
        editable=False,
        floatingFilter=True,
        wrapText=False,
        autoHeight=False,
        suppressMenu=False,
    )

    gb.configure_column("ID", pinned="left", width=110, checkboxSelection=True)
    gb.configure_column("Title", width=280, tooltipField="Title")
    gb.configure_column("Assigned to - Assigned To", header_name="Assignee", width=200)
    gb.configure_column("Assigned to - Assignment Group", header_name="Assignment Group", width=230)
    gb.configure_column("Requester - Requester", header_name="Requester", width=200)
    gb.configure_column("Status", width=130)
    gb.configure_column("Priority", width=105)
    gb.configure_column("Type", width=105)
    gb.configure_column("SLA Bucket", width=130)
    gb.configure_column("SLA Days", header_name="SLA Days", width=105, type=["numericColumn"])
    gb.configure_column("Opening date", width=170)
    gb.configure_column("Description", header_name="Description", width=450, tooltipField="Description")

    gb.configure_selection(
        selection_mode="single",
        use_checkbox=True,
        suppressRowDeselection=False,
    )

    gb.configure_pagination(
        enabled=True,
        paginationAutoPageSize=False,
        paginationPageSize=15,
    )

    gb.configure_grid_options(
        rowHeight=32,
        headerHeight=34,
        floatingFiltersHeight=34,
        suppressMenuHide=True,
        enableCellTextSelection=True,
        ensureDomOrder=True,
        domLayout="normal",
        alwaysShowHorizontalScroll=True,
        suppressHorizontalScroll=False,
        tooltipShowDelay=0,
    )

    AgGrid(
        display_df,
        gridOptions=gb.build(),
        height=455,
        theme="balham",
        fit_columns_on_grid_load=False,
        allow_unsafe_jscode=True,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        key="ticket_detail_grid",
    )

    st.markdown("</div>", unsafe_allow_html=True)


def main() -> None:
    create_css()
    init_state()

    st.markdown(
        '<div class="dashboard-title">GLPI Operations Dashboard</div>',
        unsafe_allow_html=True,
    )

    df = load_data(EXCEL_FILE)

    create_kpis(df)
    create_aging_matrix(df)

    left_col, right_col = st.columns([1, 1.35], gap="small")

    with left_col:
        create_status_pie(df)

    with right_col:
        create_dynatrace_section(df)

    bottom_left, bottom_right = st.columns(2, gap="small")

    with bottom_left:
        create_ticket_count_matrix(df)

    with bottom_right:
        create_priority_matrix(df)

    filtered_detail_df = apply_selected_filter(df)
    create_ticket_grid(filtered_detail_df)


if __name__ == "__main__":
    st.sidebar.success(
    "Version 20-Jul-2026 7:15 PM"
)

    main()