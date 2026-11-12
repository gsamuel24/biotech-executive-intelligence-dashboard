import io
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
from st_aggrid.shared import JsCode as _JsCode

import streamlit as st
import pandas as pd
import json
from pathlib import Path

st.set_page_config(
    page_title="Biotech Executive Intelligence",
    page_icon="🧬",
    layout="wide"
)

DATA_DIR = Path(__file__).resolve().parent

FUNCTIONAL_AREAS = [
    "Executive Leadership",
    "Business & Corporate Development",
    "Finance",
    "Medical & Clinical Development",
    "Research & Scientific",
    "Regulatory Affairs",
    "Technical Operations & Manufacturing",
    "Quality",
    "Human Resources",
    "Commercial",
    "Investor Relations & Corporate Communications",
    "Legal",
    "Program Leadership & Management",
    "Operations",
    "Data, Digital & Information Technology",
    "Other Executive Leadership"
]

@st.cache_data(ttl=60)
def load_data():
    companies = pd.read_parquet(
        DATA_DIR / "executive_matrix_precomputed.parquet"
    )
    return companies, None, None

companies, executives, board = load_data()

def clean(value):
    return "" if pd.isna(value) else str(value).strip()

def parse_functions(value):
    if pd.isna(value):
        return []
    try:
        parsed = json.loads(str(value))
        if isinstance(parsed, list):
            return [str(x).strip() for x in parsed]
    except (ValueError, TypeError):
        pass
    return [clean(value)] if clean(value) else []

@st.cache_data
def build_matrix(companies, executives):
    rows = []

    for _, person in executives.iterrows():
        functions = parse_functions(
            person["functional_groups_approved"]
        )

        name = clean(person["executive_name"])
        title = clean(person["standardized_title"])

        source = clean(person["source_url"])

        # Structured data: name, title, and source URL
        display_text = json.dumps({
            "name": name,
            "title": title,
            "source": source
        }, ensure_ascii=False)

        for function in functions:
            if function in FUNCTIONAL_AREAS:
                rows.append({
                    "company_id": person["company_id"],
                    "function": function,
                    "person": display_text
                })

    if rows:
        long_df = pd.DataFrame(rows)

        grouped = (
            long_df.groupby(
                ["company_id", "function"]
            )["person"]
            .apply(
                lambda x: json.dumps(
                    [json.loads(item) for item in dict.fromkeys(x)],
                    ensure_ascii=False
                )
            )
            .unstack(fill_value="")
            .reset_index()
        )
    else:
        grouped = pd.DataFrame(columns=["company_id"])

    matrix = companies.merge(
        grouped,
        on="company_id",
        how="left",
        validate="one_to_one"
    )

    for function in FUNCTIONAL_AREAS:
        if function not in matrix.columns:
            matrix[function] = ""
        matrix[function] = matrix[function].fillna("")

    return matrix


def style_matrix(df):
    """Color executive cells and improve matrix readability."""

    def apply_row_style(row):
        row_index = df.index.get_loc(row.name)

        base_color = (
            "#FFFFFF" if row_index % 2 == 0
            else "#F4F7FA"
        )

        styles = []

        for column, value in row.items():

            if column == "company_name":
                styles.append(
                    "background-color: #E3EDF5;"
                    "color: #102A43;"
                    "font-weight: bold;"
                )

            elif column in FUNCTIONAL_AREAS:

                if pd.isna(value) or str(value).strip() == "":
                    styles.append(
                        "background-color: #EDF1F4;"
                        "color: #94A3B8;"
                    )
                else:
                    styles.append(
                        "background-color: #E2F3F1;"
                        "color: #155E70;"
                    )

            else:
                styles.append(
                    f"background-color: {base_color};"
                    "color: #334155;"
                )

        return styles

    return df.style.apply(apply_row_style, axis=1)


def show_executive_grid(df):
    """Render a formatted company-by-function matrix."""

    grid_data = df.copy().fillna("")

    # Convert the temporary separator back into two parts.
    # Only functional-area cells are affected.
    for column in FUNCTIONAL_AREAS:
        if column in grid_data.columns:
            grid_data[column] = grid_data[column].astype(str)

    builder = GridOptionsBuilder.from_dataframe(grid_data)

    builder.configure_default_column(
        resizable=True,
        sortable=True,
        filter=False,
        wrapText=True,
        autoHeight=False,
        minWidth=110
    )

    # Company name with a clickable website icon
    company_renderer = JsCode("""
    class CompanyRenderer {
        init(params) {
            this.eGui = document.createElement('div');
            this.eGui.style.display = 'flex';
            this.eGui.style.alignItems = 'center';
            this.eGui.style.flexWrap = 'wrap';
            this.eGui.style.gap = '5px';

            const name = document.createElement('span');
            name.textContent = String(params.value || '');
            name.style.fontWeight = '700';
            name.style.color = '#102A43';

            this.eGui.appendChild(name);

            const website = String(
                params.data?.website || ''
            ).trim();

            if (website.startsWith('https://') ||
                website.startsWith('http://')) {

                const link = document.createElement('a');
                link.href = website;
                link.target = '_blank';
                link.rel = 'noopener noreferrer';
                link.title = 'Visit company website';
                link.setAttribute(
                    'aria-label',
                    'Visit website for ' + name.textContent
                );

                link.style.display = 'inline-flex';
                link.style.alignItems = 'center';
                link.style.color = '#176B80';
                link.style.flexShrink = '0';

                link.innerHTML = `
                    <svg xmlns="http://www.w3.org/2000/svg"
                         width="13" height="13"
                         viewBox="0 0 24 24"
                         fill="none"
                         stroke="currentColor"
                         stroke-width="2"
                         stroke-linecap="round"
                         stroke-linejoin="round">
                        <path d="M15 3h6v6"/>
                        <path d="M10 14 21 3"/>
                        <path d="M21 13v6a2 2 0 0 1-2 2H5
                                 a2 2 0 0 1-2-2V5
                                 a2 2 0 0 1 2-2h6"/>
                    </svg>
                `;

                link.addEventListener('click', event => {
                    event.stopPropagation();
                });

                this.eGui.appendChild(link);
            }
        }

        getGui() {
            return this.eGui;
        }
    }
    """)

    builder.configure_column(
        "company_name",
        header_name="Company",
        pinned="left",
        width=220,
        minWidth=180,
        cellRenderer=company_renderer,
        cellStyle={
            "backgroundColor": "#E3EDF5",
            "color": "#102A43",
            "fontWeight": "bold"
        }
    )

    builder.configure_column(
        "website",
        hide=True
    )

    # Custom renderer: name on first line, title below
    executive_renderer = JsCode("""
    class ExecutiveRenderer {
        init(params) {
            this.eGui = document.createElement('div');
            this.eGui.style.padding = '2px';
            this.eGui.style.whiteSpace = 'normal';

            const raw = String(params.value || '').trim();
            if (!raw) return;

            let executives;

            try {
                executives = JSON.parse(raw);
                if (!Array.isArray(executives)) {
                    executives = [executives];
                }
            } catch (error) {
                executives = [{
                    name: raw,
                    title: '',
                    source: ''
                }];
            }

            executives.forEach((person, index) => {
                const nameRow = document.createElement('div');
                nameRow.style.display = 'flex';
                nameRow.style.alignItems = 'center';
                nameRow.style.gap = '5px';
                nameRow.style.flexWrap = 'wrap';

                const name = document.createElement('span');
                name.textContent = person.name || '';
                name.style.fontWeight = '700';
                name.style.color = '#075985';
                name.style.fontSize = '13px';
                name.style.lineHeight = '17px';

                nameRow.appendChild(name);

                const source = String(person.source || '').trim();

                if (source.startsWith('https://') ||
                    source.startsWith('http://')) {

                    const link = document.createElement('a');
                    link.href = source;
                    link.target = '_blank';
                    link.rel = 'noopener noreferrer';
                    link.title = 'View source';
                    link.setAttribute(
                        'aria-label',
                        'View source for ' + (person.name || 'executive')
                    );

                    link.style.display = 'inline-flex';
                    link.style.alignItems = 'center';
                    link.style.flexShrink = '0';
                    link.style.color = '#176B80';

                    link.innerHTML = `
                        <svg xmlns="http://www.w3.org/2000/svg"
                             width="13" height="13"
                             viewBox="0 0 24 24"
                             fill="none"
                             stroke="currentColor"
                             stroke-width="2"
                             stroke-linecap="round"
                             stroke-linejoin="round">
                            <path d="M15 3h6v6"/>
                            <path d="M10 14 21 3"/>
                            <path d="M21 13v6a2 2 0 0 1-2 2H5
                                     a2 2 0 0 1-2-2V5
                                     a2 2 0 0 1 2-2h6"/>
                        </svg>
                    `;

                    link.addEventListener('click', event => {
                        event.stopPropagation();
                    });

                    nameRow.appendChild(link);
                }

                const title = document.createElement('div');
                title.textContent = person.title || '';
                title.style.color = '#64748B';
                title.style.fontSize = '12px';
                title.style.lineHeight = '16px';
                title.style.marginTop = '0px';

                this.eGui.appendChild(nameRow);
                this.eGui.appendChild(title);

                if (index < executives.length - 1) {
                    const divider = document.createElement('hr');
                    divider.style.margin = '6px 0';
                    divider.style.border = '0';
                    divider.style.borderTop = '1px solid #C6DEDF';
                    this.eGui.appendChild(divider);
                }
            });
        }

        getGui() {
            return this.eGui;
        }
    }
    """)

    for function in FUNCTIONAL_AREAS:
        if function in grid_data.columns:
            builder.configure_column(
                function,
                header_name=function,
                width=230,
                minWidth=180,
                cellRenderer=executive_renderer,
                cellStyle=JsCode("""
                    function(params) {
                        const hasExecutive =
                            params.value != null &&
                            String(params.value).trim() !== "";

                        return {
                            backgroundColor: hasExecutive
                                ? "#E2F3F1"
                                : "#EDF1F4",
                            whiteSpace: "normal",
                            padding: "4px 8px",
                            display: "flex",
                            alignItems: "flex-start"
                        };
                    }
                """)
            )


    # Friendly display names for company information
    builder.configure_column(
        "company_type",
        header_name="Company Type"
    )
    builder.configure_column(
        "latest_stage",
        header_name="Stage"
    )
    builder.configure_column(
        "primary_therapeutic_area",
        header_name="Primary TA"
    )

    options = builder.build()

    # Restrict AG Grid's built-in export and copy features.
    options["suppressCsvExport"] = True
    options["suppressExcelExport"] = True
    options["suppressClipboardPaste"] = True
    options["suppressCopyRowsToClipboard"] = True
    options["suppressCopySingleCellRanges"] = True
    options["suppressContextMenu"] = True
    options["enableRangeSelection"] = False
    options["cellSelection"] = False

    options["rowHeight"] = 90
    options["headerHeight"] = 48
    options["suppressRowClickSelection"] = True
    options["ensureDomOrder"] = True

    AgGrid(
        grid_data,
        gridOptions=options,
        height=650,
        width="100%",
        allow_unsafe_jscode=True,
        theme="streamlit",
        fit_columns_on_grid_load=False,
        key="executive_matrix_aggrid"
    )


# ============================================================
# BOARD OF DIRECTORS MATRIX
# ============================================================

@st.cache_data
def build_board_matrix(companies, board):
    """Group all board members into one column per company."""

    board_rows = []

    for _, person in board.iterrows():
        company_id = clean(person["company_id"])
        name = clean(person["executive_name"])
        title = clean(person["board_title_display"])
        source = clean(person["source_url"])

        if not company_id or not name:
            continue

        board_rows.append({
            "company_id": company_id,
            "board_member": json.dumps({
                "name": name,
                "title": title,
                "source": source
            }, ensure_ascii=False)
        })

    if board_rows:
        board_df = pd.DataFrame(board_rows)

        grouped = (
            board_df.groupby("company_id")["board_member"]
            .apply(
                lambda x: json.dumps(
                    [json.loads(item) for item in dict.fromkeys(x)],
                    ensure_ascii=False
                )
            )
            .reset_index(name="board_members")
        )
    else:
        grouped = pd.DataFrame(
            columns=["company_id", "board_members"]
        )

    result = companies.merge(
        grouped,
        on="company_id",
        how="left",
        validate="one_to_one"
    )

    result["board_members"] = (
        result["board_members"].fillna("")
    )

    return result



# ============================================================
# BOARD OF DIRECTORS GRID
# ============================================================

def show_board_grid(df):
    """Render the Board of Directors matrix."""

    grid_data = df.copy().fillna("")

    builder = GridOptionsBuilder.from_dataframe(grid_data)

    builder.configure_default_column(
        resizable=True,
        sortable=True,
        filter=False,
        wrapText=True,
        autoHeight=False,
        minWidth=110
    )

    # Company name and website link
    company_renderer = JsCode("""
    class BoardCompanyRenderer {
        init(params) {
            this.eGui = document.createElement('div');
            this.eGui.style.display = 'flex';
            this.eGui.style.alignItems = 'center';
            this.eGui.style.flexWrap = 'wrap';
            this.eGui.style.gap = '5px';

            const name = document.createElement('span');
            name.textContent = String(params.value || '');
            name.style.fontWeight = '700';
            name.style.color = '#102A43';
            this.eGui.appendChild(name);

            const website = String(params.data?.website || '').trim();

            if (website.startsWith('https://') ||
                website.startsWith('http://')) {

                const link = document.createElement('a');
                link.href = website;
                link.target = '_blank';
                link.rel = 'noopener noreferrer';
                link.title = 'Visit company website';
                link.style.display = 'inline-flex';
                link.style.alignItems = 'center';
                link.style.color = '#176B80';

                link.innerHTML = `
                    <svg xmlns="http://www.w3.org/2000/svg"
                         width="13" height="13" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor"
                         stroke-width="2" stroke-linecap="round"
                         stroke-linejoin="round">
                        <path d="M15 3h6v6"/>
                        <path d="M10 14 21 3"/>
                        <path d="M21 13v6a2 2 0 0 1-2 2H5
                                 a2 2 0 0 1-2-2V5
                                 a2 2 0 0 1 2-2h6"/>
                    </svg>
                `;

                link.addEventListener('click', event => {
                    event.stopPropagation();
                });

                this.eGui.appendChild(link);
            }
        }
        getGui() { return this.eGui; }
    }
    """)

    builder.configure_column(
        "company_name",
        header_name="Company",
        pinned="left",
        width=240,
        minWidth=180,
        cellRenderer=company_renderer,
        cellStyle={
            "backgroundColor": "#E3EDF5",
            "color": "#102A43",
            "fontWeight": "bold"
        }
    )

    builder.configure_column("website", hide=True)
    builder.configure_column(
        "company_type", header_name="Company Type", width=160
    )
    builder.configure_column(
        "latest_stage", header_name="Stage", width=150
    )
    builder.configure_column(
        "primary_therapeutic_area",
        header_name="Primary TA",
        width=210
    )

    # Board member names, titles, and clickable source icons
    board_member_renderer = JsCode("""
    class BoardMemberRenderer {
        init(params) {
            this.eGui = document.createElement('div');
            this.eGui.style.padding = '4px';
            this.eGui.style.whiteSpace = 'normal';

            const raw = String(params.value || '').trim();
            if (!raw) return;

            let members;

            try {
                members = JSON.parse(raw);
                if (!Array.isArray(members)) members = [members];
            } catch (error) {
                members = [{name: raw, title: '', source: ''}];
            }

            members.forEach((person, index) => {
                const nameRow = document.createElement('div');
                nameRow.style.display = 'flex';
                nameRow.style.alignItems = 'center';
                nameRow.style.flexWrap = 'wrap';
                nameRow.style.gap = '5px';

                const name = document.createElement('span');
                name.textContent = person.name || '';
                name.style.fontWeight = '700';
                name.style.color = '#075985';
                name.style.fontSize = '13px';
                name.style.lineHeight = '17px';
                nameRow.appendChild(name);

                const source = String(person.source || '').trim();

                if (source.startsWith('https://') ||
                    source.startsWith('http://')) {

                    const link = document.createElement('a');
                    link.href = source;
                    link.target = '_blank';
                    link.rel = 'noopener noreferrer';
                    link.title = 'View board member source';
                    link.style.display = 'inline-flex';
                    link.style.alignItems = 'center';
                    link.style.color = '#176B80';
                    link.style.flexShrink = '0';

                    link.innerHTML = `
                        <svg xmlns="http://www.w3.org/2000/svg"
                             width="13" height="13" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor"
                             stroke-width="2" stroke-linecap="round"
                             stroke-linejoin="round">
                            <path d="M15 3h6v6"/>
                            <path d="M10 14 21 3"/>
                            <path d="M21 13v6a2 2 0 0 1-2 2H5
                                     a2 2 0 0 1-2-2V5
                                     a2 2 0 0 1 2-2h6"/>
                        </svg>
                    `;

                    link.addEventListener('click', event => {
                        event.stopPropagation();
                    });

                    nameRow.appendChild(link);
                }

                const title = document.createElement('div');
                title.textContent = person.title || '';
                title.style.color = '#64748B';
                title.style.fontSize = '12px';
                title.style.lineHeight = '16px';

                this.eGui.appendChild(nameRow);
                this.eGui.appendChild(title);

                if (index < members.length - 1) {
                    const divider = document.createElement('hr');
                    divider.style.margin = '6px 0';
                    divider.style.border = '0';
                    divider.style.borderTop = '1px solid #C6DEDF';
                    this.eGui.appendChild(divider);
                }
            });
        }
        getGui() { return this.eGui; }
    }
    """)

    builder.configure_column(
        "board_members",
        header_name="Board Members",
        width=500,
        minWidth=300,
        cellRenderer=board_member_renderer,
        cellStyle=JsCode("""
            function(params) {
                const hasMembers = params.value != null &&
                    String(params.value).trim() !== "";
                return {
                    backgroundColor: hasMembers ? "#E2F3F1" : "#EDF1F4",
                    whiteSpace: "normal",
                    padding: "4px 8px",
                    display: "flex",
                    alignItems: "flex-start"
                };
            }
        """)
    )

    options = builder.build()

    # Size each row based on number of board members.
    options["getRowHeight"] = JsCode("""
        function(params) {
            const raw = String(params.data?.board_members || '').trim();
            if (!raw) return 70;

            try {
                const members = JSON.parse(raw);
                const count = Array.isArray(members) ? members.length : 1;
                return Math.max(70, count * 48 + 16);
            } catch (error) {
                return 70;
            }
        }
    """)

    options["headerHeight"] = 48
    options["suppressRowClickSelection"] = True

    # Keep export restrictions consistent with executive matrix.
    options["suppressCsvExport"] = True
    options["suppressExcelExport"] = True
    options["suppressClipboardPaste"] = True
    options["suppressCopyRowsToClipboard"] = True
    options["suppressCopySingleCellRanges"] = True
    options["suppressContextMenu"] = True
    options["enableRangeSelection"] = False
    options["cellSelection"] = False

    AgGrid(
        grid_data,
        gridOptions=options,
        height=650,
        width="100%",
        allow_unsafe_jscode=True,
        theme="streamlit",
        fit_columns_on_grid_load=False,
        key="board_matrix_aggrid"
    )


matrix = pd.read_parquet(
    DATA_DIR / "executive_matrix_precomputed.parquet"
)


st.markdown("""
<style>
    .stApp {
        background-color: #F3F7FA;
        color: #102A43;
    }

    .block-container {
        max-width: 98%;
        padding-top: 1rem;
        padding-bottom: 2rem;
    }

    h1, h2, h3 {
        color: #102A43 !important;
        font-family: Arial, sans-serif;
    }

    h1 {
        font-weight: 750 !important;
        letter-spacing: -0.8px;
    }

    [data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid #DCE6ED;
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: 0 2px 6px rgba(16, 42, 67, 0.04);
    }

    [data-testid="stMetricValue"] {
        color: #176B80;
        font-weight: 700;
    }

    [data-testid="stTextInput"] input {
        background-color: #FFFFFF;
        border-radius: 8px;
    }

    hr {
        border-color: #DCE6ED !important;
    }

    .stButton button {
        border-radius: 8px;
        border-color: #176B80;
    }
</style>
""", unsafe_allow_html=True)

st.title("🧬 Biotech Executive Intelligence")
st.caption(
    "Company-by-function leadership matrix | "
    "Research-derived records, not independently verified"
)



st.markdown("""
<style>
    /* Year Founded slider: blue/teal */
    [data-testid="stSlider"] [role="slider"] {
        background-color: #176B80 !important;
        border-color: #176B80 !important;
    }

    [data-testid="stSlider"] [data-baseweb="slider"] {
        accent-color: #176B80 !important;
    }

    [data-testid="stSlider"] [data-baseweb="slider"] > div {
        accent-color: #176B80 !important;
    }
</style>
""", unsafe_allow_html=True)

st.subheader("Company Filters")

search_col, type_col, stage_col, area_col = st.columns(4)

with search_col:
    search = st.text_input(
        "Search companies",
        placeholder="Company name..."
    )

with type_col:
    types = st.multiselect(
        "Company type",
        sorted(companies["company_type"].dropna().unique())
    )

with stage_col:
    stages = st.multiselect(
        "Development stage",
        sorted(companies["latest_stage"].dropna().unique())
    )

with area_col:
    areas = sorted({
        a.strip()
        for value in companies["all_therapeutic_areas"].dropna()
        for a in str(value).split(";")
        if a.strip()
    })
    selected_areas = st.multiselect(
        "Therapeutic area",
        areas
    )


# ADVANCED FILTERS — START

with st.expander("Advanced Filters", expanded=False):

    st.caption(
        "Refine companies using additional business, "
        "technology, financial, and geographic information."
    )

    def filter_options(field, multi_value=False):
        """Return searchable filter options from company data."""
        values = companies[field].dropna().astype(str)

        if multi_value:
            return sorted({
                item.strip()
                for value in values
                for item in value.split(";")
                if item.strip()
            }, key=str.casefold)

        return sorted({
            value.strip()
            for value in values
            if value.strip()
        }, key=str.casefold)

    def multivalue_match(value, selected):
        """Match any selected item in a semicolon-delimited field."""
        items = {
            item.strip()
            for item in str(value).split(";")
            if item.strip()
        }
        return bool(items.intersection(selected))

    adv_col1, adv_col2, adv_col3 = st.columns(3)

    with adv_col1:
        adv_business = st.multiselect(
            "Business Model",
            filter_options("business_model"),
            key="adv_business"
        )

        adv_therapeutic = st.multiselect(
            "All Therapeutic Areas",
            filter_options("all_therapeutic_areas", True),
            key="adv_therapeutic"
        )

        adv_indications = st.multiselect(
            "Indications",
            filter_options("indications", True),
            key="adv_indications"
        )

    with adv_col2:
        adv_primary_tech = st.multiselect(
            "Primary Technology",
            filter_options("primary_tech"),
            key="adv_primary_tech"
        )

        adv_all_tech = st.multiselect(
            "All Technology Areas",
            filter_options("all_tech_areas", True),
            key="adv_all_tech"
        )

        adv_ticker = st.multiselect(
            "Ticker",
            filter_options("ticker"),
            key="adv_ticker"
        )

    with adv_col3:
        adv_state = st.multiselect(
            "State",
            filter_options("state"),
            key="adv_state"
        )

        founded_values = pd.to_numeric(
            companies["year_founded"],
            errors="coerce"
        ).dropna()

        if not founded_values.empty:
            min_year = int(founded_values.min())
            max_year = int(founded_values.max())

            adv_year = st.slider(
                "Year Founded",
                min_value=min_year,
                max_value=max_year,
                value=(min_year, max_year),
                key="adv_year"
            )
        else:
            adv_year = None
            st.caption("Year Founded: No available data")

        ipo_values = pd.to_datetime(
            companies["ipo_date"],
            errors="coerce"
        ).dropna()

        if not ipo_values.empty:
            ipo_min = ipo_values.min().date()
            ipo_max = ipo_values.max().date()

            adv_ipo = st.date_input(
                "IPO Date",
                value=(),
                min_value=ipo_min,
                max_value=ipo_max,
                key="adv_ipo",
                help="Select a start and end date to filter IPO dates."
            )
        else:
            adv_ipo = ()
            st.caption("IPO Date: No available data")

# ADVANCED FILTERS — END

filtered = matrix.copy()

if search:
    filtered = filtered[
        filtered["company_name"].str.contains(
            search, case=False, na=False, regex=False
        )
    ]

if types:
    filtered = filtered[
        filtered["company_type"].isin(types)
    ]

if stages:
    filtered = filtered[
        filtered["latest_stage"].isin(stages)
    ]

if selected_areas:
    selected_set = set(selected_areas)
    filtered = filtered[
        filtered["all_therapeutic_areas"].fillna("").apply(
            lambda value: bool(
                {
                    a.strip()
                    for a in str(value).split(";")
                } & selected_set
            )
        )
    ]


# ADVANCED FILTER LOGIC — START

if adv_business:
    filtered = filtered[
        filtered["business_model"].isin(adv_business)
    ]

if adv_therapeutic:
    selected = set(adv_therapeutic)
    filtered = filtered[
        filtered["all_therapeutic_areas"].fillna("").apply(
            lambda value: multivalue_match(value, selected)
        )
    ]

if adv_indications:
    selected = set(adv_indications)
    filtered = filtered[
        filtered["indications"].fillna("").apply(
            lambda value: multivalue_match(value, selected)
        )
    ]

if adv_primary_tech:
    filtered = filtered[
        filtered["primary_tech"].isin(adv_primary_tech)
    ]

if adv_all_tech:
    selected = set(adv_all_tech)
    filtered = filtered[
        filtered["all_tech_areas"].fillna("").apply(
            lambda value: multivalue_match(value, selected)
        )
    ]

if adv_ticker:
    filtered = filtered[
        filtered["ticker"].isin(adv_ticker)
    ]

if adv_state:
    filtered = filtered[
        filtered["state"].isin(adv_state)
    ]

if adv_year is not None:
    full_year_range = (
        min_year,
        max_year
    )

    if tuple(adv_year) != full_year_range:
        years = pd.to_numeric(
            filtered["year_founded"],
            errors="coerce"
        )

        filtered = filtered[
            years.between(adv_year[0], adv_year[1])
        ]

if len(adv_ipo) == 2:
    ipo_dates = pd.to_datetime(
        filtered["ipo_date"],
        errors="coerce"
    ).dt.date

    filtered = filtered[
        (ipo_dates >= adv_ipo[0]) &
        (ipo_dates <= adv_ipo[1])
    ]

# ADVANCED FILTER LOGIC — END

st.markdown("### Dashboard View")

selected_view = st.segmented_control(
    "Select dashboard view",
    options=["Executive Leadership", "Board of Directors"],
    default="Executive Leadership",
    label_visibility="collapsed",
    key="dashboard_view"
)

if selected_view == "Executive Leadership":
    st.subheader("Executive Leadership Matrix")
    st.caption(
        f"{len(filtered):,} companies displayed. "
        "Scroll horizontally to see all functional areas."
    )

    display_columns = [
        "company_name",
        "website",
        "company_type",
        "latest_stage",
        "primary_therapeutic_area"
    ] + FUNCTIONAL_AREAS

    column_config = {
        "company_name": st.column_config.TextColumn(
            "Company", width="medium", pinned=True
        ),
        "company_type": st.column_config.TextColumn(
            "Type", width="small"
        ),
        "latest_stage": st.column_config.TextColumn(
            "Stage", width="small"
        ),
        "primary_therapeutic_area": st.column_config.TextColumn(
            "Therapeutic Area", width="medium"
        )
    }

    for function in FUNCTIONAL_AREAS:
        column_config[function] = st.column_config.TextColumn(
            function,
            width="medium"
        )

    matrix_display = (
        filtered[display_columns]
        .sort_values("company_name")
        .reset_index(drop=True)
        .fillna("")
    )

    show_executive_grid(matrix_display)

elif selected_view == "Board of Directors":
    st.subheader("Board of Directors Matrix")

    st.caption(
        f"{len(filtered):,} companies displayed. "
        "Board members are grouped by company."
    )

    # Use the same filtered company IDs as the executive matrix.
    filtered_company_ids = set(
        filtered["company_id"].astype(str)
    )

    board_matrix = pd.read_parquet(
        DATA_DIR / "board_matrix_precomputed.parquet"
    )

    board_filtered = board_matrix[
        board_matrix["company_id"].astype(str).isin(
            filtered_company_ids
        )
    ].copy()

    board_display_columns = [
        "company_name",
        "website",
        "company_type",
        "latest_stage",
        "primary_therapeutic_area",
        "board_members"
    ]

    board_matrix_display = (
        board_filtered[board_display_columns]
        .sort_values("company_name")
        .reset_index(drop=True)
        .fillna("")
    )

    show_board_grid(board_matrix_display)
