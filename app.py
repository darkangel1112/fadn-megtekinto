from __future__ import annotations

from datetime import datetime
import hashlib
import sys
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

APP_DIR = Path(__file__).resolve().parent
SRC_DIR = APP_DIR / "src"
HELP_FILE = APP_DIR / "sugo.md"
WEEKDAYS_HU = {
    0: "Hétfő",
    1: "Kedd",
    2: "Szerda",
    3: "Csütörtök",
    4: "Péntek",
    5: "Szombat",
    6: "Vasárnap",
}
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from workbook_mapper import (
    attach_sheet_names,
    build_sheet_view,
    list_farms,
    load_bulk_data,
    load_template,
    normalize_display_value,
    summarize_template,
)


st.set_page_config(
    page_title="FADN Megtekintő",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded",
)


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
            --bg-main: #071411;
            --bg-panel: rgba(7, 32, 28, 0.88);
            --line: rgba(0, 255, 208, 0.18);
            --text-main: #dbfff8;
            --text-soft: #89d8ca;
            --accent: #00f7c2;
            --accent-2: #3dd9ff;
            --accent-3: #95ff62;
        }
        .stApp {
            background:
                radial-gradient(circle at top right, rgba(61, 217, 255, 0.14), transparent 30%),
                radial-gradient(circle at top left, rgba(0, 247, 194, 0.12), transparent 28%),
                linear-gradient(180deg, #05110f 0%, #081916 45%, #06100f 100%);
            color: var(--text-main);
        }
        [data-testid="stSidebar"] {
            background:
                linear-gradient(180deg, rgba(4, 20, 17, 0.98), rgba(5, 30, 26, 0.96));
            border-right: 1px solid var(--line);
        }
        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
        .hero {
            padding: 1.3rem 1.4rem;
            border: 1px solid var(--line);
            border-radius: 18px;
            background: linear-gradient(135deg, rgba(0, 247, 194, 0.08), rgba(61, 217, 255, 0.06));
            box-shadow: 0 0 40px rgba(0, 247, 194, 0.06);
            margin-bottom: 1rem;
        }
        .hero h1 {
            margin: 0;
            font-size: 2rem;
            color: var(--text-main);
            letter-spacing: 0.02em;
        }
        .hero p {
            margin: 0.55rem 0 0;
            color: var(--text-soft);
        }
        .hero-warning {
            margin-top: 0.85rem;
            padding: 0.8rem 0.95rem;
            border-radius: 14px;
            border: 1px solid rgba(255, 215, 0, 0.18);
            background: rgba(255, 215, 0, 0.06);
            color: #fff0a8;
            line-height: 1.45;
        }
        .chip {
            display: inline-block;
            padding: 0.3rem 0.65rem;
            margin-right: 0.45rem;
            border-radius: 999px;
            border: 1px solid rgba(149, 255, 98, 0.28);
            color: #d6ffbf;
            background: rgba(149, 255, 98, 0.08);
            font-size: 0.84rem;
        }
        .stMetric {
            background: rgba(5, 28, 24, 0.72);
            border: 1px solid var(--line);
            padding: 0.75rem;
            border-radius: 16px;
        }
        .stDataFrame, .stTable {
            border: 1px solid var(--line);
            border-radius: 14px;
            overflow: hidden;
        }
        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        .stTextInput > div > div,
        .stNumberInput > div > div {
            background: rgba(5, 28, 24, 0.9);
            border-color: rgba(0, 247, 194, 0.22);
        }
        .stButton > button, .stDownloadButton > button {
            background: linear-gradient(90deg, rgba(0, 247, 194, 0.18), rgba(61, 217, 255, 0.16));
            color: var(--text-main);
            border: 1px solid rgba(61, 217, 255, 0.25);
            border-radius: 12px;
        }
        .hero-side {
            display: flex;
            flex-direction: column;
            align-items: flex-end;
            gap: 0.7rem;
        }
        .hero-clock {
            width: 100%;
            max-width: 220px;
            padding: 0.9rem 1rem;
            border-radius: 16px;
            border: 1px solid rgba(61, 217, 255, 0.22);
            background: rgba(5, 28, 24, 0.7);
            text-align: center;
        }
        .hero-clock-head {
            display: flex;
            justify-content: center;
            align-items: baseline;
            gap: 0.45rem;
            flex-wrap: wrap;
        }
        .hero-clock-date {
            color: var(--text-main);
            font-size: 1.02rem;
            font-weight: 600;
        }
        .hero-clock-day {
            color: #9cf6df;
            font-size: 0.98rem;
        }
        .hero-clock-time {
            color: #d6ffbf;
            margin-top: 0.45rem;
            font-size: 1.35rem;
            font-weight: 700;
            letter-spacing: 0.04em;
        }
        .hero-clock-author {
            margin-top: 0.5rem;
            color: var(--text-soft);
            font-size: 0.82rem;
            line-height: 1.35;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def inject_language_hints() -> None:
    components.html(
        """
        <script>
        (function () {
            const doc = window.parent.document;
            const root = doc.documentElement;
            const body = doc.body;
            root.lang = "hu";
            root.setAttribute("translate", "no");
            root.classList.add("notranslate");
            if (body) {
                body.lang = "hu";
                body.setAttribute("translate", "no");
                body.classList.add("notranslate");
            }

            const head = doc.head;

            function ensureMeta(name, content) {
                let meta = head.querySelector(`meta[name="${name}"]`);
                if (!meta) {
                    meta = doc.createElement("meta");
                    meta.setAttribute("name", name);
                    head.appendChild(meta);
                }
                meta.setAttribute("content", content);
            }

            ensureMeta("google", "notranslate");
            ensureMeta("content-language", "hu");

            const targets = doc.querySelectorAll(
                ".stApp, .main, [data-testid='stAppViewContainer'], [data-testid='stSidebar'], .block-container"
            );
            targets.forEach((node) => {
                node.setAttribute("translate", "no");
                node.classList.add("notranslate");
                node.lang = "hu";
            });
        })();
        </script>
        """,
        height=0,
        width=0,
    )


def clear_uploaded_sources() -> None:
    st.session_state["uploader_nonce"] = st.session_state.get("uploader_nonce", 0) + 1
    for key in [
        "loaded_source_signature",
        "loaded_bulk_data",
        "loaded_template_bundle",
        "loaded_bulk_name",
        "loaded_template_name",
    ]:
        st.session_state.pop(key, None)


def read_help_text() -> str:
    if HELP_FILE.exists():
        return HELP_FILE.read_text(encoding="utf-8").strip()
    return "A súgó szövege még nincs megadva."


@st.dialog("Súgó")
def show_help_dialog() -> None:
    st.markdown(read_help_text())
    if st.button("Bezárás", use_container_width=True):
        st.rerun()


def _source_signature(bulk_bytes: bytes, template_bytes: bytes) -> str:
    digest = hashlib.sha1()
    digest.update(bulk_bytes)
    digest.update(b"::")
    digest.update(template_bytes)
    return digest.hexdigest()


def load_session_sources(
    bulk_bytes: bytes,
    template_bytes: bytes,
    bulk_name: str,
    template_name: str,
) -> tuple[dict, pd.DataFrame]:
    signature = _source_signature(bulk_bytes, template_bytes)
    if st.session_state.get("loaded_source_signature") != signature:
        template_bundle = load_template(template_bytes)
        bulk_data = load_bulk_data(bulk_bytes)
        bulk_data = attach_sheet_names(bulk_data, template_bundle["row_to_sheet"])

        st.session_state["loaded_source_signature"] = signature
        st.session_state["loaded_template_bundle"] = template_bundle
        st.session_state["loaded_bulk_data"] = bulk_data
        st.session_state["loaded_bulk_name"] = bulk_name
        st.session_state["loaded_template_name"] = template_name

    return (
        st.session_state["loaded_template_bundle"],
        st.session_state["loaded_bulk_data"],
    )


def resolve_sources():
    if "uploader_nonce" not in st.session_state:
        st.session_state["uploader_nonce"] = 0

    st.sidebar.subheader("Adatforrások")
    st.sidebar.caption("Az app csak a feltöltött fájlokból dolgozik.")

    bulk_upload = st.sidebar.file_uploader(
        "Ömlesztett Excel",
        type=["xlsx", "xlsm", "xls"],
        key=f"bulk_upload_{st.session_state['uploader_nonce']}",
    )
    template_upload = st.sidebar.file_uploader(
        "Táblázatos sablon Excel",
        type=["xlsx", "xlsm", "xls"],
        key=f"template_upload_{st.session_state['uploader_nonce']}",
    )

    st.sidebar.button(
        "Források törlése",
        on_click=clear_uploaded_sources,
        use_container_width=True,
    )

    if bulk_upload is None or template_upload is None:
        st.info("A kezdéshez tölts fel egy ömlesztett és egy táblázatos Excel-fájlt a bal oldali sávban.")
        st.stop()

    bulk_bytes = bulk_upload.getvalue()
    template_bytes = template_upload.getvalue()
    bulk_name = bulk_upload.name
    template_name = template_upload.name
    return bulk_bytes, template_bytes, bulk_name, template_name


def render_header() -> None:
    now = datetime.now()
    date_text = now.strftime("%Y.%m.%d.")
    day_text = WEEKDAYS_HU[now.weekday()]
    time_text = now.strftime("%H:%M")

    left_col, right_col = st.columns([6, 1.6], gap="large")

    with left_col:
        st.markdown(
            """
            <div class="hero">
                <div class="chip">Alapverzió</div>
                <div class="chip">Ömlesztett -> Táblázatos nézet</div>
                <div class="chip">Streamlit alap</div>
                <h1>FADN Megtekintő</h1>
                <p>
                    Válassz üzemet és munkalapot, az app pedig a sablonmunkafüzet
                    alapján olvasható táblázatot épít az ömlesztett adatokból.
                </p>
                <div class="hero-warning">
                    Figyelem! A feltöltött file-ok a szerveren nem kerülnek tárolásra,
                    az ablak bezárásával minden adat törlődik.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with right_col:
        st.markdown(
            f"""
            <div class="hero hero-side">
                <div class="hero-clock">
                    <div class="hero-clock-head">
                        <div class="hero-clock-date">{date_text}</div>
                        <div class="hero-clock-day">{day_text}</div>
                    </div>
                    <div class="hero-clock-time">{time_text}</div>
                    <div class="hero-clock-author">Készítette: Kovács Péter DDRI</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Súgó", use_container_width=True):
            show_help_dialog()


def dataframe_to_csv_bytes(dataframe: pd.DataFrame) -> bytes:
    return dataframe.to_csv(index=False).encode("utf-8-sig")


def main() -> None:
    inject_styles()
    inject_language_hints()
    render_header()

    bulk_bytes, template_bytes, bulk_name, template_name = resolve_sources()

    with st.spinner("Munkafüzetek beolvasása..."):
        template_bundle, bulk_data = load_session_sources(
            bulk_bytes=bulk_bytes,
            template_bytes=template_bytes,
            bulk_name=bulk_name,
            template_name=template_name,
        )

    farms = list_farms(bulk_data)
    sheets = template_bundle["sheets"]
    template_summary = summarize_template(template_bundle)

    st.sidebar.subheader("Nézet")
    farm_code = st.sidebar.selectbox("Üzem", farms, index=0 if farms else None)
    selected_sheet_name = st.sidebar.selectbox(
        "Munkalap",
        options=list(sheets.keys()),
        format_func=lambda key: f"{key} - {sheets[key].title}",
    )
    filled_only = st.sidebar.checkbox("Csak kitöltött sorok", value=False)
    closing_only = st.sidebar.checkbox("Kísérleti: csak záró/összesítő sorok", value=False)
    show_hidden_technical = st.sidebar.checkbox("Rejtett technikai oszlopok megnyitása", value=False)

    unknown_count = int((bulk_data["sheet_name"] == "").sum())
    if unknown_count:
        st.sidebar.warning(f"{unknown_count} rekordhoz nem találtam sablonmunkalapot.")

    current_sheet = sheets[selected_sheet_name]
    view_df = build_sheet_view(
        dataframe=bulk_data,
        template=current_sheet,
        farm_code=farm_code,
        filled_only=filled_only,
        closing_only=closing_only,
    )
    full_view_df = build_sheet_view(
        dataframe=bulk_data,
        template=current_sheet,
        farm_code=farm_code,
        filled_only=False,
        closing_only=closing_only,
    )

    metric_col_1, metric_col_2, metric_col_3, metric_col_4 = st.columns(4)
    metric_col_1.metric("Üzemek száma", len(farms))
    metric_col_2.metric("Sablonlapok", len(sheets))
    metric_col_3.metric("Ömlesztett rekordok", len(bulk_data))
    metric_col_4.metric("Aktuális nézet sorai", len(view_df))

    st.caption(f"Források: `{bulk_name}` és `{template_name}`")

    tab_view, tab_mapping, tab_raw = st.tabs(["Nézet", "Leképzés", "Nyers adatok"])

    with tab_view:
        st.subheader(current_sheet.title)
        st.caption(
            f"Munkalap azonosító: `{current_sheet.sheet_name}`. "
            f"Értékoszlopok: {len(current_sheet.column_codes)}."
        )
        if filled_only:
            hidden_rows = len(full_view_df) - len(view_df)
            if hidden_rows == 0:
                st.info("Ezen a lapon most minden sorban van érdemi adat, ezért a szűrés nem változtat a listán.")
            else:
                st.caption(f"A szűrés {hidden_rows} üres vagy csak technikai `= 0` értéket tartalmazó sort rejtett el.")

        display_df = view_df.copy()
        display_df = display_df.map(normalize_display_value)

        preferred_order = [
            column
            for column in ["FarmCode", "RowCode", "RowTitle"]
            if column in display_df.columns
        ]
        value_columns = [
            column
            for column in display_df.columns
            if column not in ["FarmCode", "RowCode", "RowTitle", "DynRowSerial", "Dimension1"]
        ]
        technical_tail = []
        if show_hidden_technical:
            technical_tail = [
                column
                for column in ["DynRowSerial", "Dimension1"]
                if column in display_df.columns
            ]

        display_df = display_df[preferred_order + value_columns + technical_tail]

        st.dataframe(display_df, use_container_width=True, height=680)
        st.download_button(
            "Aktuális nézet letöltése CSV-ben",
            data=dataframe_to_csv_bytes(display_df),
            file_name=f"{farm_code}_{current_sheet.sheet_name}.csv",
            mime="text/csv",
        )

    with tab_mapping:
        st.subheader("Sablonösszefoglaló")
        st.dataframe(template_summary, use_container_width=True, height=420)

        st.subheader("Aktuális lap oszloptérképe")
        column_map_df = pd.DataFrame(
            {
                "Oszlop kód": current_sheet.column_codes,
                "Megjelenített címke": current_sheet.column_labels,
            }
        )
        st.dataframe(column_map_df, use_container_width=True, height=320)

    with tab_raw:
        st.subheader("Nyers ömlesztett rekordok az aktuális laphoz")
        raw_df = bulk_data[
            (bulk_data["akod"] == farm_code) & (bulk_data["sheet_name"] == current_sheet.sheet_name)
        ].copy()
        st.dataframe(raw_df, use_container_width=True, height=620)


if __name__ == "__main__":
    main()
