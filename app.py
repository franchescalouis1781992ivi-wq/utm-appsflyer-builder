import re
import json
import traceback
from io import BytesIO
from pathlib import Path
from datetime import datetime
import csv
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

import pandas as pd
import streamlit as st

from adapters import ADAPTER_REGISTRY
from archive_store import (
    create_archive,
    delete_archive,
    get_archive,
    init_archive_db,
    list_archives,
    update_archive,
)

st.set_page_config(page_title="UTM / AppsFlyer Builder", layout="wide")

DEFAULT_EMARSYS_PATH = (
    "/Users/chengxinyuan/patpat邮件/codex邮件看板搭建实验/"
    "Emarsys_Campaigns for Email_20260301_20260307.xlsx"
)
DEFAULT_ATTRIBULY_PATH = (
    "/Users/chengxinyuan/patpat邮件/codex邮件看板搭建实验/"
    "Attribuly更新 3.1-3.7日.xlsx"
)
MAPPING_PATH = Path("data/campaign_mapping_master.csv")
LINK_INPUT_COLUMNS = ["sequence_no", "raw_web_url", "raw_app_deeplink", "optional_label"]
DEFAULT_LINK_ROWS = [
    {
        "sequence_no": "01",
        "raw_web_url": "https://www.example.com/collections/paw",
        "raw_app_deeplink": "patpat://?action=collection_detail&collection_id=453890769139",
        "optional_label": "Hero CTA",
    },
    {
        "sequence_no": "02",
        "raw_web_url": "https://www.example.com/product/sku-123",
        "raw_app_deeplink": "patpat://?action=product_detail_new&product_id=9217060569331",
        "optional_label": "PDP CTA",
    },
    {
        "sequence_no": "03",
        "raw_web_url": "",
        "raw_app_deeplink": "",
        "optional_label": "",
    },
]
BUILDER_DEFAULTS = {
    "archive_name": "",
    "pid": "email",
    "medium": "email",
    "utm_source": "manual",
    "campaign_prefix": "emanual",
    "type": "Sale",
    "theme": "paw",
    "sender": "team",
    "year": "2026",
    "date": "0320",
    "utm_term": "0320paw",
    "af_click_lookback": "30d",
    "is_retargeting": "TRUE",
    "campaign_template": "{campaign_prefix}_{type}_{theme}_{sender}_{year}_{date}",
}
MEDIUM_OPTIONS = ["email", "app-push", "sms"]
UTM_SOURCE_OPTIONS = ["manual", "automation"]
TYPE_OPTIONS = ["Matching", "Sale", "Bamboo", "Disney", "Ip", "Kids", "Holiday", "Cruise", "Test"]
CAMPAIGN_PREFIX_OPTIONS = [
    "emanual",
    "eautomation",
    "smanual",
    "sautomation",
    "pmanual",
    "pautomation",
]
CUSTOM_OPTION_LABEL = "自定义输入..."


def inject_apple_panel_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
            --app-bg: linear-gradient(180deg, #f5f5f7 0%, #eef1f5 100%);
            --panel-bg: rgba(255, 255, 255, 0.9);
            --panel-border: rgba(15, 23, 42, 0.08);
            --panel-shadow: 0 18px 50px rgba(15, 23, 42, 0.08);
            --text-strong: #111827;
            --text-soft: #6b7280;
            --accent: #111111;
            --accent-soft: #e8eefc;
            --line: rgba(15, 23, 42, 0.08);
        }

        .stApp {
            background: var(--app-bg);
            color: var(--text-strong);
        }

        html, body, [class*="css"] {
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Helvetica Neue", sans-serif;
        }

        .block-container {
            max-width: 1440px;
            padding-top: 28px;
            padding-bottom: 64px;
        }

        h1, h2, h3 {
            letter-spacing: -0.03em;
            color: var(--text-strong);
        }

        div[data-testid="stMetric"] {
            background: var(--panel-bg);
            border: 1px solid var(--panel-border);
            border-radius: 22px;
            padding: 18px 18px 12px 18px;
            box-shadow: var(--panel-shadow);
        }

        div[data-testid="stMetricLabel"] {
            color: var(--text-soft);
            font-weight: 600;
        }

        div[data-testid="stMetricValue"] {
            font-size: 1.6rem;
            font-weight: 650;
        }

        div[data-testid="stTextInputRootElement"] > div,
        div[data-testid="stTextAreaRootElement"] > div,
        div[data-baseweb="select"] > div,
        div[data-testid="stFileUploaderDropzone"],
        div[data-testid="stExpander"] {
            background: var(--panel-bg);
            border: 1px solid var(--panel-border);
            border-radius: 18px;
            box-shadow: 0 10px 30px rgba(15, 23, 42, 0.05);
        }

        div[data-testid="stTextInputRootElement"] input,
        div[data-testid="stTextAreaRootElement"] textarea {
            font-size: 15px;
        }

        div[data-testid="stButton"] > button,
        div[data-testid="stDownloadButton"] > button {
            border-radius: 999px;
            border: 1px solid rgba(17, 24, 39, 0.08);
            background: rgba(255, 255, 255, 0.9);
            color: var(--text-strong);
            font-weight: 600;
            min-height: 46px;
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.06);
        }

        div[data-testid="stButton"] > button[kind="primary"] {
            background: linear-gradient(180deg, #111827 0%, #1f2937 100%);
            color: white;
            border: none;
        }

        div[data-testid="stTabs"] button {
            border-radius: 999px;
            padding: 10px 18px;
            color: var(--text-soft);
        }

        div[data-testid="stTabs"] button[aria-selected="true"] {
            background: white;
            color: var(--text-strong);
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08);
        }

        div[data-testid="stDataFrame"] {
            background: var(--panel-bg);
            border: 1px solid var(--panel-border);
            border-radius: 24px;
            padding: 10px;
            box-shadow: var(--panel-shadow);
        }

        div[data-testid="stAlert"] {
            border-radius: 20px;
            border: 1px solid var(--panel-border);
        }

        .apple-hero {
            background: linear-gradient(135deg, rgba(255,255,255,0.95) 0%, rgba(246,247,251,0.92) 100%);
            border: 1px solid rgba(255,255,255,0.7);
            border-radius: 30px;
            padding: 28px 30px;
            box-shadow: 0 24px 60px rgba(15, 23, 42, 0.08);
            margin-bottom: 18px;
        }

        .apple-kicker {
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            color: #6b7280;
            margin-bottom: 10px;
            font-weight: 700;
        }

        .apple-title {
            font-size: 2rem;
            line-height: 1.1;
            margin: 0 0 10px 0;
            color: #111827;
            font-weight: 700;
        }

        .apple-subtitle {
            margin: 0;
            color: #6b7280;
            font-size: 1rem;
            line-height: 1.6;
            max-width: 760px;
        }

        .apple-section {
            background: var(--panel-bg);
            border: 1px solid var(--panel-border);
            border-radius: 28px;
            padding: 22px 24px 8px 24px;
            box-shadow: var(--panel-shadow);
            margin: 14px 0 18px 0;
        }

        .apple-section h3 {
            margin-top: 0;
            margin-bottom: 4px;
        }

        .apple-section p {
            color: var(--text-soft);
            margin-top: 0;
            margin-bottom: 18px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_apple_hero(title: str, subtitle: str, kicker: str = "Link Builder") -> None:
    st.markdown(
        f"""
        <div class="apple-hero">
            <div class="apple-kicker">{kicker}</div>
            <div class="apple-title">{title}</div>
            <p class="apple-subtitle">{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_select_with_custom_input(
    container,
    label: str,
    options: list[str],
    current_value: str,
    key_prefix: str,
    form_revision: int,
    placeholder: str = "输入自定义值",
) -> str:
    current_text = coerce_text(current_value)
    select_options = options + [CUSTOM_OPTION_LABEL]
    default_option = current_text if current_text in options else CUSTOM_OPTION_LABEL

    selected = container.selectbox(
        label,
        select_options,
        index=select_options.index(default_option),
        key=f"{key_prefix}_mode_widget_{form_revision}",
    )
    if selected == CUSTOM_OPTION_LABEL:
        custom_value = container.text_input(
            f"{label}（自定义）",
            value="" if current_text in options else current_text,
            key=f"{key_prefix}_custom_widget_{form_revision}",
            placeholder=placeholder,
        )
        return coerce_text(custom_value)

    return selected


def render_section_shell(title: str, subtitle: str) -> None:
    st.markdown(
        f"""
        <div class="apple-section">
            <h3>{title}</h3>
            <p>{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_step_shell(step_no: str, title: str, subtitle: str, tone: str = "#111827") -> None:
    st.markdown(
        f"""
        <div class="apple-section" style="border-color: rgba(15, 23, 42, 0.06);">
            <div style="display:flex;align-items:center;gap:12px;margin-bottom:8px;">
                <div style="width:34px;height:34px;border-radius:999px;background:{tone};color:white;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;">
                    {step_no}
                </div>
                <h3 style="margin:0;">{title}</h3>
            </div>
            <p>{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def normalize_text(value: str) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip().lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    return re.sub(r"_+", "_", s).strip("_")


def parse_campaign_date(name: str):
    if pd.isna(name):
        return pd.NaT
    s = str(name)
    patterns = [
        r"(20\d{2})[-_]?([01]\d)[-_]?([0-3]\d)",  # 20260307 / 2026_03_07
        r"\b(\d{2})([01]\d)([0-3]\d)\b",  # 260307
    ]

    m = re.search(patterns[0], s)
    if m:
        y, mo, d = m.groups()
        return pd.to_datetime(f"{y}-{mo}-{d}", errors="coerce")

    m = re.search(patterns[1], s)
    if m:
        y2, mo, d = m.groups()
        return pd.to_datetime(f"20{y2}-{mo}-{d}", errors="coerce")

    return pd.NaT


@st.cache_data
def load_emarsys(path_or_file) -> pd.DataFrame:
    raw = pd.read_excel(path_or_file)
    header = raw.iloc[0].tolist()
    df = raw.iloc[1:].copy()
    df.columns = header

    numeric_cols = [
        "Sent",
        "Delivered",
        "Opened",
        "Clicked",
        "Unsubscribed",
        "Bounced",
        "Open Rate(%)",
        "Click Rate(%)",
        "Unsubscribe Rate(%)",
    ]
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["campaign_name"] = df["Campaign Name"].astype(str)
    df["campaign_norm"] = df["campaign_name"].map(normalize_text)
    df["send_date"] = df["campaign_name"].map(parse_campaign_date)

    df["CTOR(%)"] = (df["Clicked"] / df["Opened"].replace(0, pd.NA) * 100).fillna(0)
    return df


@st.cache_data
def load_attribuly(path_or_file) -> pd.DataFrame:
    df = pd.read_excel(path_or_file)

    email_df = df[df["Medium"].astype(str).str.lower() == "email"].copy()
    email_df["campaign_name"] = email_df["Campaign(Source/Medium)"].astype(str)
    email_df["campaign_norm"] = email_df["campaign_name"].map(normalize_text)
    email_df["send_date"] = email_df["campaign_name"].map(parse_campaign_date)

    for c in ["Clicks(UV)", "Attribuly tracked conversions", "Attribuly tracked conversion value"]:
        if c in email_df.columns:
            email_df[c] = pd.to_numeric(email_df[c], errors="coerce").fillna(0)

    email_df = email_df.rename(
        columns={
            "Clicks(UV)": "UV",
            "Attribuly tracked conversions": "Orders",
            "Attribuly tracked conversion value": "GMV",
        }
    )

    email_df["UV-Order CVR(%)"] = (email_df["Orders"] / email_df["UV"].replace(0, pd.NA) * 100).fillna(0)
    email_df["AOV"] = (email_df["GMV"] / email_df["Orders"].replace(0, pd.NA)).fillna(0)

    return email_df


def ensure_mapping(emarsys_df: pd.DataFrame, attribuly_df: pd.DataFrame) -> pd.DataFrame:
    if MAPPING_PATH.exists():
        mp = pd.read_csv(MAPPING_PATH)
        return mp

    em = emarsys_df[["campaign_norm", "campaign_name", "send_date"]].drop_duplicates().copy()
    at = attribuly_df[["campaign_norm", "campaign_name", "send_date"]].drop_duplicates().copy()

    merged = em.merge(
        at,
        on="campaign_norm",
        how="outer",
        suffixes=("_emarsys", "_attribuly"),
    )
    merged = merged.rename(
        columns={
            "campaign_norm": "normalized_campaign_id",
            "campaign_name_emarsys": "emarsys_campaign_name",
            "campaign_name_attribuly": "attribuly_campaign_name",
        }
    )

    merged["normalized_campaign_name"] = merged["normalized_campaign_id"]
    merged["send_date"] = merged["send_date_emarsys"].combine_first(merged["send_date_attribuly"])
    merged["site_country"] = ""
    merged["campaign_type"] = ""
    merged["theme"] = ""
    merged["audience_type"] = ""
    merged["offer_type"] = ""
    merged["version"] = ""
    merged["owner"] = ""
    merged["mapping_status"] = merged.apply(
        lambda r: "matched"
        if pd.notna(r.get("emarsys_campaign_name")) and pd.notna(r.get("attribuly_campaign_name"))
        else "unmatched",
        axis=1,
    )
    merged["notes"] = ""

    out_cols = [
        "normalized_campaign_id",
        "normalized_campaign_name",
        "emarsys_campaign_name",
        "attribuly_campaign_name",
        "send_date",
        "site_country",
        "campaign_type",
        "theme",
        "audience_type",
        "offer_type",
        "version",
        "owner",
        "mapping_status",
        "notes",
    ]
    merged = merged[out_cols].sort_values("normalized_campaign_id")

    MAPPING_PATH.parent.mkdir(parents=True, exist_ok=True)
    merged.to_csv(MAPPING_PATH, index=False)
    return merged


def metric_card(col, label: str, value, fmt: str = "{:.2f}"):
    with col:
        if isinstance(value, (int, float)):
            st.metric(label, fmt.format(value))
        else:
            st.metric(label, value)


def safe_div(a, b):
    return a / b if b else 0


def coerce_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_sequence(value) -> str:
    raw = coerce_text(value)
    if not raw:
        return ""
    digits = re.sub(r"\D", "", raw)
    if digits:
        return digits.zfill(2)
    return raw


def build_utm_campaign(template: str, values: dict[str, str]) -> str:
    campaign = template.format(**values)
    campaign = re.sub(r"\s+", "_", campaign.strip())
    campaign = re.sub(r"_+", "_", campaign)
    return campaign.strip("_")


def is_patpat_deeplink(url: str) -> bool:
    value = coerce_text(url)
    return bool(re.match(r"^patpat:/?", value, flags=re.IGNORECASE))


def normalize_web_url(url: str) -> str:
    value = coerce_text(url)
    if not value or is_patpat_deeplink(value):
        return value

    parsed = urlsplit(value)
    if parsed.scheme.lower() in {"http", "https"}:
        return value

    if re.match(r"^www\.", value, flags=re.IGNORECASE):
        return f"https://{value}"

    if re.match(r"^(?:[a-z0-9-]+\.)+[a-z]{2,}(?:/.*)?$", value, flags=re.IGNORECASE):
        return f"https://{value}"

    return value


def restore_browser_rewritten_patpat_url(url: str) -> str:
    value = coerce_text(url)
    if not value:
        return ""
    restored = re.sub(r"^https?://patpat//", "patpat://", value, flags=re.IGNORECASE)
    restored = re.sub(r"^https?://patpat/\?", "patpat://?", restored, flags=re.IGNORECASE)
    return restored


def normalize_app_deeplink(url: str) -> str:
    value = coerce_text(url)
    if not value:
        return ""
    value = restore_browser_rewritten_patpat_url(value)
    if not is_patpat_deeplink(value):
        return value

    rest = re.sub(r"^patpat:", "", value, flags=re.IGNORECASE)
    rest = rest.lstrip("/")
    if rest.startswith("?"):
        return f"patpat://{rest}"
    if rest:
        return f"patpat://{rest}"
    return "patpat://"


def append_query_params(raw_url: str, params: dict[str, str]) -> str:
    url = normalize_web_url(raw_url)
    if not url:
        return ""

    parsed = urlsplit(url)
    if parsed.scheme and parsed.scheme.lower() not in {"http", "https"}:
        return url

    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    for key, value in params.items():
        if value != "":
            query[key] = value

    return urlunsplit(
        (
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            urlencode(query, doseq=True),
            parsed.fragment,
        )
    )


def append_app_tracking_params(raw_deeplink: str, params: dict[str, str]) -> str:
    deeplink = normalize_app_deeplink(raw_deeplink)
    if not deeplink or not is_patpat_deeplink(deeplink):
        return deeplink

    body = deeplink[len("patpat://") :]
    if "?" in body:
        path, query_string = body.split("?", 1)
    else:
        path, query_string = body, ""

    query = dict(parse_qsl(query_string, keep_blank_values=True))
    for key, value in params.items():
        if value != "":
            query[key] = value

    rebuilt = f"patpat://{path}"
    encoded_query = urlencode(query, doseq=True)
    if encoded_query:
        rebuilt = f"{rebuilt}?{encoded_query}"
    return rebuilt


def build_outputs(
    campaign_config: dict[str, str],
    link_rows: pd.DataFrame,
    adapter_key: str = "local_passthrough",
) -> tuple[pd.DataFrame, pd.DataFrame]:
    raw_rows = normalize_input_rows(link_rows)
    working_rows = raw_rows.copy()
    working_rows["normalized_web_url"] = working_rows["raw_web_url"].map(normalize_web_url)
    working_rows["final_app_deeplink"] = working_rows["raw_app_deeplink"].map(normalize_app_deeplink)

    working_rows = working_rows[
        get_active_row_mask(working_rows)
    ].copy()
    working_rows["sequence_no"] = working_rows["sequence_no"].map(normalize_sequence)
    auto_counter = 1
    generated_sequence: list[str] = []
    for value in working_rows["sequence_no"].tolist():
        if value:
            generated_sequence.append(value)
        else:
            generated_sequence.append(str(auto_counter).zfill(2))
        auto_counter += 1
    working_rows["sequence_no"] = generated_sequence

    if working_rows.empty:
        preview_cols = [
            "sequence_no",
            "optional_label",
            "raw_web_url",
            "full_web_url",
            "raw_app_deeplink",
            "final_app_deeplink",
            "utm_campaign",
            "utm_term",
            "c",
        ]
        af_cols = [
            "pid",
            "c",
            "af_click_lookback",
            "deep_link_value",
            "af_dp",
            "af_web_dp",
            "af_android_url",
            "af_ios_url",
            "is_retargeting",
        ]
        return pd.DataFrame(columns=preview_cols), pd.DataFrame(columns=af_cols)

    utm_campaign = build_utm_campaign(campaign_config["campaign_template"], campaign_config)
    utm_term = campaign_config["utm_term"]
    tracking_params = {
        "utm_medium": campaign_config["medium"],
        "utm_source": campaign_config["utm_source"],
        "utm_campaign": utm_campaign,
        "utm_term": utm_term,
    }

    working_rows["utm_campaign"] = utm_campaign
    working_rows["utm_term"] = utm_term
    working_rows["c"] = working_rows["sequence_no"].map(lambda seq: f"{utm_term}{seq}" if seq else utm_term)
    working_rows["full_web_url"] = working_rows["normalized_web_url"].map(
        lambda raw_url: append_query_params(
            raw_url,
            tracking_params,
        )
    )
    working_rows["final_app_deeplink"] = working_rows["raw_app_deeplink"].map(
        lambda deeplink: append_app_tracking_params(deeplink, tracking_params)
    )

    adapter = ADAPTER_REGISTRY[adapter_key]
    transformed_rows = working_rows.apply(
        lambda row: adapter.transform(
            web_url=row["full_web_url"],
            app_deeplink=row["final_app_deeplink"],
        ),
        axis=1,
        result_type="expand",
    )
    appsflyer_df = pd.DataFrame(
        {
            "pid": campaign_config["pid"],
            "c": working_rows["c"],
            "af_click_lookback": campaign_config["af_click_lookback"],
            "deep_link_value": transformed_rows["deep_link_value"],
            "af_dp": transformed_rows["af_dp"],
            "af_web_dp": transformed_rows["af_web_dp"],
            "af_android_url": transformed_rows["af_android_url"],
            "af_ios_url": transformed_rows["af_ios_url"],
            "is_retargeting": campaign_config["is_retargeting"],
        }
    )

    preview_df = working_rows[
        [
            "sequence_no",
            "optional_label",
            "raw_web_url",
            "full_web_url",
            "raw_app_deeplink",
            "final_app_deeplink",
            "utm_campaign",
            "utm_term",
            "c",
        ]
    ].copy()

    return preview_df, appsflyer_df


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def dataframe_to_xlsx_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()


def dataframe_to_tsv(df: pd.DataFrame) -> str:
    return df.to_csv(index=False, sep="\t")


def now_local_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_default_archive_name(values: dict[str, str]) -> str:
    parts = [
        coerce_text(values.get("date")),
        coerce_text(values.get("medium")),
        coerce_text(values.get("type")),
        coerce_text(values.get("theme")),
        coerce_text(values.get("sender")),
    ]
    return "_".join([part for part in parts if part]) or "untitled_campaign"


def ensure_link_builder_state() -> None:
    for field, default_value in BUILDER_DEFAULTS.items():
        state_key = f"draft_{field}"
        if state_key not in st.session_state:
            st.session_state[state_key] = default_value
    if "link_builder_rows" not in st.session_state:
        st.session_state["link_builder_rows"] = DEFAULT_LINK_ROWS.copy()
    if "current_archive_id" not in st.session_state:
        st.session_state["current_archive_id"] = None
    if "form_revision" not in st.session_state:
        st.session_state["form_revision"] = st.session_state.get("editor_revision", 0)
    if "link_builder_last_saved_at" not in st.session_state:
        st.session_state["link_builder_last_saved_at"] = None
    if "link_builder_saved_signature" not in st.session_state:
        st.session_state["link_builder_saved_signature"] = ""
    if "link_builder_saved_result_json" not in st.session_state:
        st.session_state["link_builder_saved_result_json"] = ""
    if "link_builder_pending_action" not in st.session_state:
        st.session_state["link_builder_pending_action"] = None
    if "link_builder_confirm_delete_id" not in st.session_state:
        st.session_state["link_builder_confirm_delete_id"] = None
    if "link_builder_history_search" not in st.session_state:
        st.session_state["link_builder_history_search"] = ""
    if "link_builder_validation_errors" not in st.session_state:
        st.session_state["link_builder_validation_errors"] = []
    if "link_builder_ui_error" not in st.session_state:
        st.session_state["link_builder_ui_error"] = ""
    if "link_builder_opened_archive_result" not in st.session_state:
        st.session_state["link_builder_opened_archive_result"] = None
    if "link_builder_show_live_results" not in st.session_state:
        st.session_state["link_builder_show_live_results"] = True
    if "link_builder_last_draft_snapshot" not in st.session_state:
        st.session_state["link_builder_last_draft_snapshot"] = None
    if "link_builder_bulk_paste" not in st.session_state:
        st.session_state["link_builder_bulk_paste"] = ""
    if "link_builder_column_sequence" not in st.session_state:
        st.session_state["link_builder_column_sequence"] = ""
    if "link_builder_column_web" not in st.session_state:
        st.session_state["link_builder_column_web"] = ""
    if "link_builder_column_app" not in st.session_state:
        st.session_state["link_builder_column_app"] = ""
    if "link_builder_column_label" not in st.session_state:
        st.session_state["link_builder_column_label"] = ""


def collect_campaign_context() -> dict[str, str]:
    return {
        "archive_name": coerce_text(st.session_state.get("draft_archive_name")),
        "pid": coerce_text(st.session_state.get("draft_pid")),
        "medium": coerce_text(st.session_state.get("draft_medium")),
        "utm_source": coerce_text(st.session_state.get("draft_utm_source")),
        "campaign_prefix": coerce_text(st.session_state.get("draft_campaign_prefix")),
        "type": coerce_text(st.session_state.get("draft_type")),
        "theme": coerce_text(st.session_state.get("draft_theme")),
        "sender": coerce_text(st.session_state.get("draft_sender")),
        "year": coerce_text(st.session_state.get("draft_year")),
        "date": coerce_text(st.session_state.get("draft_date")),
        "utm_term": coerce_text(st.session_state.get("draft_utm_term")),
        "af_click_lookback": coerce_text(st.session_state.get("draft_af_click_lookback")),
        "is_retargeting": coerce_text(st.session_state.get("draft_is_retargeting")),
        "campaign_template": coerce_text(st.session_state.get("draft_campaign_template")),
    }


def normalize_input_rows(rows) -> pd.DataFrame:
    if isinstance(rows, pd.DataFrame):
        df = rows.copy()
    else:
        df = pd.DataFrame(rows)
    for col in LINK_INPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    for col in LINK_INPUT_COLUMNS:
        df[col] = df[col].map(coerce_text)
    df["raw_app_deeplink"] = df["raw_app_deeplink"].map(restore_browser_rewritten_patpat_url)
    return df[LINK_INPUT_COLUMNS].fillna("").copy()


def get_display_rows(rows) -> pd.DataFrame:
    df = normalize_input_rows(rows)
    used_mask = df.apply(lambda row: any(bool(coerce_text(v)) for v in row), axis=1)
    used_rows = df[used_mask].copy()
    if used_rows.empty:
        return pd.DataFrame(DEFAULT_LINK_ROWS)
    return used_rows.reset_index(drop=True)


def add_blank_row() -> None:
    rows_df = normalize_input_rows(st.session_state.get("link_builder_rows", DEFAULT_LINK_ROWS))
    rows_df.loc[len(rows_df)] = {col: "" for col in LINK_INPUT_COLUMNS}
    st.session_state["link_builder_rows"] = rows_df.to_dict(orient="records")
    st.session_state["form_revision"] += 1


def delete_row_at(index_to_delete: int) -> None:
    rows_df = get_display_rows(st.session_state.get("link_builder_rows", DEFAULT_LINK_ROWS))
    if 0 <= index_to_delete < len(rows_df):
        rows_df = rows_df.drop(index=index_to_delete).reset_index(drop=True)
    if rows_df.empty:
        rows_df = pd.DataFrame(DEFAULT_LINK_ROWS)
    st.session_state["link_builder_rows"] = normalize_input_rows(rows_df).to_dict(orient="records")
    st.session_state["form_revision"] += 1


def ensure_minimum_rows(rows, minimum: int = 3) -> pd.DataFrame:
    df = normalize_input_rows(rows)
    while len(df) < minimum:
        df.loc[len(df)] = {col: "" for col in LINK_INPUT_COLUMNS}
    return df.reset_index(drop=True)


def parse_bulk_rows(raw_text: str) -> pd.DataFrame:
    text = coerce_text(raw_text)
    if not text:
        return pd.DataFrame(columns=LINK_INPUT_COLUMNS)

    lines = [line for line in text.splitlines() if line.strip()]
    parsed_rows: list[dict[str, str]] = []

    for line in lines:
        if "\t" in line:
            parts = next(csv.reader([line], delimiter="\t"))
        else:
            parts = next(csv.reader([line]))

        parts = [coerce_text(part) for part in parts]
        if len(parts) == 1:
            value = parts[0]
            if is_patpat_deeplink(value) or re.match(r"^https?://", value, flags=re.IGNORECASE):
                parsed_rows.append(
                    {
                        "sequence_no": "",
                        "raw_web_url": "" if is_patpat_deeplink(value) else value,
                        "raw_app_deeplink": value if is_patpat_deeplink(value) else "",
                        "optional_label": "",
                    }
                )
            continue

        if len(parts) == 2:
            first, second = parts
            parsed_rows.append(
                {
                    "sequence_no": "",
                    "raw_web_url": first if not is_patpat_deeplink(first) else "",
                    "raw_app_deeplink": second if is_patpat_deeplink(second) else (first if is_patpat_deeplink(first) else ""),
                    "optional_label": second if (not is_patpat_deeplink(second) and not re.match(r"^https?://", second, flags=re.IGNORECASE)) else "",
                }
            )
            continue

        if len(parts) == 3:
            parsed_rows.append(
                {
                    "sequence_no": parts[0] if not re.match(r"^https?://|^patpat:/?", parts[0], flags=re.IGNORECASE) else "",
                    "raw_web_url": parts[1] if len(parts) > 1 else "",
                    "raw_app_deeplink": parts[2] if len(parts) > 2 else "",
                    "optional_label": "",
                }
            )
            continue

        row = {
            "sequence_no": parts[0] if len(parts) > 0 else "",
            "raw_web_url": parts[1] if len(parts) > 1 else "",
            "raw_app_deeplink": parts[2] if len(parts) > 2 else "",
            "optional_label": parts[3] if len(parts) > 3 else "",
        }
        parsed_rows.append(row)

    return normalize_input_rows(parsed_rows)


def split_multiline_values(raw_text: str) -> list[str]:
    text = coerce_text(raw_text)
    if not text:
        return []
    return [coerce_text(line) for line in text.splitlines()]


def build_rows_from_columns(
    sequence_text: str,
    web_text: str,
    app_text: str,
    label_text: str,
) -> pd.DataFrame:
    sequence_values = split_multiline_values(sequence_text)
    web_values = split_multiline_values(web_text)
    app_values = split_multiline_values(app_text)
    label_values = split_multiline_values(label_text)

    max_len = max(len(sequence_values), len(web_values), len(app_values), len(label_values), 0)
    rows: list[dict[str, str]] = []
    for index in range(max_len):
        rows.append(
            {
                "sequence_no": sequence_values[index] if index < len(sequence_values) else "",
                "raw_web_url": web_values[index] if index < len(web_values) else "",
                "raw_app_deeplink": app_values[index] if index < len(app_values) else "",
                "optional_label": label_values[index] if index < len(label_values) else "",
            }
        )
    return normalize_input_rows(rows)


def get_active_row_mask(df: pd.DataFrame) -> pd.Series:
    normalized = normalize_input_rows(df)
    return normalized[["raw_web_url", "raw_app_deeplink", "optional_label"]].apply(
        lambda row: any(bool(coerce_text(v)) for v in row),
        axis=1,
    )


def validate_link_rows(rows) -> list[str]:
    df = normalize_input_rows(rows)
    active_mask = get_active_row_mask(df)
    errors: list[str] = []

    for idx, row in df[active_mask].reset_index(drop=False).iterrows():
        row_no = int(row["index"]) + 1
        raw_web_url = coerce_text(row["raw_web_url"])
        raw_app_deeplink = coerce_text(row["raw_app_deeplink"])

        if raw_web_url:
            normalized_web_url = normalize_web_url(raw_web_url)
            parsed = urlsplit(normalized_web_url)
            if parsed.scheme.lower() not in {"http", "https"} or not parsed.netloc:
                errors.append(f"Row {row_no}: raw_web_url is invalid")

        if raw_app_deeplink:
            if re.match(r"^https?://patpat", raw_app_deeplink, flags=re.IGNORECASE):
                errors.append(
                    f"Row {row_no}: App deeplink was incorrectly normalized. Please keep it as patpat://..."
                )
                continue
            normalized_deeplink = normalize_app_deeplink(raw_app_deeplink)
            if not is_patpat_deeplink(normalized_deeplink):
                errors.append(f"Row {row_no}: raw_app_deeplink must start with patpat://")

    return errors


def build_archive_payload(
    campaign_context: dict[str, str],
    input_df: pd.DataFrame,
    preview_df: pd.DataFrame,
    appsflyer_df: pd.DataFrame,
) -> dict:
    archive_name = campaign_context["archive_name"] or get_default_archive_name(campaign_context)
    generated_at = now_local_text()
    return {
        "archive_name": archive_name,
        "utm_campaign": preview_df["utm_campaign"].iloc[0] if not preview_df.empty else "",
        "theme": campaign_context["theme"],
        "sender": campaign_context["sender"],
        "campaign_date": campaign_context["date"],
        "utm_term": campaign_context["utm_term"],
        "metadata": campaign_context,
        "rows": normalize_input_rows(input_df).fillna("").to_dict(orient="records"),
        "result": {
            "generated_at": generated_at,
            "preview_rows": preview_df.fillna("").to_dict(orient="records"),
            "appsflyer_rows": appsflyer_df.fillna("").to_dict(orient="records"),
        },
    }


def payload_signature(campaign_context: dict[str, str], input_df: pd.DataFrame) -> str:
    payload = {
        "metadata": campaign_context,
        "rows": normalize_input_rows(input_df).fillna("").to_dict(orient="records"),
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True)


def load_archive_into_state(archive: dict, as_duplicate: bool = False) -> None:
    metadata = archive.get("metadata", {})
    if "utm_term" not in metadata and "utm_term_base" in metadata:
        metadata["utm_term"] = metadata["utm_term_base"]
    for field in BUILDER_DEFAULTS:
        st.session_state[f"draft_{field}"] = metadata.get(field, BUILDER_DEFAULTS[field])

    rows_df = normalize_input_rows(archive.get("rows", DEFAULT_LINK_ROWS))
    st.session_state["link_builder_rows"] = rows_df.to_dict(orient="records")
    st.session_state["form_revision"] += 1

    if as_duplicate:
        archive_name = metadata.get("archive_name") or archive.get("archive_name") or get_default_archive_name(metadata)
        st.session_state["draft_archive_name"] = f"{archive_name}_copy"
        st.session_state["current_archive_id"] = None
        st.session_state["link_builder_last_saved_at"] = None
        st.session_state["link_builder_saved_signature"] = ""
        st.session_state["link_builder_saved_result_json"] = ""
        st.session_state["link_builder_validation_errors"] = []
        st.session_state["link_builder_ui_error"] = f"已复制活动：{archive_name}"
        st.session_state["link_builder_opened_archive_result"] = archive.get("result", {})
        st.session_state["link_builder_show_live_results"] = False
        return

    st.session_state["current_archive_id"] = archive.get("archive_id")
    st.session_state["link_builder_last_saved_at"] = archive.get("updated_at")
    st.session_state["link_builder_saved_signature"] = json.dumps(
        {
            "metadata": archive.get("metadata", {}),
            "rows": archive.get("rows", []),
        },
        ensure_ascii=False,
        sort_keys=True,
    )
    st.session_state["link_builder_saved_result_json"] = json.dumps(archive.get("result", {}), ensure_ascii=False, sort_keys=True)
    st.session_state["link_builder_validation_errors"] = []
    st.session_state["link_builder_ui_error"] = f"已打开活动：{archive.get('archive_name', archive.get('archive_id'))}"
    st.session_state["link_builder_opened_archive_result"] = archive.get("result", {})
    st.session_state["link_builder_show_live_results"] = False


def get_current_activity_name() -> str:
    archive_name = coerce_text(st.session_state.get("draft_archive_name"))
    if archive_name:
        return archive_name
    return get_default_archive_name(collect_campaign_context())


def snapshot_current_draft() -> dict:
    metadata = {field: st.session_state.get(f"draft_{field}", BUILDER_DEFAULTS[field]) for field in BUILDER_DEFAULTS}
    return {
        "metadata": metadata,
        "rows": normalize_input_rows(st.session_state.get("link_builder_rows", DEFAULT_LINK_ROWS)).to_dict(orient="records"),
        "current_archive_id": st.session_state.get("current_archive_id"),
        "last_saved_at": st.session_state.get("link_builder_last_saved_at"),
        "saved_signature": st.session_state.get("link_builder_saved_signature", ""),
        "saved_result_json": st.session_state.get("link_builder_saved_result_json", ""),
        "opened_archive_result": st.session_state.get("link_builder_opened_archive_result"),
        "show_live_results": st.session_state.get("link_builder_show_live_results", True),
    }


def restore_draft_snapshot(snapshot: dict | None) -> None:
    if not snapshot:
        return
    metadata = snapshot.get("metadata", {})
    for field in BUILDER_DEFAULTS:
        st.session_state[f"draft_{field}"] = metadata.get(field, BUILDER_DEFAULTS[field])
    st.session_state["link_builder_rows"] = snapshot.get("rows", DEFAULT_LINK_ROWS)
    st.session_state["current_archive_id"] = snapshot.get("current_archive_id")
    st.session_state["link_builder_last_saved_at"] = snapshot.get("last_saved_at")
    st.session_state["link_builder_saved_signature"] = snapshot.get("saved_signature", "")
    st.session_state["link_builder_saved_result_json"] = snapshot.get("saved_result_json", "")
    st.session_state["link_builder_opened_archive_result"] = snapshot.get("opened_archive_result")
    st.session_state["link_builder_show_live_results"] = snapshot.get("show_live_results", True)
    st.session_state["form_revision"] += 1


def save_archive_payload(payload: dict, mode: str) -> int | None:
    archive_name = payload["archive_name"]
    timestamp = now_local_text()
    current_archive_id = st.session_state.get("current_archive_id")

    if mode == "update":
        if current_archive_id is None:
            return None
        update_archive(
            archive_id=current_archive_id,
            archive_name=archive_name,
            utm_campaign=payload["utm_campaign"],
            theme=payload["theme"],
            sender=payload["sender"],
            campaign_date=payload["campaign_date"],
            utm_term_base=payload["utm_term"],
            updated_at=timestamp,
            metadata=payload["metadata"],
            rows=payload["rows"],
            result=payload["result"],
        )
        st.session_state["link_builder_last_saved_at"] = timestamp
        saved_id = current_archive_id
    else:
        saved_id = create_archive(
            archive_name=archive_name,
            utm_campaign=payload["utm_campaign"],
            theme=payload["theme"],
            sender=payload["sender"],
            campaign_date=payload["campaign_date"],
            utm_term_base=payload["utm_term"],
            created_at=timestamp,
            updated_at=timestamp,
            metadata=payload["metadata"],
            rows=payload["rows"],
            result=payload["result"],
        )
        st.session_state["current_archive_id"] = saved_id
        st.session_state["link_builder_last_saved_at"] = timestamp

    st.session_state["draft_archive_name"] = archive_name
    st.session_state["link_builder_saved_signature"] = json.dumps(
        {"metadata": payload["metadata"], "rows": payload["rows"]},
        ensure_ascii=False,
        sort_keys=True,
    )
    st.session_state["link_builder_saved_result_json"] = json.dumps(payload["result"], ensure_ascii=False, sort_keys=True)
    st.session_state["link_builder_opened_archive_result"] = payload["result"]
    st.session_state["link_builder_show_live_results"] = True
    return saved_id


def reset_to_new_activity() -> None:
    for field, default_value in BUILDER_DEFAULTS.items():
        st.session_state[f"draft_{field}"] = default_value
    st.session_state["draft_archive_name"] = ""
    st.session_state["link_builder_rows"] = DEFAULT_LINK_ROWS.copy()
    st.session_state["current_archive_id"] = None
    st.session_state["form_revision"] += 1
    st.session_state["link_builder_last_saved_at"] = None
    st.session_state["link_builder_saved_signature"] = ""
    st.session_state["link_builder_saved_result_json"] = ""
    st.session_state["link_builder_validation_errors"] = []
    st.session_state["link_builder_opened_archive_result"] = None
    st.session_state["link_builder_show_live_results"] = False


init_archive_db()
ensure_link_builder_state()
inject_apple_panel_styles()

try:
        render_apple_hero(
            "UTM / AppsFlyer Bulk Builder",
            "专注做一件事：把 campaign 参数、web 链接、app deeplink 和 AppsFlyer 批量表整合成一个可直接交付给运营同学使用的工具页。",
            kicker="Team Tool",
        )

        campaign_context = collect_campaign_context()
        if not campaign_context["archive_name"]:
            st.session_state["draft_archive_name"] = get_default_archive_name(campaign_context)
            campaign_context = collect_campaign_context()

        current_rows = ensure_minimum_rows(st.session_state.get("link_builder_rows", DEFAULT_LINK_ROWS), minimum=3)
        form_revision = st.session_state["form_revision"]

        status_cols = st.columns(3)
        current_archive_id = st.session_state.get("current_archive_id")
        status_cols[0].metric("当前活动", get_current_activity_name())
        status_cols[1].metric("当前模式", "编辑已有活动" if current_archive_id else "新活动")
        status_cols[2].metric("上次保存时间", st.session_state.get("link_builder_last_saved_at") or "-")
        if st.session_state.get("link_builder_ui_error"):
            st.success(st.session_state["link_builder_ui_error"])
            st.session_state["link_builder_ui_error"] = ""

        render_step_shell("1", "Campaign 参数区", "先定义这次活动的命名规则和追踪参数。", tone="#111827")
        with st.expander("查看参数规则说明", expanded=False):
            st.caption("medium：默认建议 email / app-push / sms，也可手动输入新值")
            st.caption("utm_source：默认建议 manual / automation，也可手动输入新值")
            st.caption("campaign prefix：默认建议渠道首字母 + 发送方式（emanual / eautomation / smanual / sautomation / pmanual / pautomation），也可手动输入")
            st.caption("type：默认建议 Matching / Sale / Bamboo / Disney / Ip / Kids / Holiday / Cruise / Test，也可手动输入新类型")
            st.caption("theme / sender：自定义填写")
            st.caption("utm_term：本次活动的 term 基础值，例如 0320paw")
        c1, c2, c3 = st.columns(3)
        archive_name = c1.text_input("活动名称", value=st.session_state.get("draft_archive_name", ""), key=f"lb_archive_name_widget_{form_revision}")
        pid = c2.text_input("pid", value=st.session_state.get("draft_pid", ""), key=f"lb_pid_widget_{form_revision}")
        medium = render_select_with_custom_input(
            c3,
            "medium",
            MEDIUM_OPTIONS,
            st.session_state.get("draft_medium", MEDIUM_OPTIONS[0]),
            "lb_medium",
            form_revision,
            placeholder="例如：in-app-message",
        )

        c4, c5, c6 = st.columns(3)
        utm_source = render_select_with_custom_input(
            c4,
            "utm_source",
            UTM_SOURCE_OPTIONS,
            st.session_state.get("draft_utm_source", UTM_SOURCE_OPTIONS[0]),
            "lb_utm_source",
            form_revision,
            placeholder="例如：crm-journey",
        )
        campaign_prefix = render_select_with_custom_input(
            c5,
            "campaign prefix",
            CAMPAIGN_PREFIX_OPTIONS,
            st.session_state.get("draft_campaign_prefix", CAMPAIGN_PREFIX_OPTIONS[0]),
            "lb_campaign_prefix",
            form_revision,
            placeholder="例如：xmanual",
        )
        campaign_type = render_select_with_custom_input(
            c6,
            "type",
            TYPE_OPTIONS,
            st.session_state.get("draft_type", TYPE_OPTIONS[0]),
            "lb_type",
            form_revision,
            placeholder="例如：Shipping",
        )

        c7, c8, c9 = st.columns(3)
        theme = c7.text_input("theme", value=st.session_state.get("draft_theme", ""), key=f"lb_theme_widget_{form_revision}")
        sender = c8.text_input("sender", value=st.session_state.get("draft_sender", ""), key=f"lb_sender_widget_{form_revision}")
        year = c9.text_input("year", value=st.session_state.get("draft_year", ""), key=f"lb_year_widget_{form_revision}")

        c10, c11, c12 = st.columns(3)
        date = c10.text_input("date", value=st.session_state.get("draft_date", ""), key=f"lb_date_widget_{form_revision}")
        utm_term = c11.text_input("utm_term", value=st.session_state.get("draft_utm_term", ""), key=f"lb_utm_term_widget_{form_revision}")
        click_lookback = c12.text_input("click lookback", value=st.session_state.get("draft_af_click_lookback", ""), key=f"lb_af_click_lookback_widget_{form_revision}")

        c13, c14, c15 = st.columns([2, 2, 2])
        retarget_index = 0 if st.session_state.get("draft_is_retargeting", "TRUE") == "TRUE" else 1
        is_retargeting = c13.selectbox("is_retargeting", ["TRUE", "FALSE"], index=retarget_index, key=f"lb_is_retargeting_widget_{form_revision}")
        if c14.button("New Activity", use_container_width=True):
            st.session_state["link_builder_last_draft_snapshot"] = snapshot_current_draft()
            reset_to_new_activity()
            st.session_state["link_builder_ui_error"] = "已新建空白活动。可点击“恢复上一步”撤销。"
            st.rerun()
        if c15.button("恢复上一步", use_container_width=True, disabled=st.session_state.get("link_builder_last_draft_snapshot") is None):
            restore_draft_snapshot(st.session_state.get("link_builder_last_draft_snapshot"))
            st.session_state["link_builder_ui_error"] = "已恢复上一步草稿。"
            st.rerun()

        rule_copy_text = (
            "utm_campaign 命名规则\n"
            "生成顺序：campaign prefix _ type _ theme _ sender _ year _ date\n"
            "示例：emanual_sale_flash_cxy_2026_0320\n"
            "字段说明：campaign prefix=发送渠道+方式，type=活动类型，theme=主题，sender=负责人，year=年份，date=MMDD"
        )
        st.session_state["draft_archive_name"] = archive_name
        st.session_state["draft_pid"] = pid
        st.session_state["draft_medium"] = medium
        st.session_state["draft_utm_source"] = utm_source
        st.session_state["draft_campaign_prefix"] = campaign_prefix
        st.session_state["draft_type"] = campaign_type
        st.session_state["draft_theme"] = theme
        st.session_state["draft_sender"] = sender
        st.session_state["draft_year"] = year
        st.session_state["draft_date"] = date
        st.session_state["draft_utm_term"] = utm_term
        st.session_state["draft_af_click_lookback"] = click_lookback
        st.session_state["draft_is_retargeting"] = is_retargeting
        st.session_state["draft_campaign_template"] = st.session_state.get("draft_campaign_template", BUILDER_DEFAULTS["campaign_template"])
        campaign_context = collect_campaign_context()
        campaign_preview = build_utm_campaign(campaign_context["campaign_template"], campaign_context)

        preview_parts = [
            ("campaign prefix", campaign_prefix),
            ("type", campaign_type),
            ("theme", theme),
            ("sender", sender),
            ("year", year),
            ("date", date),
        ]
        preview_badges = " ".join(
            [
                f"<span style='display:inline-block;padding:6px 12px;margin:4px 6px 4px 0;border-radius:999px;background:#eef2ff;color:#1f2937;font-size:14px;font-weight:600;' title='{label}'>{value or '-'}</span>"
                for label, value in preview_parts
            ]
        )
        st.markdown("#### utm_campaign 命名规则")
        st.markdown(f"**当前预览：** `{campaign_preview}`")
        st.markdown(preview_badges, unsafe_allow_html=True)

        with st.expander("展开说明", expanded=False):
            st.write("系统会按以下顺序自动生成 utm_campaign：")
            st.markdown("`campaign prefix` + `_` + `type` + `_` + `theme` + `_` + `sender` + `_` + `year` + `_` + `date`")
            st.write("例如，如果填写：")
            st.markdown(
                "- `campaign prefix = emanual`\n"
                "- `type = Sale`\n"
                "- `theme = flash`\n"
                "- `sender = cxy`\n"
                "- `year = 2026`\n"
                "- `date = 0320`"
            )
            st.markdown(f"则生成：`{campaign_preview}`")

            explanation_df = pd.DataFrame(
                [
                    {"段落": "第 1 段", "字段": "campaign prefix", "来源": "Campaign Prefix", "作用": "区分渠道与发送方式", "示例": "emanual / eautomation / pmanual"},
                    {"段落": "第 2 段", "字段": "type", "来源": "Type", "作用": "区分活动类型", "示例": "Sale / Ip / Holiday / Test"},
                    {"段落": "第 3 段", "字段": "theme", "来源": "Theme", "作用": "区分活动主题", "示例": "flash / paw / easter"},
                    {"段落": "第 4 段", "字段": "sender", "来源": "Sender", "作用": "标记负责人或归属", "示例": "cxy"},
                    {"段落": "第 5 段", "字段": "year", "来源": "Year", "作用": "年份", "示例": "2026"},
                    {"段落": "第 6 段", "字段": "date", "来源": "Date", "作用": "活动日期", "示例": "0320"},
                ]
            )
            st.dataframe(explanation_df, use_container_width=True, hide_index=True)
            st.info(
                "命名小贴士\n"
                "• 全部使用小写英文和数字\n"
                "• 不要加空格\n"
                "• 用下划线连接\n"
                "• theme 尽量简短清晰\n"
                "• date 统一用 MMDD 格式，例如 0320"
            )

        with st.expander("高级规则设置 / 开发模式", expanded=False):
            campaign_template = st.text_input(
                "utm_campaign 模板字符串",
                value=st.session_state.get("draft_campaign_template", ""),
                key=f"lb_campaign_template_widget_{form_revision}",
                help="默认不需要修改。只有在你想调整命名顺序时才使用。",
            )
            st.session_state["draft_campaign_template"] = campaign_template
            st.text_area(
                "复制规则说明",
                value=rule_copy_text,
                height=120,
                help="可直接复制发给同事。",
            )

        render_step_shell("2", "链接输入区", "支持单条录入，也支持按整块或按列批量粘贴。", tone="#2563eb")
        with st.expander("批量粘贴导入", expanded=False):
            st.caption("支持直接粘贴从 Excel / 表格复制的多行数据。推荐列顺序：sequence_no, raw_web_url, raw_app_deeplink, optional_label。支持 TSV 或 CSV。")
            bulk_paste_text = st.text_area(
                "Bulk Paste",
                value=st.session_state.get("link_builder_bulk_paste", ""),
                key=f"link_builder_bulk_paste_widget_{form_revision}",
                height=180,
                placeholder="01\thttps://www.patpat.com/...\tpatpat://?action=...\tHero CTA",
            )
            st.session_state["link_builder_bulk_paste"] = bulk_paste_text
            bp1, bp2 = st.columns(2)
            if bp1.button("导入并覆盖当前行", use_container_width=True):
                parsed_df = parse_bulk_rows(bulk_paste_text)
                if parsed_df.empty:
                    st.session_state["link_builder_validation_errors"] = ["Bulk paste 内容为空或无法解析。"]
                else:
                    st.session_state["link_builder_rows"] = parsed_df.to_dict(orient="records")
                    st.session_state["form_revision"] += 1
                    st.session_state["link_builder_validation_errors"] = []
                    st.rerun()
            if bp2.button("追加到当前行", use_container_width=True):
                parsed_df = parse_bulk_rows(bulk_paste_text)
                if parsed_df.empty:
                    st.session_state["link_builder_validation_errors"] = ["Bulk paste 内容为空或无法解析。"]
                else:
                    existing_df = normalize_input_rows(st.session_state.get("link_builder_rows", DEFAULT_LINK_ROWS))
                    active_existing = existing_df[get_active_row_mask(existing_df)].copy()
                    merged_df = pd.concat([active_existing, parsed_df], ignore_index=True)
                    st.session_state["link_builder_rows"] = normalize_input_rows(merged_df).to_dict(orient="records")
                    st.session_state["form_revision"] += 1
                    st.session_state["link_builder_validation_errors"] = []
                    st.rerun()

        with st.expander("按列批量粘贴", expanded=True):
            st.caption("适合直接复制整列内容。每列按行对齐，空白会自动补齐。最常用的是只粘贴 raw_app_deeplink 或 web + app 两列。")
            col1, col2, col3, col4 = st.columns(4)
            sequence_column_text = col1.text_area(
                "sequence_no 列",
                value=st.session_state.get("link_builder_column_sequence", ""),
                key=f"link_builder_column_sequence_widget_{form_revision}",
                height=180,
                placeholder="01\n02\n03",
            )
            web_column_text = col2.text_area(
                "raw_web_url 列",
                value=st.session_state.get("link_builder_column_web", ""),
                key=f"link_builder_column_web_widget_{form_revision}",
                height=180,
                placeholder="https://www.patpat.com/...\nhttps://www.patpat.com/...",
            )
            app_column_text = col3.text_area(
                "raw_app_deeplink 列",
                value=st.session_state.get("link_builder_column_app", ""),
                key=f"link_builder_column_app_widget_{form_revision}",
                height=180,
                placeholder="patpat://?action=...\npatpat://?action=...",
            )
            label_column_text = col4.text_area(
                "optional_label 列",
                value=st.session_state.get("link_builder_column_label", ""),
                key=f"link_builder_column_label_widget_{form_revision}",
                height=180,
                placeholder="Hero CTA\nPDP CTA",
            )
            st.session_state["link_builder_column_sequence"] = sequence_column_text
            st.session_state["link_builder_column_web"] = web_column_text
            st.session_state["link_builder_column_app"] = app_column_text
            st.session_state["link_builder_column_label"] = label_column_text

            cp1, cp2 = st.columns(2)
            if cp1.button("按列导入并覆盖当前行", use_container_width=True):
                parsed_df = build_rows_from_columns(
                    sequence_column_text,
                    web_column_text,
                    app_column_text,
                    label_column_text,
                )
                if parsed_df.empty:
                    st.session_state["link_builder_validation_errors"] = ["按列粘贴内容为空，无法导入。"]
                else:
                    st.session_state["link_builder_rows"] = parsed_df.to_dict(orient="records")
                    st.session_state["form_revision"] += 1
                    st.session_state["link_builder_validation_errors"] = []
                    st.rerun()
            if cp2.button("按列追加到当前行", use_container_width=True):
                parsed_df = build_rows_from_columns(
                    sequence_column_text,
                    web_column_text,
                    app_column_text,
                    label_column_text,
                )
                if parsed_df.empty:
                    st.session_state["link_builder_validation_errors"] = ["按列粘贴内容为空，无法追加。"]
                else:
                    existing_df = normalize_input_rows(st.session_state.get("link_builder_rows", DEFAULT_LINK_ROWS))
                    active_existing = existing_df[get_active_row_mask(existing_df)].copy()
                    merged_df = pd.concat([active_existing, parsed_df], ignore_index=True)
                    st.session_state["link_builder_rows"] = normalize_input_rows(merged_df).to_dict(orient="records")
                    st.session_state["form_revision"] += 1
                    st.session_state["link_builder_validation_errors"] = []
                    st.rerun()

        edit_cols = st.columns([1, 1, 4])
        if edit_cols[0].button("Add Row", use_container_width=True):
            add_blank_row()
            st.rerun()
        delete_options = [str(i + 1) for i in range(len(current_rows))]
        delete_choice = edit_cols[1].selectbox("Delete Row", delete_options, label_visibility="collapsed") if delete_options else None
        if edit_cols[1].button("Delete Row", use_container_width=True, disabled=not delete_options):
            delete_row_at(int(delete_choice) - 1)
            st.rerun()

        header_cols = st.columns([1, 3, 3, 2])
        header_cols[0].markdown("`sequence_no`")
        header_cols[1].markdown("`raw_web_url`")
        header_cols[2].markdown("`raw_app_deeplink`")
        header_cols[3].markdown("`optional_label`")

        edited_rows: list[dict[str, str]] = []
        for row_index, row in current_rows.iterrows():
            row_cols = st.columns([1, 3, 3, 2])
            sequence_no = row_cols[0].text_input(
                f"sequence_no_{row_index + 1}",
                value=coerce_text(row["sequence_no"]),
                key=f"lb_sequence_no_widget_{form_revision}_{row_index}",
                label_visibility="collapsed",
                placeholder="01",
            )
            raw_web_url = row_cols[1].text_input(
                f"raw_web_url_{row_index + 1}",
                value=coerce_text(row["raw_web_url"]),
                key=f"lb_raw_web_url_widget_{form_revision}_{row_index}",
                label_visibility="collapsed",
                placeholder="https://www.patpat.com/...",
            )
            raw_app_deeplink = row_cols[2].text_input(
                f"raw_app_deeplink_{row_index + 1}",
                value=coerce_text(row["raw_app_deeplink"]),
                key=f"lb_raw_app_deeplink_widget_{form_revision}_{row_index}",
                label_visibility="collapsed",
                placeholder="patpat://?action=...",
            )
            optional_label = row_cols[3].text_input(
                f"optional_label_{row_index + 1}",
                value=coerce_text(row["optional_label"]),
                key=f"lb_optional_label_widget_{form_revision}_{row_index}",
                label_visibility="collapsed",
                placeholder="Hero CTA",
            )
            edited_rows.append(
                {
                    "sequence_no": sequence_no,
                    "raw_web_url": raw_web_url,
                    "raw_app_deeplink": raw_app_deeplink,
                    "optional_label": optional_label,
                }
            )

        normalized_edited_df = normalize_input_rows(pd.DataFrame(edited_rows))
        st.session_state["link_builder_rows"] = normalized_edited_df.to_dict(orient="records")

        validation_errors = st.session_state.get("link_builder_validation_errors", [])
        preview_df, appsflyer_df = build_outputs(
            campaign_context,
            normalized_edited_df,
            adapter_key="local_passthrough",
        )
        current_payload = build_archive_payload(campaign_context, normalized_edited_df, preview_df, appsflyer_df)
        current_signature = payload_signature(campaign_context, normalized_edited_df)
        is_dirty = current_signature != st.session_state.get("link_builder_saved_signature", "")

        render_step_shell("3", "生成结果区", "生成 web 追踪链接、app deeplink 转链结果，以及 AppsFlyer 批量表。", tone="#0f766e")
        action_cols = st.columns(3)
        if action_cols[0].button("Generate & Save", use_container_width=True, type="primary"):
            validation_errors = validate_link_rows(normalized_edited_df)
            st.session_state["link_builder_validation_errors"] = validation_errors
            if validation_errors:
                st.warning("校验未通过，请先修正后再生成。")
            else:
                st.session_state["link_builder_show_live_results"] = True
                mode = "update" if current_archive_id else "create"
                save_archive_payload(current_payload, mode=mode)
                st.rerun()
        if action_cols[1].button("Save as New", use_container_width=True):
            validation_errors = validate_link_rows(normalized_edited_df)
            st.session_state["link_builder_validation_errors"] = validation_errors
            if validation_errors:
                st.warning("校验未通过，请先修正后再另存。")
            else:
                st.session_state["link_builder_show_live_results"] = True
                save_archive_payload(current_payload, mode="create")
                st.rerun()
        if action_cols[2].button("Update Current", use_container_width=True, disabled=current_archive_id is None):
            validation_errors = validate_link_rows(normalized_edited_df)
            st.session_state["link_builder_validation_errors"] = validation_errors
            if validation_errors:
                st.warning("校验未通过，请先修正后再更新。")
            else:
                st.session_state["link_builder_show_live_results"] = True
                save_archive_payload(current_payload, mode="update")
                st.rerun()

        if is_dirty:
            st.info("当前有未保存修改。")
        else:
            st.caption("当前内容已与最近一次保存保持一致。")

        st.write("Validation Summary")
        if validation_errors:
            for message in validation_errors:
                st.error(message)
        else:
            st.caption("当前没有校验错误。只有活跃行会在 Generate / Save 时校验。")

        result_cols = st.columns(3)
        result_cols[0].metric("有效链接行数", f"{len(preview_df)}")
        result_cols[1].metric("utm_term", campaign_context["utm_term"])
        result_cols[2].metric("AppsFlyer 导出行数", f"{len(appsflyer_df)}")

        show_live_results = st.session_state.get("link_builder_show_live_results", True)
        if show_live_results:
            st.dataframe(
                preview_df[["raw_web_url", "full_web_url", "raw_app_deeplink", "final_app_deeplink", "c"]] if not preview_df.empty else preview_df,
                use_container_width=True,
                column_config={
                    "raw_web_url": st.column_config.TextColumn("raw_web_url", width="large"),
                    "full_web_url": st.column_config.TextColumn("full_web_url", width="large"),
                    "raw_app_deeplink": st.column_config.TextColumn("raw_app_deeplink", width="large"),
                    "final_app_deeplink": st.column_config.TextColumn("final_app_deeplink", width="large"),
                },
            )
            st.write("AppsFlyer 输出预览")
            st.dataframe(appsflyer_df, use_container_width=True)
        else:
            st.caption("当前主结果区只展示重新 Generate 后的最新结果。打开历史活动后，可在历史区查看该活动的历史快照。")

        csv_bytes = dataframe_to_csv_bytes(appsflyer_df)
        xlsx_bytes = dataframe_to_xlsx_bytes({"link_preview": preview_df, "appsflyer_bulk_sheet": appsflyer_df})
        d1, d2 = st.columns(2)
        d1.download_button("下载 AppsFlyer CSV", data=csv_bytes, file_name=f"{current_payload['archive_name']}_appsflyer.csv", mime="text/csv", use_container_width=True)
        d2.download_button("下载 XLSX（含两张表）", data=xlsx_bytes, file_name=f"{current_payload['archive_name']}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        copy_target = st.selectbox("复制内容", ["full_web_url 列", "af_dp 列", "AppsFlyer 整表 TSV", "结果预览整表 TSV"])
        if copy_target == "full_web_url 列":
            copy_value = "\n".join(preview_df["full_web_url"].tolist()) if not preview_df.empty else ""
        elif copy_target == "af_dp 列":
            copy_value = "\n".join(appsflyer_df["af_dp"].tolist()) if not appsflyer_df.empty else ""
        elif copy_target == "AppsFlyer 整表 TSV":
            copy_value = dataframe_to_tsv(appsflyer_df)
        else:
            copy_value = dataframe_to_tsv(preview_df)
        st.text_area("复制区", value=copy_value, height=180)

        render_step_shell("4", "历史存档区", "保存、打开、复制和回查过往活动，避免重复劳动。", tone="#7c3aed")
        st.text_input("搜索历史活动", key="link_builder_history_search", placeholder="campaign / theme / sender / date / utm_term")
        history_items = list_archives(st.session_state.get("link_builder_history_search", ""))
        st.caption(f"共 {len(history_items)} 个活动存档")

        opened_archive_result = st.session_state.get("link_builder_opened_archive_result")
        if opened_archive_result:
            with st.expander("当前打开历史活动的结果快照", expanded=True):
                st.caption(f"生成时间：{opened_archive_result.get('generated_at', '-')}")
                preview_snapshot = pd.DataFrame(opened_archive_result.get("preview_rows", []))
                appsflyer_snapshot = pd.DataFrame(opened_archive_result.get("appsflyer_rows", []))
                if not preview_snapshot.empty:
                    st.write("历史链接结果")
                    st.dataframe(preview_snapshot, use_container_width=True)
                if not appsflyer_snapshot.empty:
                    st.write("历史 AppsFlyer 结果")
                    st.dataframe(appsflyer_snapshot, use_container_width=True)

        confirm_delete_id = st.session_state.get("link_builder_confirm_delete_id")
        if confirm_delete_id is not None:
            st.error(f"确认删除活动 #{confirm_delete_id}？")
            dc1, dc2 = st.columns(2)
            if dc1.button("确认删除", key="confirm_delete_yes", use_container_width=True):
                delete_archive(confirm_delete_id)
                if st.session_state.get("current_archive_id") == confirm_delete_id:
                    reset_to_new_activity()
                st.session_state["link_builder_confirm_delete_id"] = None
                st.rerun()
            if dc2.button("取消删除", key="confirm_delete_no", use_container_width=True):
                st.session_state["link_builder_confirm_delete_id"] = None
                st.rerun()

        for item in history_items:
            with st.container(border=True):
                st.markdown(f"**{item['archive_name']}**")
                st.caption(
                    f"campaign: {item['utm_campaign'] or '-'} | theme: {item['theme'] or '-'} | "
                    f"sender: {item['sender'] or '-'} | date: {item['campaign_date'] or '-'}\n\n"
                    f"created: {item['created_at']} | updated: {item['updated_at']}"
                )
                hc1, hc2, hc3 = st.columns(3)
                if hc1.button("打开", key=f"open_{item['archive_id']}", use_container_width=True):
                    archive = get_archive(item["archive_id"])
                    if archive:
                        load_archive_into_state(archive, as_duplicate=False)
                    st.rerun()
                if hc2.button("复制", key=f"duplicate_{item['archive_id']}", use_container_width=True):
                    archive = get_archive(item["archive_id"])
                    if archive:
                        load_archive_into_state(archive, as_duplicate=True)
                    st.rerun()
                if hc3.button("删除", key=f"delete_{item['archive_id']}", use_container_width=True):
                    st.session_state["link_builder_confirm_delete_id"] = item["archive_id"]
                    st.rerun()
except Exception as exc:
    st.error(f"Link Builder 暂时无法渲染：{exc}")
    with st.expander("Show debug details"):
        st.code(traceback.format_exc())
