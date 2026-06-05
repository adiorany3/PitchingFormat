import html
import re
from io import BytesIO
from typing import Any

import streamlit as st

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


st.set_page_config(
    page_title="Seed Investor Pitch Deck Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==============================
# Streamlit Chrome Removal
# ==============================
def hide_streamlit_emblems():
    """Hide Streamlit chrome and apply theme-aware readable UI styling."""
    st.markdown(
        """
        <style>
            :root {
                --deck-bg: var(--background-color, #f8fafc);
                --deck-surface: var(--secondary-background-color, #ffffff);
                --deck-surface-soft: rgba(148, 163, 184, 0.10);
                --deck-surface-strong: rgba(148, 163, 184, 0.16);
                --deck-text: var(--text-color, #0f172a);
                --deck-primary: var(--primary-color, #2563eb);
                --deck-muted: rgba(100, 116, 139, 0.96);
                --deck-muted: color-mix(in srgb, var(--text-color, #0f172a) 68%, transparent);
                --deck-border: rgba(148, 163, 184, 0.35);
                --deck-border-strong: rgba(148, 163, 184, 0.55);
                --deck-shadow: 0 14px 35px rgba(15, 23, 42, 0.08);
                --deck-radius: 16px;
            }

            /* Remove Streamlit chrome / emblem */
            #MainMenu {visibility: hidden !important;}
            footer {visibility: hidden !important;}
            header {visibility: hidden !important;}
            [data-testid="stHeader"] {display: none !important;}
            [data-testid="stToolbar"] {display: none !important;}
            [data-testid="stDecoration"] {display: none !important;}
            [data-testid="stStatusWidget"] {display: none !important;}
            [data-testid="stDeployButton"] {display: none !important;}
            [data-testid="collapsedControl"] {display: none !important;}
            .stAppDeployButton {display: none !important;}
            .st-emotion-cache-1dp5vir {display: none !important;}
            .st-emotion-cache-14xtw13 {display: none !important;}
            .viewerBadge_container__1QSob,
            .viewerBadge_link__1S137,
            .viewerBadge_text__1JaDK {display: none !important;}

            /* App layout */
            .stApp {
                background: var(--deck-bg) !important;
                color: var(--deck-text) !important;
            }
            .block-container {
                max-width: 1320px;
                padding-top: 1.35rem !important;
                padding-bottom: 5.8rem !important;
            }
            [data-testid="stSidebar"] > div:first-child {
                background: var(--deck-surface) !important;
                border-right: 1px solid var(--deck-border);
            }
            [data-testid="stSidebar"] * {
                color: var(--deck-text);
            }

            /* Text readability across light/dark themes */
            h1, h2, h3, h4, h5, h6,
            [data-testid="stMarkdownContainer"] p,
            [data-testid="stMarkdownContainer"] li,
            [data-testid="stCaptionContainer"],
            label {
                color: var(--deck-text) !important;
            }
            small, .caption, [data-testid="stCaptionContainer"] {
                color: var(--deck-muted) !important;
            }
            hr {
                border-color: var(--deck-border) !important;
            }

            /* Tabs */
            [data-testid="stTabs"] [role="tablist"] {
                gap: 0.35rem;
                border-bottom: 1px solid var(--deck-border);
                flex-wrap: wrap;
            }
            [data-testid="stTabs"] button[role="tab"] {
                min-height: 42px;
                padding: 0.55rem 0.85rem;
                border-radius: 999px 999px 0 0;
                color: var(--deck-muted) !important;
                background: transparent !important;
                border: 1px solid transparent !important;
            }
            [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
                color: var(--deck-text) !important;
                background: var(--deck-surface) !important;
                border: 1px solid var(--deck-border) !important;
                border-bottom-color: var(--deck-surface) !important;
                font-weight: 700;
            }

            /* Form controls */
            [data-testid="stTextInput"] input,
            [data-testid="stNumberInput"] input,
            [data-testid="stTextArea"] textarea,
            [data-baseweb="select"] > div,
            [data-testid="stFileUploader"] section,
            [data-testid="stColorPicker"] input {
                background: var(--deck-surface) !important;
                color: var(--deck-text) !important;
                border-color: var(--deck-border) !important;
                caret-color: var(--deck-primary) !important;
            }
            [data-testid="stTextInput"] input::placeholder,
            [data-testid="stTextArea"] textarea::placeholder {
                color: var(--deck-muted) !important;
                opacity: 0.82;
            }
            [data-testid="stTextInput"] input:focus,
            [data-testid="stNumberInput"] input:focus,
            [data-testid="stTextArea"] textarea:focus,
            [data-baseweb="select"] > div:focus-within {
                border-color: var(--deck-primary) !important;
                box-shadow: 0 0 0 2px color-mix(in srgb, var(--deck-primary) 24%, transparent) !important;
            }
            [data-testid="stWidgetLabel"] p {
                color: var(--deck-text) !important;
                font-weight: 650;
            }
            [data-testid="InputInstructions"],
            [data-testid="stWidgetLabel"] small,
            [data-testid="stHelp"] {
                color: var(--deck-muted) !important;
            }

            /* Buttons */
            .stButton > button,
            .stDownloadButton > button {
                border-radius: 12px !important;
                border: 1px solid var(--deck-border-strong) !important;
                background: var(--deck-surface) !important;
                color: var(--deck-text) !important;
                font-weight: 700 !important;
                min-height: 44px;
            }
            .stButton > button[kind="primary"],
            .stButton > button[data-testid="baseButton-primary"] {
                background: var(--deck-primary) !important;
                color: #ffffff !important;
                border-color: var(--deck-primary) !important;
            }
            .stDownloadButton > button {
                background: var(--deck-primary) !important;
                color: #ffffff !important;
                border-color: var(--deck-primary) !important;
            }

            /* Native Streamlit metric readability */
            [data-testid="stMetric"] {
                background: var(--deck-surface) !important;
                border: 1px solid var(--deck-border);
                border-radius: var(--deck-radius);
                padding: 0.9rem 1rem;
                box-shadow: var(--deck-shadow);
            }
            [data-testid="stMetricLabel"] p {
                color: var(--deck-muted) !important;
                font-weight: 700;
            }
            [data-testid="stMetricValue"] {
                color: var(--deck-text) !important;
            }

            /* Custom reusable blocks */
            .developer-footer {
                position: fixed;
                left: 0;
                bottom: 0;
                width: 100%;
                background: color-mix(in srgb, var(--deck-surface) 94%, transparent);
                backdrop-filter: blur(10px);
                border-top: 1px solid var(--deck-border);
                color: var(--deck-muted) !important;
                text-align: center;
                padding: 8px 0;
                font-size: 12px;
                line-height: 1.4;
                z-index: 9999;
            }
            .guide-box,
            .insight-card,
            .readable-panel {
                background: var(--deck-surface);
                border: 1px solid var(--deck-border);
                border-radius: var(--deck-radius);
                padding: 15px 17px;
                margin: 8px 0 18px 0;
                color: var(--deck-text) !important;
                box-shadow: var(--deck-shadow);
            }
            .guide-box {
                border-left: 4px solid var(--deck-primary);
            }
            .guide-box strong,
            .guide-box .guide-title,
            .insight-card h4 {
                color: var(--deck-text) !important;
                font-weight: 800;
            }
            .guide-box p,
            .guide-box span,
            .insight-card p,
            .insight-card li,
            .readable-panel p,
            .readable-panel li {
                color: var(--deck-muted) !important;
                font-size: 0.94rem;
                line-height: 1.58;
            }
            .insight-card h4 {
                margin: 0 0 10px 0;
                font-size: 1.02rem;
            }
            .insight-card ul {
                margin: 0.35rem 0 0 1.1rem;
                padding: 0;
            }
            .insight-card li {
                margin-bottom: 0.48rem;
            }

            /* Alerts and upload blocks should keep contrast */
            [data-testid="stAlert"] {
                border-radius: var(--deck-radius) !important;
                border: 1px solid var(--deck-border) !important;
            }
            [data-testid="stFileUploader"] section {
                border-radius: var(--deck-radius) !important;
            }

            @media (max-width: 900px) {
                .block-container {
                    padding-left: 1rem !important;
                    padding-right: 1rem !important;
                }
                .developer-footer {
                    font-size: 11px;
                    padding: 7px 0;
                }
                [data-testid="stTabs"] button[role="tab"] {
                    min-height: 38px;
                    padding: 0.45rem 0.65rem;
                }
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

hide_streamlit_emblems()


# ==============================
# Professional Seed Deck Theme
# ==============================
THEME = {
    "bg": RGBColor(248, 250, 252),
    "white": RGBColor(255, 255, 255),
    "ink": RGBColor(15, 23, 42),
    "muted": RGBColor(100, 116, 139),
    "line": RGBColor(226, 232, 240),
    "dark": RGBColor(2, 6, 23),
    "green": RGBColor(22, 163, 74),
    "amber": RGBColor(217, 119, 6),
    "red": RGBColor(220, 38, 38),
    "blue_soft": RGBColor(239, 246, 255),
    "blue_line": RGBColor(191, 219, 254),
}

FONT_HEAD = "Aptos Display"
FONT_BODY = "Aptos"
DEVELOPER_FOOTER = "Developed by Galuh Adi Insani"

SLIDE_W = 13.333
SLIDE_H = 7.5


# ==============================
# Utility Helpers
# ==============================
def rgb(hex_color: str) -> RGBColor:
    value = (hex_color or "#2563EB").replace("#", "").strip()

    if len(value) != 6:
        value = "2563EB"

    return RGBColor(
        int(value[:2], 16),
        int(value[2:4], 16),
        int(value[4:], 16),
    )


def set_bg(slide, color: RGBColor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def lines(text: str, limit: int = 5) -> list[str]:
    result = []

    for row in (text or "").splitlines():
        cleaned = row.strip().lstrip("-•*0123456789. ").strip()

        if cleaned:
            result.append(cleaned)

    return result[:limit]


def money(value: int | float, currency: str) -> str:
    if currency == "Rp":
        return f"Rp {value:,.0f}".replace(",", ".")

    return f"{currency} {value:,.0f}"


def filename(value: str) -> str:
    value = re.sub(
        r"[^a-zA-Z0-9 _-]",
        "",
        value or "pitch-deck",
    )

    return value.strip().lower().replace(" ", "-") or "pitch-deck"


def parse_percent(value: str) -> float | None:
    if not value:
        return None

    match = re.search(r"-?\d+(?:[\.,]\d+)?", str(value))

    if not match:
        return None

    return float(match.group(0).replace(",", "."))


def safe_div(numerator: float, denominator: float) -> float | None:
    if denominator == 0:
        return None

    return numerator / denominator


def pct(value: float | None) -> str:
    if value is None:
        return "N/A"

    return f"{value * 100:.1f}%"


def multiple(value: float | None) -> str:
    if value is None:
        return "N/A"

    return f"{value:.1f}x"


def short_money(value: float, currency: str) -> str:
    sign = "-" if value < 0 else ""
    value = abs(value)

    if currency == "Rp":
        if value >= 1_000_000_000_000:
            return f"{sign}Rp {value / 1_000_000_000_000:.1f}T"
        if value >= 1_000_000_000:
            return f"{sign}Rp {value / 1_000_000_000:.1f}M"
        if value >= 1_000_000:
            return f"{sign}Rp {value / 1_000_000:.1f}Jt"
        return f"{sign}Rp {value:,.0f}".replace(",", ".")

    if value >= 1_000_000_000:
        return f"{sign}{currency} {value / 1_000_000_000:.1f}B"
    if value >= 1_000_000:
        return f"{sign}{currency} {value / 1_000_000:.1f}M"
    if value >= 1_000:
        return f"{sign}{currency} {value / 1_000:.1f}K"

    return f"{sign}{currency} {value:,.0f}"



def normalize_competitors(data: dict[str, Any]) -> list[dict[str, str]]:
    """Return clean competitor rows from dynamic UI data or legacy fields."""
    rows = data.get("competitors") or []
    cleaned = []

    for row in rows:
        if not isinstance(row, dict):
            continue

        name = str(row.get("name", "")).strip()
        weakness = str(row.get("weakness", "")).strip()
        advantage = str(row.get("advantage", "")).strip()
        category = str(row.get("category", "Alternative")).strip() or "Alternative"

        if name or weakness or advantage:
            cleaned.append(
                {
                    "name": name or "Alternative",
                    "category": category,
                    "weakness": weakness or "Belum diisi",
                    "advantage": advantage or "Belum diisi",
                }
            )

    if cleaned:
        return cleaned

    legacy_rows = [
        {
            "name": data.get("competitor_1", ""),
            "category": "Direct competitor",
            "weakness": data.get("weakness_1", ""),
            "advantage": data.get("advantage_1", ""),
        },
        {
            "name": data.get("competitor_2", ""),
            "category": "Indirect competitor",
            "weakness": data.get("weakness_2", ""),
            "advantage": data.get("advantage_2", ""),
        },
        {
            "name": "Status quo",
            "category": "Status quo",
            "weakness": data.get("status_quo", ""),
            "advantage": data.get("advantage_3", ""),
        },
    ]

    return [row for row in legacy_rows if str(row.get("name", "")).strip()]


def normalize_milestones(data: dict[str, Any]) -> list[dict[str, str]]:
    """Return clean milestone rows from structured UI data or milestone headline."""
    rows = data.get("milestones") or []
    cleaned = []

    for row in rows:
        if not isinstance(row, dict):
            continue

        period = str(row.get("period", "")).strip()
        target = str(row.get("target", "")).strip()
        metric = str(row.get("metric", "")).strip()
        owner = str(row.get("owner", "")).strip()

        if period or target or metric or owner:
            cleaned.append(
                {
                    "period": period or "Next",
                    "target": target or "Target belum diisi",
                    "metric": metric or "Metric belum diisi",
                    "owner": owner or "Team",
                }
            )

    if cleaned:
        return cleaned

    fallback = str(data.get("milestone", "")).strip()

    if fallback:
        return [
            {
                "period": "12-18 bulan",
                "target": fallback,
                "metric": "Funding milestone",
                "owner": "Core team",
            }
        ]

    return []


def milestone_headline(data: dict[str, Any]) -> str:
    milestones = normalize_milestones(data)

    if not milestones:
        return str(data.get("milestone", "Milestone belum diisi"))

    first = milestones[0]
    period = first.get("period", "Next")
    target = first.get("target", "Milestone belum diisi")

    return f"{period}: {target}"


# ==============================
# Investor Insight Engine
# ==============================
def generate_investor_insights(data: dict[str, Any]) -> dict[str, Any]:
    rev1 = float(data.get("rev1", 0) or 0)
    rev2 = float(data.get("rev2", 0) or 0)
    rev3 = float(data.get("rev3", 0) or 0)
    cost1 = float(data.get("cost1", 0) or 0)
    cost2 = float(data.get("cost2", 0) or 0)
    cost3 = float(data.get("cost3", 0) or 0)
    profit1 = float(data.get("profit1", 0) or 0)
    profit2 = float(data.get("profit2", 0) or 0)
    profit3 = float(data.get("profit3", 0) or 0)
    ask = float(data.get("ask", 0) or 0)
    runway = int(data.get("runway", 0) or 0)
    currency = data.get("currency", "Rp")
    competitor_count = len(normalize_competitors(data))
    milestone_count = len(normalize_milestones(data))

    growth_y2 = safe_div(rev2 - rev1, rev1) if rev1 else None
    growth_y3 = safe_div(rev3 - rev2, rev2) if rev2 else None
    gross_margin = parse_percent(data.get("gross_margin", ""))
    retention = parse_percent(data.get("retention", ""))
    payback = parse_percent(data.get("payback", ""))

    burn_y1 = max(cost1 - rev1, 0)
    monthly_burn_y1 = burn_y1 / 12 if burn_y1 else 0
    estimated_runway = safe_div(ask, monthly_burn_y1) if monthly_burn_y1 else None
    funding_to_y1_cost = safe_div(ask, cost1) if cost1 else None
    year3_profit_margin = safe_div(profit3, rev3) if rev3 else None
    revenue_multiple_y3 = safe_div(rev3, rev1) if rev1 else None

    completeness_fields = [
        "company", "one_liner", "problem", "problem_evidence", "solution", "value_prop",
        "product_flow", "product_benefit", "market_notes", "business_model", "traction_notes",
        "gtm", "team", "founder_fit", "use_of_funds", "next_round", "milestone",
    ]
    filled = sum(1 for field in completeness_fields if str(data.get(field, "")).strip())
    if competitor_count >= 2:
        filled += 1
    if milestone_count >= 2:
        filled += 1
    completeness = filled / (len(completeness_fields) + 2)

    score = 45
    score += int(completeness * 20)

    if growth_y2 is not None and growth_y2 > 1:
        score += 8
    elif growth_y2 is not None and growth_y2 > 0.3:
        score += 4

    if growth_y3 is not None and growth_y3 > 0.75:
        score += 6
    elif growth_y3 is not None and growth_y3 > 0.25:
        score += 3

    if gross_margin is not None and gross_margin >= 70:
        score += 6
    elif gross_margin is not None and gross_margin >= 50:
        score += 3

    if retention is not None and retention >= 70:
        score += 5
    elif retention is not None and retention >= 50:
        score += 2

    if estimated_runway is not None and estimated_runway >= 15:
        score += 5
    elif estimated_runway is not None and estimated_runway >= 9:
        score += 2

    if year3_profit_margin is not None and year3_profit_margin > 0:
        score += 5

    if competitor_count >= 3:
        score += 3

    if milestone_count >= 3:
        score += 4

    score = min(score, 100)

    strengths = []
    risks = []
    recommendations = []

    if growth_y2 is not None and growth_y2 > 0:
        strengths.append(
            f"Proyeksi revenue naik {pct(growth_y2)} dari Year 1 ke Year 2 dan {pct(growth_y3)} dari Year 2 ke Year 3. Narasikan sebagai momentum pertumbuhan, bukan hanya tabel angka."
        )
    else:
        risks.append(
            "Pertumbuhan revenue belum terlihat kuat dari proyeksi. Investor seed perlu melihat jalur pertumbuhan yang masuk akal dan berulang."
        )

    if gross_margin is not None:
        if gross_margin >= 70:
            strengths.append(
                f"Gross margin {gross_margin:.0f}% memberi sinyal model bisnis berpotensi scalable jika CAC dan retention terjaga."
            )
        else:
            risks.append(
                f"Gross margin {gross_margin:.0f}% perlu dijelaskan: apakah margin akan membaik karena automation, volume, atau perubahan pricing."
            )

    if retention is not None:
        if retention >= 70:
            strengths.append(
                f"Retention {retention:.0f}% dapat dipakai sebagai bukti awal bahwa produk mulai menjadi habit atau workflow penting."
            )
        else:
            risks.append(
                f"Retention {retention:.0f}% perlu diimbangi dengan strategi aktivasi dan customer success agar investor tidak membaca traction sebagai sementara."
            )

    if estimated_runway is not None:
        if estimated_runway >= runway * 1.25:
            recommendations.append(
                f"Runway estimasi dari ask dan burn Year 1 sekitar {estimated_runway:.0f} bulan, lebih tinggi dari input {runway} bulan. Jelaskan buffer atau percepatan hiring agar angka tetap kredibel."
            )
        elif estimated_runway < runway * 0.75:
            risks.append(
                f"Runway input {runway} bulan terlihat agresif dibanding estimasi burn Year 1 sekitar {estimated_runway:.0f} bulan. Perlu validasi ulang use of funds."
            )
        else:
            strengths.append(
                f"Ask pendanaan relatif konsisten dengan burn Year 1 dan runway sekitar {runway} bulan."
            )

    if year3_profit_margin is not None:
        if year3_profit_margin > 0:
            strengths.append(
                f"Profit margin Year 3 sekitar {pct(year3_profit_margin)}. Ini membantu menunjukkan jalur menuju efisiensi, walaupun seed deck tetap harus fokus pada growth."
            )
        else:
            risks.append(
                "Year 3 masih belum menunjukkan profit positif. Jelaskan kenapa loss tersebut disengaja untuk growth dan kapan operating leverage muncul."
            )

    if funding_to_y1_cost is not None:
        recommendations.append(
            f"Funding yang diminta setara {multiple(funding_to_y1_cost)} operating cost Year 1. Hubungkan angka ini dengan hiring, distribusi, dan milestone yang bisa diverifikasi."
        )

    if revenue_multiple_y3 is not None:
        recommendations.append(
            f"Revenue Year 3 adalah {multiple(revenue_multiple_y3)} dibanding Year 1. Tambahkan asumsi utama: jumlah customer, ARPU, churn, dan channel conversion."
        )

    if competitor_count < 3:
        risks.append(
            "Kompetitor/alternatif masih terlalu sedikit. Tambahkan kompetitor langsung, tidak langsung, dan status quo agar positioning terlihat matang."
        )
    else:
        recommendations.append(
            f"Slide kompetisi memuat {competitor_count} alternatif. Gunakan narasi: customer saat ini memakai apa, kelemahannya apa, dan kenapa produk Anda menang."
        )

    if milestone_count < 3:
        risks.append(
            "Milestone masih kurang detail. Seed investor biasanya ingin melihat target 3-4 tahap: product, traction, revenue, dan readiness untuk round berikutnya."
        )
    else:
        recommendations.append(
            f"Gunakan {milestone_count} milestone sebagai jembatan antara ask funding, runway, dan next round story."
        )

    if not data.get("problem_evidence", "").strip():
        risks.append(
            "Evidence problem belum kuat. Tambahkan data interview, pilot, waiting list, atau biaya masalah yang dialami customer."
        )

    if not data.get("competition_summary", "").strip():
        risks.append(
            "Slide kompetisi belum punya narrative advantage. Investor biasanya ingin tahu kenapa startup ini bisa menang, bukan hanya siapa pesaingnya."
        )

    if len(str(data.get("one_liner", ""))) > 130:
        recommendations.append(
            "One-liner terlalu panjang. Buat menjadi: untuk siapa, masalah apa, solusi apa, hasil bisnis apa."
        )

    if not strengths:
        strengths.append(
            "Deck sudah memiliki struktur seed standar. Fokus berikutnya adalah memperkuat bukti pasar, angka traction, dan narrative advantage."
        )

    if not risks:
        risks.append(
            "Risiko utama belum terlihat dari input. Tetap siapkan jawaban untuk CAC, churn, defensibility, dan asumsi proyeksi."
        )

    if not recommendations:
        recommendations.append(
            "Tambahkan asumsi di speaker notes: sumber data, metode perhitungan market size, dan alasan channel GTM dipilih."
        )

    headline = (
        f"Deck readiness {score}/100. "
        f"Fokus pitching: buktikan urgency problem, kualitas traction, dan hubungan funding dengan milestone."
    )

    return {
        "score": score,
        "headline": headline,
        "metrics": {
            "growth_y2": pct(growth_y2),
            "growth_y3": pct(growth_y3),
            "burn_y1": short_money(burn_y1, currency),
            "estimated_runway": f"{estimated_runway:.0f} bulan" if estimated_runway else "N/A",
            "year3_profit_margin": pct(year3_profit_margin),
            "funding_to_y1_cost": multiple(funding_to_y1_cost),
        },
        "strengths": strengths[:4],
        "risks": risks[:4],
        "recommendations": recommendations[:4],
    }


# ==============================
# UI Guidance Helpers
# ==============================
def guide(title: str, body: str):
    st.markdown(
        f"""
        <div class="guide-box">
            <strong>{title}</strong><br>
            {body}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_insight_card(title: str, items: list[str]):
    safe_items = "".join(
        f"<li>{html.escape(str(item))}</li>"
        for item in items
    )

    st.markdown(
        f"""
        <div class="insight-card">
            <h4>{html.escape(title)}</h4>
            <ul>{safe_items}</ul>
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_insights(insights: dict[str, Any]):
    st.subheader("📈 Analisa Investor Readiness")
    st.caption(
        "Analisa ini membaca data yang Anda isi dan memberi gambaran apakah narasi pitching sudah kuat, "
        "bagian mana yang perlu dipertajam, serta insight yang bisa dipakai saat presentasi."
    )

    st.progress(insights["score"] / 100)
    st.markdown(
        f"""
        <div class="readable-panel">
            <p><strong>{html.escape(str(insights['headline']))}</strong></p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    metric_cols = st.columns(6)
    metric_labels = [
        ("Growth Y1→Y2", "growth_y2"),
        ("Growth Y2→Y3", "growth_y3"),
        ("Burn Y1", "burn_y1"),
        ("Runway estimasi", "estimated_runway"),
        ("Margin Y3", "year3_profit_margin"),
        ("Funding/Cost Y1", "funding_to_y1_cost"),
    ]

    for col, (label, key) in zip(metric_cols, metric_labels):
        col.metric(label, insights["metrics"][key])

    c1, c2, c3 = st.columns(3)

    with c1:
        render_insight_card("Strength", insights["strengths"])

    with c2:
        render_insight_card("Risk / Investor Question", insights["risks"])

    with c3:
        render_insight_card("Rekomendasi Pitching", insights["recommendations"])


# ==============================
# Slide Drawing Helpers
# ==============================
def add_text(
    slide,
    text,
    x,
    y,
    w,
    h,
    size=18,
    color=None,
    bold=False,
    align=PP_ALIGN.LEFT,
):
    color = color or THEME["ink"]

    box = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )

    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)

    paragraph = tf.paragraphs[0]
    paragraph.alignment = align

    run = paragraph.add_run()
    run.text = text or ""
    run.font.name = FONT_HEAD if bold else FONT_BODY
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color

    return box


def add_seed_header(slide, eyebrow, title, subtitle, data):
    accent = rgb(data["accent_color"])

    add_text(
        slide,
        eyebrow.upper(),
        0.7,
        0.35,
        3.8,
        0.25,
        8,
        accent,
        True,
    )

    add_text(
        slide,
        title,
        0.7,
        0.67,
        10.2,
        0.58,
        28,
        THEME["ink"],
        True,
    )

    add_text(
        slide,
        subtitle,
        0.72,
        1.25,
        11.1,
        0.32,
        10,
        THEME["muted"],
    )

    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.7),
        Inches(1.7),
        Inches(11.95),
        Inches(0.01),
    )

    line.fill.solid()
    line.fill.fore_color.rgb = THEME["line"]
    line.line.fill.background()


def add_footer(slide, data, page=None, dark=False):
    color = RGBColor(148, 163, 184) if dark else THEME["muted"]

    add_text(
        slide,
        data.get("company", "Pitch Deck"),
        0.7,
        7.08,
        3.6,
        0.22,
        8,
        color,
    )

    add_text(
        slide,
        DEVELOPER_FOOTER,
        4.25,
        7.08,
        4.8,
        0.22,
        8,
        color,
        align=PP_ALIGN.CENTER,
    )

    if page is not None:
        add_text(
            slide,
            str(page).zfill(2),
            12.0,
            7.08,
            0.65,
            0.22,
            8,
            color,
            align=PP_ALIGN.RIGHT,
        )


def add_card(slide, title, body, x, y, w, h, data):
    accent = rgb(data["accent_color"])

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )

    shape.fill.solid()
    shape.fill.fore_color.rgb = THEME["white"]
    shape.line.color.rgb = THEME["line"]

    add_text(
        slide,
        title.upper(),
        x + 0.22,
        y + 0.20,
        w - 0.44,
        0.25,
        8,
        accent,
        True,
    )

    add_text(
        slide,
        body,
        x + 0.22,
        y + 0.58,
        w - 0.44,
        h - 0.78,
        15,
        THEME["ink"],
        True,
    )


def add_big_metric(slide, label, value, x, y, w, data):
    accent = rgb(data["accent_color"])

    add_text(
        slide,
        value,
        x,
        y,
        w,
        0.55,
        28,
        accent,
        True,
    )

    add_text(
        slide,
        label.upper(),
        x,
        y + 0.62,
        w,
        0.22,
        8,
        THEME["muted"],
        True,
    )


def add_bullets(slide, items, x, y, w, h, size=18):
    box = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )

    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)

    clean_items = items or ["Lengkapi poin utama slide ini."]

    for i, item in enumerate(clean_items[:5]):
        paragraph = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        paragraph.text = item
        paragraph.level = 0
        paragraph.space_after = Pt(12)
        paragraph.font.name = FONT_BODY
        paragraph.font.size = Pt(size)
        paragraph.font.color.rgb = THEME["ink"]


def add_takeaway(slide, text, data):
    accent = rgb(data["accent_color"])

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.7),
        Inches(6.25),
        Inches(11.95),
        Inches(0.52),
    )

    shape.fill.solid()
    shape.fill.fore_color.rgb = THEME["blue_soft"]
    shape.line.color.rgb = THEME["blue_line"]

    add_text(
        slide,
        f"Investor takeaway: {text}",
        0.95,
        6.38,
        11.3,
        0.25,
        10,
        accent,
        True,
    )


def add_notes(slide, body):
    try:
        slide.notes_slide.notes_text_frame.text = body
    except Exception:
        pass


def add_table(slide, headers, rows, x, y, w, h, data):
    accent = rgb(data["accent_color"])

    table_shape = slide.shapes.add_table(
        len(rows) + 1,
        len(headers),
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )

    tbl = table_shape.table

    for col, value in enumerate(headers):
        cell = tbl.cell(0, col)
        cell.text = value
        cell.fill.solid()
        cell.fill.fore_color.rgb = accent

        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.name = FONT_BODY
            paragraph.font.size = Pt(10)
            paragraph.font.bold = True
            paragraph.font.color.rgb = THEME["white"]

    for row_i, row in enumerate(rows, 1):
        for col_i, value in enumerate(row):
            cell = tbl.cell(row_i, col_i)
            cell.text = str(value)
            cell.fill.solid()
            cell.fill.fore_color.rgb = THEME["white"]

            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.name = FONT_BODY
                paragraph.font.size = Pt(10)
                paragraph.font.color.rgb = THEME["ink"]

    return table_shape


def add_competitor_card(slide, row, x, y, w, h, data):
    accent = rgb(data["accent_color"])

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = THEME["white"]
    shape.line.color.rgb = THEME["line"]

    add_text(slide, row.get("category", "Alternative").upper(), x + 0.22, y + 0.18, w - 0.44, 0.22, 7, accent, True)
    add_text(slide, row.get("name", "Alternative"), x + 0.22, y + 0.43, w - 0.44, 0.32, 14, THEME["ink"], True)
    add_text(slide, "Weakness", x + 0.22, y + 0.86, 1.25, 0.18, 7, THEME["muted"], True)
    add_text(slide, row.get("weakness", "Belum diisi"), x + 0.22, y + 1.08, w / 2 - 0.34, 0.43, 8.5, THEME["ink"])
    add_text(slide, "Our advantage", x + w / 2 + 0.03, y + 0.86, 1.75, 0.18, 7, THEME["muted"], True)
    add_text(slide, row.get("advantage", "Belum diisi"), x + w / 2 + 0.03, y + 1.08, w / 2 - 0.30, 0.43, 8.5, THEME["ink"])


def add_competition_slides(prs, data, page):
    competitors = normalize_competitors(data)

    if not competitors:
        competitors = [
            {
                "name": "Status quo",
                "category": "Status quo",
                "weakness": "Customer tetap memakai cara lama.",
                "advantage": "Produk memberi workflow baru yang lebih cepat dan terukur.",
            }
        ]

    chunk_size = 4

    for chunk_index in range(0, len(competitors), chunk_size):
        chunk = competitors[chunk_index:chunk_index + chunk_size]
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        set_bg(slide, THEME["bg"])

        suffix = "" if len(competitors) <= chunk_size else f" ({chunk_index // chunk_size + 1}/{(len(competitors) + chunk_size - 1) // chunk_size})"
        add_seed_header(
            slide,
            "Competition",
            f"We win through focus, speed, and distribution{suffix}",
            "Layout otomatis menyesuaikan jumlah kompetitor: langsung, tidak langsung, dan status quo.",
            data,
        )

        positions = [
            (0.85, 2.00),
            (6.78, 2.00),
            (0.85, 4.05),
            (6.78, 4.05),
        ]

        for idx, row in enumerate(chunk):
            x, y = positions[idx]
            add_competitor_card(slide, row, x, y, 5.45, 1.68, data)

        add_takeaway(
            slide,
            data.get("competition_summary", "Jelaskan kenapa startup ini menang dibanding alternatif yang sudah dipakai customer."),
            data,
        )
        add_footer(slide, data, page)
        add_notes(slide, "Kompetitor bisa lebih dari satu. Gunakan slide ini untuk menunjukkan positioning dan alasan customer akan berpindah.")
        page += 1

    return page


def add_milestone_card(slide, row, x, y, w, h, data):
    accent = rgb(data["accent_color"])

    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = THEME["white"]
    shape.line.color.rgb = THEME["line"]

    add_text(slide, row.get("period", "Next"), x + 0.22, y + 0.20, w - 0.44, 0.28, 12, accent, True)
    add_text(slide, row.get("target", "Target belum diisi"), x + 0.22, y + 0.60, w - 0.44, 0.46, 13, THEME["ink"], True)
    add_text(slide, "Success metric", x + 0.22, y + 1.14, 1.65, 0.18, 7, THEME["muted"], True)
    add_text(slide, row.get("metric", "Metric belum diisi"), x + 0.22, y + 1.36, w - 0.44, 0.28, 8.5, THEME["ink"])
    add_text(slide, f"Owner: {row.get('owner', 'Team')}", x + 0.22, y + h - 0.27, w - 0.44, 0.18, 7, THEME["muted"])


def add_milestone_slides(prs, data, page):
    milestones = normalize_milestones(data)

    if not milestones:
        return page

    chunk_size = 4

    for chunk_index in range(0, len(milestones), chunk_size):
        chunk = milestones[chunk_index:chunk_index + chunk_size]
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        set_bg(slide, THEME["bg"])

        suffix = "" if len(milestones) <= chunk_size else f" ({chunk_index // chunk_size + 1}/{(len(milestones) + chunk_size - 1) // chunk_size})"
        add_seed_header(
            slide,
            "Milestones",
            f"Funding converts into measurable execution milestones{suffix}",
            "Tunjukkan target bertahap yang menghubungkan pendanaan, runway, traction, dan next round readiness.",
            data,
        )

        positions = [
            (0.85, 2.00),
            (6.78, 2.00),
            (0.85, 4.05),
            (6.78, 4.05),
        ]

        for idx, row in enumerate(chunk):
            x, y = positions[idx]
            add_milestone_card(slide, row, x, y, 5.45, 1.68, data)

        add_takeaway(
            slide,
            "Milestone yang baik harus measurable: angka revenue, user, retention, partnership, atau product release.",
            data,
        )
        add_footer(slide, data, page)
        add_notes(slide, "Gunakan milestone untuk menjelaskan kenapa jumlah funding, runway, dan next round logic saling konsisten.")
        page += 1

    return page


def seed_content_slide(
    prs,
    data,
    page,
    eyebrow,
    title,
    subtitle,
    bullets,
    takeaway,
    metric_items=None,
    side_title=None,
    side_body=None,
    speaker_note="",
):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME["bg"])

    add_seed_header(
        slide,
        eyebrow,
        title,
        subtitle,
        data,
    )

    if metric_items:
        count = len(metric_items)
        width = 11.5 / max(count, 1)

        for idx, item in enumerate(metric_items):
            label, value = item
            add_big_metric(
                slide,
                label,
                value,
                0.85 + idx * width,
                2.15,
                width - 0.2,
                data,
            )

        add_bullets(
            slide,
            bullets,
            0.85,
            3.55,
            10.8,
            2.2,
            18,
        )

    elif side_title:
        add_bullets(
            slide,
            bullets,
            0.85,
            2.05,
            6.3,
            3.7,
            20,
        )

        add_card(
            slide,
            side_title,
            side_body,
            7.7,
            2.15,
            4.35,
            2.15,
            data,
        )

    else:
        add_bullets(
            slide,
            bullets,
            0.85,
            2.05,
            10.8,
            3.8,
            21,
        )

    add_takeaway(
        slide,
        takeaway,
        data,
    )

    add_footer(
        slide,
        data,
        page,
    )

    add_notes(
        slide,
        speaker_note,
    )

    return slide


def add_insight_slide(prs, data, page):
    insights = generate_investor_insights(data)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME["bg"])

    add_seed_header(
        slide,
        "Investor Readiness",
        "What the data says about this pitch",
        "Analisa otomatis untuk membantu founder membaca kekuatan, risiko, dan pesan utama sebelum pitching.",
        data,
    )

    add_big_metric(
        slide,
        "Deck readiness score",
        f"{insights['score']}/100",
        0.85,
        2.05,
        3.2,
        data,
    )

    add_card(
        slide,
        "Growth signal",
        f"Y1→Y2: {insights['metrics']['growth_y2']}\nY2→Y3: {insights['metrics']['growth_y3']}",
        4.2,
        2.0,
        2.55,
        1.45,
        data,
    )

    add_card(
        slide,
        "Runway signal",
        f"Burn Y1: {insights['metrics']['burn_y1']}\nRunway est.: {insights['metrics']['estimated_runway']}",
        7.0,
        2.0,
        2.55,
        1.45,
        data,
    )

    add_card(
        slide,
        "Efficiency signal",
        f"Margin Y3: {insights['metrics']['year3_profit_margin']}\nFunding/Cost Y1: {insights['metrics']['funding_to_y1_cost']}",
        9.8,
        2.0,
        2.55,
        1.45,
        data,
    )

    add_text(
        slide,
        "Key strengths",
        0.85,
        3.8,
        3.3,
        0.32,
        13,
        THEME["green"],
        True,
    )
    add_bullets(slide, insights["strengths"][:3], 0.85, 4.18, 3.55, 1.75, 10)

    add_text(
        slide,
        "Investor questions",
        4.75,
        3.8,
        3.3,
        0.32,
        13,
        THEME["amber"],
        True,
    )
    add_bullets(slide, insights["risks"][:3], 4.75, 4.18, 3.55, 1.75, 10)

    add_text(
        slide,
        "Pitching recommendations",
        8.6,
        3.8,
        3.3,
        0.32,
        13,
        rgb(data["accent_color"]),
        True,
    )
    add_bullets(slide, insights["recommendations"][:3], 8.6, 4.18, 3.55, 1.75, 10)

    add_footer(slide, data, page)
    add_notes(
        slide,
        "Gunakan slide ini sebagai internal readiness view. Bila deck dibagikan ke investor, slide ini bisa dipertahankan atau dihapus sesuai kebutuhan.",
    )


# ==============================
# Deck Generator
# ==============================
def build_deck(data, image_buffer=None):
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    accent = rgb(data["accent_color"])
    page = 1

    # 1. Cover
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME["dark"])

    add_text(slide, data["company"], 0.8, 0.85, 8.5, 0.65, 36, THEME["white"], True)
    add_text(slide, data["one_liner"], 0.82, 1.75, 8.9, 0.85, 24, RGBColor(226, 232, 240), True)
    add_text(slide, f'{data["round"]} • {money(data["ask"], data["currency"])}', 0.85, 4.85, 7.5, 0.35, 15, accent, True)
    add_text(slide, f'{data["presenter"]} | {data["contact"]}', 0.85, 5.35, 7.5, 0.28, 11, RGBColor(203, 213, 225))
    add_text(slide, "Seed Investor Pitch", 9.6, 6.72, 2.8, 0.3, 10, RGBColor(148, 163, 184), True, PP_ALIGN.RIGHT)

    add_footer(slide, data, page, dark=True)
    add_notes(slide, "Open with a concise narrative: customer, pain, solution, and why this can become venture-scale.")
    page += 1

    seed_content_slide(
        prs,
        data,
        page,
        "Problem",
        "The current workflow is broken and expensive",
        "Tunjukkan pain point yang sering terjadi, mahal, dan cukup besar untuk menjadi venture-scale opportunity.",
        lines(data["problem"], 5),
        "Problem harus terasa urgent, bukan sekadar nice-to-have.",
        side_title="Proof of pain",
        side_body=data["problem_evidence"],
        speaker_note="Mulai dari pain point, frekuensi masalah, dan biaya masalah. Jangan mulai dari fitur.",
    )
    page += 1

    seed_content_slide(
        prs,
        data,
        page,
        "Solution",
        "A simpler way to solve the problem",
        "Jelaskan perubahan sebelum dan sesudah produk dipakai.",
        lines(data["solution"], 5),
        "Solusi harus langsung menjawab problem utama.",
        side_title="Core value proposition",
        side_body=data["value_prop"],
        speaker_note="Jelaskan perubahan kondisi pelanggan sebelum dan sesudah memakai produk.",
    )
    page += 1

    # Product
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME["bg"])
    add_seed_header(slide, "Product", "The product turns user activity into business outcomes", "Gunakan slide ini untuk demo flow utama, bukan daftar fitur panjang.", data)

    if image_buffer:
        image_buffer.seek(0)
        slide.shapes.add_picture(image_buffer, Inches(0.85), Inches(2.05), width=Inches(6.2))
        add_card(slide, "Product flow", data["product_flow"], 7.55, 2.05, 4.55, 1.65, data)
        add_card(slide, "Product benefit", data["product_benefit"], 7.55, 4.05, 4.55, 1.25, data)
    else:
        add_bullets(slide, lines(data["features"], 5), 0.85, 2.05, 6.4, 3.7, 20)
        add_card(slide, "Product flow", data["product_flow"], 7.7, 2.15, 4.35, 1.75, data)
        add_card(slide, "Product benefit", data["product_benefit"], 7.7, 4.2, 4.35, 1.15, data)

    add_takeaway(slide, "Investor harus paham cara produk menciptakan nilai dalam 30 detik.", data)
    add_footer(slide, data, page)
    add_notes(slide, "Demo singkat: input, proses, output, dan dampak untuk pelanggan.")
    page += 1

    seed_content_slide(
        prs,
        data,
        page,
        "Market",
        "A large market with a focused entry wedge",
        "TAM menunjukkan potensi; SOM menunjukkan fokus eksekusi awal.",
        lines(data["market_notes"], 5),
        "Pasar awal harus spesifik dan bisa dimenangkan.",
        metric_items=[("TAM", data["tam"]), ("SAM", data["sam"]), ("SOM", data["som"])],
        speaker_note="Tekankan wedge pasar awal, bukan hanya ukuran pasar yang besar.",
    )
    page += 1

    seed_content_slide(
        prs,
        data,
        page,
        "Business Model",
        "Revenue can scale with attractive unit economics",
        "Tampilkan cara menghasilkan uang dan indikasi kualitas bisnis.",
        lines(data["business_model"], 5),
        "Seed investor mencari sinyal bahwa bisnis bisa menjadi scalable.",
        metric_items=[("ARPU", data["arpu"]), ("Gross Margin", data["gross_margin"]), ("CAC", data["cac"]), ("Payback", data["payback"])],
        speaker_note="Jelaskan siapa yang membayar, kapan membayar, dan kenapa margin bisa naik saat skala naik.",
    )
    page += 1

    seed_content_slide(
        prs,
        data,
        page,
        "Traction",
        "Early demand is already visible",
        "Untuk seed, traction bisa berupa revenue, active usage, retention, pilot, LOI, atau pipeline.",
        lines(data["traction_notes"], 5),
        "Traction harus membuktikan demand, bukan vanity metrics.",
        metric_items=[("Users", data["users"]), ("Revenue", data["revenue"]), ("Growth", data["growth"]), ("Retention", data["retention"])],
        speaker_note="Gunakan metrik demand: revenue, active usage, retention, pilot, LOI, atau pipeline.",
    )
    page += 1

    seed_content_slide(
        prs,
        data,
        page,
        "Go-To-Market",
        "A repeatable acquisition motion is emerging",
        "Jelaskan ICP, channel utama, dan cara distribusi bisa berulang.",
        lines(data["gtm"], 5),
        "GTM harus menunjukkan mesin akuisisi, bukan hanya rencana marketing.",
        side_title="ICP & channel",
        side_body=f'{data["icp"]}\n\nPrimary channel: {data["channel"]}',
        speaker_note="Tunjukkan lead source, conversion, dan biaya akuisisi pelanggan.",
    )
    page += 1

    page = add_competition_slides(prs, data, page)

    # Financials
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME["bg"])
    add_seed_header(slide, "Financials", "Funding converts into measurable milestones", "Seed financials harus sederhana, berbasis asumsi, dan terhubung ke runway.", data)
    add_table(
        slide,
        ["Metric", "Year 1", "Year 2", "Year 3"],
        [
            ["Revenue", money(data["rev1"], data["currency"]), money(data["rev2"], data["currency"]), money(data["rev3"], data["currency"])],
            ["Operating Cost", money(data["cost1"], data["currency"]), money(data["cost2"], data["currency"]), money(data["cost3"], data["currency"])],
            ["EBITDA / Profit", money(data["profit1"], data["currency"]), money(data["profit2"], data["currency"]), money(data["profit3"], data["currency"])],
        ],
        0.85,
        2.05,
        11.65,
        2.7,
        data,
    )
    add_big_metric(slide, "Runway", f'{data["runway"]} months', 0.95, 5.3, 3.0, data)
    add_big_metric(slide, "Next milestone", milestone_headline(data), 4.25, 5.3, 7.4, data)
    add_footer(slide, data, page)
    add_notes(slide, "Jelaskan asumsi utama. Hubungkan funding ke milestone berikutnya.")
    page += 1

    page = add_milestone_slides(prs, data, page)

    if data.get("include_insight_slide", True):
        add_insight_slide(prs, data, page)
        page += 1

    seed_content_slide(
        prs,
        data,
        page,
        "Team",
        "The right team for this market",
        "Di seed stage, investor banyak menilai founder-market fit.",
        lines(data["team"], 5),
        "Tim harus terlihat punya unfair advantage untuk mengeksekusi bisnis ini.",
        side_title="Founder-market fit",
        side_body=data["founder_fit"],
        speaker_note="Hubungkan pengalaman tim dengan problem dan akses ke market.",
    )
    page += 1

    seed_content_slide(
        prs,
        data,
        page,
        "Fundraising Ask",
        "We are raising to reach the next inflection point",
        "Hubungkan jumlah dana dengan runway, penggunaan dana, dan milestone berikutnya.",
        lines(data["use_of_funds"], 5),
        "Ask harus jelas: jumlah, penggunaan, runway, dan target pembuktian.",
        side_title=data["round"],
        side_body=f'{money(data["ask"], data["currency"])}\n\nNext milestone: {milestone_headline(data)}\n\n{data["next_round"]}',
        speaker_note="Tutup dengan jumlah dana, runway, use of funds, dan target 12-18 bulan.",
    )
    page += 1

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME["dark"])
    add_text(slide, data["closing"], 0.85, 1.25, 9.8, 1.1, 34, THEME["white"], True)
    add_text(slide, f'{data["presenter"]}\n{data["contact"]}', 0.9, 5.35, 7.0, 0.6, 14, RGBColor(203, 213, 225))
    add_footer(slide, data, page, dark=True)
    add_notes(slide, "Akhiri dengan ajakan follow-up: demo, data room, atau meeting berikutnya.")

    output = BytesIO()
    prs.save(output)
    output.seek(0)

    return output


# ==============================
# Streamlit UI
# ==============================
st.title("Seed Investor Pitch Deck Generator")
st.caption(
    "Generator PPTX pitch deck profesional untuk seed-stage startup: story-led, data-first, "
    "minim teks, kompetitor dinamis, milestone execution plan, dan analisa investor readiness otomatis."
)

st.markdown(f'<div class="developer-footer">{DEVELOPER_FOOTER}</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("Brand & Format")
    guide(
        "Pengaturan visual deck",
        "Pilih warna aksen yang sesuai brand. Tampilan aplikasi sudah theme-aware: ketika Streamlit memakai light, dark, atau system theme, warna teks, kartu, input, tab, dan panel analisa akan menyesuaikan agar tetap terbaca.",
    )
    accent_color = st.color_picker(
        "Warna aksen",
        "#2563EB",
        help="Warna ini dipakai untuk headline kecil, metric utama, tabel, dan highlight investor takeaway.",
    )
    currency = st.selectbox(
        "Mata uang",
        ["Rp", "USD"],
        help="Pilih mata uang yang dipakai pada ask funding dan proyeksi finansial.",
    )
    include_insight_slide = st.checkbox(
        "Tambahkan slide analisa otomatis",
        value=True,
        help="Jika aktif, PPTX akan menyertakan slide Investor Readiness berisi insight dari data yang diinput.",
    )
    st.success(DEVELOPER_FOOTER)

identity, story, market, finance, team_asset, insight_tab = st.tabs(
    [
        "Identitas",
        "Story",
        "Market & Traction",
        "Financial & Funding",
        "Team & Competition",
        "Analisa",
    ]
)

with identity:
    guide(
        "Identitas deck",
        "Bagian ini membentuk slide pembuka. Investor harus langsung paham: nama startup, siapa yang presentasi, startup menyelesaikan masalah apa, dan berapa pendanaan yang sedang dicari.",
    )

    col1, col2 = st.columns(2)

    with col1:
        company = st.text_input(
            "Nama startup",
            "Rex AI",
            help="Tulis nama legal/brand startup. Nama ini juga muncul di footer setiap slide.",
        )
        one_liner = st.text_input(
            "One-liner",
            "AI copilot untuk membantu UMKM membuat laporan keuangan otomatis.",
            help="Formula praktis: [produk] untuk [customer] agar [hasil bisnis]. Hindari slogan abstrak.",
        )
        presenter = st.text_input(
            "Presenter",
            "Rex Founder",
            help="Nama founder atau orang yang akan melakukan pitching.",
        )

    with col2:
        contact = st.text_input(
            "Kontak",
            "founder@rex.ai",
            help="Email atau kontak bisnis yang akan muncul di cover dan closing slide.",
        )
        round_name = st.text_input(
            "Round",
            "Seed Round",
            help="Contoh: Pre-Seed Round, Seed Round, Bridge Round. Sesuaikan dengan tahap fundraising.",
        )
        ask = st.number_input(
            "Jumlah pendanaan",
            min_value=0,
            value=1_500_000_000,
            step=50_000_000,
            help="Jumlah dana yang dicari. Pastikan selaras dengan runway dan milestone 12-18 bulan.",
        )

with story:
    guide(
        "Story: problem, solution, dan product",
        "Bagian ini adalah jantung pitch. Investor perlu melihat masalah yang urgent, solusi yang tajam, dan produk yang bisa menjawab masalah dengan cara lebih baik dari alternatif yang ada.",
    )

    col1, col2 = st.columns(2)

    with col1:
        problem = st.text_area(
            "Problem",
            "UMKM sulit mencatat transaksi harian.\nLaporan keuangan masih manual dan rawan salah.\nPemilik usaha tidak punya insight real-time.",
            height=130,
            help="Tulis 3-5 pain point utama. Fokus pada masalah yang sering, mahal, dan penting untuk customer.",
        )
        problem_evidence = st.text_area(
            "Evidence problem",
            "Interview awal: 24 dari 30 UMKM masih memakai catatan manual atau spreadsheet.",
            height=90,
            help="Bukti bahwa problem nyata: interview, pilot, data market, waiting list, churn dari solusi lama, atau biaya masalah.",
        )
        solution = st.text_area(
            "Solution",
            "Input transaksi via chat.\nOCR struk otomatis.\nDashboard cashflow dan laba rugi.\nRekomendasi aksi berbasis data.",
            height=130,
            help="Jelaskan cara solusi bekerja dalam 3-5 poin. Hindari terlalu banyak fitur teknis.",
        )
        value_prop = st.text_area(
            "Value proposition",
            "Menghemat waktu pencatatan dan membantu owner memahami kondisi keuangan tanpa akuntan penuh waktu.",
            height=90,
            help="Tulis manfaat utama yang dirasakan customer: hemat waktu, naik revenue, turun biaya, lebih cepat, lebih akurat, atau lebih mudah.",
        )

    with col2:
        product_flow = st.text_area(
            "Cara kerja produk",
            "User kirim transaksi/foto struk → sistem klasifikasi otomatis → dashboard dan laporan dibuat real-time.",
            height=90,
            help="Tulis alur produk dari input → proses → output → dampak. Ini membantu saat demo pitching.",
        )
        product_benefit = st.text_area(
            "Benefit produk",
            "Owner melihat omzet, margin, utang-piutang, dan cashflow dalam satu tempat.",
            height=90,
            help="Tulis outcome produk, bukan fitur. Contoh: keputusan lebih cepat, error turun, cashflow lebih jelas.",
        )
        features = st.text_area(
            "Fitur utama",
            "OCR struk\nAuto-categorization\nCashflow dashboard\nExport laporan PDF/Excel\nReminder piutang",
            height=130,
            help="Pilih maksimal 5 fitur yang paling mendukung value proposition. Jangan jadikan slide sebagai daftar fitur panjang.",
        )
        product_image = st.file_uploader(
            "Screenshot/mockup produk",
            type=["png", "jpg", "jpeg"],
            help="Upload mockup, screenshot, atau visual produk. Untuk seed deck, visual produk sangat membantu investor memahami demo secara cepat.",
        )

with market:
    guide(
        "Market, business model, traction, dan GTM",
        "Bagian ini menjawab: seberapa besar peluangnya, siapa customer awal, bagaimana startup menghasilkan uang, apakah ada demand awal, dan bagaimana akuisisi pelanggan akan diulang.",
    )

    col1, col2 = st.columns(2)

    with col1:
        tam = st.text_input("TAM", "Rp 120T", help="Total Addressable Market: keseluruhan peluang pasar jika startup menang besar.")
        sam = st.text_input("SAM", "Rp 18T", help="Serviceable Available Market: pasar yang realistis dijangkau oleh produk dan model bisnis saat ini.")
        som = st.text_input("SOM", "Rp 450M", help="Serviceable Obtainable Market: target pasar awal yang realistis dimenangkan dalam beberapa tahun pertama.")
        market_notes = st.text_area(
            "Market notes",
            "Target awal: UMKM F&B dan ritel.\nWedge: bisnis yang sudah memakai WhatsApp untuk operasional.\nEkspansi: inventory, payroll, dan embedded financing.",
            height=120,
            help="Jelaskan segmen awal, alasan memilih segmen tersebut, dan rencana ekspansi pasar.",
        )
        arpu = st.text_input("ARPU", "Rp 99.000/bulan", help="Average Revenue Per User. Tulis pendapatan rata-rata per customer/user.")
        gross_margin = st.text_input("Gross margin", "78%", help="Margin kotor. Untuk software/SaaS, margin tinggi menjadi sinyal scalability.")
        cac = st.text_input("CAC", "Rp 140.000", help="Customer Acquisition Cost. Tulis biaya rata-rata untuk mendapatkan satu customer.")
        payback = st.text_input("Payback", "< 2 bulan", help="Periode balik modal CAC. Semakin pendek, semakin kuat unit economics.")
        business_model = st.text_area(
            "Business model",
            "Subscription SaaS bulanan.\nPaket Pro untuk multi-outlet.\nAdd-on laporan pajak dan inventory.\nRevenue share dari pembiayaan UMKM.",
            height=120,
            help="Jelaskan siapa yang membayar, berapa, kapan membayar, dan potensi expansion revenue.",
        )

    with col2:
        users = st.text_input("Users/customers", "1.200 users", help="Jumlah user/customer aktif, customer berbayar, pilot, atau akun terdaftar. Sebutkan konteksnya.")
        revenue = st.text_input("Revenue/GMV", "Rp 85 juta MRR", help="Revenue bulanan, ARR, GMV, atau pipeline. Untuk investor, revenue berulang lebih kuat daripada vanity metrics.")
        growth = st.text_input("Growth", "+28% MoM", help="Pertumbuhan month-over-month atau quarter-over-quarter. Jelaskan basis perhitungannya.")
        retention = st.text_input("Retention", "72% D30", help="Retention D7/D30, logo retention, revenue retention, atau repeat usage. Retention menjawab apakah produk benar-benar dipakai ulang.")
        traction_notes = st.text_area(
            "Traction notes",
            "Pilot berbayar dengan 35 UMKM.\nPipeline 180 UMKM dari partner.\nChurn turun setelah fitur reminder piutang dirilis.",
            height=120,
            help="Tulis bukti demand: revenue, active usage, LOI, pilot, partnership, pipeline, atau waiting list.",
        )
        icp = st.text_area(
            "ICP",
            "Pemilik UMKM F&B/ritel dengan 3-50 transaksi per hari.",
            height=80,
            help="Ideal Customer Profile. Spesifikkan jenis customer yang paling sakit masalahnya dan paling mudah dijangkau.",
        )
        channel = st.text_area(
            "Channel utama",
            "Komunitas UMKM, reseller akuntansi, konten edukasi, partnership POS.",
            height=80,
            help="Channel akuisisi utama: komunitas, sales, referral, partnership, content, outbound, marketplace, atau channel lain.",
        )
        gtm = st.text_area(
            "GTM",
            "Akuisisi lewat komunitas UMKM.\nReferral program.\nPartnership konsultan pajak dan POS.\nKonten edukasi cashflow sebagai lead magnet.",
            height=110,
            help="Go-To-Market: jelaskan proses akuisisi dari lead source, conversion, onboarding, sampai retention.",
        )

with finance:
    guide(
        "Financial, funding ask, dan milestone",
        "Bagian ini membantu investor melihat apakah jumlah pendanaan masuk akal. Hubungkan funding dengan runway, penggunaan dana, dan milestone yang bisa dicapai dalam 12-18 bulan.",
    )

    col1, col2 = st.columns(2)

    with col1:
        rev1 = st.number_input("Revenue Y1", value=1_200_000_000, step=100_000_000, help="Proyeksi revenue tahun pertama setelah funding. Gunakan asumsi realistis.")
        rev2 = st.number_input("Revenue Y2", value=4_800_000_000, step=100_000_000, help="Proyeksi revenue tahun kedua. Pastikan growth didukung asumsi customer, ARPU, dan channel.")
        rev3 = st.number_input("Revenue Y3", value=13_500_000_000, step=100_000_000, help="Proyeksi revenue tahun ketiga. Gunakan untuk menunjukkan potensi skala.")
        cost1 = st.number_input("Operating Cost Y1", value=2_000_000_000, step=100_000_000, help="Biaya operasional tahun pertama: tim, produk, marketing, sales, tools, legal, operasional.")
        cost2 = st.number_input("Operating Cost Y2", value=5_200_000_000, step=100_000_000, help="Biaya operasional tahun kedua.")
        cost3 = st.number_input("Operating Cost Y3", value=10_000_000_000, step=100_000_000, help="Biaya operasional tahun ketiga.")

    with col2:
        profit1 = st.number_input("EBITDA/Profit Y1", value=-800_000_000, step=100_000_000, help="EBITDA/profit tahun pertama. Untuk seed, negatif masih wajar jika diarahkan ke growth.")
        profit2 = st.number_input("EBITDA/Profit Y2", value=-400_000_000, step=100_000_000, help="EBITDA/profit tahun kedua.")
        profit3 = st.number_input("EBITDA/Profit Y3", value=3_500_000_000, step=100_000_000, help="EBITDA/profit tahun ketiga. Ini membantu menunjukkan operating leverage.")
        runway = st.number_input("Runway / bulan", min_value=1, value=18, help="Berapa bulan startup bisa berjalan dengan dana yang dicari. Seed biasanya ditargetkan 12-18 bulan, bergantung strategi.")
        milestone = st.text_input(
            "Milestone utama / headline",
            "Rp 500 juta MRR, 8.000 active businesses, gross margin >80%",
            help="Ringkasan milestone paling penting yang muncul di slide Financials dan Ask.",
        )
        use_of_funds = st.text_area(
            "Use of funds",
            "Product & engineering: 40%\nSales & marketing: 35%\nCustomer success: 15%\nOperations & legal: 10%",
            height=110,
            help="Pembagian penggunaan dana. Buat logis dan langsung terhubung ke milestone.",
        )
        next_round = st.text_area(
            "Next round logic",
            "Raise Seed/Series A setelah akuisisi repeatable dan retention terbukti.",
            height=80,
            help="Jelaskan kondisi apa yang membuat startup siap raise round berikutnya.",
        )

    guide(
        "Detail milestone",
        "Tambahkan milestone bertahap agar investor melihat rute eksekusi. Setiap milestone sebaiknya punya periode, target, success metric, dan owner. Jika lebih dari 4 milestone, PPT akan otomatis memecahnya menjadi beberapa slide.",
    )

    milestone_defaults = [
        {
            "period": "0-3 bulan",
            "target": "Rilis MVP stabil dan onboarding 100 customer awal",
            "metric": "100 active customers, activation >60%",
            "owner": "Product & Growth",
        },
        {
            "period": "4-6 bulan",
            "target": "Validasi channel akuisisi utama",
            "metric": "CAC payback <3 bulan, 500 customers",
            "owner": "Growth",
        },
        {
            "period": "7-12 bulan",
            "target": "Capai repeatable revenue motion",
            "metric": "Rp 250 juta MRR, retention D30 >70%",
            "owner": "Sales & CS",
        },
        {
            "period": "13-18 bulan",
            "target": "Siap raise next round dengan traction kuat",
            "metric": "Rp 500 juta MRR, gross margin >80%",
            "owner": "Leadership",
        },
    ]

    milestone_count = int(
        st.number_input(
            "Jumlah milestone detail",
            min_value=1,
            max_value=8,
            value=4,
            step=1,
            help="PPT akan menyesuaikan layout otomatis. Maksimal 4 milestone per slide.",
        )
    )

    milestones = []

    for idx in range(milestone_count):
        default = milestone_defaults[idx] if idx < len(milestone_defaults) else {
            "period": f"Tahap {idx + 1}",
            "target": "Target milestone",
            "metric": "Metric keberhasilan",
            "owner": "Team",
        }

        with st.expander(f"Milestone {idx + 1}: {default['period']}", expanded=idx < 2):
            m1, m2 = st.columns(2)
            with m1:
                period = st.text_input(
                    f"Periode milestone {idx + 1}",
                    default["period"],
                    key=f"milestone_period_{idx}",
                    help="Contoh: 0-3 bulan, Q1 2027, sebelum Seed extension, sebelum Series A.",
                )
                target = st.text_area(
                    f"Target milestone {idx + 1}",
                    default["target"],
                    height=80,
                    key=f"milestone_target_{idx}",
                    help="Target eksekusi yang ingin dicapai pada periode ini.",
                )
            with m2:
                metric = st.text_input(
                    f"Success metric {idx + 1}",
                    default["metric"],
                    key=f"milestone_metric_{idx}",
                    help="Buat measurable: MRR, active user, retention, CAC payback, gross margin, jumlah partnership, atau product release.",
                )
                owner = st.text_input(
                    f"Owner milestone {idx + 1}",
                    default["owner"],
                    key=f"milestone_owner_{idx}",
                    help="Tim/role yang bertanggung jawab: Product, Growth, Sales, CS, Engineering, atau Leadership.",
                )

        milestones.append(
            {
                "period": period,
                "target": target,
                "metric": metric,
                "owner": owner,
            }
        )

with team_asset:
    guide(
        "Team dan competition",
        "Investor seed sangat memperhatikan founder-market fit dan alasan startup ini bisa menang. Tambahkan kompetitor langsung, tidak langsung, dan status quo. PPT akan otomatis menyesuaikan layout jika kompetitor lebih dari satu atau lebih dari empat.",
    )

    col1, col2 = st.columns(2)

    with col1:
        team = st.text_area(
            "Team",
            "Rex - CEO - 6 tahun membangun SaaS UMKM.\nNadia - CTO - ex fintech data engineer.\nBima - Growth - pernah scale komunitas UMKM 50k members.",
            height=120,
            help="Tulis nama, role, dan pengalaman relevan. Fokus pada kemampuan yang langsung berhubungan dengan problem, produk, dan market.",
        )
        founder_fit = st.text_area(
            "Founder-market fit",
            "Tim pernah membangun tools operasional untuk UMKM dan punya akses langsung ke komunitas target.",
            height=90,
            help="Jelaskan unfair advantage tim: pengalaman domain, akses distribusi, technical edge, atau insight unik dari market.",
        )
        closing = st.text_input(
            "Closing line",
            "Let’s help millions of small businesses understand their money.",
            help="Kalimat penutup yang merangkum visi besar startup.",
        )

    with col2:
        competition_summary = st.text_area(
            "Ringkasan kompetisi / narrative advantage",
            "Kami menang di simplicity, distribusi lokal, dan data workflow UMKM.",
            height=100,
            help="Satu kalimat narrative advantage: kenapa startup ini bisa menang dibanding alternatif yang sudah dipakai customer.",
        )
        competitor_count = int(
            st.number_input(
                "Jumlah kompetitor / alternatif",
                min_value=1,
                max_value=10,
                value=4,
                step=1,
                help="Masukkan kompetitor langsung, tidak langsung, dan status quo. Maksimal 4 kompetitor per slide; jika lebih, PPT otomatis memecah slide.",
            )
        )

    competitor_defaults = [
        {
            "name": "Aplikasi Akuntansi A",
            "category": "Direct competitor",
            "weakness": "Terlalu kompleks untuk mikro-UMKM",
            "advantage": "Workflow chat-first dan onboarding cepat",
        },
        {
            "name": "Spreadsheet/manual",
            "category": "Indirect competitor",
            "weakness": "Tidak real-time dan rawan error",
            "advantage": "Otomatisasi laporan dan insight",
        },
        {
            "name": "Status quo",
            "category": "Status quo",
            "weakness": "Owner tidak punya data untuk keputusan cepat",
            "advantage": "Data harian langsung menjadi rekomendasi aksi",
        },
        {
            "name": "POS tanpa analitik keuangan",
            "category": "Adjacent tool",
            "weakness": "Mencatat transaksi tetapi tidak memberi insight cashflow",
            "advantage": "Mengubah transaksi menjadi laporan dan rekomendasi bisnis",
        },
    ]

    competitors = []

    st.markdown("### Peta kompetitor")
    st.caption(
        "Untuk deck seed, kompetitor tidak harus hanya produk yang sama. Masukkan juga cara manual, spreadsheet, agency, status quo, atau adjacent tool yang saat ini dipakai customer."
    )

    for idx in range(competitor_count):
        default = competitor_defaults[idx] if idx < len(competitor_defaults) else {
            "name": f"Kompetitor {idx + 1}",
            "category": "Alternative",
            "weakness": "Kelemahan dari sudut pandang customer",
            "advantage": "Keunggulan produk Anda",
        }

        with st.expander(f"Kompetitor / alternatif {idx + 1}: {default['name']}", expanded=idx < 3):
            c1, c2 = st.columns(2)
            with c1:
                comp_name = st.text_input(
                    f"Nama kompetitor / alternatif {idx + 1}",
                    default["name"],
                    key=f"competitor_name_{idx}",
                    help="Nama produk, perusahaan, workflow manual, spreadsheet, agency, atau status quo.",
                )
                comp_category = st.selectbox(
                    f"Kategori {idx + 1}",
                    ["Direct competitor", "Indirect competitor", "Status quo", "Adjacent tool", "Alternative"],
                    index=["Direct competitor", "Indirect competitor", "Status quo", "Adjacent tool", "Alternative"].index(default["category"]) if default["category"] in ["Direct competitor", "Indirect competitor", "Status quo", "Adjacent tool", "Alternative"] else 4,
                    key=f"competitor_category_{idx}",
                    help="Kategorisasi membantu investor memahami medan kompetisi.",
                )
            with c2:
                comp_weakness = st.text_area(
                    f"Kelemahan {idx + 1}",
                    default["weakness"],
                    height=80,
                    key=f"competitor_weakness_{idx}",
                    help="Tulis kelemahan dari perspektif customer: mahal, lambat, rumit, tidak real-time, sulit diadopsi, tidak terintegrasi, dsb.",
                )
                comp_advantage = st.text_area(
                    f"Keunggulan kita vs alternatif {idx + 1}",
                    default["advantage"],
                    height=80,
                    key=f"competitor_advantage_{idx}",
                    help="Tulis keunggulan spesifik dan defensible: distribusi, data, UX, cost, speed, domain expertise, switching workflow, dsb.",
                )

        competitors.append(
            {
                "name": comp_name,
                "category": comp_category,
                "weakness": comp_weakness,
                "advantage": comp_advantage,
            }
        )

# Collect current data for analysis and generation.
data = {
    "company": company,
    "one_liner": one_liner,
    "presenter": presenter,
    "contact": contact,
    "round": round_name,
    "ask": ask,
    "currency": currency,
    "accent_color": accent_color,
    "include_insight_slide": include_insight_slide,
    "problem": problem,
    "problem_evidence": problem_evidence,
    "solution": solution,
    "value_prop": value_prop,
    "product_flow": product_flow,
    "product_benefit": product_benefit,
    "features": features,
    "tam": tam,
    "sam": sam,
    "som": som,
    "market_notes": market_notes,
    "arpu": arpu,
    "gross_margin": gross_margin,
    "cac": cac,
    "payback": payback,
    "business_model": business_model,
    "users": users,
    "revenue": revenue,
    "growth": growth,
    "retention": retention,
    "traction_notes": traction_notes,
    "icp": icp,
    "channel": channel,
    "gtm": gtm,
    "competitors": competitors,
    "competitor_1": competitors[0]["name"] if len(competitors) > 0 else "",
    "weakness_1": competitors[0]["weakness"] if len(competitors) > 0 else "",
    "advantage_1": competitors[0]["advantage"] if len(competitors) > 0 else "",
    "competitor_2": competitors[1]["name"] if len(competitors) > 1 else "",
    "weakness_2": competitors[1]["weakness"] if len(competitors) > 1 else "",
    "advantage_2": competitors[1]["advantage"] if len(competitors) > 1 else "",
    "status_quo": competitors[2]["weakness"] if len(competitors) > 2 else "",
    "advantage_3": competitors[2]["advantage"] if len(competitors) > 2 else "",
    "competition_summary": competition_summary,
    "rev1": rev1,
    "rev2": rev2,
    "rev3": rev3,
    "cost1": cost1,
    "cost2": cost2,
    "cost3": cost3,
    "profit1": profit1,
    "profit2": profit2,
    "profit3": profit3,
    "runway": runway,
    "milestone": milestone,
    "milestones": milestones,
    "use_of_funds": use_of_funds,
    "next_round": next_round,
    "team": team,
    "founder_fit": founder_fit,
    "closing": closing,
}

with insight_tab:
    guide(
        "Analisa otomatis dari data pitching",
        "Bagian ini tidak perlu diisi manual. Sistem membaca input Anda dan memberi ringkasan kekuatan, risiko pertanyaan investor, serta rekomendasi perbaikan narasi sebelum deck di-download.",
    )
    show_insights(generate_investor_insights(data))

st.divider()

col_generate, col_hint = st.columns([1, 2])

with col_generate:
    generate = st.button("Generate Seed Investor Pitch Deck", type="primary", use_container_width=True)

with col_hint:
    st.caption(
        "Sebelum generate, cek tab Analisa untuk memastikan narasi problem, traction, financial, dan ask sudah konsisten. "
        "PPTX akan otomatis membawa footer Developed by Galuh Adi Insani."
    )

if generate:
    if not company.strip():
        st.error("Nama startup wajib diisi.")
        st.stop()

    image_buffer = BytesIO(product_image.read()) if product_image else None
    pptx = build_deck(data, image_buffer)

    st.success("Pitch deck berhasil dibuat.")
    show_insights(generate_investor_insights(data))

    st.download_button(
        "📥 Download Seed Investor Pitch Deck (.pptx)",
        data=pptx,
        file_name=f"{filename(company)}-seed-investor-pitch-deck.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )
