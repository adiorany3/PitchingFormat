import base64
import html
import json
import math
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from typing import Any

import streamlit as st
import streamlit.components.v1 as components
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.util import Inches, Pt
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table as RLTable,
    TableStyle,
)

DEVELOPER = "Developed by Galuh Adi Insani"
APP_VERSION = "v10.6 - Forced Readable Theme"

st.set_page_config(
    page_title="Seed Investor Pitch Deck Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


def hide_streamlit_emblems() -> None:
    st.markdown(
        """
        <style>
            :root {
                --deck-bg: var(--background-color, #ffffff);
                --deck-surface: var(--secondary-background-color, #f8fafc);
                --deck-text: var(--text-color, #0f172a);
                --deck-primary: var(--primary-color, #2563eb);
                --deck-border: rgba(128, 128, 128, 0.38);
                --deck-border-strong: rgba(128, 128, 128, 0.58);
                --deck-radius: 16px;
                --deck-shadow: 0 12px 30px rgba(0, 0, 0, 0.10);
            }

            #MainMenu, footer, header,
            [data-testid="stHeader"],
            [data-testid="stToolbar"],
            [data-testid="stDecoration"],
            [data-testid="stStatusWidget"],
            [data-testid="stDeployButton"],
            [data-testid="collapsedControl"],
            .stAppDeployButton,
            .viewerBadge_container__1QSob,
            .viewerBadge_link__1S137,
            .viewerBadge_text__1JaDK {
                display: none !important;
                visibility: hidden !important;
            }

            .stApp {
                background: var(--deck-bg) !important;
                color: var(--deck-text) !important;
            }

            .block-container {
                max-width: 1360px;
                padding-top: 1.2rem !important;
                padding-bottom: 5.8rem !important;
            }

            [data-testid="stSidebar"] > div:first-child {
                background: var(--deck-surface) !important;
                border-right: 1px solid var(--deck-border) !important;
            }

            /* Global text contrast guard */
            .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
            .stApp p, .stApp li, .stApp label, .stApp span,
            [data-testid="stMarkdownContainer"],
            [data-testid="stMarkdownContainer"] *,
            [data-testid="stCaptionContainer"],
            [data-testid="stWidgetLabel"],
            [data-testid="stWidgetLabel"] *,
            [data-testid="stExpander"] *,
            [data-testid="stMetric"] *,
            [data-testid="stFileUploader"] * {
                color: var(--deck-text) !important;
            }

            [data-testid="stCaptionContainer"],
            [data-testid="InputInstructions"],
            [data-testid="stHelp"],
            .caption, small {
                color: var(--deck-text) !important;
                opacity: 0.88 !important;
            }

            /* Tabs */
            [data-testid="stTabs"] [role="tablist"] {
                gap: .35rem;
                border-bottom: 1px solid var(--deck-border) !important;
                flex-wrap: wrap;
            }
            [data-testid="stTabs"] button[role="tab"] {
                min-height: 42px;
                padding: .55rem .85rem;
                border-radius: 999px 999px 0 0;
                color: var(--deck-text) !important;
                background: transparent !important;
                border: 1px solid transparent !important;
                opacity: 0.85 !important;
            }
            [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
                color: var(--deck-text) !important;
                background: var(--deck-surface) !important;
                border: 1px solid var(--deck-border) !important;
                border-bottom-color: var(--deck-surface) !important;
                font-weight: 800 !important;
                opacity: 1 !important;
            }

            /* Inputs and BaseWeb widgets */
            [data-testid="stTextInput"] input,
            [data-testid="stNumberInput"] input,
            [data-testid="stTextArea"] textarea,
            [data-baseweb="input"] input,
            [data-baseweb="textarea"] textarea,
            [data-baseweb="select"] > div,
            [data-testid="stFileUploader"] section,
            [data-testid="stColorPicker"] input {
                background: var(--deck-surface) !important;
                color: var(--deck-text) !important;
                border-color: var(--deck-border-strong) !important;
                caret-color: var(--deck-primary) !important;
            }
            [data-testid="stTextInput"] input::placeholder,
            [data-testid="stTextArea"] textarea::placeholder,
            [data-baseweb="input"] input::placeholder,
            [data-baseweb="textarea"] textarea::placeholder {
                color: var(--deck-text) !important;
                opacity: 0.70 !important;
            }
            [data-baseweb="select"] *,
            [data-baseweb="popover"] *,
            [data-baseweb="menu"] *,
            [data-baseweb="option"] * {
                color: var(--deck-text) !important;
            }
            [data-baseweb="popover"],
            [data-baseweb="menu"],
            [data-baseweb="option"] {
                background: var(--deck-surface) !important;
                color: var(--deck-text) !important;
            }
            [data-baseweb="option"]:hover,
            [data-baseweb="option"][aria-selected="true"] {
                background: var(--deck-bg) !important;
                color: var(--deck-text) !important;
            }

            /* Buttons: use text/background pair for guaranteed contrast in any theme */
            .stButton > button,
            .stDownloadButton > button,
            button[data-testid^="baseButton"] {
                border-radius: 12px !important;
                border: 1px solid var(--deck-border-strong) !important;
                background: var(--deck-surface) !important;
                color: var(--deck-text) !important;
                font-weight: 800 !important;
                min-height: 44px;
            }
            .stButton > button[kind="primary"],
            .stButton > button[data-testid="baseButton-primary"],
            .stDownloadButton > button {
                background: var(--deck-text) !important;
                color: var(--deck-bg) !important;
                border-color: var(--deck-text) !important;
            }
            .stButton > button[kind="primary"] *,
            .stButton > button[data-testid="baseButton-primary"] *,
            .stDownloadButton > button * {
                color: var(--deck-bg) !important;
            }

            /* Cards, panels, expanders, metrics */
            [data-testid="stMetric"],
            [data-testid="stExpander"],
            .guide-box,
            .insight-card,
            .readable-panel,
            .score-card {
                background: var(--deck-surface) !important;
                border: 1px solid var(--deck-border) !important;
                border-radius: var(--deck-radius) !important;
                color: var(--deck-text) !important;
                box-shadow: var(--deck-shadow) !important;
            }
            [data-testid="stMetric"] {
                padding: .9rem 1rem !important;
            }
            .guide-box,
            .insight-card,
            .readable-panel,
            .score-card {
                padding: 15px 17px;
                margin: 8px 0 18px 0;
            }
            .guide-box {
                border-left: 4px solid var(--deck-primary) !important;
            }
            .guide-box strong,
            .guide-title,
            .insight-card h4,
            .score-card strong,
            .readable-panel strong {
                color: var(--deck-text) !important;
                font-weight: 850 !important;
            }
            .guide-box p, .guide-box span,
            .insight-card p, .insight-card li,
            .readable-panel p, .readable-panel li,
            .score-card p,
            .score-card li {
                color: var(--deck-text) !important;
                font-size: .94rem;
                line-height: 1.58;
                opacity: 0.94 !important;
            }

            .pill {
                display: inline-block;
                padding: 4px 10px;
                border: 1px solid var(--deck-border-strong) !important;
                border-radius: 999px;
                margin: 3px;
                background: var(--deck-surface) !important;
                color: var(--deck-text) !important;
            }
            .risk {border-left: 4px solid #ef4444 !important;}
            .ok {border-left: 4px solid #22c55e !important;}
            .warn {border-left: 4px solid #f59e0b !important;}

            /* Alerts and progress labels */
            [data-testid="stAlert"],
            [data-testid="stAlert"] *,
            [data-testid="stProgress"],
            [data-testid="stProgress"] * {
                color: var(--deck-text) !important;
            }

            /* Footer */
            .developer-footer {
                position: fixed;
                left: 0;
                bottom: 0;
                width: 100%;
                background: var(--deck-surface) !important;
                border-top: 1px solid var(--deck-border) !important;
                color: var(--deck-text) !important;
                text-align: center;
                padding: 8px 0;
                font-size: 12px;
                line-height: 1.4;
                z-index: 9999;
            }

            /* Data editor/table safety */
            [data-testid="stDataFrame"],
            [data-testid="stDataFrame"] *,
            [data-testid="stTable"],
            [data-testid="stTable"] * {
                color: var(--deck-text) !important;
            }
            [data-testid="stDataFrame"] div,
            [data-testid="stDataFrame"] span,
            [data-testid="stDataFrame"] p {
                white-space: normal !important;
                overflow: visible !important;
                text-overflow: clip !important;
                line-height: 1.35 !important;
            }
            .stMarkdown, .stMarkdown p, .stMarkdown li,
            .guide-box, .insight-card, .readable-panel, .score-card {
                overflow-wrap: anywhere !important;
                word-break: normal !important;
                white-space: normal !important;
                text-overflow: clip !important;
            }

            .prompter-frame-note {
                padding: 12px 14px;
                border: 1px solid var(--deck-border);
                border-radius: 14px;
                background: var(--deck-surface);
                color: var(--deck-text);
                margin-bottom: 12px;
                line-height: 1.55;
            }


            /* v10.2 Safe Contrast Guard
               Prevents black text on black backgrounds when Streamlit theme/browser theme changes.
               The editor UI intentionally uses light high-contrast surfaces; generated PPT/PDF keep selected brand colors. */
            html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stSidebar"], .block-container {
                color-scheme: light !important;
            }
            .stApp,
            [data-testid="stAppViewContainer"],
            [data-testid="stMain"],
            .main,
            .block-container {
                background: #f8fafc !important;
                color: #0f172a !important;
            }
            [data-testid="stSidebar"] > div:first-child,
            section[data-testid="stSidebar"] {
                background: #ffffff !important;
                color: #0f172a !important;
                border-right: 1px solid #cbd5e1 !important;
            }
            .stApp :where(h1,h2,h3,h4,h5,h6,p,li,label,span,div,strong,em,small,code,pre),
            .stApp [data-testid="stMarkdownContainer"],
            .stApp [data-testid="stMarkdownContainer"] *,
            .stApp [data-testid="stCaptionContainer"],
            .stApp [data-testid="stWidgetLabel"],
            .stApp [data-testid="stWidgetLabel"] *,
            .stApp [data-testid="InputInstructions"],
            .stApp [data-testid="stHelp"],
            .stApp [data-testid="stExpander"] *,
            .stApp [data-testid="stMetric"] *,
            .stApp [data-testid="stFileUploader"] * {
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                text-shadow: none !important;
            }
            .stApp a, .stApp a *,
            .stApp [data-testid="stMarkdownContainer"] a,
            .stApp [data-testid="stMarkdownContainer"] a * {
                color: #1d4ed8 !important;
                -webkit-text-fill-color: #1d4ed8 !important;
            }
            .stApp [data-testid="stTextInput"] input,
            .stApp [data-testid="stNumberInput"] input,
            .stApp [data-testid="stTextArea"] textarea,
            .stApp [data-baseweb="input"] input,
            .stApp [data-baseweb="textarea"] textarea,
            .stApp [data-baseweb="select"] > div,
            .stApp [data-testid="stFileUploader"] section,
            .stApp [data-testid="stColorPicker"] input {
                background: #ffffff !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                border-color: #94a3b8 !important;
                box-shadow: none !important;
            }
            .stApp input::placeholder,
            .stApp textarea::placeholder,
            .stApp [data-baseweb="input"] input::placeholder,
            .stApp [data-baseweb="textarea"] textarea::placeholder {
                color: #475569 !important;
                -webkit-text-fill-color: #475569 !important;
                opacity: 1 !important;
            }
            [data-baseweb="popover"],
            [data-baseweb="menu"],
            [data-baseweb="option"],
            [role="listbox"],
            [role="option"] {
                background: #ffffff !important;
                color: #0f172a !important;
            }
            [data-baseweb="popover"] *,
            [data-baseweb="menu"] *,
            [data-baseweb="option"] *,
            [role="listbox"] *,
            [role="option"] * {
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
            }
            .stApp [data-testid="stTabs"] button[role="tab"],
            .stApp [data-testid="stTabs"] button[role="tab"] * {
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                background: #e2e8f0 !important;
            }
            .stApp [data-testid="stTabs"] button[role="tab"][aria-selected="true"],
            .stApp [data-testid="stTabs"] button[role="tab"][aria-selected="true"] * {
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                background: #ffffff !important;
                font-weight: 850 !important;
            }
            .stApp [data-testid="stExpander"],
            .stApp [data-testid="stMetric"],
            .guide-box,
            .insight-card,
            .readable-panel,
            .score-card,
            .prompter-frame-note {
                background: #ffffff !important;
                color: #0f172a !important;
                border: 1px solid #cbd5e1 !important;
                box-shadow: 0 10px 24px rgba(15, 23, 42, 0.08) !important;
            }
            .guide-box *,
            .insight-card *,
            .readable-panel *,
            .score-card *,
            .prompter-frame-note * {
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
            }
            .pill,
            .stApp .pill {
                background: #ffffff !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                border: 1px solid #94a3b8 !important;
            }
            .stApp [data-testid="stAlert"],
            .stApp [data-testid="stAlert"] *,
            .stApp [data-testid="stDataFrame"],
            .stApp [data-testid="stDataFrame"] *,
            .stApp [data-testid="stTable"],
            .stApp [data-testid="stTable"] * {
                background-color: #ffffff !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
            }
            .stApp .stButton > button,
            .stApp .stDownloadButton > button,
            .stApp button[data-testid^="baseButton"] {
                background: #0f172a !important;
                color: #ffffff !important;
                -webkit-text-fill-color: #ffffff !important;
                border: 1px solid #0f172a !important;
            }
            .stApp .stButton > button *,
            .stApp .stDownloadButton > button *,
            .stApp button[data-testid^="baseButton"] * {
                color: #ffffff !important;
                -webkit-text-fill-color: #ffffff !important;
            }
            .developer-footer {
                background: #ffffff !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                border-top: 1px solid #cbd5e1 !important;
            }

            /* v10.3 Dropdown / selectbox contrast hardening
               BaseWeb renders opened dropdowns in a portal outside .stApp, so
               these selectors are intentionally global and include listbox,
               option, popover, menu, svg icons, and search input states. */
            div[data-baseweb="select"],
            div[data-baseweb="select"] > div,
            div[data-baseweb="select"] div,
            div[data-baseweb="select"] span,
            div[data-baseweb="select"] input {
                background-color: #ffffff !important;
                background-image: none !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                opacity: 1 !important;
                text-shadow: none !important;
            }
            div[data-baseweb="select"] svg,
            div[data-baseweb="select"] svg *,
            div[data-baseweb="select"] path {
                color: #0f172a !important;
                fill: #0f172a !important;
                stroke: #0f172a !important;
            }
            div[data-baseweb="select"] > div {
                border: 1px solid #94a3b8 !important;
                box-shadow: none !important;
            }
            div[data-baseweb="select"]:hover > div,
            div[data-baseweb="select"]:focus-within > div {
                border-color: #2563eb !important;
                box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.22) !important;
            }
            div[data-baseweb="popover"],
            div[data-baseweb="popover"] > div,
            div[data-baseweb="menu"],
            ul[role="listbox"],
            div[role="listbox"],
            li[role="option"],
            div[role="option"] {
                background: #ffffff !important;
                background-color: #ffffff !important;
                background-image: none !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                border-color: #94a3b8 !important;
                opacity: 1 !important;
                text-shadow: none !important;
            }
            div[data-baseweb="popover"] *,
            div[data-baseweb="menu"] *,
            ul[role="listbox"] *,
            div[role="listbox"] *,
            li[role="option"] *,
            div[role="option"] * {
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                opacity: 1 !important;
                text-shadow: none !important;
            }
            li[role="option"][aria-selected="true"],
            div[role="option"][aria-selected="true"],
            li[role="option"]:hover,
            div[role="option"]:hover,
            li[role="option"][data-highlighted="true"],
            div[role="option"][data-highlighted="true"] {
                background: #dbeafe !important;
                background-color: #dbeafe !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
            }
            li[role="option"][aria-selected="true"] *,
            div[role="option"][aria-selected="true"] *,
            li[role="option"]:hover *,
            div[role="option"]:hover * {
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
            }
            input[aria-autocomplete="list"],
            input[role="combobox"],
            div[role="combobox"],
            [aria-haspopup="listbox"] {
                background-color: #ffffff !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
            }
            .stSelectbox label,
            .stSelectbox label *,
            .stMultiSelect label,
            .stMultiSelect label * {
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
            }


            /* v10.6 Forced Readable Theme
               Single light editor theme for Streamlit UI. Generated PPT/PDF still use selected brand colors.
               This prevents unreadable sections such as Q&A investor and Cara menghitung. */
            :root,
            html,
            body,
            .stApp,
            [data-testid="stAppViewContainer"] {
                color-scheme: light !important;
                --deck-bg: #f8fafc !important;
                --deck-surface: #ffffff !important;
                --deck-text: #0f172a !important;
                --deck-muted: #334155 !important;
                --deck-primary: #2563eb !important;
                --deck-border: #cbd5e1 !important;
                --deck-border-strong: #94a3b8 !important;
            }
            .stApp,
            [data-testid="stAppViewContainer"],
            [data-testid="stMain"],
            .main,
            .block-container {
                background: #f8fafc !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
            }
            .stApp *:not(svg):not(path) {
                text-shadow: none !important;
            }
            .stApp :where(h1,h2,h3,h4,h5,h6,p,li,label,span,div,strong,em,small,blockquote,td,th),
            .stApp [data-testid="stMarkdownContainer"],
            .stApp [data-testid="stMarkdownContainer"] *,
            .stApp [data-testid="stWidgetLabel"],
            .stApp [data-testid="stWidgetLabel"] *,
            .stApp [data-testid="stCaptionContainer"],
            .stApp [data-testid="InputInstructions"],
            .stApp [data-testid="stHelp"] {
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                opacity: 1 !important;
            }
            .stApp code,
            .stApp pre,
            .stApp kbd,
            .stApp [data-testid="stCodeBlock"],
            .stApp [data-testid="stCodeBlock"] * {
                background: #e2e8f0 !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                border-color: #94a3b8 !important;
                text-shadow: none !important;
            }
            .stApp [data-testid="stExpander"],
            .stApp [data-testid="stExpander"] > details,
            .stApp [data-testid="stExpander"] summary,
            .stApp [data-testid="stExpander"] div,
            .stApp [data-testid="stExpander"] p,
            .stApp [data-testid="stExpander"] span,
            .stApp [data-testid="stExpander"] code,
            .stApp [data-testid="stMetric"],
            .stApp [data-testid="stFileUploader"] section,
            .guide-box,
            .insight-card,
            .readable-panel,
            .score-card,
            .prompter-frame-note {
                background: #ffffff !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                border-color: #cbd5e1 !important;
            }
            .stApp [data-testid="stExpander"] *,
            .stApp [data-testid="stMetric"] *,
            .stApp [data-testid="stFileUploader"] *,
            .guide-box *,
            .insight-card *,
            .readable-panel *,
            .score-card *,
            .prompter-frame-note * {
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                opacity: 1 !important;
            }
            .stApp [data-testid="stDataFrame"],
            .stApp [data-testid="stDataFrame"] *,
            .stApp [data-testid="stTable"],
            .stApp [data-testid="stTable"] *,
            .stApp iframe {
                color-scheme: light !important;
            }
            .stApp [data-testid="stTextInput"] input,
            .stApp [data-testid="stNumberInput"] input,
            .stApp [data-testid="stTextArea"] textarea,
            .stApp [data-baseweb="input"] input,
            .stApp [data-baseweb="textarea"] textarea,
            .stApp [data-baseweb="select"] > div,
            .stApp [data-testid="stColorPicker"] input {
                background: #ffffff !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                border-color: #94a3b8 !important;
                caret-color: #2563eb !important;
            }
            .stApp input::placeholder,
            .stApp textarea::placeholder {
                color: #475569 !important;
                -webkit-text-fill-color: #475569 !important;
                opacity: 1 !important;
            }
            [data-baseweb="popover"],
            [data-baseweb="popover"] *,
            [data-baseweb="menu"],
            [data-baseweb="menu"] *,
            [role="listbox"],
            [role="listbox"] *,
            [role="option"],
            [role="option"] * {
                background-color: #ffffff !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
                opacity: 1 !important;
            }
            [role="option"]:hover,
            [role="option"][aria-selected="true"],
            li[role="option"]:hover,
            li[role="option"][aria-selected="true"] {
                background-color: #dbeafe !important;
                color: #0f172a !important;
                -webkit-text-fill-color: #0f172a !important;
            }
            .stApp .stButton > button,
            .stApp .stDownloadButton > button,
            .stApp button[data-testid^="baseButton"] {
                background: #0f172a !important;
                color: #ffffff !important;
                -webkit-text-fill-color: #ffffff !important;
                border-color: #0f172a !important;
            }
            .stApp .stButton > button *,
            .stApp .stDownloadButton > button *,
            .stApp button[data-testid^="baseButton"] * {
                color: #ffffff !important;
                -webkit-text-fill-color: #ffffff !important;
            }

        </style>
        """,
        unsafe_allow_html=True,
    )


hide_streamlit_emblems()


THEME = {
    "bg": RGBColor(248, 250, 252),
    "white": RGBColor(255, 255, 255),
    "ink": RGBColor(15, 23, 42),
    "muted": RGBColor(100, 116, 139),
    "line": RGBColor(226, 232, 240),
    "dark": RGBColor(2, 6, 23),
    "blue_soft": RGBColor(239, 246, 255),
    "green": RGBColor(22, 163, 74),
    "red": RGBColor(220, 38, 38),
    "yellow": RGBColor(245, 158, 11),
}

FONT_HEAD = "Aptos Display"
FONT_BODY = "Aptos"


def rgb(hex_color: str) -> RGBColor:
    value = (hex_color or "#2563EB").replace("#", "").strip()
    if len(value) != 6:
        value = "2563EB"
    try:
        return RGBColor(int(value[:2], 16), int(value[2:4], 16), int(value[4:], 16))
    except ValueError:
        return RGBColor(37, 99, 235)


def set_bg(slide, color: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def clean_lines(text: str, limit: int = 5) -> list[str]:
    result: list[str] = []
    for row in (text or "").splitlines():
        cleaned = row.strip().lstrip("-•*0123456789. ").strip()
        if cleaned:
            result.append(cleaned)
    return result[:limit]


def truncate(text: Any, max_chars: int = 150) -> str:
    """Return full text.

    Earlier versions shortened long copy with an ellipsis. For pitching material,
    long text is more useful when preserved and fitted with smaller typography.
    The max_chars argument is kept for backward compatibility with older calls.
    """
    return str(text or "").strip()


def estimate_fit_font_size(
    text: Any,
    box_w: float,
    box_h: float,
    base_size: float,
    min_size: float = 6.5,
) -> float:
    """Approximate a readable font size that keeps long text inside a PPT box."""
    value = str(text or "").strip()
    if not value:
        return base_size

    explicit_lines = max(1, value.count("\n") + 1)
    chars = len(value)
    area = max(box_w * box_h, 0.15)

    # Capacity estimate at base size. Larger boxes and smaller fonts hold more text.
    capacity = max(18, area * 58 * (18 / max(base_size, 1)))
    line_pressure = explicit_lines / max(box_h * 3.3, 1)
    char_pressure = chars / capacity
    pressure = max(char_pressure, line_pressure)

    adjusted = base_size
    if pressure > 1:
        adjusted = base_size / (pressure ** 0.50)

    # Additional reduction for very long single-line values such as milestone text.
    longest_line = max((len(row) for row in value.splitlines()), default=0)
    single_line_capacity = max(8, box_w * 9.5)
    if longest_line > single_line_capacity:
        adjusted = min(adjusted, base_size * (single_line_capacity / longest_line) ** 0.35)

    return max(min_size, min(base_size, adjusted))


def table_font_size(rows: list[list[Any]], base_size: float = 8.5, min_size: float = 6.2) -> float:
    longest = 0
    total = 0
    cells = 0
    for row in rows:
        for cell in row:
            text = str(cell or "")
            longest = max(longest, len(text))
            total += len(text)
            cells += 1
    avg = total / max(cells, 1)
    size = base_size
    if longest > 90 or avg > 55:
        size -= 1.2
    if longest > 140 or avg > 80:
        size -= 1.0
    if len(rows) >= 5:
        size -= 0.7
    return max(min_size, size)


def filename(value: str) -> str:
    value = re.sub(r"[^a-zA-Z0-9 _-]", "", value or "pitch-deck")
    return value.strip().lower().replace(" ", "-") or "pitch-deck"


def money(value: int | float, currency: str) -> str:
    if currency == "Rp":
        return f"Rp {float(value):,.0f}".replace(",", ".")
    return f"{currency} {float(value):,.0f}"


def short_money(value: float, currency: str) -> str:
    sign = "-" if value < 0 else ""
    value = abs(float(value or 0))
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


def parse_percent(value: str | float | int | None) -> float | None:
    if value is None:
        return None
    match = re.search(r"-?\d+(?:[\.,]\d+)?", str(value))
    if not match:
        return None
    return float(match.group(0).replace(",", ".")) / 100


def safe_div(num: float, den: float) -> float | None:
    return None if not den else num / den


def pct(value: float | None) -> str:
    return "N/A" if value is None else f"{value * 100:.1f}%"


def multiple(value: float | None) -> str:
    return "N/A" if value is None else f"{value:.1f}x"


def to_jsonable(data: dict[str, Any]) -> dict[str, Any]:
    result = {}
    for key, value in data.items():
        if key == "logo_bytes" or isinstance(value, (BytesIO, bytes, bytearray)):
            continue
        try:
            json.dumps(value)
            result[key] = value
        except TypeError:
            result[key] = str(value)
    return result


LANGUAGE_OPTIONS = ["Bahasa Indonesia", "English", "Bilingual"]
PITCH_TYPE_OPTIONS = [
    "Investor Seed Round",
    "Demo Day",
    "Pitch Competition",
    "Corporate Partnership",
    "Grant / Hibah",
    "Incubator / Accelerator",
    "Internal Business Proposal",
]

PITCH_TYPE_GUIDES = {
    "Investor Seed Round": {
        "focus": "venture-scale opportunity, traction, growth engine, defensibility, funding-to-milestone logic",
        "tone": "data-first, concise, confident",
        "slides": "Problem, Solution, Market, Traction, Business Model, Competition, Team, Financials, Ask",
    },
    "Demo Day": {
        "focus": "clear story, strong hook, traction signal, memorable ask",
        "tone": "fast, high-energy, simple",
        "slides": "Problem, Solution, Product, Traction, Market, Team, Ask",
    },
    "Pitch Competition": {
        "focus": "clarity, impact, differentiation, feasibility, presentation delivery",
        "tone": "persuasive, accessible, story-led",
        "slides": "Problem, Solution, Impact, Market, Product, Traction, Team, Ask",
    },
    "Corporate Partnership": {
        "focus": "strategic fit, pilot design, business value, operational feasibility",
        "tone": "professional, risk-aware, implementation-focused",
        "slides": "Problem, Solution, Product, Use Case, Partnership Value, Pilot Plan, Team, Next Step",
    },
    "Grant / Hibah": {
        "focus": "impact, beneficiary, implementation plan, budget accountability, measurement",
        "tone": "credible, impact-oriented, transparent",
        "slides": "Problem, Impact, Solution, Beneficiary, Implementation, Budget, Milestones, Team",
    },
    "Incubator / Accelerator": {
        "focus": "founder-market fit, learning velocity, opportunity, coachability, milestone plan",
        "tone": "clear, reflective, growth-oriented",
        "slides": "Problem, Solution, Product, Market, Traction, Team, Milestones, Support Needed",
    },
    "Internal Business Proposal": {
        "focus": "business case, risk, cost, ROI, execution plan, decision needed",
        "tone": "structured, analytical, practical",
        "slides": "Context, Problem, Proposal, Business Case, Financials, Risks, Timeline, Decision",
    },
}

PITCH_DURATION_OPTIONS = [3, 5, 8, 10, 15, 20]
PITCH_DURATION_PROFILES = {
    3: {
        "name": "3-minute elevator pitch",
        "scheme": "7 slide ringkas. Konteks tetap lengkap, tetapi product, market, business model, dan traction dipadatkan.",
        "slide_goal": "20-35 detik per slide. Satu pesan utama per slide.",
        "max_competitors": 3,
        "max_milestones": 2,
    },
    5: {
        "name": "5-minute demo day pitch",
        "scheme": "8 slide kompak. Semua konteks investor muncul, tetapi financial dan ask digabung.",
        "slide_goal": "30-45 detik per slide. Jangan membaca semua bullet.",
        "max_competitors": 4,
        "max_milestones": 3,
    },
    8: {
        "name": "8-minute seed pitch",
        "scheme": "10 slide seed ringkas. Cukup detail untuk investor meeting awal.",
        "slide_goal": "35-55 detik per slide. Tekankan bukti dan asumsi utama.",
        "max_competitors": 5,
        "max_milestones": 4,
    },
    10: {
        "name": "10-minute seed standard",
        "scheme": "11-12 slide standar seed. Semua konteks utama muncul dengan kompetisi dan milestone terpisah.",
        "slide_goal": "45-65 detik per slide. Sisakan waktu Q&A.",
        "max_competitors": 5,
        "max_milestones": 4,
    },
    15: {
        "name": "15-minute detailed seed meeting",
        "scheme": "Full deck. Competition dan milestones bisa otomatis dipaginasi bila data banyak.",
        "slide_goal": "50-80 detik per slide. Cocok untuk first investor meeting yang butuh pendalaman.",
        "max_competitors": 10,
        "max_milestones": 8,
    },
    20: {
        "name": "20-minute deep-dive seed meeting",
        "scheme": "Full deck dengan ruang narasi lebih longgar. PDF memberi timing lebih detail untuk Q&A dan asumsi.",
        "slide_goal": "60-100 detik per slide. Gunakan waktu tambahan untuk asumsi, competition, dan milestone.",
        "max_competitors": 10,
        "max_milestones": 8,
    },
}

BUSINESS_MODEL_TEMPLATES = {
    "SaaS / Subscription": {
        "description": "Customer membayar biaya berlangganan bulanan/tahunan.",
        "best_for": "B2B/B2C software, workflow tool, vertical SaaS, AI tool.",
        "pitch_focus": "MRR/ARR, retention, churn, ARPU, CAC payback, expansion revenue.",
        "formula": "MRR = pelanggan berbayar x harga bulanan rata-rata; ARR = MRR x 12.",
        "metrics": ["MRR / ARR", "ARPU", "Gross Margin", "CAC Payback"],
        "defaults": ["Rp 85 juta MRR", "Rp 99.000/bulan", "78%", "< 2 bulan"],
        "lines": "Subscription SaaS bulanan.\nPaket Pro untuk multi-outlet.\nAdd-on premium untuk automation dan analytics.\nExpansion revenue melalui seat, outlet, usage, atau modul tambahan.",
    },
    "Marketplace / Take Rate": {
        "description": "Platform mempertemukan supply dan demand lalu mengambil komisi transaksi.",
        "best_for": "Marketplace jasa, produk, B2B procurement, booking, merchant platform.",
        "pitch_focus": "GMV, take rate, liquidity, repeat transaction, unit economics transaksi.",
        "formula": "Revenue = GMV x take rate; GMV = jumlah transaksi x nilai transaksi rata-rata.",
        "metrics": ["GMV", "Take Rate", "Repeat Rate", "Liquidity"],
        "defaults": ["Rp 1,5M/bulan", "5%", "42%", "70% matched"],
        "lines": "Marketplace mengambil komisi dari transaksi.\nSupply diakuisisi melalui partner dan komunitas.\nDemand datang dari channel digital dan referral.\nRevenue naik melalui GMV, take rate, dan repeat transaction.",
    },
    "E-commerce / D2C": {
        "description": "Menjual produk langsung ke customer dengan margin produk/logistik.",
        "best_for": "Brand consumer, D2C, retail online, produk fisik.",
        "pitch_focus": "AOV, gross margin, repeat purchase, CAC, payback, inventory turn.",
        "formula": "Revenue = jumlah order x AOV; gross margin = (revenue - COGS) / revenue.",
        "metrics": ["AOV", "Gross Margin", "Repeat Purchase", "CAC Payback"],
        "defaults": ["Rp 180.000", "45%", "35%", "< 3 bulan"],
        "lines": "Penjualan langsung ke customer melalui website, marketplace, dan komunitas.\nMargin berasal dari produk inti dan bundle.\nGrowth ditopang repeat purchase dan channel partnership.",
    },
    "Transaction Fee / Fintech": {
        "description": "Revenue dari fee pembayaran, pembiayaan, transfer, insurance, atau transaksi finansial.",
        "best_for": "Fintech, payment, lending enablement, insurtech, embedded finance.",
        "pitch_focus": "Transaction volume, fee rate, risk cost/default, compliance, active merchants.",
        "formula": "Revenue = transaction volume x fee rate; untuk lending tambahkan default/risk cost.",
        "metrics": ["Transaction Volume", "Fee Rate", "Risk Cost", "Active Merchants"],
        "defaults": ["Rp 3M/bulan", "1,5%", "< 2%", "1.200"],
        "lines": "Revenue berasal dari fee transaksi dan revenue share partner finansial.\nProduk tertanam di workflow customer.\nRisk, compliance, dan distribution partner menjadi kunci scale.",
    },
    "Usage-Based / API": {
        "description": "Customer membayar sesuai pemakaian: API call, token, storage, transaksi, atau volume data.",
        "best_for": "Developer tool, AI API, infrastructure, data platform, cloud service.",
        "pitch_focus": "Usage growth, revenue/unit, gross margin per usage, NRR, developer adoption.",
        "formula": "Revenue = volume usage x harga per unit.",
        "metrics": ["Usage Volume", "Revenue / Unit", "Gross Margin", "NRR"],
        "defaults": ["2 juta API calls/bulan", "Rp 12/call", "72%", "115%"],
        "lines": "Customer membayar berdasarkan pemakaian.\nEntry point murah, revenue naik seiring volume.\nExpansion terjadi ketika produk tertanam dalam workflow customer.",
    },
    "Freemium": {
        "description": "Sebagian user memakai gratis lalu sebagian dikonversi ke paket berbayar.",
        "best_for": "Consumer app, productivity, creator tool, education, AI tool.",
        "pitch_focus": "Active users, activation, conversion rate, paid retention, ARPU, viral/referral loop.",
        "formula": "Paid users = free users x conversion rate; MRR = paid users x ARPU.",
        "metrics": ["Active Users", "Free-to-Paid", "ARPU", "Retention"],
        "defaults": ["25.000 MAU", "4%", "Rp 49.000", "68%"],
        "lines": "Produk gratis dipakai sebagai acquisition engine.\nMonetisasi melalui premium features dan team plan.\nKunci scale adalah activation, conversion, dan retention.",
    },
    "Enterprise / Licensing": {
        "description": "Perusahaan membayar kontrak lisensi, implementation fee, atau annual contract.",
        "best_for": "B2B enterprise, govtech, cybersecurity, HR/ERP, vertical software.",
        "pitch_focus": "ACV, pipeline, sales cycle, pilot conversion, gross retention, implementation cost.",
        "formula": "ARR = jumlah customer enterprise x ACV.",
        "metrics": ["ACV", "Pipeline", "Sales Cycle", "Pilot Conversion"],
        "defaults": ["Rp 350 juta", "Rp 4,2M", "3-6 bulan", "35%"],
        "lines": "Revenue berasal dari annual license dan implementation fee.\nMasuk melalui pilot berbayar lalu expand ke multi-department.\nCustomer success dan compliance menjadi kunci renewal.",
    },
    "Service-Enabled Software": {
        "description": "Software dikombinasikan dengan layanan untuk memberi hasil lebih cepat.",
        "best_for": "AI service, managed platform, ops automation, agency-to-SaaS transition.",
        "pitch_focus": "Service margin, automation rate, gross margin expansion, delivery speed, repeatability.",
        "formula": "Gross margin naik jika porsi automation meningkat dan jam manual/customer turun.",
        "metrics": ["Automation Rate", "Service Margin", "Delivery Time", "Repeat Revenue"],
        "defaults": ["60%", "42%", "3 hari", "70%"],
        "lines": "Software mempercepat delivery layanan.\nDi awal ada human-in-the-loop untuk quality.\nMargin membaik saat automation meningkat dan playbook menjadi repeatable.",
    },
    "Advertising / Media": {
        "description": "Revenue dari iklan, sponsorship, affiliate, atau monetisasi audience.",
        "best_for": "Media, creator platform, community, content network, consumer app.",
        "pitch_focus": "MAU/DAU, engagement, ad inventory, CPM, fill rate, audience quality.",
        "formula": "Revenue = impressions / 1000 x CPM x fill rate.",
        "metrics": ["MAU", "Engagement", "CPM", "Fill Rate"],
        "defaults": ["120.000 MAU", "18 menit/sesi", "Rp 45.000", "65%"],
        "lines": "Audience dibangun melalui content dan community.\nRevenue berasal dari sponsorship, ads, dan affiliate.\nKunci bisnis adalah engagement, inventory, dan kualitas segmen audience.",
    },
    "Hybrid / Other": {
        "description": "Gabungan beberapa revenue stream atau model yang masih bereksperimen.",
        "best_for": "Startup awal yang menguji pricing, multi-sided platform, atau bisnis kompleks.",
        "pitch_focus": "Jelaskan revenue stream utama, mana yang terbukti, dan mana yang masih eksperimen.",
        "formula": "Total revenue = revenue stream 1 + revenue stream 2 + revenue stream 3.",
        "metrics": ["Main Revenue", "Secondary Revenue", "Gross Margin", "Payback"],
        "defaults": ["Subscription", "Transaction fee", "65%", "< 4 bulan"],
        "lines": "Revenue utama berasal dari produk inti.\nRevenue tambahan berasal dari layanan, fee transaksi, atau partnership.\nEksperimen pricing diprioritaskan pada stream yang paling repeatable.",
    },
}

GLOSSARY = [
    ("Market", "TAM", "Total Addressable Market: total pasar maksimum jika semua target memakai solusi.", "TAM = jumlah seluruh target customer x potensi belanja tahunan rata-rata", "10 juta UMKM x Rp 1 juta/tahun = Rp 10T"),
    ("Market", "SAM", "Serviceable Available Market: bagian TAM yang realistis dilayani oleh produk saat ini.", "SAM = segmen target yang cocok x belanja tahunan rata-rata", "1 juta UMKM digital x Rp 1 juta = Rp 1T"),
    ("Market", "SOM", "Serviceable Obtainable Market: bagian pasar yang realistis direbut dalam 2-3 tahun.", "SOM = target customer yang bisa diakuisisi x ARPU tahunan", "10.000 customer x Rp 1,2 juta = Rp 12M"),
    ("Revenue", "MRR", "Monthly Recurring Revenue: revenue berulang per bulan.", "MRR = customer berbayar x harga bulanan rata-rata", "1.000 customer x Rp 99.000 = Rp 99 juta"),
    ("Revenue", "ARR", "Annual Recurring Revenue: MRR disetahunkan.", "ARR = MRR x 12", "Rp 99 juta x 12 = Rp 1,188M"),
    ("Revenue", "GMV", "Gross Merchandise Value: total nilai transaksi yang lewat platform.", "GMV = jumlah transaksi x nilai transaksi rata-rata", "10.000 transaksi x Rp 150.000 = Rp 1,5M"),
    ("Revenue", "Take Rate", "Persentase GMV yang menjadi revenue platform.", "Take Rate = revenue platform / GMV x 100%", "Rp 75 juta / Rp 1,5M = 5%"),
    ("Unit Economics", "ARPU", "Average Revenue Per User: rata-rata revenue per user/customer.", "ARPU = total revenue / jumlah customer", "Rp 99 juta / 1.000 = Rp 99.000"),
    ("Unit Economics", "CAC", "Customer Acquisition Cost: biaya untuk mendapatkan customer baru.", "CAC = biaya sales & marketing / customer baru", "Rp 20 juta / 100 customer = Rp 200.000"),
    ("Unit Economics", "CAC Payback", "Waktu untuk menutup biaya akuisisi dari margin customer.", "CAC Payback = CAC / gross profit bulanan per customer", "Rp 200.000 / Rp 80.000 = 2,5 bulan"),
    ("Unit Economics", "Gross Margin", "Persentase revenue yang tersisa setelah biaya langsung/COGS.", "Gross Margin = (Revenue - COGS) / Revenue x 100%", "(100 juta - 25 juta) / 100 juta = 75%"),
    ("Retention", "Retention", "Persentase customer/user yang tetap aktif setelah periode tertentu.", "Retention = user yang tetap aktif / user awal x 100%", "700 aktif dari 1.000 = 70%"),
    ("Retention", "Churn", "Persentase customer yang berhenti memakai/membayar.", "Churn = customer hilang / customer awal x 100%", "50 churn dari 1.000 = 5%"),
    ("Financial", "Burn Rate", "Uang bersih yang habis setiap bulan untuk operasional.", "Burn Rate = biaya bulanan - revenue bulanan", "Biaya 200 juta - revenue 80 juta = 120 juta"),
    ("Financial", "Runway", "Berapa bulan startup dapat bertahan dengan kas saat ini.", "Runway = cash tersedia / burn rate bulanan", "Rp 1,8M / Rp 100 juta = 18 bulan"),
    ("Fundraising", "Valuation", "Nilai perusahaan yang menjadi dasar negosiasi saham.", "Post-money = pre-money + investasi", "Pre-money Rp 20M + investasi Rp 5M = post-money Rp 25M"),
    ("Fundraising", "Dilution", "Persentase kepemilikan founder yang berkurang karena investor masuk.", "Dilution = investasi / post-money valuation", "Rp 5M / Rp 25M = 20%"),
    ("Execution", "Milestone", "Target pembuktian yang harus dicapai dengan dana yang diperoleh.", "Milestone = target waktu + outcome + success metric", "Dalam 12 bulan mencapai Rp 500 juta MRR"),
    ("Execution", "GTM", "Go-To-Market: strategi mendapatkan, mengonversi, dan mempertahankan customer.", "Tidak ada rumus tunggal. Ukur funnel: lead -> trial -> paid -> retained.", "1.000 leads -> 200 demo -> 50 paid"),
    ("Competition", "Moat", "Alasan bisnis sulit ditiru atau dikalahkan saat tumbuh.", "Moat bisa berasal dari data, network effect, switching cost, distribution, brand, atau regulasi.", "Data transaksi historis membuat rekomendasi makin akurat."),
]


def business_model_template(model_type: str | dict[str, Any]) -> dict[str, Any]:
    if isinstance(model_type, dict):
        model_type = model_type.get("business_model_type", "SaaS / Subscription")
    return BUSINESS_MODEL_TEMPLATES.get(str(model_type), BUSINESS_MODEL_TEMPLATES["SaaS / Subscription"])


def guide(title: str, body: str) -> None:
    st.markdown(
        f"""
        <div class="guide-box">
            <strong>{html.escape(title)}</strong><br>
            <span>{body}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def footer_ui() -> None:
    st.markdown(f'<div class="developer-footer">{DEVELOPER} • {APP_VERSION}</div>', unsafe_allow_html=True)


def get_uploaded_logo_bytes() -> bytes | None:
    file = st.session_state.get("logo_file")
    if file is None:
        return None
    try:
        return file.getvalue()
    except Exception:
        return None


def initialize_defaults() -> None:
    defaults = {
        "company": "Rex AI",
        "one_liner": "AI copilot untuk membantu UMKM membuat laporan keuangan otomatis.",
        "presenter": "Rex Founder",
        "contact": "founder@rex.ai",
        "round_name": "Pre-Seed Round",
        "ask": 750_000_000,
        "currency": "Rp",
        "pitch_duration_minutes": 10,
        "pitch_type": "Investor Seed Round",
        "output_language": "Bahasa Indonesia",
        "business_model_type": "SaaS / Subscription",
        "accent_color": "#2563EB",
        "secondary_color": "#0F172A",
        "deck_style": "Minimal VC",
        "font_style": "Modern Sans",
        "include_insight_slide": True,
        "ui_mode": "Tabs",
        "problem": "UMKM sulit mencatat transaksi harian.\nLaporan keuangan masih manual dan rawan salah.\nPemilik usaha tidak punya insight real-time.",
        "problem_evidence": "Interview awal: 24 dari 30 UMKM masih memakai catatan manual atau spreadsheet.",
        "solution": "Input transaksi via chat.\nOCR struk otomatis.\nDashboard cashflow dan laba rugi.\nRekomendasi aksi berbasis data.",
        "value_prop": "Menghemat waktu pencatatan dan membantu owner memahami kondisi keuangan tanpa akuntan penuh waktu.",
        "product_flow": "User kirim transaksi/foto struk -> sistem klasifikasi otomatis -> dashboard dan laporan dibuat real-time.",
        "product_benefit": "Owner melihat omzet, margin, utang-piutang, dan cashflow dalam satu tempat.",
        "features": "OCR struk\nAuto-categorization\nCashflow dashboard\nExport laporan PDF/Excel\nReminder piutang",
        "tam": "Rp 120T",
        "sam": "Rp 18T",
        "som": "Rp 450M",
        "market_notes": "Target awal: UMKM F&B dan ritel.\nWedge: bisnis yang sudah memakai WhatsApp untuk operasional.\nEkspansi: inventory, payroll, dan embedded financing.",
        "arpu": "Rp 99.000/bulan",
        "gross_margin": "78%",
        "cac": "Rp 140.000",
        "payback": "< 2 bulan",
        "business_model": BUSINESS_MODEL_TEMPLATES["SaaS / Subscription"]["lines"],
        "users": "1.200 users",
        "revenue": "Rp 85 juta MRR",
        "growth": "+28% MoM",
        "retention": "72% D30",
        "traction_notes": "Pilot berbayar dengan 35 UMKM.\nPipeline 180 UMKM dari partner.\nChurn turun setelah fitur reminder piutang dirilis.",
        "icp": "Pemilik UMKM F&B/ritel dengan 3-50 transaksi per hari.",
        "channel": "Komunitas UMKM, reseller akuntansi, konten edukasi, partnership POS.",
        "gtm": "Akuisisi lewat komunitas UMKM.\nReferral program.\nPartnership konsultan pajak dan POS.\nKonten edukasi cashflow sebagai lead magnet.",
        "rev1": 1_200_000_000,
        "rev2": 4_800_000_000,
        "rev3": 13_500_000_000,
        "cost1": 2_000_000_000,
        "cost2": 5_200_000_000,
        "cost3": 10_000_000_000,
        "profit1": -800_000_000,
        "profit2": -400_000_000,
        "profit3": 3_500_000_000,
        "runway": 18,
        "milestone": "Rp 500 juta MRR, 8.000 active businesses, gross margin >80%",
        "use_of_funds": "Product & engineering: 40%\nSales & marketing: 35%\nCustomer success: 15%\nOperations & legal: 10%",
        "next_round": "Raise Seed/Series A setelah akuisisi repeatable dan retention terbukti.",
        "team": "Rex - CEO - 6 tahun membangun SaaS UMKM.\nNadia - CTO - ex fintech data engineer.\nBima - Growth - pernah scale komunitas UMKM 50k members.",
        "founder_fit": "Tim pernah membangun tools operasional untuk UMKM dan punya akses langsung ke komunitas target.",
        "closing": "Let’s help millions of small businesses understand their money.",
        "competition_summary": "Kami menang di simplicity, distribusi lokal, dan data workflow UMKM.",
        "competitor_count": 4,
        "milestone_count": 4,
        "auto_calc_metrics": True,
        "paid_customers": 1000,
        "monthly_price": 99_000,
        "monthly_revenue_calc": 99_000_000,
        "marketing_spend": 20_000_000,
        "new_customers": 100,
        "cogs": 25_000_000,
        "cash_available": 1_800_000_000,
        "monthly_cost": 200_000_000,
        "gmv_calc": 1_500_000_000,
        "platform_revenue_calc": 75_000_000,
        "average_order_value": 150_000,
        "orders_count": 10_000,
        "use_custom_prompter": True,
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)
    template = business_model_template(st.session_state["business_model_type"])
    for idx, label in enumerate(template["metrics"]):
        st.session_state.setdefault(f"model_metric_label_{idx}", label)
        st.session_state.setdefault(f"model_metric_value_{idx}", template["defaults"][idx])
    competitor_defaults = [
        ("Aplikasi Akuntansi A", "Direct competitor", "Terlalu kompleks untuk mikro-UMKM", "Workflow chat-first dan onboarding cepat"),
        ("Spreadsheet/manual", "Indirect competitor", "Tidak real-time dan rawan error", "Otomatisasi laporan dan insight"),
        ("Status quo", "Status quo", "Owner tidak punya data untuk keputusan cepat", "Data harian langsung menjadi rekomendasi aksi"),
        ("POS tanpa analitik", "Adjacent tool", "Mencatat transaksi tetapi tidak memberi insight cashflow", "Mengubah transaksi menjadi laporan dan rekomendasi bisnis"),
    ]
    for i, (name, category, weakness, advantage) in enumerate(competitor_defaults):
        st.session_state.setdefault(f"competitor_name_{i}", name)
        st.session_state.setdefault(f"competitor_category_{i}", category)
        st.session_state.setdefault(f"competitor_weakness_{i}", weakness)
        st.session_state.setdefault(f"competitor_advantage_{i}", advantage)
    milestone_defaults = [
        ("0-3 bulan", "Launch paid beta", "35 customer berbayar", "Product/Growth"),
        ("4-6 bulan", "Repeatable acquisition", "CAC payback < 3 bulan", "Growth"),
        ("7-12 bulan", "Scale revenue", "Rp 500 juta MRR", "CEO"),
        ("12-18 bulan", "Next round readiness", "8.000 active businesses", "CEO/Finance"),
    ]
    for i, (period, target, metric, owner) in enumerate(milestone_defaults):
        st.session_state.setdefault(f"milestone_period_{i}", period)
        st.session_state.setdefault(f"milestone_target_{i}", target)
        st.session_state.setdefault(f"milestone_metric_{i}", metric)
        st.session_state.setdefault(f"milestone_owner_{i}", owner)


def load_project_from_json(file) -> None:
    try:
        payload = json.loads(file.getvalue().decode("utf-8"))
    except Exception as exc:
        st.error(f"File JSON tidak bisa dibaca: {exc}")
        return
    if not isinstance(payload, dict):
        st.error("Format JSON tidak valid. Gunakan file project yang diekspor dari aplikasi ini.")
        return
    payload = payload.get("project", payload)
    for key, value in payload.items():
        if key == "competitors" and isinstance(value, list):
            st.session_state["competitor_count"] = min(max(len(value), 1), 10)
            for idx, item in enumerate(value[:10]):
                st.session_state[f"competitor_name_{idx}"] = item.get("name", "")
                st.session_state[f"competitor_category_{idx}"] = item.get("category", "Alternative")
                st.session_state[f"competitor_weakness_{idx}"] = item.get("weakness", "")
                st.session_state[f"competitor_advantage_{idx}"] = item.get("advantage", "")
        elif key == "milestones" and isinstance(value, list):
            st.session_state["milestone_count"] = min(max(len(value), 1), 8)
            for idx, item in enumerate(value[:8]):
                st.session_state[f"milestone_period_{idx}"] = item.get("period", "")
                st.session_state[f"milestone_target_{idx}"] = item.get("target", "")
                st.session_state[f"milestone_metric_{idx}"] = item.get("metric", "")
                st.session_state[f"milestone_owner_{idx}"] = item.get("owner", "")
        elif key == "custom_prompter_scripts" and isinstance(value, list):
            st.session_state["prompter_slide_count"] = len(value)
            for idx, item in enumerate(value[:30]):
                st.session_state[f"prompter_script_{idx}"] = str(item or "")
        elif key == "use_custom_prompter":
            st.session_state["use_custom_prompter"] = bool(value)
        elif key == "model_metric_labels" and isinstance(value, list):
            for idx, item in enumerate(value[:4]):
                st.session_state[f"model_metric_label_{idx}"] = item
        elif key == "model_metric_values" and isinstance(value, list):
            for idx, item in enumerate(value[:4]):
                st.session_state[f"model_metric_value_{idx}"] = item
        elif key in st.session_state or key in {"company", "one_liner", "business_model_type"}:
            st.session_state[key] = value
    st.success("Project berhasil dimuat. Form akan mengikuti data dari JSON.")
    st.rerun()


def build_project_json(data: dict[str, Any]) -> BytesIO:
    payload = {
        "app": "Seed Investor Pitch Deck Generator",
        "version": APP_VERSION,
        "exported_at": datetime.utcnow().isoformat() + "Z",
        "project": to_jsonable(data),
    }
    out = BytesIO(json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"))
    out.seek(0)
    return out


def calculate_auto_metrics(data: dict[str, Any]) -> dict[str, Any]:
    currency = data.get("currency", "Rp")
    paid = float(data.get("paid_customers", 0) or 0)
    price = float(data.get("monthly_price", 0) or 0)
    monthly_revenue = float(data.get("monthly_revenue_calc", paid * price) or 0)
    marketing_spend = float(data.get("marketing_spend", 0) or 0)
    new_customers = float(data.get("new_customers", 0) or 0)
    cogs = float(data.get("cogs", 0) or 0)
    cash_available = float(data.get("cash_available", 0) or 0)
    monthly_cost = float(data.get("monthly_cost", 0) or 0)
    gmv = float(data.get("gmv_calc", 0) or 0)
    platform_revenue = float(data.get("platform_revenue_calc", 0) or 0)
    orders = float(data.get("orders_count", 0) or 0)
    aov = float(data.get("average_order_value", 0) or 0)
    mrr = paid * price if paid and price else monthly_revenue
    arr = mrr * 12
    arpu = safe_div(monthly_revenue, paid)
    cac = safe_div(marketing_spend, new_customers)
    gross_margin = safe_div(monthly_revenue - cogs, monthly_revenue)
    monthly_burn = max(monthly_cost - monthly_revenue, 0)
    runway = safe_div(cash_available, monthly_burn) if monthly_burn else None
    take_rate = safe_div(platform_revenue, gmv)
    calculated_gmv = orders * aov if orders and aov else gmv
    return {
        "MRR": short_money(mrr, currency),
        "ARR": short_money(arr, currency),
        "ARPU": short_money(arpu or 0, currency),
        "CAC": short_money(cac or 0, currency),
        "Gross Margin": pct(gross_margin),
        "Burn Rate": short_money(monthly_burn, currency),
        "Runway": f"{runway:.1f} bulan" if runway else "N/A",
        "Take Rate": pct(take_rate),
        "GMV": short_money(calculated_gmv, currency),
    }


def get_model_metrics(data: dict[str, Any]) -> list[tuple[str, str]]:
    labels = data.get("model_metric_labels", []) or []
    values = data.get("model_metric_values", []) or []
    metrics = []
    for label, value in zip(labels, values):
        if str(label).strip() or str(value).strip():
            metrics.append((str(label), str(value)))
    if len(metrics) >= 3:
        return metrics[:4]
    template = business_model_template(data)
    return list(zip(template["metrics"], template["defaults"]))


def build_slide_plan(data: dict[str, Any]) -> list[dict[str, Any]]:
    duration = int(data.get("pitch_duration_minutes", 10))
    pitch_type = data.get("pitch_type", "Investor Seed Round")
    include_insight = bool(data.get("include_insight_slide", True))
    plan: list[dict[str, Any]] = [
        {"key": "cover", "title": "Opening", "purpose": "Buka dengan one-liner, kategori bisnis, round, dan alasan investor harus mendengar lebih lanjut."},
    ]
    if duration <= 3:
        plan += [
            {"key": "problem_solution", "title": "Problem & Solution", "purpose": "Gabungkan pain point dan solusi agar konteks cepat tertangkap."},
            {"key": "product", "title": "Product", "purpose": "Tunjukkan cara kerja produk dan value yang dihasilkan."},
            {"key": "market_model", "title": "Market & Business Model", "purpose": "Ringkas ukuran pasar, wedge, dan revenue engine."},
            {"key": "traction", "title": "Traction", "purpose": "Beri bukti demand awal."},
            {"key": "team_ask", "title": "Team & Ask", "purpose": "Tutup dengan founder-market fit, dana yang diminta, dan milestone."},
        ]
    elif duration <= 5:
        plan += [
            {"key": "problem", "title": "Problem", "purpose": "Buktikan masalah urgent dan mahal."},
            {"key": "solution_product", "title": "Solution & Product", "purpose": "Tunjukkan solusi dan flow produk."},
            {"key": "market", "title": "Market", "purpose": "Tunjukkan peluang pasar dan entry wedge."},
            {"key": "traction_model", "title": "Traction & Business Model", "purpose": "Gabungkan bukti demand dan cara monetisasi."},
            {"key": "competition", "title": "Competition", "purpose": "Jelaskan alternatif dan narrative advantage."},
            {"key": "financial_ask", "title": "Financials & Ask", "purpose": "Hubungkan funding dengan runway dan milestone."},
            {"key": "team", "title": "Team", "purpose": "Tunjukkan founder-market fit."},
        ]
    elif duration <= 10:
        plan += [
            {"key": "problem", "title": "Problem", "purpose": "Buktikan masalah yang urgent."},
            {"key": "solution", "title": "Solution", "purpose": "Jelaskan perubahan sebelum/sesudah produk."},
            {"key": "product", "title": "Product", "purpose": "Tunjukkan cara kerja produk."},
            {"key": "market", "title": "Market", "purpose": "Tunjukkan TAM/SAM/SOM dan wedge."},
            {"key": "business_model", "title": "Business Model", "purpose": "Jelaskan revenue engine sesuai model bisnis."},
            {"key": "traction", "title": "Traction", "purpose": "Bukti demand dan kualitas retention/growth."},
            {"key": "gtm", "title": "Go-To-Market", "purpose": "Jelaskan mesin akuisisi."},
            {"key": "competition", "title": "Competition", "purpose": "Posisi terhadap alternatif."},
            {"key": "financial_ask", "title": "Financials & Ask", "purpose": "Hubungkan angka, runway, dan funding ask."},
            {"key": "team", "title": "Team", "purpose": "Founder-market fit."},
        ]
    else:
        plan += [
            {"key": "problem", "title": "Problem", "purpose": "Buktikan masalah yang urgent."},
            {"key": "solution", "title": "Solution", "purpose": "Jelaskan solusi."},
            {"key": "product", "title": "Product", "purpose": "Demo flow produk."},
            {"key": "market", "title": "Market Opportunity", "purpose": "Pasar dan wedge."},
            {"key": "business_model", "title": "Business Model", "purpose": "Revenue engine."},
            {"key": "traction", "title": "Traction", "purpose": "Demand signal."},
            {"key": "gtm", "title": "Go-To-Market", "purpose": "Distribution engine."},
            {"key": "competition", "title": "Competition", "purpose": "Alternatif dan narrative advantage."},
            {"key": "milestones", "title": "Milestones", "purpose": "Target pembuktian per periode."},
            {"key": "financials", "title": "Financials", "purpose": "Proyeksi dan asumsi."},
            {"key": "fundraising", "title": "Fundraising Ask", "purpose": "Dana, use of funds, runway, next round."},
            {"key": "team", "title": "Team", "purpose": "Founder-market fit."},
        ]
    if include_insight and duration >= 8:
        plan.append({"key": "readiness", "title": "Investor Readiness", "purpose": "Ringkas kekuatan, risiko, dan fokus diskusi investor."})
    plan.append({"key": "closing", "title": "Closing", "purpose": "Akhiri dengan visi dan next step."})
    # Pitch-type adjustments without losing core context.
    if pitch_type in {"Grant / Hibah", "Pitch Competition"}:
        for item in plan:
            if item["key"] == "market":
                item["title"] = "Market & Impact"
                item["purpose"] += " Tambahkan dampak bagi beneficiary/customer."
    if pitch_type == "Corporate Partnership":
        for item in plan:
            if item["key"] == "gtm":
                item["title"] = "Pilot & Partnership Plan"
                item["purpose"] = "Tunjukkan bentuk pilot, value untuk partner, dan metrik keberhasilan."
    return plan


def timing_for_plan(plan: list[dict[str, Any]], minutes: int) -> list[int]:
    total_seconds = minutes * 60
    weights = []
    for item in plan:
        key = item["key"]
        if key in {"cover", "closing"}:
            weights.append(0.6)
        elif key in {"traction", "business_model", "financials", "financial_ask", "fundraising", "competition"}:
            weights.append(1.25)
        elif key in {"problem_solution", "market_model", "team_ask"}:
            weights.append(1.35)
        else:
            weights.append(1.0)
    total_weight = sum(weights)
    seconds = [max(18, int(total_seconds * w / total_weight)) for w in weights]
    diff = total_seconds - sum(seconds)
    if seconds:
        seconds[-1] += diff
    return seconds


def category_score(data: dict[str, Any]) -> dict[str, int]:
    def score_text(key: str, min_lines: int = 1, min_chars: int = 40) -> int:
        value = str(data.get(key, "")).strip()
        line_count = len(clean_lines(value, 10))
        score = 25 if value else 0
        if len(value) >= min_chars:
            score += 35
        if line_count >= min_lines:
            score += 25
        if len(value) > 500:
            score -= 10
        return max(0, min(score + 15, 100))
    competitors = data.get("competitors", [])
    milestones = data.get("milestones", [])
    gm = parse_percent(data.get("gross_margin", ""))
    retention = parse_percent(data.get("retention", ""))
    financial_score = 55
    if data.get("rev1") and data.get("rev2") and data.get("rev3"):
        financial_score += 15
    if (data.get("ask") or 0) > 0 and str(data.get("use_of_funds", "")).strip():
        financial_score += 15
    if len(milestones) >= 3:
        financial_score += 15
    return {
        "Problem clarity": score_text("problem", 2, 70),
        "Solution clarity": score_text("solution", 2, 70),
        "Market logic": min(100, 45 + 15 * sum(1 for x in ["tam", "sam", "som"] if str(data.get(x, "")).strip()) + (20 if str(data.get("market_notes", "")).strip() else 0)),
        "Traction strength": min(100, 30 + 15 * sum(1 for x in ["users", "revenue", "growth", "retention"] if str(data.get(x, "")).strip()) + (10 if retention and retention >= 0.6 else 0)),
        "Business model quality": min(100, 40 + 10 * len(get_model_metrics(data)) + (20 if gm and gm >= 0.5 else 0)),
        "Financial consistency": min(100, financial_score),
        "Funding ask clarity": min(100, 30 + (25 if data.get("ask") else 0) + (25 if str(data.get("use_of_funds", "")).strip() else 0) + (20 if len(milestones) >= 3 else 0)),
        "Team strength": score_text("team", 2, 60),
        "Competition clarity": min(100, 35 + 15 * min(len(competitors), 4) + (20 if str(data.get("competition_summary", "")).strip() else 0)),
    }


def validate_inputs(data: dict[str, Any]) -> list[dict[str, str]]:
    issues: list[dict[str, str]] = []
    if len(clean_lines(data.get("problem", ""), 10)) < 2:
        issues.append({"level": "warning", "area": "Problem", "message": "Problem minimal berisi 2-3 poin agar investor memahami pain point, frekuensi, dan biaya masalah."})
    if not str(data.get("problem_evidence", "")).strip():
        issues.append({"level": "warning", "area": "Problem evidence", "message": "Tambahkan bukti problem: interview, waiting list, pilot, data pasar, atau biaya masalah."})
    if len(clean_lines(data.get("solution", ""), 10)) < 2:
        issues.append({"level": "warning", "area": "Solution", "message": "Solution terlalu pendek. Jelaskan mekanisme produk dan hasil yang dirasakan customer."})
    for key, name in [("tam", "TAM"), ("sam", "SAM"), ("som", "SOM")]:
        if not str(data.get(key, "")).strip():
            issues.append({"level": "error", "area": "Market", "message": f"{name} belum diisi. Investor perlu memahami ukuran pasar dan fokus segmen awal."})
    if not str(data.get("traction_notes", "")).strip():
        issues.append({"level": "warning", "area": "Traction", "message": "Traction belum punya narasi. Tambahkan revenue, active usage, retention, pilot, LOI, pipeline, atau repeat usage."})
    if len(data.get("competitors", [])) < 2:
        issues.append({"level": "warning", "area": "Competition", "message": "Tambahkan minimal 2 alternatif, termasuk status quo. Tidak ada kompetitor biasanya dianggap red flag."})
    if len(data.get("milestones", [])) < 3:
        issues.append({"level": "warning", "area": "Milestone", "message": "Tambahkan minimal 3 milestone: product, traction/revenue, dan readiness untuk round berikutnya."})
    if not str(data.get("use_of_funds", "")).strip():
        issues.append({"level": "error", "area": "Funding ask", "message": "Use of funds wajib jelas agar investor tahu dana akan mengubah apa menjadi milestone apa."})
    if (data.get("ask") or 0) <= 0:
        issues.append({"level": "error", "area": "Funding ask", "message": "Jumlah pendanaan harus lebih dari 0."})
    if not str(data.get("founder_fit", "")).strip():
        issues.append({"level": "warning", "area": "Team", "message": "Founder-market fit belum dijelaskan. Hubungkan pengalaman tim dengan problem dan akses market."})
    return issues


def generate_investor_insights(data: dict[str, Any]) -> dict[str, Any]:
    currency = data.get("currency", "Rp")
    rev1 = float(data.get("rev1", 0) or 0)
    rev2 = float(data.get("rev2", 0) or 0)
    rev3 = float(data.get("rev3", 0) or 0)
    cost1 = float(data.get("cost1", 0) or 0)
    profit3 = float(data.get("profit3", 0) or 0)
    ask = float(data.get("ask", 0) or 0)
    runway_input = float(data.get("runway", 0) or 0)
    growth_y2 = safe_div(rev2 - rev1, rev1)
    growth_y3 = safe_div(rev3 - rev2, rev2)
    burn_y1 = max(cost1 - rev1, 0)
    monthly_burn_y1 = burn_y1 / 12 if burn_y1 else 0
    estimated_runway = safe_div(ask, monthly_burn_y1) if monthly_burn_y1 else None
    year3_margin = safe_div(profit3, rev3) if rev3 else None
    funding_to_cost = safe_div(ask, cost1) if cost1 else None
    scores = category_score(data)
    score = int(sum(scores.values()) / len(scores)) if scores else 0
    issues = validate_inputs(data)
    auto_metrics = calculate_auto_metrics(data)
    strengths: list[str] = []
    risks: list[str] = []
    recommendations: list[str] = []
    if growth_y2 and growth_y2 > 0.3:
        strengths.append(f"Revenue diproyeksikan tumbuh {pct(growth_y2)} dari Year 1 ke Year 2. Narasikan asumsi channel dan conversion yang membuat growth ini realistis.")
    else:
        risks.append("Proyeksi revenue belum menunjukkan growth yang kuat. Jelaskan strategi akuisisi dan expansion agar growth story tidak terlihat linear biasa.")
    gm = parse_percent(data.get("gross_margin", ""))
    if gm and gm >= 0.65:
        strengths.append(f"Gross margin {pct(gm)} memberi sinyal model bisnis bisa scalable bila CAC dan retention terkendali.")
    elif gm:
        risks.append(f"Gross margin {pct(gm)} perlu dijelaskan: bagaimana margin membaik saat scale, automation, pricing, atau supplier terms meningkat.")
    retention = parse_percent(data.get("retention", ""))
    if retention and retention >= 0.65:
        strengths.append(f"Retention {pct(retention)} dapat dipakai sebagai bukti awal produk mulai menjadi habit atau workflow penting.")
    elif retention:
        risks.append(f"Retention {pct(retention)} harus diimbangi dengan strategi activation, onboarding, dan customer success.")
    if estimated_runway:
        if runway_input and estimated_runway < runway_input * 0.75:
            risks.append(f"Runway input {runway_input:.0f} bulan terlihat agresif dibanding estimasi burn Year 1 sekitar {estimated_runway:.0f} bulan.")
        else:
            strengths.append(f"Ask pendanaan relatif konsisten dengan burn Year 1 dan memberi estimasi runway sekitar {estimated_runway:.0f} bulan.")
    if year3_margin and year3_margin > 0:
        strengths.append(f"Profit margin Year 3 sekitar {pct(year3_margin)}. Ini membantu menunjukkan potensi operating leverage.")
    model = business_model_template(data)
    recommendations.append(f"Untuk model {data.get('business_model_type')}, fokuskan narasi pada: {model['pitch_focus']}")
    recommendations.append(f"Gunakan milestone sebagai jembatan antara ask {money(ask, currency)} dan next round: apa yang akan terbukti dalam {data.get('runway')} bulan.")
    recommendations.append("Siapkan asumsi angka: sumber data market size, conversion funnel, CAC, churn/retention, dan alasan gross margin bisa naik.")
    qa = build_investor_qa(data)
    return {
        "score": score,
        "category_scores": scores,
        "issues": issues,
        "metrics": {
            "Growth Y1→Y2": pct(growth_y2),
            "Growth Y2→Y3": pct(growth_y3),
            "Burn Year 1": short_money(burn_y1, currency),
            "Estimated Runway": f"{estimated_runway:.0f} bulan" if estimated_runway else "N/A",
            "Year 3 Margin": pct(year3_margin),
            "Funding/Cost Y1": multiple(funding_to_cost),
            **auto_metrics,
        },
        "strengths": strengths[:4] or ["Deck sudah memiliki struktur seed lengkap. Fokus berikutnya adalah memperkuat bukti demand dan asumsi angka."],
        "risks": risks[:4] or ["Risiko belum terlihat jelas dari input. Tetap siapkan jawaban untuk CAC, churn, defensibility, dan asumsi financial."],
        "recommendations": recommendations[:5],
        "qa": qa,
        "headline": f"Deck readiness {score}/100. Fokus: problem urgency, traction quality, revenue engine, dan funding-to-milestone logic.",
    }


def build_investor_qa(data: dict[str, Any]) -> list[tuple[str, str]]:
    model = business_model_template(data)
    return [
        ("Kenapa masalah ini urgent sekarang?", f"Gunakan evidence problem: {truncate(data.get('problem_evidence'), 160)}. Hubungkan dengan biaya waktu/uang jika customer tetap memakai status quo."),
        ("Kenapa solusi ini bisa menang dibanding alternatif?", f"Tekankan narrative advantage: {truncate(data.get('competition_summary'), 160)}."),
        ("Bagaimana startup menghasilkan uang?", f"Model bisnis: {data.get('business_model_type')}. Fokus metrik: {model['pitch_focus']}"),
        ("Apa asumsi utama proyeksi finansial?", "Siapkan asumsi jumlah customer, pricing/ARPU, conversion channel, CAC, retention, COGS, dan hiring plan."),
        ("Apa yang akan dicapai dengan dana ini?", f"Jawab dengan milestone: {truncate(data.get('milestone'), 150)} dan use of funds yang spesifik."),
        ("Apa red flag terbesar bisnis ini?", "Jawab dengan jujur: acquisition, retention, margin, compliance, atau execution risk - lalu jelaskan mitigasinya."),
        ("Mengapa tim ini tepat?", f"Hubungkan pengalaman tim dengan founder-market fit: {truncate(data.get('founder_fit'), 160)}."),
    ]


def render_scorecards(insights: dict[str, Any]) -> None:
    st.subheader("📈 Analisa Investor Readiness")
    st.progress(insights["score"] / 100)
    st.markdown(f"<div class='readable-panel'><p><strong>{html.escape(insights['headline'])}</strong></p></div>", unsafe_allow_html=True)
    cols = st.columns(3)
    for idx, (label, value) in enumerate(list(insights["metrics"].items())[:9]):
        with cols[idx % 3]:
            st.metric(label, value)
    st.markdown("#### Skor per kategori")
    score_cols = st.columns(3)
    for idx, (label, value) in enumerate(insights["category_scores"].items()):
        with score_cols[idx % 3]:
            st.markdown(f"<div class='score-card'><strong>{html.escape(label)}</strong><p>{value}/100</p></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        render_list_card("Kekuatan", insights["strengths"], "ok")
    with c2:
        render_list_card("Risiko / pertanyaan investor", insights["risks"], "risk")
    with c3:
        render_list_card("Rekomendasi", insights["recommendations"], "warn")


def render_list_card(title: str, items: list[str], css_class: str = "") -> None:
    safe_items = "".join(f"<li>{html.escape(str(item))}</li>" for item in items)
    st.markdown(f"<div class='insight-card {css_class}'><h4>{html.escape(title)}</h4><ul>{safe_items}</ul></div>", unsafe_allow_html=True)


def render_validation_panel(issues: list[dict[str, str]]) -> None:
    if not issues:
        st.success("Validasi input aman. Tidak ada isu besar yang terdeteksi.")
        return
    for issue in issues:
        if issue["level"] == "error":
            st.error(f"{issue['area']}: {issue['message']}")
        else:
            st.warning(f"{issue['area']}: {issue['message']}")



def _isolated_page_html(title: str, body_html: str, subtitle: str = "") -> str:
    """Return a light-theme isolated HTML document for Streamlit components.

    This iframe is intentionally independent from Streamlit theme CSS so black text
    cannot be rendered on black backgrounds in Q&A, glossary, and formula sections.
    """
    safe_title = html.escape(title)
    safe_subtitle = html.escape(subtitle)
    return f"""
    <!doctype html>
    <html lang="id">
    <head>
      <meta charset="utf-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1" />
      <style>
        :root {{ color-scheme: light; }}
        * {{ box-sizing: border-box; }}
        html, body {{
          margin: 0;
          padding: 0;
          background: #f8fafc !important;
          color: #0f172a !important;
          -webkit-text-fill-color: #0f172a !important;
          font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
          line-height: 1.55;
        }}
        .shell {{
          background: #f8fafc;
          color: #0f172a;
          padding: 2px 2px 12px 2px;
        }}
        .title {{
          font-size: 18px;
          font-weight: 900;
          color: #0f172a !important;
          -webkit-text-fill-color: #0f172a !important;
          margin: 0 0 4px 0;
        }}
        .subtitle {{
          font-size: 13px;
          color: #334155 !important;
          -webkit-text-fill-color: #334155 !important;
          margin: 0 0 12px 0;
        }}
        .grid {{
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
          gap: 12px;
        }}
        .card {{
          background: #ffffff !important;
          color: #0f172a !important;
          -webkit-text-fill-color: #0f172a !important;
          border: 1px solid #cbd5e1;
          border-radius: 16px;
          padding: 14px 16px;
          box-shadow: 0 8px 18px rgba(15, 23, 42, 0.08);
          overflow-wrap: anywhere;
        }}
        .card * {{
          color: #0f172a !important;
          -webkit-text-fill-color: #0f172a !important;
          opacity: 1 !important;
          text-shadow: none !important;
        }}
        .card h4 {{
          margin: 0 0 8px 0;
          font-size: 15px;
          line-height: 1.35;
          font-weight: 900;
        }}
        .label {{
          display: inline-block;
          font-size: 11px;
          font-weight: 900;
          letter-spacing: .04em;
          text-transform: uppercase;
          color: #1d4ed8 !important;
          -webkit-text-fill-color: #1d4ed8 !important;
          margin-bottom: 6px;
        }}
        .body {{
          margin: 0;
          font-size: 13px;
          color: #0f172a !important;
          -webkit-text-fill-color: #0f172a !important;
        }}
        .formula {{
          margin-top: 8px;
          padding: 8px 10px;
          border-radius: 10px;
          background: #e2e8f0 !important;
          color: #0f172a !important;
          -webkit-text-fill-color: #0f172a !important;
          border: 1px solid #94a3b8;
          font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
          font-size: 12px;
          font-weight: 800;
          white-space: normal;
          overflow-wrap: anywhere;
        }}
        .example {{
          margin-top: 8px;
          color: #334155 !important;
          -webkit-text-fill-color: #334155 !important;
          font-size: 12px;
          font-weight: 650;
        }}
        table {{
          width: 100%;
          border-collapse: collapse;
          background: #ffffff !important;
          color: #0f172a !important;
          -webkit-text-fill-color: #0f172a !important;
          border: 1px solid #cbd5e1;
          border-radius: 16px;
          overflow: hidden;
          box-shadow: 0 8px 18px rgba(15, 23, 42, 0.08);
        }}
        th {{
          background: #0f172a !important;
          color: #ffffff !important;
          -webkit-text-fill-color: #ffffff !important;
          text-align: left;
          padding: 12px;
          font-size: 12px;
          text-transform: uppercase;
          letter-spacing: .04em;
        }}
        td {{
          color: #0f172a !important;
          -webkit-text-fill-color: #0f172a !important;
          background: #ffffff !important;
          border-bottom: 1px solid #e2e8f0;
          padding: 12px;
          vertical-align: top;
          font-size: 13px;
          overflow-wrap: anywhere;
        }}
        tr:nth-child(even) td {{ background: #f8fafc !important; }}
        tr:last-child td {{ border-bottom: 0; }}
        strong, b {{ color: #0f172a !important; -webkit-text-fill-color: #0f172a !important; font-weight: 900; }}
      </style>
    </head>
    <body>
      <div class="shell">
        <div class="title">{safe_title}</div>
        {f'<div class="subtitle">{safe_subtitle}</div>' if safe_subtitle else ''}
        {body_html}
      </div>
    </body>
    </html>
    """


def render_investor_qa(qa: list[tuple[str, str]]) -> None:
    rows = []
    for idx, (question, answer) in enumerate(qa, 1):
        rows.append(
            "<tr>"
            f"<td style='width:44px;font-weight:900;color:#1d4ed8;-webkit-text-fill-color:#1d4ed8;'>{idx}</td>"
            f"<td><strong>{html.escape(str(question))}</strong></td>"
            f"<td>{html.escape(str(answer))}</td>"
            "</tr>"
        )
    body = (
        "<table aria-label='Q&A investor'>"
        "<thead><tr><th>No</th><th>Pertanyaan investor</th><th>Cara menjawab</th></tr></thead>"
        f"<tbody>{''.join(rows)}</tbody></table>"
    )
    iframe = _isolated_page_html(
        "Q&A investor",
        body,
        "Bagian ini memakai tema terang terisolasi agar selalu terbaca di semua tema Streamlit.",
    )
    height = min(760, max(320, 110 + len(qa) * 88))
    components.html(iframe, height=height, scrolling=True)


def render_glossary_cards(rows: list[tuple[str, str, str, str, str]]) -> None:
    if not rows:
        st.warning("Tidak ada istilah yang cocok dengan filter.")
        return
    cards = []
    for cat, term, simple, formula, example in rows:
        cards.append(
            "<div class='card'>"
            f"<span class='label'>{html.escape(cat)}</span>"
            f"<h4>{html.escape(term)}</h4>"
            f"<p class='body'><strong>Arti sederhana:</strong> {html.escape(simple)}</p>"
            f"<div class='formula'>Cara menghitung: {html.escape(formula)}</div>"
            f"<div class='example'><strong>Contoh:</strong> {html.escape(example)}</div>"
            "</div>"
        )
    body = f"<div class='grid'>{''.join(cards)}</div>"
    iframe = _isolated_page_html(
        "Istilah investor dan cara menghitungnya",
        body,
        "Semua kartu memakai background terang dan teks gelap agar rumus tetap terbaca.",
    )
    height = min(900, max(360, 160 + math.ceil(len(rows) / 2) * 230))
    components.html(iframe, height=height, scrolling=True)

def render_glossary() -> None:
    guide(
        "Istilah & rumus",
        "Founder baru bisa belajar istilah investor langsung dari aplikasi. Gunakan pencarian untuk mencari TAM, CAC, runway, dilution, GMV, dan istilah lain.",
    )
    q = st.text_input("Cari istilah", "", help="Contoh: TAM, CAC, runway, gross margin, take rate")
    categories = sorted({item[0] for item in GLOSSARY})
    category = st.selectbox("Filter kategori", ["Semua"] + categories)
    rows = []
    for cat, term, simple, formula, example in GLOSSARY:
        text = " ".join([cat, term, simple, formula, example]).lower()
        if q and q.lower() not in text:
            continue
        if category != "Semua" and cat != category:
            continue
        rows.append((cat, term, simple, formula, example))
    render_glossary_cards(rows)


def render_calculator(data: dict[str, Any] | None = None) -> None:
    guide(
        "Kalkulator metrik otomatis",
        "Gunakan bagian ini untuk menghitung metrik umum. Hasilnya membantu mengisi Business Model, Financials, dan PDF scenario guide.",
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        st.number_input("Customer berbayar", min_value=0, step=10, key="paid_customers", help="Dipakai untuk menghitung MRR, ARR, dan ARPU.")
        st.number_input("Harga bulanan rata-rata", min_value=0, step=10_000, key="monthly_price", help="Average subscription price per month.")
        st.number_input("Revenue bulanan aktual", min_value=0, step=1_000_000, key="monthly_revenue_calc", help="Masukkan jika revenue tidak hanya subscription.")
    with c2:
        st.number_input("Biaya sales & marketing", min_value=0, step=1_000_000, key="marketing_spend", help="Dipakai menghitung CAC.")
        st.number_input("Customer baru", min_value=0, step=1, key="new_customers", help="Customer baru dari periode biaya marketing di atas.")
        st.number_input("COGS / biaya langsung", min_value=0, step=1_000_000, key="cogs", help="Dipakai menghitung gross margin.")
    with c3:
        st.number_input("Kas tersedia", min_value=0, step=10_000_000, key="cash_available", help="Dipakai menghitung runway.")
        st.number_input("Biaya bulanan", min_value=0, step=1_000_000, key="monthly_cost", help="Biaya operasional per bulan.")
        st.number_input("GMV / volume transaksi", min_value=0, step=1_000_000, key="gmv_calc", help="Untuk marketplace/fintech.")
        st.number_input("Revenue platform dari transaksi", min_value=0, step=1_000_000, key="platform_revenue_calc", help="Dipakai menghitung take rate.")
    calc_data = collect_data_preview_only()
    metrics = calculate_auto_metrics(calc_data)
    metric_cols = st.columns(5)
    for idx, (label, value) in enumerate(metrics.items()):
        with metric_cols[idx % 5]:
            st.metric(label, value)


def label_text(text_id: str, language: str) -> str:
    labels = {
        "cover": ("Opening", "Pembuka"),
        "problem": ("Problem", "Masalah"),
        "solution": ("Solution", "Solusi"),
        "product": ("Product", "Produk"),
        "market": ("Market Opportunity", "Peluang Pasar"),
        "business_model": ("Business Model", "Model Bisnis"),
        "traction": ("Traction", "Traksi"),
        "gtm": ("Go-To-Market", "Go-To-Market"),
        "competition": ("Competition", "Kompetisi"),
        "milestones": ("Milestones", "Milestone"),
        "financials": ("Financials", "Finansial"),
        "fundraising": ("Fundraising Ask", "Permintaan Pendanaan"),
        "team": ("Team", "Tim"),
        "readiness": ("Investor Readiness", "Kesiapan Investor"),
        "closing": ("Closing", "Penutup"),
    }
    en, idn = labels.get(text_id, (text_id, text_id))
    if language == "English":
        return en
    if language == "Bilingual":
        return f"{en} / {idn}" if en != idn else en
    return idn


def add_text(slide, text, x, y, w, h, size=18, color=None, bold=False, align=PP_ALIGN.LEFT):
    color = color or THEME["ink"]
    value = str(text or "")
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = value
    r.font.name = FONT_HEAD if bold and len(value) < 90 else FONT_BODY
    fitted_size = estimate_fit_font_size(value, w, h, size, min_size=6.5 if size <= 12 else 7.0)
    r.font.size = Pt(fitted_size)
    r.font.bold = bold
    r.font.color.rgb = color
    return box


def add_logo(slide, data, x=11.55, y=0.35, w=0.65):
    logo_bytes = data.get("logo_bytes")
    if not logo_bytes:
        return
    try:
        slide.shapes.add_picture(BytesIO(logo_bytes), Inches(x), Inches(y), width=Inches(w))
    except Exception:
        pass


def add_header(slide, title, subtitle, data, page: int):
    accent = rgb(data.get("accent_color", "#2563EB"))
    add_text(slide, str(title).upper(), 0.7, 0.34, 3.8, 0.25, 8, accent, True)
    add_text(slide, title, 0.7, 0.66, 9.7, 0.58, 27, THEME["ink"], True)
    add_text(slide, truncate(subtitle, 180), 0.72, 1.25, 10.9, 0.33, 10, THEME["muted"])
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(1.7), Inches(11.95), Inches(0.01))
    line.fill.solid(); line.fill.fore_color.rgb = THEME["line"]; line.line.fill.background()
    add_logo(slide, data)
    add_footer(slide, data, page)


def add_footer(slide, data, page):
    add_text(slide, truncate(data.get("company", ""), 40), 0.7, 7.08, 3.1, 0.22, 7.5, THEME["muted"])
    add_text(slide, "Confidential Investor Deck", 3.75, 7.08, 3.2, 0.22, 7.5, THEME["muted"], align=PP_ALIGN.CENTER)
    add_text(slide, DEVELOPER, 6.95, 7.08, 3.7, 0.22, 7.2, THEME["muted"], align=PP_ALIGN.CENTER)
    add_text(slide, str(page).zfill(2), 12.05, 7.08, 0.55, 0.22, 7.5, THEME["muted"], align=PP_ALIGN.RIGHT)


def add_bullets(slide, items, x, y, w, h, size=18, max_items=5):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear(); tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = Inches(0); tf.margin_right = Inches(0)
    clean_items = (items or ["Lengkapi poin utama slide ini."])[:max_items]
    total_text = "\n".join(str(item or "") for item in clean_items)
    fitted_size = estimate_fit_font_size(total_text, w, h, size, min_size=7.0)
    for idx, item in enumerate(clean_items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = str(item or "")
        p.level = 0
        p.space_after = Pt(7 if fitted_size < 12 else 10)
        p.font.name = FONT_BODY
        p.font.size = Pt(fitted_size)
        p.font.color.rgb = THEME["ink"]


def add_card(slide, title, body, x, y, w, h, data, title_size=8, body_size=15):
    accent = rgb(data.get("accent_color", "#2563EB"))
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid(); shape.fill.fore_color.rgb = THEME["white"]; shape.line.color.rgb = THEME["line"]
    add_text(slide, str(title).upper(), x + 0.2, y + 0.16, w - 0.4, 0.25, title_size, accent, True)
    body_text = str(body or "")
    adjusted_body_size = estimate_fit_font_size(body_text, w - 0.4, h - 0.65, body_size, min_size=6.8)
    add_text(slide, body_text, x + 0.2, y + 0.52, w - 0.4, h - 0.65, adjusted_body_size, THEME["ink"], True)


def add_metric(slide, label, value, x, y, w, data, size=24):
    accent = rgb(data.get("accent_color", "#2563EB"))
    value_text = str(value or "")
    value_h = 0.66 if len(value_text) > 28 else 0.45
    add_text(slide, value_text, x, y, w, value_h, size, accent, True)
    add_text(slide, str(label).upper(), x, y + value_h + 0.10, w, 0.23, 8, THEME["muted"], True)


def add_takeaway(slide, text, data):
    accent = rgb(data.get("accent_color", "#2563EB"))
    takeaway = "Investor takeaway: " + str(text or "")
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(6.12), Inches(11.95), Inches(0.66))
    shape.fill.solid(); shape.fill.fore_color.rgb = THEME["blue_soft"]; shape.line.color.rgb = RGBColor(191, 219, 254)
    add_text(slide, takeaway, 0.95, 6.25, 11.2, 0.38, 9.5, accent, True)


def add_table(slide, headers, rows, x, y, w, h, data):
    accent = rgb(data.get("accent_color", "#2563EB"))
    shape = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = shape.table
    body_size = table_font_size(rows, 8.5, 6.2)
    for c, header in enumerate(headers):
        cell = tbl.cell(0, c); cell.text = str(header); cell.fill.solid(); cell.fill.fore_color.rgb = accent
        cell.text_frame.word_wrap = True
        for p in cell.text_frame.paragraphs:
            p.font.name = FONT_BODY; p.font.size = Pt(8.5); p.font.bold = True; p.font.color.rgb = THEME["white"]
    for r_idx, row in enumerate(rows, 1):
        for c_idx, value in enumerate(row):
            cell = tbl.cell(r_idx, c_idx); cell.text = str(value or ""); cell.fill.solid(); cell.fill.fore_color.rgb = THEME["white"]
            cell.text_frame.word_wrap = True
            for p in cell.text_frame.paragraphs:
                p.font.name = FONT_BODY; p.font.size = Pt(body_size); p.font.color.rgb = THEME["ink"]


def new_slide(prs, data, title, subtitle, page):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, THEME["bg"])
    add_header(slide, title, subtitle, data, page)
    return slide


def content_slide(prs, data, page, title, subtitle, bullets, takeaway, side_title=None, side_body=None, metrics=None):
    slide = new_slide(prs, data, title, subtitle, page)
    if metrics:
        width = 11.4 / max(len(metrics), 1)
        for idx, (label, value) in enumerate(metrics[:4]):
            add_metric(slide, label, value, 0.85 + idx * width, 2.1, width - 0.25, data)
        add_bullets(slide, bullets, 0.85, 3.45, 10.9, 2.2, 17, 5)
    elif side_title:
        add_bullets(slide, bullets, 0.85, 2.05, 6.25, 3.65, 19, 5)
        add_card(slide, side_title, side_body, 7.65, 2.15, 4.45, 2.15, data)
    else:
        add_bullets(slide, bullets, 0.85, 2.05, 10.9, 3.7, 20, 5)
    add_takeaway(slide, takeaway, data)
    return slide


def add_cover_slide(prs, data, page):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(data.get("secondary_color", "#0F172A")))
    accent = rgb(data.get("accent_color", "#2563EB"))
    add_logo(slide, data, 11.25, 0.75, 1.0)
    add_text(slide, data.get("company", ""), 0.85, 0.85, 8.7, 0.7, 38, THEME["white"], True)
    add_text(slide, truncate(data.get("one_liner", ""), 150), 0.88, 1.75, 8.8, 0.85, 24, RGBColor(226, 232, 240), True)
    add_text(slide, f"{data.get('pitch_type')} • {data.get('business_model_type')}", 0.9, 3.65, 8.2, 0.3, 11, RGBColor(203, 213, 225))
    add_text(slide, f"{data.get('round')} • {money(data.get('ask', 0), data.get('currency', 'Rp'))}", 0.9, 4.85, 7.6, 0.35, 15, accent, True)
    add_text(slide, f"{data.get('presenter')} | {data.get('contact')}", 0.9, 5.35, 7.6, 0.28, 11, RGBColor(203, 213, 225))
    add_text(slide, f"Duration: {data.get('pitch_duration_minutes')} min", 9.3, 6.62, 2.8, 0.25, 9.5, RGBColor(148, 163, 184), True, PP_ALIGN.RIGHT)
    add_text(slide, DEVELOPER, 0.85, 6.95, 5.6, 0.22, 7.5, RGBColor(148, 163, 184))
    return slide


def add_competition_slides(prs, data, page):
    competitors = data.get("competitors", []) or []
    if not competitors:
        competitors = [{"name": "Status quo", "category": "Status quo", "weakness": "Workflow manual", "advantage": data.get("competition_summary", "") }]
    duration = int(data.get("pitch_duration_minutes", 10))
    # Keep fewer rows per slide so full text can wrap without ellipsis.
    per_slide = 2
    pages = []
    for chunk_idx in range(0, len(competitors), per_slide):
        chunk = competitors[chunk_idx: chunk_idx + per_slide]
        slide = new_slide(prs, data, "Competition", "Alternatif yang dipakai customer, kelemahannya, dan kenapa kita menang.", page + len(pages))
        rows = [[c.get("name", ""), c.get("category", ""), c.get("weakness", ""), c.get("advantage", "")] for c in chunk]
        add_table(slide, ["Alternative", "Type", "Weakness", "Our edge"], rows, 0.8, 1.95, 11.75, 4.05, data)
        add_takeaway(slide, data.get("competition_summary", "Tunjukkan narrative advantage yang spesifik."), data)
        pages.append(slide)
    return len(pages)


def add_milestone_slides(prs, data, page):
    milestones = data.get("milestones", []) or []
    if not milestones:
        milestones = [{"period": "0-12 bulan", "target": data.get("milestone", "Milestone utama"), "metric": "Success metric", "owner": "Team"}]
    duration = int(data.get("pitch_duration_minutes", 10))
    per_slide = 2
    pages = []
    for chunk_idx in range(0, len(milestones), per_slide):
        chunk = milestones[chunk_idx: chunk_idx + per_slide]
        slide = new_slide(prs, data, "Milestones", "Target pembuktian yang menghubungkan funding dengan next round readiness.", page + len(pages))
        card_w = 11.6 / max(len(chunk), 1)
        for idx, item in enumerate(chunk):
            x = 0.85 + idx * card_w
            body = f"Target: {item.get('target', '')}\nMetric: {item.get('metric', '')}\nOwner: {item.get('owner', '')}"
            add_card(slide, item.get("period", f"Step {idx + 1}"), body, x, 2.05, card_w - 0.25, 3.15, data, 9, 13)
        add_takeaway(slide, data.get("milestone", "Milestone harus spesifik, terukur, dan terkait ask pendanaan."), data)
        pages.append(slide)
    return len(pages)


def build_deck(data: dict[str, Any], image_buffer: BytesIO | None = None) -> BytesIO:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    plan = build_slide_plan(data)
    insights = generate_investor_insights(data)
    page = 1
    for item in plan:
        key = item["key"]
        title = label_text(key, data.get("output_language", "Bahasa Indonesia")) if key in {"problem", "solution", "product", "market", "business_model", "traction", "gtm", "competition", "milestones", "financials", "fundraising", "team", "readiness", "closing", "cover"} else item["title"]
        if key == "cover":
            add_cover_slide(prs, data, page)
        elif key == "problem_solution":
            content_slide(prs, data, page, "Problem & Solution", item["purpose"], clean_lines(data.get("problem"), 3) + clean_lines(data.get("solution"), 3), "Investor harus paham masalah dan solusi dalam 60 detik.", "Proof of pain", data.get("problem_evidence"))
        elif key == "problem":
            content_slide(prs, data, page, title, item["purpose"], clean_lines(data.get("problem"), 5), "Problem harus terasa urgent, sering terjadi, dan mahal.", "Proof of pain", data.get("problem_evidence"))
        elif key == "solution_product":
            content_slide(prs, data, page, "Solution & Product", item["purpose"], clean_lines(data.get("solution"), 3) + clean_lines(data.get("features"), 3), "Solusi harus langsung menjawab problem utama.", "Product flow", data.get("product_flow"))
        elif key == "solution":
            content_slide(prs, data, page, title, item["purpose"], clean_lines(data.get("solution"), 5), "Solusi harus menjelaskan kondisi sebelum dan sesudah memakai produk.", "Value proposition", data.get("value_prop"))
        elif key == "product":
            slide = new_slide(prs, data, title, item["purpose"], page)
            if image_buffer:
                try:
                    image_buffer.seek(0)
                    slide.shapes.add_picture(image_buffer, Inches(0.85), Inches(2.0), width=Inches(6.2), height=Inches(3.4))
                    add_card(slide, "Product flow", data.get("product_flow"), 7.55, 2.0, 4.55, 1.55, data)
                    add_card(slide, "Benefit", data.get("product_benefit"), 7.55, 3.85, 4.55, 1.45, data)
                except Exception:
                    add_bullets(slide, clean_lines(data.get("features"), 5), 0.85, 2.05, 10.8, 3.65, 19)
            else:
                add_bullets(slide, clean_lines(data.get("features"), 5), 0.85, 2.05, 6.1, 3.65, 19)
                add_card(slide, "Product flow", data.get("product_flow"), 7.55, 2.0, 4.55, 1.55, data)
                add_card(slide, "Benefit", data.get("product_benefit"), 7.55, 3.85, 4.55, 1.45, data)
            add_takeaway(slide, "Investor harus paham value produk tanpa demo panjang.", data)
        elif key == "market_model":
            content_slide(prs, data, page, "Market & Business Model", item["purpose"], clean_lines(data.get("market_notes"), 3) + clean_lines(data.get("business_model"), 3), "Pasar awal harus spesifik dan revenue engine harus jelas.", metrics=[("TAM", data.get("tam")), ("SAM", data.get("sam")), ("SOM", data.get("som"))])
        elif key == "market":
            content_slide(prs, data, page, title, item["purpose"], clean_lines(data.get("market_notes"), 5), "TAM besar penting, tetapi wedge awal yang bisa dimenangkan lebih penting.", metrics=[("TAM", data.get("tam")), ("SAM", data.get("sam")), ("SOM", data.get("som"))])
        elif key == "business_model":
            model = business_model_template(data)
            content_slide(prs, data, page, title, f"{data.get('business_model_type')} - {model['description']}", clean_lines(data.get("business_model"), 5), f"Pitch focus: {model['pitch_focus']}", metrics=get_model_metrics(data))
        elif key == "traction_model":
            content_slide(prs, data, page, "Traction & Business Model", item["purpose"], clean_lines(data.get("traction_notes"), 3) + clean_lines(data.get("business_model"), 3), "Tunjukkan demand dan bagaimana demand berubah menjadi revenue.", metrics=[("Users", data.get("users")), ("Revenue", data.get("revenue")), ("Growth", data.get("growth")), ("Retention", data.get("retention"))])
        elif key == "traction":
            content_slide(prs, data, page, title, item["purpose"], clean_lines(data.get("traction_notes"), 5), "Traction harus membuktikan demand, bukan hanya vanity metrics.", metrics=[("Users", data.get("users")), ("Revenue", data.get("revenue")), ("Growth", data.get("growth")), ("Retention", data.get("retention"))])
        elif key == "gtm":
            content_slide(prs, data, page, title, item["purpose"], clean_lines(data.get("gtm"), 5), "GTM harus menunjukkan acquisition motion yang bisa berulang.", "ICP & Channel", f"{data.get('icp')}\n\nChannel utama: {data.get('channel')}")
        elif key == "competition":
            added = add_competition_slides(prs, data, page)
            page += added - 1
        elif key == "milestones":
            added = add_milestone_slides(prs, data, page)
            page += added - 1
        elif key == "financials":
            slide = new_slide(prs, data, title, item["purpose"], page)
            rows = [
                ["Revenue", money(data.get("rev1", 0), data.get("currency", "Rp")), money(data.get("rev2", 0), data.get("currency", "Rp")), money(data.get("rev3", 0), data.get("currency", "Rp"))],
                ["Operating Cost", money(data.get("cost1", 0), data.get("currency", "Rp")), money(data.get("cost2", 0), data.get("currency", "Rp")), money(data.get("cost3", 0), data.get("currency", "Rp"))],
                ["EBITDA / Profit", money(data.get("profit1", 0), data.get("currency", "Rp")), money(data.get("profit2", 0), data.get("currency", "Rp")), money(data.get("profit3", 0), data.get("currency", "Rp"))],
            ]
            add_table(slide, ["Metric", "Year 1", "Year 2", "Year 3"], rows, 0.85, 2.0, 11.65, 2.5, data)
            add_metric(slide, "Runway", f"{data.get('runway')} bulan", 0.95, 5.2, 2.3, data)
            add_card(slide, "Next Milestone", data.get("milestone"), 3.9, 5.05, 8.2, 1.0, data, 8, 12)
        elif key in {"financial_ask", "fundraising"}:
            content_slide(prs, data, page, title if key == "fundraising" else "Financials & Ask", item["purpose"], clean_lines(data.get("use_of_funds"), 5), "Ask harus jelas: jumlah, runway, use of funds, dan target pembuktian.", "Ask & next round", f"{money(data.get('ask', 0), data.get('currency', 'Rp'))} {data.get('round')}\n\nRunway: {data.get('runway')} bulan\n{data.get('next_round')}")
        elif key == "team_ask":
            content_slide(prs, data, page, "Team & Ask", item["purpose"], clean_lines(data.get("team"), 3) + clean_lines(data.get("use_of_funds"), 3), "Tim dan funding harus terlihat mampu mencapai milestone berikutnya.", "Funding", f"{money(data.get('ask', 0), data.get('currency', 'Rp'))}\n{data.get('milestone')}")
        elif key == "team":
            content_slide(prs, data, page, title, item["purpose"], clean_lines(data.get("team"), 5), "Tim harus terlihat punya unfair advantage untuk mengeksekusi peluang ini.", "Founder-market fit", data.get("founder_fit"))
        elif key == "readiness":
            content_slide(prs, data, page, title, item["purpose"], insights["strengths"][:2] + insights["risks"][:2], insights["headline"], metrics=[("Score", f"{insights['score']}/100"), ("Runway", insights["metrics"].get("Estimated Runway", "N/A")), ("Burn", insights["metrics"].get("Burn Year 1", "N/A"))])
        elif key == "closing":
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_bg(slide, rgb(data.get("secondary_color", "#0F172A")))
            add_logo(slide, data, 11.2, 0.75, 1.0)
            add_text(slide, truncate(data.get("closing", "Thank you"), 120), 0.85, 1.25, 9.5, 1.1, 34, THEME["white"], True)
            add_text(slide, f"{data.get('presenter')}\n{data.get('contact')}", 0.9, 5.35, 7.0, 0.6, 14, RGBColor(203, 213, 225))
            add_text(slide, DEVELOPER, 0.85, 6.95, 5.6, 0.22, 7.5, RGBColor(148, 163, 184))
        page += 1
    out = BytesIO()
    prs.save(out)
    out.seek(0)
    return out


def pdf_styles():
    styles = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("TitleCustom", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=20, leading=24, alignment=TA_CENTER, textColor=colors.HexColor("#0F172A")),
        "h1": ParagraphStyle("H1Custom", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=15, leading=18, textColor=colors.HexColor("#0F172A"), spaceAfter=8),
        "h2": ParagraphStyle("H2Custom", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=colors.HexColor("#2563EB"), spaceAfter=6),
        "body": ParagraphStyle("BodyCustom", parent=styles["BodyText"], fontName="Helvetica", fontSize=9.5, leading=13, textColor=colors.HexColor("#334155"), spaceAfter=6),
        "small": ParagraphStyle("SmallCustom", parent=styles["BodyText"], fontName="Helvetica", fontSize=8, leading=10, textColor=colors.HexColor("#64748B"), spaceAfter=4),
    }


def p(text: Any, style) -> Paragraph:
    return Paragraph(html.escape(str(text or "")).replace("\n", "<br/>"), style)


def make_rl_table(rows, widths=None):
    table = RLTable(rows, colWidths=widths)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2563EB")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return table



def talk_track_for_slide(item: dict[str, Any], data: dict[str, Any], insights: dict[str, Any] | None = None) -> str:
    """Create the same practical speaking guidance used by the PDF and prompter."""
    insights = insights or generate_investor_insights(data)
    key = item.get("key", "")
    if key == "cover":
        return f"Perkenalkan {data.get('company')} dalam satu kalimat: {data.get('one_liner')}. Sebutkan {data.get('round')} dan ask {money(data.get('ask'), data.get('currency'))} secara ringkas."
    if "problem" in key:
        return f"Mulai dari pain point customer: {truncate(data.get('problem'), 260)} Evidence: {truncate(data.get('problem_evidence'), 180)}"
    if "solution" in key or key == "product":
        return f"Jelaskan solusi dan flow produk: {truncate(data.get('solution'), 240)} Product flow: {truncate(data.get('product_flow'), 180)}"
    if "market" in key:
        return f"Sebutkan TAM {data.get('tam')}, SAM {data.get('sam')}, SOM {data.get('som')}. Jelaskan wedge awal: {truncate(data.get('market_notes'), 220)}"
    if "business" in key or "model" in key:
        model = business_model_template(data)
        metrics = ", ".join([a + ": " + b for a, b in get_model_metrics(data)])
        return f"Model bisnis: {data.get('business_model_type')}. Formula utama: {model['formula']}. Metrik utama: {metrics}."
    if "traction" in key:
        return f"Bukti demand: users {data.get('users')}, revenue {data.get('revenue')}, growth {data.get('growth')}, retention {data.get('retention')}. {truncate(data.get('traction_notes'), 240)}"
    if key == "gtm":
        return f"ICP: {data.get('icp')}. Channel utama: {data.get('channel')}. GTM: {truncate(data.get('gtm'), 240)}"
    if key == "competition":
        competitors = data.get("competitors", []) or []
        comp_names = ", ".join([str(c.get("name", "")) for c in competitors[:4] if c.get("name")])
        return f"Jelaskan alternatif customer saat ini ({comp_names or 'status quo dan alternatif utama'}) dan narrative advantage: {data.get('competition_summary')}"
    if key == "milestones":
        milestones = data.get("milestones", []) or []
        details = "; ".join([f"{m.get('period')}: {m.get('target')} ({m.get('metric')})" for m in milestones[:4]])
        return f"Hubungkan dana dengan milestone berikut: {details or data.get('milestone')}. Tekankan apa yang akan terbukti sebelum round berikutnya."
    if "financial" in key or "fundraising" in key:
        return f"Ask {money(data.get('ask'), data.get('currency'))}, runway {data.get('runway')} bulan, use of funds: {truncate(data.get('use_of_funds'), 240)}. Next round logic: {truncate(data.get('next_round'), 160)}"
    if "team" in key:
        return f"Kenapa tim ini tepat: {truncate(data.get('team'), 240)} Founder-market fit: {truncate(data.get('founder_fit'), 180)}"
    if key == "readiness":
        return insights.get("headline", "Ringkas kesiapan deck, kekuatan, dan risiko utama.")
    return data.get("closing", "Tutup dengan visi dan next step yang jelas.")


def screen_summary_for_slide(item: dict[str, Any], data: dict[str, Any]) -> str:
    key = item.get("key", "")
    if key == "cover":
        return f"{data.get('company')}\n{data.get('one_liner')}\n{data.get('round')} • {money(data.get('ask'), data.get('currency'))}"
    if "problem" in key and "solution" in key:
        return f"Problem:\n{data.get('problem')}\n\nSolution:\n{data.get('solution')}"
    if key == "problem":
        return f"{data.get('problem')}\n\nProof: {data.get('problem_evidence')}"
    if "solution" in key or key == "product":
        return f"{data.get('solution')}\n\nFlow: {data.get('product_flow')}\n\nBenefit: {data.get('product_benefit')}"
    if "market" in key:
        return f"TAM: {data.get('tam')} | SAM: {data.get('sam')} | SOM: {data.get('som')}\n\n{data.get('market_notes')}"
    if "business" in key or "model" in key:
        metrics = "\n".join([f"{a}: {b}" for a, b in get_model_metrics(data)])
        return f"{data.get('business_model_type')}\n{data.get('business_model')}\n\n{metrics}"
    if "traction" in key:
        return f"Users: {data.get('users')} | Revenue: {data.get('revenue')} | Growth: {data.get('growth')} | Retention: {data.get('retention')}\n\n{data.get('traction_notes')}"
    if key == "gtm":
        return f"ICP: {data.get('icp')}\nChannel: {data.get('channel')}\n\n{data.get('gtm')}"
    if key == "competition":
        competitors = data.get("competitors", []) or []
        rows = [f"• {c.get('name')} — {c.get('advantage')}" for c in competitors[:5]]
        return "\n".join(rows) + f"\n\nTakeaway: {data.get('competition_summary')}"
    if key == "milestones":
        milestones = data.get("milestones", []) or []
        return "\n".join([f"• {m.get('period')}: {m.get('target')} — {m.get('metric')}" for m in milestones[:5]])
    if "financial" in key or "fundraising" in key:
        return f"Ask: {money(data.get('ask'), data.get('currency'))}\nRunway: {data.get('runway')} bulan\nMilestone: {data.get('milestone')}\n\nUse of funds:\n{data.get('use_of_funds')}"
    if "team" in key:
        return f"{data.get('team')}\n\nFounder-market fit: {data.get('founder_fit')}"
    if key == "readiness":
        insights = generate_investor_insights(data)
        return f"Score: {insights.get('score')}/100\n{insights.get('headline')}"
    return data.get("closing", "Thank you")


def normalize_custom_prompter_scripts(data: dict[str, Any], plan_len: int) -> list[str]:
    """Return user-edited teleprompter scripts aligned to the current slide plan."""
    raw = data.get("custom_prompter_scripts", []) or []
    if not isinstance(raw, list):
        raw = []

    scripts: list[str] = []
    for idx in range(plan_len):
        value = raw[idx] if idx < len(raw) else ""
        scripts.append(str(value or "").strip())
    return scripts


def build_rehearsal_items(data: dict[str, Any]) -> list[dict[str, Any]]:
    plan = build_slide_plan(data)
    timings = timing_for_plan(plan, int(data.get("pitch_duration_minutes", 10)))
    insights = generate_investor_insights(data)
    custom_scripts = normalize_custom_prompter_scripts(data, len(plan))
    use_custom = bool(data.get("use_custom_prompter", True))
    items: list[dict[str, Any]] = []
    for idx, item in enumerate(plan):
        generated_talk = talk_track_for_slide(item, data, insights)
        custom_talk = custom_scripts[idx] if idx < len(custom_scripts) else ""
        talk = custom_talk if use_custom and custom_talk else generated_talk
        items.append({
            "slide": idx + 1,
            "title": str(item.get("title", f"Slide {idx + 1}")),
            "key": str(item.get("key", "")),
            "duration": int(timings[idx]),
            "purpose": str(item.get("purpose", "")),
            "talk": talk,
            "generated_talk": generated_talk,
            "screen": screen_summary_for_slide(item, data),
            "transition": "Hubungkan poin ini ke slide berikutnya; jangan membaca semua teks, tekankan bukti dan keputusan yang ingin investor ingat.",
        })
    return items


def sync_prompter_script_defaults(data: dict[str, Any], *, force: bool = False) -> list[dict[str, Any]]:
    """Prepare editable Streamlit text areas for the current slide plan."""
    base_data = dict(data)
    base_data["use_custom_prompter"] = False
    base_data["custom_prompter_scripts"] = []
    items = build_rehearsal_items(base_data)
    st.session_state["prompter_slide_count"] = len(items)

    for idx, item in enumerate(items):
        key = f"prompter_script_{idx}"
        if force or key not in st.session_state:
            st.session_state[key] = item.get("generated_talk", item.get("talk", ""))

    return items


def build_rehearsal_html(data: dict[str, Any], *, standalone: bool = True) -> str:
    items = build_rehearsal_items(data)
    payload = json.dumps({
        "company": data.get("company", "Startup"),
        "duration": int(data.get("pitch_duration_minutes", 10)),
        "pitchType": data.get("pitch_type", "Investor Seed Round"),
        "businessModel": data.get("business_model_type", ""),
        "developer": DEVELOPER,
        "items": items,
    }, ensure_ascii=False)
    return f"""
<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Pitch Rehearsal Prompter</title>
<style>
:root {{
  --bg: #0f172a;
  --panel: #111827;
  --panel2: #1e293b;
  --text: #f8fafc;
  --muted: #e2e8f0;
  --accent: #60a5fa;
  --line: rgba(255,255,255,.16);
  --ok: #22c55e;
  --warn: #f59e0b;
}}
* {{ box-sizing: border-box; }}
html {{ color-scheme: dark; }}
body {{
  margin: 0;
  font-family: Inter, Aptos, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  background: radial-gradient(circle at top left, rgba(96,165,250,.24), transparent 36%), var(--bg);
  color: var(--text);
}}
.app {{ min-height: 100vh; padding: 18px; }}
.topbar {{
  display: flex; align-items: center; justify-content: space-between; gap: 12px;
  padding: 12px 14px; border: 1px solid var(--line); border-radius: 18px;
  background: rgba(15,23,42,.72); backdrop-filter: blur(12px);
}}
.brand h1 {{ margin: 0; font-size: 18px; line-height: 1.25; }}
.brand p {{ margin: 4px 0 0; color: var(--muted); font-size: 12px; }}
.controls {{ display: flex; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }}
button, select {{
  border: 1px solid rgba(255,255,255,.28); background: #f8fafc; color: #0f172a;
  -webkit-text-fill-color: #0f172a;
  border-radius: 12px; padding: 9px 12px; font-weight: 800; cursor: pointer;
}}
button *, select * {{ color: #0f172a; -webkit-text-fill-color: #0f172a; }}
button.secondary {{ background: rgba(255,255,255,.10); color: var(--text); -webkit-text-fill-color: var(--text); }}
button.danger {{ background: #fee2e2; color: #991b1b; -webkit-text-fill-color: #991b1b; }}
.grid {{ display: grid; grid-template-columns: 1.05fr .95fr; gap: 16px; margin-top: 16px; }}
.card {{ border: 1px solid var(--line); border-radius: 22px; background: rgba(15,23,42,.78); overflow: hidden; }}
.card-header {{ padding: 14px 16px; border-bottom: 1px solid var(--line); display:flex; justify-content:space-between; gap:12px; align-items:center; }}
.eyebrow {{ color: var(--accent); text-transform: uppercase; letter-spacing: .12em; font-size: 11px; font-weight: 900; }}
.slide-title {{ font-size: clamp(28px, 4vw, 54px); line-height: 1.05; font-weight: 950; margin: 0; color: var(--text); -webkit-text-fill-color: var(--text); }}
.slide-screen {{ min-height: 475px; padding: 28px; display: flex; flex-direction: column; justify-content: center; gap: 18px; }}
.screen-body {{ white-space: pre-wrap; font-size: clamp(18px, 2.0vw, 30px); line-height: 1.35; color: var(--text); }}
.meta-row {{ display:flex; gap:8px; flex-wrap: wrap; margin-top: 8px; }}
.pill {{ border:1px solid var(--line); border-radius:999px; padding:5px 9px; color:var(--muted); font-size:12px; }}
.timer {{ font-size: clamp(38px, 7vw, 88px); font-weight: 950; letter-spacing: -.04em; }}
.timer.warn {{ color: var(--warn); }}
.timer.ok {{ color: var(--ok); }}
.progress-wrap {{ height: 10px; background: rgba(255,255,255,.10); border-radius: 999px; overflow: hidden; }}
.progress {{ height: 100%; width: 0%; background: linear-gradient(90deg, var(--accent), var(--ok)); }}
.prompter {{ height: 420px; overflow: hidden; position: relative; padding: 20px 24px; }}
.prompt-scroll {{ position: absolute; left: 24px; right: 24px; top: 20px; transition: transform .12s linear; }}
.prompt-text {{ font-size: clamp(25px, 3.5vw, 44px); line-height: 1.42; font-weight: 800; white-space: pre-wrap; color: var(--text); -webkit-text-fill-color: var(--text); }}
.purpose {{ padding: 0 24px 18px; color: var(--muted); font-size: 14px; line-height: 1.55; }}
.next {{ padding: 14px 16px; color: var(--muted); border-top:1px solid var(--line); font-size: 14px; line-height:1.45; }}
.timeline {{ display:grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap:8px; margin-top:16px; }}
.step {{ border:1px solid var(--line); border-radius: 14px; padding:8px; color: var(--muted); font-size:11px; min-height:54px; }}
.step.active {{ background: rgba(96,165,250,.20); color: var(--text); border-color: rgba(96,165,250,.65); }}
.footer {{ text-align:center; color:var(--muted); font-size:12px; padding: 14px 0 4px; }}
@media (max-width: 920px) {{ .grid {{ grid-template-columns: 1fr; }} .slide-screen {{ min-height: 360px; }} .prompter {{ height: 350px; }} }}
</style>
</head>
<body>
<div class="app" id="app">
  <div class="topbar">
    <div class="brand">
      <h1 id="company"></h1>
      <p id="subtitle"></p>
    </div>
    <div class="controls">
      <button id="startBtn">▶ Start</button>
      <button id="pauseBtn" class="secondary">⏸ Pause</button>
      <button id="prevBtn" class="secondary">‹ Prev</button>
      <button id="nextBtn" class="secondary">Next ›</button>
      <button id="resetBtn" class="danger">Reset</button>
      <button id="fullBtn" class="secondary">Fullscreen</button>
      <select id="speed"><option value="0.85">0.85x</option><option value="1" selected>1x</option><option value="1.15">1.15x</option><option value="1.3">1.3x</option></select>
    </div>
  </div>

  <div class="grid">
    <section class="card">
      <div class="card-header">
        <div><div class="eyebrow">Presentation view</div><h2 class="slide-title" id="slideTitle"></h2></div>
        <div class="timer" id="timer">00:00</div>
      </div>
      <div class="slide-screen">
        <div class="screen-body" id="screenBody"></div>
        <div class="meta-row"><span class="pill" id="slideCount"></span><span class="pill" id="slidePurposeShort"></span></div>
        <div class="progress-wrap"><div class="progress" id="progress"></div></div>
      </div>
    </section>

    <section class="card">
      <div class="card-header"><div><div class="eyebrow">Teleprompter</div><h2 class="slide-title" style="font-size:28px">Skenario bicara</h2></div></div>
      <div class="prompter" id="prompter"><div class="prompt-scroll" id="promptScroll"><div class="prompt-text" id="promptText"></div></div></div>
      <div class="purpose" id="purpose"></div>
      <div class="next" id="nextSlide"></div>
    </section>
  </div>

  <div class="timeline" id="timeline"></div>
  <div class="footer" id="footer"></div>
</div>
<script>
const data = {payload};
let idx = 0;
let elapsed = 0;
let running = false;
let last = null;
let raf = null;
const $ = (id) => document.getElementById(id);
function fmt(sec) {{ sec=Math.max(0, Math.ceil(sec)); return String(Math.floor(sec/60)).padStart(2,'0') + ':' + String(sec%60).padStart(2,'0'); }}
function currentDuration() {{ return Math.max(1, data.items[idx].duration / parseFloat($('speed').value || '1')); }}
function render() {{
  const item = data.items[idx];
  const dur = currentDuration();
  const remain = dur - elapsed;
  $('company').textContent = data.company + ' — Pitch Rehearsal';
  $('subtitle').textContent = data.pitchType + ' • ' + data.duration + ' menit • ' + data.businessModel;
  $('slideTitle').textContent = item.title;
  $('screenBody').textContent = item.screen || '';
  $('slideCount').textContent = 'Slide ' + item.slide + ' / ' + data.items.length;
  $('slidePurposeShort').textContent = item.key;
  $('purpose').textContent = 'Tujuan slide: ' + item.purpose + ' Transisi: ' + item.transition;
  $('promptText').textContent = item.talk || '';
  $('timer').textContent = fmt(remain);
  $('timer').className = 'timer' + (remain <= 10 ? ' warn' : remain > dur * .55 ? ' ok' : '');
  $('progress').style.width = Math.min(100, Math.max(0, elapsed / dur * 100)) + '%';
  const next = data.items[idx + 1];
  $('nextSlide').textContent = next ? 'Berikutnya: ' + next.title + ' — ' + next.purpose : 'Slide terakhir. Tutup dengan next step dan ajakan follow-up.';
  const prompter = $('prompter');
  const scroll = $('promptScroll');
  const maxScroll = Math.max(0, scroll.scrollHeight - prompter.clientHeight + 45);
  scroll.style.transform = 'translateY(' + (-maxScroll * Math.min(1, elapsed / dur)) + 'px)';
  document.querySelectorAll('.step').forEach((el, i) => el.classList.toggle('active', i === idx));
}}
function buildTimeline() {{
  $('timeline').innerHTML = data.items.map((it, i) => '<div class="step" data-i="'+i+'"><b>' + it.slide + '. ' + it.title + '</b><br>' + it.duration + ' detik</div>').join('');
  document.querySelectorAll('.step').forEach(el => el.onclick = () => {{ idx = Number(el.dataset.i); elapsed = 0; render(); }});
}}
function tick(ts) {{
  if (!last) last = ts;
  const delta = (ts - last) / 1000;
  last = ts;
  if (running) {{
    elapsed += delta;
    if (elapsed >= currentDuration()) {{
      if (idx < data.items.length - 1) {{ idx += 1; elapsed = 0; }}
      else {{ elapsed = currentDuration(); running = false; }}
    }}
    render();
  }}
  raf = requestAnimationFrame(tick);
}}
$('startBtn').onclick = () => {{ running = true; last = null; }};
$('pauseBtn').onclick = () => {{ running = false; }};
$('resetBtn').onclick = () => {{ running = false; idx = 0; elapsed = 0; render(); }};
$('prevBtn').onclick = () => {{ idx = Math.max(0, idx - 1); elapsed = 0; render(); }};
$('nextBtn').onclick = () => {{ idx = Math.min(data.items.length - 1, idx + 1); elapsed = 0; render(); }};
$('fullBtn').onclick = () => {{ document.documentElement.requestFullscreen && document.documentElement.requestFullscreen(); }};
$('speed').onchange = () => {{ elapsed = Math.min(elapsed, currentDuration()); render(); }};
document.addEventListener('keydown', (e) => {{
  if (e.code === 'Space') {{ running = !running; e.preventDefault(); }}
  if (e.code === 'ArrowRight') {{ idx = Math.min(data.items.length - 1, idx + 1); elapsed = 0; render(); }}
  if (e.code === 'ArrowLeft') {{ idx = Math.max(0, idx - 1); elapsed = 0; render(); }}
}});
$('footer').textContent = data.developer + ' • Space: start/pause • Arrow keys: prev/next';
buildTimeline(); render(); raf = requestAnimationFrame(tick);
</script>
</body>
</html>
"""


def render_rehearsal_section(data: dict[str, Any]) -> None:
    guide(
        "Simulasi pitching otomatis",
        "Gunakan mode ini untuk latihan seperti teleprompter. Presentation view menampilkan ringkasan slide, sementara panel kanan menjalankan skenario bicara otomatis sesuai timing pitch. Teks teleprompter bisa diedit manual per slide.",
    )
    st.markdown(
        """
        <div class="prompter-frame-note">
            <strong>Kontrol latihan:</strong> klik Start untuk mulai, Pause untuk berhenti, Prev/Next untuk pindah slide, Reset untuk mengulang, dan Fullscreen untuk mode layar penuh. Di keyboard, gunakan Space untuk start/pause dan tombol panah untuk pindah slide.
        </div>
        """,
        unsafe_allow_html=True,
    )

    base_items = sync_prompter_script_defaults(data)

    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        st.checkbox(
            "Gunakan teks teleprompter custom",
            key="use_custom_prompter",
            help="Jika aktif, simulasi, HTML prompter, dan PDF scenario guide memakai teks yang Anda edit di bawah.",
        )
    with c2:
        if st.button("Reset teks dari data deck", use_container_width=True):
            sync_prompter_script_defaults(data, force=True)
            st.success("Teks teleprompter dikembalikan dari data deck.")
            st.rerun()
    with c3:
        st.caption("Edit kalimat agar sesuai gaya bicara presenter. Gunakan bahasa natural, pendek, dan mudah dibaca saat latihan.")

    with st.expander("✍️ Edit teks teleprompter per slide", expanded=False):
        st.markdown(
            """
            <div class="readable-panel">
                <strong>Tips:</strong> tulis seperti Anda berbicara, bukan seperti laporan. Satu slide idealnya berisi 2-5 kalimat utama. Jika pitch pendek, gunakan kalimat yang lebih langsung.
            </div>
            """,
            unsafe_allow_html=True,
        )
        for idx, item in enumerate(base_items):
            st.text_area(
                f"Slide {idx + 1} — {item.get('title', 'Slide')}",
                key=f"prompter_script_{idx}",
                height=135,
                help="Teks ini akan tampil di teleprompter dan ikut tersimpan ke JSON project/export ZIP.",
            )

    live_data = collect_data_preview_only()
    st.markdown("#### Urutan slide dan timing simulasi")
    sim_plan = build_slide_plan(live_data)
    sim_timings = timing_for_plan(sim_plan, int(live_data.get("pitch_duration_minutes", 10)))
    render_slide_timing_table(sim_plan, sim_timings)
    components.html(build_rehearsal_html(live_data, standalone=False), height=900, scrolling=False)

def build_scenario_pdf(data: dict[str, Any]) -> BytesIO:
    out = BytesIO()
    doc = SimpleDocTemplate(out, pagesize=A4, rightMargin=1.4 * cm, leftMargin=1.4 * cm, topMargin=1.2 * cm, bottomMargin=1.2 * cm)
    styles = pdf_styles()
    story = []
    insights = generate_investor_insights(data)
    plan = build_slide_plan(data)
    timings = timing_for_plan(plan, int(data.get("pitch_duration_minutes", 10)))
    profile = PITCH_DURATION_PROFILES[int(data.get("pitch_duration_minutes", 10))]
    type_guide = PITCH_TYPE_GUIDES[data.get("pitch_type", "Investor Seed Round")]
    story += [
        p(f"Pitch Scenario Guide - {data.get('company')}", styles["title"]),
        Spacer(1, 0.25 * cm),
        p(f"{data.get('pitch_type')} • {profile['name']} • {data.get('business_model_type')}", styles["body"]),
        p(f"Focus: {type_guide['focus']}", styles["body"]),
        p(f"Deck readiness: {insights['score']}/100", styles["body"]),
        p(DEVELOPER, styles["small"]),
        Spacer(1, 0.35 * cm),
        p("Ringkasan skema", styles["h1"]),
        p(profile["scheme"], styles["body"]),
        p(profile["slide_goal"], styles["body"]),
    ]
    story.append(make_rl_table([["Slide", "Timing", "Tujuan", "Talk Track", "Transisi"]] + [
        [str(i + 1), f"{timings[i]} detik", p(item["title"], styles["small"]), p(item["purpose"], styles["small"]), p("Lanjutkan dengan bukti yang menguatkan konteks slide berikutnya.", styles["small"])]
        for i, item in enumerate(plan)
    ], [1.1 * cm, 1.7 * cm, 3.4 * cm, 6.0 * cm, 4.1 * cm]))
    story += [PageBreak(), p("Skenario per slide", styles["h1"])]
    rehearsal_items = build_rehearsal_items(data)
    for idx, item in enumerate(plan):
        talk_item = rehearsal_items[idx] if idx < len(rehearsal_items) else {}
        talk = str(talk_item.get("talk") or talk_track_for_slide(item, data, insights))
        story.append(p(f"{idx + 1}. {item['title']} - {timings[idx]} detik", styles["h2"]))
        story.append(p(talk, styles["body"]))
        story.append(p("Transisi: hubungkan slide ini dengan bukti berikutnya, jangan membaca seluruh teks slide.", styles["small"]))
    story += [PageBreak(), p("Pertanyaan Investor & Jawaban Latihan", styles["h1"])]
    qa_rows = [["Pertanyaan", "Cara menjawab"]] + [[p(q, styles["small"]), p(a, styles["small"])] for q, a in insights["qa"]]
    story.append(make_rl_table(qa_rows, [6.0 * cm, 10.0 * cm]))
    story += [PageBreak(), p("Istilah investor dan cara menghitungnya", styles["h1"])]
    gloss_rows = [["Istilah", "Arti", "Rumus", "Contoh"]]
    for _, term, simple, formula, example in GLOSSARY:
        gloss_rows.append([p(term, styles["small"]), p(simple, styles["small"]), p(formula, styles["small"]), p(example, styles["small"])])
    story.append(make_rl_table(gloss_rows, [2.2 * cm, 5.2 * cm, 5.0 * cm, 3.8 * cm]))
    story += [PageBreak(), p("Checklist latihan pitching", styles["h1"])]
    checklist = [
        "Opening dapat disampaikan dalam 15 detik.",
        "Problem dijelaskan dari sudut pandang customer, bukan fitur produk.",
        "Market size punya metode hitung yang bisa dijelaskan.",
        "Model bisnis punya rumus revenue yang jelas.",
        "Traction membuktikan demand, bukan hanya vanity metrics.",
        "Funding ask terhubung dengan runway dan milestone.",
        "Siapkan jawaban CAC, churn, margin, competition, moat, dan financial assumptions.",
    ]
    for item in checklist:
        story.append(p("□ " + item, styles["body"]))
    story.append(Spacer(1, 0.4 * cm))
    story.append(p(DEVELOPER, styles["small"]))
    doc.build(story)
    out.seek(0)
    return out


def build_download_zip(company_name: str, pptx_file: BytesIO, pdf_file: BytesIO, project_json: BytesIO, prompter_html: str | None = None) -> BytesIO:
    base = filename(company_name)
    out = BytesIO()
    pptx_file.seek(0); pdf_file.seek(0); project_json.seek(0)
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(f"{base}-seed-investor-pitch-deck.pptx", pptx_file.read())
        archive.writestr(f"{base}-pitch-scenario-guide.pdf", pdf_file.read())
        archive.writestr(f"{base}-project-data.json", project_json.read())
        if prompter_html:
            archive.writestr(f"{base}-pitch-prompter.html", prompter_html.encode("utf-8"))
        archive.writestr("README.txt", f"Generated by {APP_VERSION}\n{DEVELOPER}\n\nIsi ZIP:\n- PPTX investor pitch deck\n- PDF pitch scenario guide\n- HTML pitch prompter untuk simulasi otomatis/offline dengan teks teleprompter custom\n- JSON project data untuk diedit ulang\n")
    out.seek(0)
    return out


def collect_data_preview_only() -> dict[str, Any]:
    data = {k: st.session_state.get(k) for k in st.session_state.keys() if not k.startswith("FormSubmitter")}
    template = business_model_template(data.get("business_model_type", "SaaS / Subscription"))
    labels = [st.session_state.get(f"model_metric_label_{i}", template["metrics"][i] if i < len(template["metrics"]) else "") for i in range(4)]
    values = [st.session_state.get(f"model_metric_value_{i}", template["defaults"][i] if i < len(template["defaults"]) else "") for i in range(4)]
    data["model_metric_labels"] = labels
    data["model_metric_values"] = values
    comp_count = int(st.session_state.get("competitor_count", 1) or 1)
    data["competitors"] = [
        {
            "name": st.session_state.get(f"competitor_name_{i}", ""),
            "category": st.session_state.get(f"competitor_category_{i}", "Alternative"),
            "weakness": st.session_state.get(f"competitor_weakness_{i}", ""),
            "advantage": st.session_state.get(f"competitor_advantage_{i}", ""),
        }
        for i in range(comp_count)
        if str(st.session_state.get(f"competitor_name_{i}", "")).strip()
    ]
    ms_count = int(st.session_state.get("milestone_count", 1) or 1)
    data["milestones"] = [
        {
            "period": st.session_state.get(f"milestone_period_{i}", ""),
            "target": st.session_state.get(f"milestone_target_{i}", ""),
            "metric": st.session_state.get(f"milestone_metric_{i}", ""),
            "owner": st.session_state.get(f"milestone_owner_{i}", ""),
        }
        for i in range(ms_count)
        if str(st.session_state.get(f"milestone_target_{i}", "")).strip()
    ]
    data["round"] = data.get("round_name", data.get("round", ""))
    slide_count = int(st.session_state.get("prompter_slide_count", 0) or 0)
    if slide_count:
        data["custom_prompter_scripts"] = [
            st.session_state.get(f"prompter_script_{i}", "")
            for i in range(slide_count)
        ]
    else:
        data["custom_prompter_scripts"] = []
    data["use_custom_prompter"] = bool(st.session_state.get("use_custom_prompter", True))
    logo = get_uploaded_logo_bytes()
    if logo:
        data["logo_bytes"] = logo
    return data


def update_model_defaults_if_needed() -> None:
    model = st.session_state.get("business_model_type", "SaaS / Subscription")
    last = st.session_state.get("_last_model_type")
    if last != model:
        template = business_model_template(model)
        st.session_state["business_model"] = template["lines"]
        for idx in range(4):
            st.session_state[f"model_metric_label_{idx}"] = template["metrics"][idx]
            st.session_state[f"model_metric_value_{idx}"] = template["defaults"][idx]
        st.session_state["_last_model_type"] = model


def render_identity_section() -> None:
    guide("Identitas & konteks pitch", "Isi bagian ini untuk menentukan siapa yang pitching, jenis pitch, bahasa output, durasi, dan format deck yang akan dihasilkan.")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.text_input("Nama startup", key="company", help="Nama startup yang akan muncul di cover, footer, dan nama file.")
        st.text_input("One-liner", key="one_liner", help="Format ideal: membantu [target customer] menyelesaikan [problem] dengan [solusi] sehingga [hasil].")
        st.text_input("Presenter", key="presenter", help="Nama founder/presenter utama.")
        st.text_input("Kontak", key="contact", help="Email/WhatsApp/LinkedIn untuk follow-up investor.")
    with c2:
        st.text_input("Round", key="round_name", help="Contoh: Pre-Seed, Seed, Bridge, Grant, Partnership Pilot.")
        st.number_input("Jumlah pendanaan", min_value=0, step=50_000_000, key="ask", help="Dana yang diminta. Harus terhubung ke runway dan milestone.")
        st.selectbox("Jenis pitch", PITCH_TYPE_OPTIONS, key="pitch_type", help="Jenis pitch mengubah tekanan narasi di PPT dan PDF.")
        st.selectbox("Durasi pitching", PITCH_DURATION_OPTIONS, key="pitch_duration_minutes", help="Jumlah slide dan skenario PDF menyesuaikan durasi.")
    with c3:
        st.selectbox("Bahasa output", LANGUAGE_OPTIONS, key="output_language", help="Mengubah label slide dan PDF. Input bebas tetap memakai bahasa yang Anda tulis.")
        st.selectbox("Mode pengisian", ["Tabs", "Wizard"], key="ui_mode", help="Wizard cocok untuk founder baru; Tabs cocok untuk editing cepat.")
        st.checkbox("Tambahkan slide Investor Readiness", key="include_insight_slide", help="Slide analisa otomatis ditambahkan pada pitch 8 menit ke atas.")
        st.file_uploader("Upload logo startup", type=["png", "jpg", "jpeg"], key="logo_file", help="Logo muncul di cover dan header slide.")
    profile = PITCH_DURATION_PROFILES[int(st.session_state.pitch_duration_minutes)]
    type_guide = PITCH_TYPE_GUIDES[st.session_state.pitch_type]
    st.markdown(f"<div class='readable-panel'><p><strong>Skema durasi:</strong> {profile['scheme']}</p><p><strong>Fokus jenis pitch:</strong> {type_guide['focus']}</p></div>", unsafe_allow_html=True)


def render_brand_section() -> None:
    guide("Brand kit", "Atur warna dan style deck. Desain PPT tetap VC-grade, tetapi warna dan logo mengikuti brand startup.")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.color_picker("Warna aksen", key="accent_color", help="Dipakai untuk headline kecil, metric, table header, dan callout.")
        st.color_picker("Warna cover/dark", key="secondary_color", help="Dipakai untuk cover dan closing slide.")
    with c2:
        st.selectbox("Style deck", ["Minimal VC", "Corporate", "Bold Startup", "Clean Academic", "Dark Premium"], key="deck_style", help="Saat ini mempengaruhi tone desain dasar dan bisa diperluas ke template visual berbeda.")
        st.selectbox("Font style", ["Modern Sans", "Classic Business", "Compact"], key="font_style", help="PPT memakai font umum agar aman dibuka di berbagai perangkat.")
    with c3:
        st.selectbox("Mata uang", ["Rp", "USD"], key="currency", help="Dipakai untuk format pendanaan dan proyeksi finansial.")


def render_story_section() -> None:
    guide("Problem, solution, dan product", "Jangan mulai dari fitur. Mulai dari problem customer, bukti problem, lalu jelaskan solusi dan perubahan yang terjadi setelah produk digunakan.")
    c1, c2 = st.columns(2)
    with c1:
        st.text_area("Problem", height=125, key="problem", help="Minimal 2-3 poin. Jelaskan siapa yang bermasalah, apa masalahnya, seberapa sering, dan biaya masalahnya.")
        st.text_area("Evidence problem", height=90, key="problem_evidence", help="Bukti: interview, survei, pilot, data operasional, waiting list, atau pain cost.")
        st.text_area("Solution", height=125, key="solution", help="Jelaskan bagaimana produk menyelesaikan problem. Hindari kalimat terlalu generik seperti 'platform inovatif'.")
        st.text_area("Value proposition", height=90, key="value_prop", help="Manfaat utama dalam bahasa customer: hemat waktu, naik revenue, turun biaya, lebih patuh, lebih cepat.")
    with c2:
        st.text_area("Cara kerja produk", height=90, key="product_flow", help="Flow sederhana: input -> proses -> output -> business outcome.")
        st.text_area("Benefit produk", height=90, key="product_benefit", help="Hasil yang dapat dilihat customer setelah menggunakan produk.")
        st.text_area("Fitur utama", height=125, key="features", help="Maksimal 5 fitur untuk slide; prioritaskan fitur yang membuktikan value proposition.")
        st.file_uploader("Screenshot/mockup produk", type=["png", "jpg", "jpeg"], key="product_image", help="Opsional. Akan muncul di slide Product.")


def render_market_model_section() -> None:
    update_model_defaults_if_needed()
    guide("Market & model bisnis", "TAM/SAM/SOM menjelaskan peluang pasar. Model bisnis menjelaskan bagaimana demand berubah menjadi revenue dan margin.")
    c1, c2 = st.columns(2)
    with c1:
        st.text_input("TAM", key="tam", help="Total pasar maksimum. Cara hitung: total target customer x potensi belanja tahunan.")
        st.text_input("SAM", key="sam", help="Bagian TAM yang bisa dilayani oleh produk/region/channel saat ini.")
        st.text_input("SOM", key="som", help="Bagian pasar yang realistis direbut dalam 2-3 tahun.")
        st.text_area("Market notes", height=115, key="market_notes", help="Jelaskan segmen awal, wedge, beachhead, dan alasan pasar ini bisa berkembang.")
        st.text_input("ARPU", key="arpu", help="Average Revenue Per User = revenue / jumlah customer atau user.")
        st.text_input("Gross Margin", key="gross_margin", help="(Revenue - COGS) / Revenue x 100%")
        st.text_input("CAC", key="cac", help="Biaya sales & marketing / customer baru.")
        st.text_input("Payback", key="payback", help="CAC / gross profit bulanan per customer.")
    with c2:
        st.selectbox("Model bisnis", list(BUSINESS_MODEL_TEMPLATES.keys()), key="business_model_type", help="Aplikasi menyesuaikan metrik, insight, dan PDF berdasarkan model bisnis.")
        template = business_model_template(st.session_state.business_model_type)
        st.markdown(f"<div class='readable-panel'><p><strong>Model:</strong> {template['description']}</p><p><strong>Cocok untuk:</strong> {template['best_for']}</p><p><strong>Rumus utama:</strong> {template['formula']}</p></div>", unsafe_allow_html=True)
        st.text_area("Narasi model bisnis", height=130, key="business_model", help="Jelaskan siapa yang membayar, kapan membayar, pricing, margin, dan expansion revenue.")
        st.markdown("#### Metrik utama model bisnis")
        for idx in range(4):
            m1, m2 = st.columns([1, 1])
            with m1:
                st.text_input(f"Label metrik {idx + 1}", key=f"model_metric_label_{idx}")
            with m2:
                st.text_input(f"Nilai metrik {idx + 1}", key=f"model_metric_value_{idx}")


def render_traction_gtm_section() -> None:
    guide("Traction & Go-To-Market", "Traction membuktikan demand. GTM menjelaskan bagaimana startup mendapatkan customer secara berulang dan efisien.")
    c1, c2 = st.columns(2)
    with c1:
        st.text_input("Users/customers", key="users", help="Jumlah user/customer aktif, paid user, pilot, atau merchant.")
        st.text_input("Revenue/GMV", key="revenue", help="MRR, ARR, GMV, transaction volume, pipeline, atau revenue aktual.")
        st.text_input("Growth", key="growth", help="MoM growth, QoQ growth, pipeline growth, atau transaction growth.")
        st.text_input("Retention", key="retention", help="Retention D30, logo retention, revenue retention, repeat purchase, atau repeat transaction.")
        st.text_area("Traction notes", height=110, key="traction_notes", help="Bukti demand: pilot berbayar, LOI, pipeline, retention, revenue, active usage, referral.")
    with c2:
        st.text_area("ICP", height=80, key="icp", help="Ideal Customer Profile: customer awal yang paling sakit masalahnya dan paling mungkin bayar.")
        st.text_area("Channel utama", height=80, key="channel", help="Channel akuisisi utama: komunitas, partnership, sales, content, marketplace, referral, outbound.")
        st.text_area("GTM", height=135, key="gtm", help="Jelaskan funnel: lead source, conversion, sales motion, activation, retention, referral.")


def render_finance_milestone_section() -> None:
    guide("Financial, funding ask, dan milestone", "Investor seed tidak menuntut akurasi sempurna, tetapi ingin melihat asumsi logis, runway, dan milestone yang bisa diverifikasi.")
    c1, c2 = st.columns(2)
    with c1:
        st.number_input("Revenue Y1", step=100_000_000, key="rev1", help="Revenue proyeksi tahun pertama.")
        st.number_input("Revenue Y2", step=100_000_000, key="rev2")
        st.number_input("Revenue Y3", step=100_000_000, key="rev3")
        st.number_input("Cost Y1", step=100_000_000, key="cost1", help="Operating cost total tahun pertama.")
        st.number_input("Cost Y2", step=100_000_000, key="cost2")
        st.number_input("Cost Y3", step=100_000_000, key="cost3")
        st.number_input("EBITDA/Profit Y1", step=100_000_000, key="profit1")
        st.number_input("EBITDA/Profit Y2", step=100_000_000, key="profit2")
        st.number_input("EBITDA/Profit Y3", step=100_000_000, key="profit3")
    with c2:
        st.number_input("Runway / bulan", min_value=1, key="runway", help="Berapa bulan dana akan membiayai operasional.")
        st.text_input("Milestone utama funding", key="milestone", help="Target utama setelah dana masuk. Contoh: Rp 500 juta MRR, 8.000 active users, CAC payback < 3 bulan.")
        st.text_area("Use of funds", height=110, key="use_of_funds", help="Pisahkan penggunaan dana: product, growth, hiring, ops, legal, customer success.")
        st.text_area("Next round logic", height=80, key="next_round", help="Kapan dan dengan bukti apa startup siap raise round berikutnya.")
        st.number_input("Jumlah milestone detail", min_value=1, max_value=8, step=1, key="milestone_count", help="PPT otomatis memecah slide jika milestone banyak.")
    st.markdown("### Detail milestone")
    for idx in range(int(st.session_state.milestone_count)):
        with st.expander(f"Milestone {idx + 1}", expanded=idx < 3):
            a, b, c, d = st.columns(4)
            with a:
                st.text_input("Periode", key=f"milestone_period_{idx}", help="Contoh: 0-3 bulan, Q1, 12 bulan.")
            with b:
                st.text_input("Target", key=f"milestone_target_{idx}", help="Outcome yang ingin dicapai.")
            with c:
                st.text_input("Success metric", key=f"milestone_metric_{idx}", help="Ukuran keberhasilan yang objektif.")
            with d:
                st.text_input("Owner", key=f"milestone_owner_{idx}", help="Penanggung jawab utama.")


def render_competition_team_section() -> None:
    guide("Competition & team", "Kompetitor tidak harus produk yang sama. Sertakan status quo, spreadsheet, agency, manual workflow, dan adjacent tool.")
    c1, c2 = st.columns(2)
    with c1:
        st.text_area("Ringkasan kompetisi / narrative advantage", height=90, key="competition_summary", help="Satu kalimat: kenapa kita bisa menang dibanding alternatif yang sudah dipakai customer.")
        st.number_input("Jumlah kompetitor / alternatif", min_value=1, max_value=10, step=1, key="competitor_count", help="PPT otomatis memecah slide jika kompetitor banyak.")
    with c2:
        st.text_area("Team", height=110, key="team", help="Nama, role, pengalaman relevan, achievement, domain expertise.")
        st.text_area("Founder-market fit", height=90, key="founder_fit", help="Kenapa tim ini punya unfair advantage untuk menyelesaikan problem ini.")
        st.text_input("Closing line", key="closing", help="Kalimat penutup yang merangkum visi besar startup.")
    categories = ["Direct competitor", "Indirect competitor", "Status quo", "Adjacent tool", "Alternative"]
    st.markdown("### Peta kompetitor")
    for idx in range(int(st.session_state.competitor_count)):
        with st.expander(f"Kompetitor / alternatif {idx + 1}", expanded=idx < 3):
            a, b = st.columns(2)
            with a:
                st.text_input("Nama", key=f"competitor_name_{idx}", help="Produk, perusahaan, workflow manual, spreadsheet, agency, atau status quo.")
                current = st.session_state.get(f"competitor_category_{idx}", "Alternative")
                index = categories.index(current) if current in categories else 4
                st.selectbox("Kategori", categories, index=index, key=f"competitor_category_{idx}")
            with b:
                st.text_area("Kelemahan alternatif", height=80, key=f"competitor_weakness_{idx}", help="Kelemahan dari sudut pandang customer.")
                st.text_area("Keunggulan kita", height=80, key=f"competitor_advantage_{idx}", help="Keunggulan spesifik: UX, cost, data, speed, distribution, switching workflow, dsb.")


def render_preview_section(data: dict[str, Any]) -> None:
    guide("Preview sebelum generate", "Cek ringkasan ini sebelum download. Sistem akan membuat satu ZIP berisi PPTX, PDF scenario guide, dan JSON project data.")
    plan = build_slide_plan(data)
    timings = timing_for_plan(plan, int(data.get("pitch_duration_minutes", 10)))
    insights = generate_investor_insights(data)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Durasi", f"{data.get('pitch_duration_minutes')} menit")
    c2.metric("Jumlah slide", len(plan))
    c3.metric("Kompetitor", len(data.get("competitors", [])))
    c4.metric("Milestone", len(data.get("milestones", [])))
    st.markdown("#### Urutan slide dan timing")
    render_slide_timing_table(plan, timings)
    render_validation_panel(insights["issues"])



def render_slide_timing_table(plan: list[dict[str, Any]], timings: list[int]) -> None:
    """Render slide order/timing inside an isolated iframe.

    The table is rendered with ``components.html`` instead of ``st.markdown`` so
    global Streamlit theme CSS cannot override the table's foreground/background
    colors. This prevents invisible text when the app theme, browser theme, or
    previous CSS rules use the same color for text and background.
    """
    if not plan:
        st.warning("Belum ada skema slide yang dapat ditampilkan. Pilih durasi pitch dan jenis pitch terlebih dahulu.")
        return

    safe_timings = list(timings or [])
    if len(safe_timings) < len(plan):
        safe_timings = safe_timings + [0] * (len(plan) - len(safe_timings))

    total_seconds = sum(max(0, int(x or 0)) for x in safe_timings[:len(plan)])
    total_minutes = total_seconds / 60 if total_seconds else 0

    rows_html = []
    for i, item in enumerate(plan):
        title = html.escape(str(item.get("title", f"Slide {i + 1}")))
        purpose = html.escape(str(item.get("purpose", "")))
        key = html.escape(str(item.get("key", "")))
        seconds = int(safe_timings[i] or 0)
        row_bg = "#ffffff" if i % 2 == 0 else "#f8fafc"
        rows_html.append(
            f"""
            <tr style="background:{row_bg};">
                <td class="num">{i + 1}</td>
                <td class="title-cell"><strong>{title}</strong><br><span>{key}</span></td>
                <td class="time-cell">{seconds} detik</td>
                <td class="purpose-cell">{purpose}</td>
            </tr>
            """
        )

    iframe_html = f"""
    <!doctype html>
    <html lang="id">
    <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <style>
            :root {{
                color-scheme: light;
                --bg: #ffffff;
                --surface: #f8fafc;
                --ink: #0f172a;
                --muted: #334155;
                --line: #cbd5e1;
                --line-soft: #e2e8f0;
                --primary: #1d4ed8;
                --header: #0f172a;
                --header-text: #ffffff;
            }}
            * {{ box-sizing: border-box; }}
            html, body {{
                margin: 0;
                padding: 0;
                background: transparent;
                color: var(--ink);
                font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
                -webkit-text-fill-color: var(--ink);
            }}
            .summary {{
                display: flex;
                flex-wrap: wrap;
                gap: 10px;
                margin: 0 0 12px 0;
            }}
            .pill {{
                display: inline-flex;
                align-items: center;
                min-height: 34px;
                background: var(--bg);
                color: var(--ink);
                -webkit-text-fill-color: var(--ink);
                border: 1px solid #94a3b8;
                border-radius: 999px;
                padding: 7px 12px;
                font-size: 13px;
                font-weight: 850;
                line-height: 1.2;
            }}
            .wrap {{
                overflow-x: auto;
                overflow-y: hidden;
                background: var(--bg);
                border: 1px solid var(--line);
                border-radius: 16px;
                box-shadow: 0 8px 22px rgba(15, 23, 42, .08);
            }}
            table {{
                width: 100%;
                min-width: 820px;
                border-collapse: collapse;
                background: var(--bg);
                color: var(--ink);
                -webkit-text-fill-color: var(--ink);
                font-size: 14px;
                line-height: 1.45;
            }}
            thead th {{
                background: var(--header);
                color: var(--header-text);
                -webkit-text-fill-color: var(--header-text);
                text-align: left;
                padding: 12px;
                font-size: 12px;
                text-transform: uppercase;
                letter-spacing: .04em;
                border-bottom: 1px solid #334155;
                font-weight: 900;
            }}
            tbody td {{
                color: var(--ink);
                -webkit-text-fill-color: var(--ink);
                padding: 12px;
                vertical-align: top;
                border-bottom: 1px solid var(--line-soft);
                white-space: normal;
                overflow-wrap: anywhere;
                word-break: normal;
            }}
            tbody tr:last-child td {{ border-bottom: 0; }}
            .num {{
                width: 64px;
                text-align: center;
                font-weight: 950;
                color: var(--primary);
                -webkit-text-fill-color: var(--primary);
            }}
            .title-cell strong {{
                display: inline-block;
                color: var(--ink);
                -webkit-text-fill-color: var(--ink);
                font-weight: 950;
                margin-bottom: 4px;
            }}
            .title-cell span {{
                color: var(--muted);
                -webkit-text-fill-color: var(--muted);
                font-size: 12px;
                font-weight: 750;
            }}
            .time-cell {{
                width: 124px;
                white-space: nowrap;
                font-weight: 950;
                color: var(--ink);
                -webkit-text-fill-color: var(--ink);
            }}
            .purpose-cell {{
                color: var(--ink);
                -webkit-text-fill-color: var(--ink);
            }}
            @media (max-width: 760px) {{
                table {{ min-width: 720px; font-size: 13px; }}
                thead th, tbody td {{ padding: 10px; }}
            }}
        </style>
    </head>
    <body>
        <div class="summary" aria-label="Ringkasan timing slide">
            <div class="pill">Total slide: {len(plan)}</div>
            <div class="pill">Total timing: {total_seconds} detik</div>
            <div class="pill">≈ {total_minutes:.1f} menit</div>
        </div>
        <div class="wrap">
            <table aria-label="Urutan slide dan timing">
                <thead>
                    <tr>
                        <th>Slide</th>
                        <th>Judul</th>
                        <th>Timing</th>
                        <th>Tujuan</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(rows_html)}
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """

    height = min(720, max(260, 122 + len(plan) * 58))
    components.html(iframe_html, height=height, scrolling=True)


def render_save_load(data: dict[str, Any]) -> None:
    with st.sidebar:
        st.header("Project")
        uploaded = st.file_uploader("Muat project JSON", type=["json"], help="Upload file JSON yang pernah diekspor dari aplikasi ini.")
        if uploaded and st.button("Load project", use_container_width=True):
            load_project_from_json(uploaded)
        project_json = build_project_json(data)
        st.download_button(
            "💾 Download Project JSON",
            data=project_json,
            file_name=f"{filename(data.get('company', 'startup'))}-project-data.json",
            mime="application/json",
            use_container_width=True,
        )
        if st.button("Reset ke contoh awal", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            initialize_defaults()
            st.rerun()


def main() -> None:
    initialize_defaults()
    st.title("📊 Seed Investor Pitch Deck Generator")
    st.caption(f"{APP_VERSION} • {DEVELOPER}")
    st.markdown("Membuat paket pitching sekali download: PPTX investor deck, PDF scenario guide, HTML prompter simulasi otomatis, dan JSON project data.")
    current_data = collect_data_preview_only()
    render_save_load(current_data)
    with st.sidebar:
        st.header("Ringkasan")
        st.caption(f"Durasi: {st.session_state.pitch_duration_minutes} menit")
        st.caption(f"Model bisnis: {st.session_state.business_model_type}")
        st.caption(f"Jenis pitch: {st.session_state.pitch_type}")
        st.caption(f"Bahasa: {st.session_state.output_language}")
    update_model_defaults_if_needed()
    section_names = ["Identitas", "Brand", "Story", "Market & Model", "Traction & GTM", "Financial & Milestone", "Competition & Team", "Kalkulator", "Analisa", "Istilah", "Simulasi", "Preview"]
    if st.session_state.ui_mode == "Wizard":
        chosen = st.selectbox("Langkah pengisian", section_names, help="Gunakan Wizard untuk mengisi bertahap.")
        tabs = None
        active_sections = [chosen]
    else:
        tabs = st.tabs(section_names)
        active_sections = section_names
    def section_container(name):
        if tabs is None:
            return st.container()
        return tabs[section_names.index(name)]
    with section_container("Identitas"):
        if "Identitas" in active_sections:
            render_identity_section()
    with section_container("Brand"):
        if "Brand" in active_sections:
            render_brand_section()
    with section_container("Story"):
        if "Story" in active_sections:
            render_story_section()
    with section_container("Market & Model"):
        if "Market & Model" in active_sections:
            render_market_model_section()
    with section_container("Traction & GTM"):
        if "Traction & GTM" in active_sections:
            render_traction_gtm_section()
    with section_container("Financial & Milestone"):
        if "Financial & Milestone" in active_sections:
            render_finance_milestone_section()
    with section_container("Competition & Team"):
        if "Competition & Team" in active_sections:
            render_competition_team_section()
    with section_container("Kalkulator"):
        if "Kalkulator" in active_sections:
            render_calculator()
    data = collect_data_preview_only()
    insights = generate_investor_insights(data)
    with section_container("Analisa"):
        if "Analisa" in active_sections:
            render_scorecards(insights)
            st.markdown("#### Validasi input")
            render_validation_panel(insights["issues"])
            st.markdown("#### Q&A investor")
            render_investor_qa(insights["qa"])
    with section_container("Istilah"):
        if "Istilah" in active_sections:
            render_glossary()
    with section_container("Simulasi"):
        if "Simulasi" in active_sections:
            render_rehearsal_section(data)
    with section_container("Preview"):
        if "Preview" in active_sections:
            render_preview_section(data)
    st.divider()
    final_data = collect_data_preview_only()
    col_generate, col_hint = st.columns([1, 2])
    with col_generate:
        generate = st.button("Generate Pitching Package", type="primary", use_container_width=True)
    with col_hint:
        st.caption("Output satu ZIP berisi PPTX, PDF Scenario Guide, HTML Prompter, dan JSON project data. Gunakan tab Simulasi untuk latihan sebelum pitching.")
    if generate:
        if not str(final_data.get("company", "")).strip():
            st.error("Nama startup wajib diisi.")
            st.stop()
        product_file = st.session_state.get("product_image")
        image_buffer = BytesIO(product_file.getvalue()) if product_file is not None else None
        pptx = build_deck(final_data, image_buffer)
        pdf = build_scenario_pdf(final_data)
        project_json = build_project_json(final_data)
        prompter_html = build_rehearsal_html(final_data)
        package = build_download_zip(final_data.get("company", "startup"), pptx, pdf, project_json, prompter_html)
        st.success("Paket pitching berhasil dibuat dalam satu file ZIP, termasuk HTML prompter untuk simulasi otomatis.")
        render_scorecards(generate_investor_insights(final_data))
        st.download_button(
            "📦 Download PPTX + PDF + Prompter + JSON (.zip)",
            data=package,
            file_name=f"{filename(final_data.get('company', 'startup'))}-pitching-package.zip",
            mime="application/zip",
            use_container_width=True,
        )
    footer_ui()


if __name__ == "__main__":
    main()
