import streamlit as st
import sqlite3
import requests
import os
import io
import re
import difflib
import html
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    KeepTogether,
)
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

import openpyxl


# ============================================================
# CONFIGURATION
# ============================================================

APP_TITLE = "Dar Makkah International - Invoice System"

DB_FILE = "invoices.db"

SHOP_NAME = "Dar Makkah International"

BANK_NAME = "Lloyds"
BANK_ACCOUNT_NAME = "Dar Makkah Intl"
SORT_CODE = "30-99-50"
ACCOUNT_NUMBER = "67944560"

GOOGLE_DRIVE_FOLDER_NAME = "Dar Makkah Invoices"

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/drive"
]

NAV_CREATE = "Create Invoice"
NAV_INVOICES = "Invoices"
NAV_EXCEL = "Excel Export"


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🕌",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# PROFESSIONAL ISLAMIC DASHBOARD CSS
# ============================================================

st.markdown(
    """
    <style>

    /* -------------------------------------------------------
       GLOBAL
    ------------------------------------------------------- */

    .stApp {
        background:
            radial-gradient(
                circle at 10% 10%,
                rgba(197, 164, 91, 0.08),
                transparent 28%
            ),
            radial-gradient(
                circle at 90% 15%,
                rgba(23, 83, 70, 0.10),
                transparent 30%
            ),
            linear-gradient(
                135deg,
                #f8f7f2 0%,
                #f3f5f1 50%,
                #faf9f5 100%
            );
    }

    .main .block-container {
        padding-top: 1.8rem;
        padding-bottom: 3rem;
        max-width: 1450px;
    }

    /* -------------------------------------------------------
       HEADER
    ------------------------------------------------------- */

    .islamic-header {
        position: relative;
        overflow: hidden;
        padding: 28px 34px;
        border-radius: 18px;
        margin-bottom: 22px;
        color: white;

        background:
            linear-gradient(
                135deg,
                #123f35 0%,
                #185846 48%,
                #0f332d 100%
            );

        box-shadow:
            0 10px 30px rgba(15, 51, 45, 0.18);
    }

    .islamic-header:before {
        content: "۞";
        position: absolute;
        right: 28px;
        top: -22px;
        font-size: 150px;
        color: rgba(211, 177, 91, 0.10);
        font-family: serif;
    }

    .islamic-header:after {
        content: "";
        position: absolute;
        left: 0;
        right: 0;
        bottom: 0;
        height: 4px;
        background: #d5b35d;
    }

    .shop-title {
        font-size: 31px;
        font-weight: 750;
        letter-spacing: 0.2px;
        margin: 0;
    }

    .shop-subtitle {
        margin-top: 5px;
        color: #e9dfc3;
        font-size: 14px;
    }

    .arabic-decoration {
        margin-top: 12px;
        font-family: serif;
        font-size: 19px;
        color: #d9bc72;
        letter-spacing: 2px;
    }

    /* -------------------------------------------------------
       SIDEBAR
    ------------------------------------------------------- */

    section[data-testid="stSidebar"] {
        background:
            linear-gradient(
                180deg,
                #103d34 0%,
                #123f35 55%,
                #0c3029 100%
            );
    }

    section[data-testid="stSidebar"] * {
        color: #f6f3e8;
    }

    section[data-testid="stSidebar"] .stRadio label {
        color: #f7f4e9 !important;
    }

    section[data-testid="stSidebar"] hr {
        border-color: rgba(218, 188, 113, 0.35);
    }

    .sidebar-brand {
        text-align: center;
        padding: 12px 5px 20px 5px;
    }

    .sidebar-icon {
        font-size: 40px;
    }

    .sidebar-title {
        font-size: 18px;
        font-weight: 700;
        color: #dfc47e;
    }

    /* -------------------------------------------------------
       CARDS
    ------------------------------------------------------- */

    .dashboard-card {
        background: rgba(255,255,255,0.90);
        border: 1px solid #e1dfd3;
        border-radius: 16px;
        padding: 20px;
        box-shadow:
            0 5px 18px rgba(24, 45, 40, 0.06);
        margin-bottom: 15px;
    }

    .section-card {
        background: rgba(255,255,255,0.94);
        border: 1px solid #e3e0d4;
        border-radius: 14px;
        padding: 18px;
        margin-bottom: 18px;
    }

    .gold-line {
        height: 3px;
        width: 75px;
        background: #c7a653;
        border-radius: 10px;
        margin: 5px 0 17px 0;
    }

    .page-heading {
        color: #164c3e;
        font-size: 26px;
        font-weight: 750;
        margin-bottom: 4px;
    }

    .page-description {
        color: #6a706c;
        font-size: 14px;
        margin-bottom: 15px;
    }

    /* -------------------------------------------------------
       METRICS
    ------------------------------------------------------- */

    [data-testid="stMetric"] {
        background: rgba(255,255,255,0.94);
        border: 1px solid #dfddd2;
        border-radius: 14px;
        padding: 14px;
        box-shadow: 0 4px 14px rgba(20, 50, 40, 0.05);
    }

    [data-testid="stMetricLabel"] {
        color: #66716c !important;
    }

    [data-testid="stMetricValue"] {
        color: #174c3e !important;
    }

    /* -------------------------------------------------------
       BUTTONS
    ------------------------------------------------------- */

    .stButton > button {
        border-radius: 9px;
        font-weight: 600;
    }

    .stButton > button[kind="primary"] {
        background:
            linear-gradient(
                135deg,
                #185b49,
                #123e34
            );
        border: 1px solid #123e34;
    }

    /* -------------------------------------------------------
       EXPANDERS
    ------------------------------------------------------- */

    .streamlit-expanderHeader {
        border-radius: 10px;
        font-weight: 650;
    }

    /* -------------------------------------------------------
       TABLE
    ------------------------------------------------------- */

    .product-result {
        background: white;
        border: 1px solid #e1dfd4;
        border-left: 4px solid #c8a957;
        border-radius: 10px;
        padding: 10px 13px;
        margin-bottom: 7px;
    }

    .product-name {
        color: #164c3e;
        font-weight: 650;
    }

    .product-meta {
        color: #777;
        font-size: 12px;
    }

    /* -------------------------------------------------------
       FOOTER
    ------------------------------------------------------- */

    .professional-footer {
        text-align: center;
        color: #747871;
        padding: 20px;
        font-size: 12px;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
    <div class="islamic-header">
        <div class="shop-title">Dar Makkah International</div>
        <div class="shop-subtitle">
            Professional Sales Invoice Management System
        </div>
        <div class="arabic-decoration">
            بسم الله الرحمن الرحيم
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# DATABASE
# ============================================================

def get_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def column_exists(cursor, table_name, column_name):
    cursor.execute(f"PRAGMA table_info({table_name})")
    columns = [row[1] for row in cursor.fetchall()]
    return column_name in columns


def init_database():

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_number TEXT UNIQUE,
            invoice_date TEXT,
            customer_name TEXT,
            customer_address TEXT,
            created_by TEXT,

            payment_method TEXT,
            cash_received REAL DEFAULT 0,
            returned_amount REAL DEFAULT 0,
            delivery_charges REAL DEFAULT 0,
            total_discount REAL DEFAULT 0,
            subtotal REAL DEFAULT 0,
            grand_total REAL DEFAULT 0,
            total_paid REAL DEFAULT 0,
            total_due REAL DEFAULT 0,

            remarks TEXT,

            google_drive_id TEXT,
            google_drive_name TEXT,

            created_at TEXT,
            updated_at TEXT
        )
        """
    )

    conn.commit()

    required_columns = {
        "invoice_number": "TEXT",
        "invoice_date": "TEXT",
        "customer_name": "TEXT",
        "customer_address": "TEXT",
        "created_by": "TEXT",
        "payment_method": "TEXT",
        "cash_received": "REAL DEFAULT 0",
        "returned_amount": "REAL DEFAULT 0",
        "delivery_charges": "REAL DEFAULT 0",
        "total_discount": "REAL DEFAULT 0",
        "subtotal": "REAL DEFAULT 0",
        "grand_total": "REAL DEFAULT 0",
        "total_paid": "REAL DEFAULT 0",
        "total_due": "REAL DEFAULT 0",
        "remarks": "TEXT",
        "google_drive_id": "TEXT",
        "google_drive_name": "TEXT",
        "created_at": "TEXT",
        "updated_at": "TEXT",
    }

    for column, definition in required_columns.items():

        if not column_exists(
            cursor,
            "invoices",
            column
        ):

            try:

                cursor.execute(
                    f"""
                    ALTER TABLE invoices
                    ADD COLUMN {column} {definition}
                    """
                )

            except Exception:
                pass

    conn.commit()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS invoice_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_id INTEGER,
            serial_number INTEGER,

            product_id TEXT,
            sku TEXT,
            item_description TEXT,

            quantity REAL DEFAULT 1,
            units TEXT,
            unit_price REAL DEFAULT 0,
            discount_percent REAL DEFAULT 0,
            discount_amount REAL DEFAULT 0,
            discounted_unit_price REAL DEFAULT 0,
            total_price REAL DEFAULT 0,

            FOREIGN KEY(invoice_id)
                REFERENCES invoices(id)
        )
        """
    )

    conn.commit()
    conn.close()


init_database()


# ============================================================
# WOOCOMMERCE SETTINGS
# ============================================================

def get_woocommerce_settings():

    try:

        url = st.secrets["woocommerce"]["url"]
        key = st.secrets["woocommerce"]["consumer_key"]
        secret = st.secrets["woocommerce"]["consumer_secret"]

        return (
            url.rstrip("/"),
            key,
            secret
        )

    except Exception:

        st.error(
            "WooCommerce secrets could not be loaded."
        )

        st.code(
            """
[woocommerce]
url = "https://fjbookstore.co.uk"
consumer_key = "ck_YOUR_KEY"
consumer_secret = "cs_YOUR_SECRET"
            """
        )

        st.stop()


STORE_URL, WC_KEY, WC_SECRET = get_woocommerce_settings()


# ============================================================
# WOOCOMMERCE LIVE SEARCH
# ============================================================

@st.cache_data(
    ttl=10,
    show_spinner=False
)
def woo_search_raw(search_text):

    if not search_text:
        return []

    search_text = search_text.strip()

    if len(search_text) < 2:
        return []

    try:

        response = requests.get(
            f"{STORE_URL}/wp-json/wc/v3/products",
            auth=(
                WC_KEY,
                WC_SECRET
            ),
            params={
                "search": search_text,
                "per_page": 20,
                "status": "publish"
            },
            timeout=15
        )

        if response.status_code != 200:

            return {
                "error": (
                    f"WooCommerce returned "
                    f"HTTP {response.status_code}"
                ),
                "message": response.text
            }

        return response.json()

    except Exception as e:

        return {
            "error": "WooCommerce connection error",
            "message": str(e)
        }


# ============================================================
# SEARCH NORMALISATION
# ============================================================

ARABIC_NORMALISATIONS = {
    "أ": "ا",
    "إ": "ا",
    "آ": "ا",
    "ٱ": "ا",
    "ى": "ي",
    "ة": "ه",
    "ؤ": "و",
    "ئ": "ي",
    "ـ": "",
    "َ": "",
    "ً": "",
    "ُ": "",
    "ٌ": "",
    "ِ": "",
    "ٍ": "",
    "ْ": "",
    "ّ": "",
}

WORD_REPLACEMENTS = {
    "sirah": "seerah",
    "seera": "seerah",
    "sira": "seerah",
    "seerah": "seerah",
    "seerat": "seerah",
    "sirat": "seerah",

    "quraan": "quran",
    "koran": "quran",
    "quran": "quran",

    "hadees": "hadith",
    "hadeeth": "hadith",
    "hadis": "hadith",
    "hadith": "hadith",

    "fiqah": "fiqh",
    "fiqh": "fiqh",

    "tafsir": "tafsir",
    "tafseer": "tafsir",
    "tafsir": "tafsir",

    "sahaba": "sahabah",
    "sahabah": "sahabah",

    "dua": "dua",
    "duaa": "dua",

    "namaz": "salah",
    "salaah": "salah",
    "salah": "salah",

    "ramzan": "ramadan",
    "ramadan": "ramadan",
}


def normalise_arabic(text):

    if not text:
        return ""

    result = str(text)

    for old, new in ARABIC_NORMALISATIONS.items():
        result = result.replace(old, new)

    return result


def normalize_text(text):

    if not text:
        return ""

    text = normalise_arabic(str(text))

    text = text.lower()

    text = re.sub(
        r"[^\w\s\u0600-\u06ff]",
        " ",
        text,
        flags=re.UNICODE
    )

    words = text.split()

    converted = []

    for word in words:

        converted.append(
            WORD_REPLACEMENTS.get(
                word,
                word
            )
        )

    text = " ".join(converted)

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def token_score(query, name):

    q = normalize_text(query)
    n = normalize_text(name)

    if not q or not n:
        return 0

    if q == n:
        return 100

    if q in n:
        return 96

    query_words = q.split()
    name_words = n.split()

    score = 0

    for qword in query_words:

        best = 0

        for nword in name_words:

            similarity = difflib.SequenceMatcher(
                None,
                qword,
                nword
            ).ratio()

            best = max(
                best,
                similarity
            )

        score += best

    score = score / max(
        len(query_words),
        1
    )

    return int(score * 100)


def fuzzy_score(search_text, product_name):

    direct = token_score(
        search_text,
        product_name
    )

    reversed_score = token_score(
        normalize_text(product_name),
        normalize_text(search_text)
    )

    return max(
        direct,
        reversed_score
    )


def get_search_variants(search_text):

    variants = []

    original = search_text.strip()

    if original:
        variants.append(original)

    normalized = normalize_text(original)

    if normalized and normalized not in variants:
        variants.append(normalized)

    words = normalized.split()

    expanded = []

    for word in words:

        replacement = WORD_REPLACEMENTS.get(
            word
        )

        if replacement:
            expanded.append(
                replacement
            )
        else:
            expanded.append(word)

    expanded_text = " ".join(expanded)

    if (
        expanded_text
        and expanded_text not in variants
    ):
        variants.append(expanded_text)

    return variants[:3]


def search_products(search_text):

    if not search_text:
        return []

    all_products = {}

    variants = get_search_variants(
        search_text
    )

    for variant in variants:

        result = woo_search_raw(
            variant
        )

        if isinstance(result, dict):

            if "error" in result:

                if not all_products:

                    return result

                continue

        if isinstance(result, list):

            for product in result:

                product_id = str(
                    product.get(
                        "id",
                        ""
                    )
                )

                if product_id:
                    all_products[
                        product_id
                    ] = product

    products = list(
        all_products.values()
    )

    scored = []

    for product in products:

        combined_text = " ".join(
            [
                str(product.get("name", "")),
                str(product.get("sku", "")),
                str(
                    product.get(
                        "short_description",
                        ""
                    )
                ),
            ]
        )

        score = fuzzy_score(
            search_text,
            combined_text
        )

        scored.append(
            (
                score,
                product
            )
        )

    scored.sort(
        key=lambda x: x[0],
        reverse=True
    )

    return [
        product
        for score, product in scored
        if score >= 18
    ]


# ============================================================
# MONEY
# ============================================================

def money(value):

    try:

        decimal_value = Decimal(
            str(value or 0)
        )

        return float(
            decimal_value.quantize(
                Decimal("0.01"),
                rounding=ROUND_HALF_UP
            )
        )

    except Exception:

        return 0.0


def calculate_item(
    quantity,
    unit_price,
    discount_percent
):

    quantity = money(quantity)
    unit_price = money(unit_price)
    discount_percent = money(
        discount_percent
    )

    discount_amount = (
        unit_price
        * discount_percent
        / 100
    )

    discounted_price = (
        unit_price
        - discount_amount
    )

    total = (
        quantity
        * discounted_price
    )

    return (
        money(discount_amount),
        money(discounted_price),
        money(total)
    )


# ============================================================
# INVOICE NUMBER
# ============================================================

def get_next_invoice_number():

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT invoice_number
        FROM invoices
        WHERE invoice_number IS NOT NULL
          AND invoice_number != ''
        ORDER BY id DESC
        LIMIT 1
        """
    )

    row = cursor.fetchone()

    conn.close()

    if not row:
        return "10001"

    try:

        return str(
            int(
                row["invoice_number"]
            ) + 1
        )

    except Exception:

        return "10001"


# ============================================================
# GOOGLE DRIVE
# ============================================================

def get_drive_service():

    creds = None

    if os.path.exists("token.json"):

        try:

            creds = Credentials.from_authorized_user_file(
                "token.json",
                GOOGLE_SCOPES
            )

        except Exception:

            creds = None

    if (
        creds
        and creds.expired
        and creds.refresh_token
    ):

        try:

            creds.refresh(
                Request()
            )

        except Exception:

            creds = None

    if not creds or not creds.valid:

        if not os.path.exists(
            "credentials.json"
        ):

            st.error(
                "credentials.json was not found."
            )

            return None

        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json",
            GOOGLE_SCOPES
        )

        creds = flow.run_local_server(
            port=0
        )

        with open(
            "token.json",
            "w",
            encoding="utf-8"
        ) as token_file:

            token_file.write(
                creds.to_json()
            )

    return build(
        "drive",
        "v3",
        credentials=creds
    )


def get_or_create_drive_folder(
    service
):

    safe_name = (
        GOOGLE_DRIVE_FOLDER_NAME
        .replace("'", "\\'")
    )

    results = service.files().list(
        q=(
            f"name = '{safe_name}' "
            "and mimeType = "
            "'application/vnd.google-apps.folder' "
            "and trashed = false"
        ),
        spaces="drive",
        fields="files(id,name)",
    ).execute()

    files = results.get(
        "files",
        []
    )

    if files:
        return files[0]["id"]

    metadata = {
        "name": GOOGLE_DRIVE_FOLDER_NAME,
        "mimeType": (
            "application/vnd.google-apps.folder"
        ),
    }

    folder = service.files().create(
        body=metadata,
        fields="id",
    ).execute()

    return folder["id"]


def upload_pdf_to_drive(
    pdf_bytes,
    filename
):

    service = get_drive_service()

    if not service:
        return None

    folder_id = get_or_create_drive_folder(
        service
    )

    metadata = {
        "name": filename,
        "parents": [folder_id],
    }

    media = MediaIoBaseUpload(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        resumable=False,
    )

    uploaded = service.files().create(
        body=metadata,
        media_body=media,
        fields="id,name",
    ).execute()

    return uploaded["id"]


def update_pdf_on_drive(
    pdf_bytes,
    drive_file_id,
    filename
):

    service = get_drive_service()

    if not service:
        return False

    media = MediaIoBaseUpload(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        resumable=False,
    )

    try:

        service.files().update(
            fileId=drive_file_id,
            body={
                "name": filename
            },
            media_body=media,
        ).execute()

        return True

    except Exception as e:

        st.warning(
            "Could not update the existing "
            f"Google Drive PDF: {e}"
        )

        return False


# ============================================================
# PDF FONT
# ============================================================

def register_fonts():

    possible_fonts = [
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/Arial.ttf",
        "C:/Windows/Fonts/calibri.ttf",
        "C:/Windows/Fonts/Calibri.ttf",
    ]

    for path in possible_fonts:

        if os.path.exists(path):

            try:

                pdfmetrics.registerFont(
                    TTFont(
                        "InvoiceMainFont",
                        path
                    )
                )

                return "InvoiceMainFont"

            except Exception:
                pass

    return "Helvetica"


# ============================================================
# PDF HELPERS
# ============================================================

def safe_pdf_text(value):

    if value is None:
        return ""

    return html.escape(
        str(value)
    )


def pdf_paragraph(
    value,
    style
):

    return Paragraph(
        safe_pdf_text(value),
        style
    )


def pdf_markup_paragraph(
    value,
    style
):

    return Paragraph(
        str(value),
        style
    )


# ============================================================
# PROFESSIONAL PDF
# ============================================================

def build_invoice_pdf(
    invoice,
    items
):

    font_name = register_fonts()

    buffer = io.BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=13 * mm,
        title=(
            f"Invoice {invoice['invoice_number']}"
        ),
        author=SHOP_NAME,
    )

    styles = getSampleStyleSheet()

    green = colors.HexColor(
        "#164C3E"
    )

    dark_green = colors.HexColor(
        "#0E332B"
    )

    gold = colors.HexColor(
        "#C6A553"
    )

    light_gold = colors.HexColor(
        "#F4EEDC"
    )

    pale_green = colors.HexColor(
        "#EFF5F1"
    )

    border = colors.HexColor(
        "#D7D7CE"
    )

    normal_style = ParagraphStyle(
        "PDFNormal",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor(
            "#333333"
        ),
    )

    small_style = ParagraphStyle(
        "PDFSmall",
        parent=normal_style,
        fontSize=7.2,
        leading=9.2,
    )

    header_style = ParagraphStyle(
        "PDFHeader",
        parent=normal_style,
        fontSize=19,
        leading=22,
        alignment=TA_CENTER,
        textColor=colors.white,
    )

    subtitle_style = ParagraphStyle(
        "PDFSubtitle",
        parent=normal_style,
        fontSize=9,
        leading=11,
        alignment=TA_CENTER,
        textColor=colors.HexColor(
            "#E9DFC3"
        ),
    )

    invoice_title_style = ParagraphStyle(
        "PDFInvoiceTitle",
        parent=normal_style,
        fontSize=13,
        leading=16,
        alignment=TA_CENTER,
        textColor=green,
    )

    section_style = ParagraphStyle(
        "PDFSection",
        parent=normal_style,
        fontSize=9,
        leading=11,
        textColor=green,
    )

    right_style = ParagraphStyle(
        "PDFRight",
        parent=normal_style,
        alignment=TA_RIGHT,
    )

    center_style = ParagraphStyle(
        "PDFCenter",
        parent=normal_style,
        alignment=TA_CENTER,
    )

    story = []

    # --------------------------------------------------------
    # Header
    # --------------------------------------------------------

    header_data = [
        [
            Paragraph(
                SHOP_NAME,
                header_style
            )
        ],
        [
            Paragraph(
                "بسم الله الرحمن الرحيم",
                ParagraphStyle(
                    "Arabic",
                    parent=subtitle_style,
                    fontSize=10,
                )
            )
        ],
        [
            Paragraph(
                "Professional Sales Invoice",
                subtitle_style
            )
        ],
    ]

    header_table = Table(
        header_data,
        colWidths=[186 * mm],
    )

    header_table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    dark_green
                ),
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    1,
                    dark_green
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),
            ]
        )
    )

    story.append(
        header_table
    )

    story.append(
        Spacer(1, 6 * mm)
    )

    story.append(
        Paragraph(
            "SALES INVOICE",
            invoice_title_style
        )
    )

    story.append(
        Spacer(1, 3 * mm)
    )

    # --------------------------------------------------------
    # Invoice information
    # --------------------------------------------------------

    info_data = [
        [
            Paragraph(
                "<b>Invoice Number</b><br/>"
                + safe_pdf_text(
                    invoice["invoice_number"]
                ),
                normal_style
            ),
            Paragraph(
                "<b>Invoice Date</b><br/>"
                + safe_pdf_text(
                    invoice["invoice_date"]
                ),
                normal_style
            ),
            Paragraph(
                "<b>Added By</b><br/>"
                + safe_pdf_text(
                    invoice["created_by"]
                ),
                normal_style
            ),
        ],
        [
            Paragraph(
                "<b>Customer</b><br/>"
                + safe_pdf_text(
                    invoice["customer_name"]
                    or "Walk-in Customer"
                ),
                normal_style
            ),
            Paragraph(
                "<b>Address</b><br/>"
                + safe_pdf_text(
                    invoice["customer_address"]
                    or "-"
                ),
                normal_style
            ),
            Paragraph(
                "<b>Payment</b><br/>"
                + safe_pdf_text(
                    invoice["payment_method"]
                    or "-"
                ),
                normal_style
            ),
        ],
    ]

    info_table = Table(
        info_data,
        colWidths=[
            62 * mm,
            62 * mm,
            62 * mm,
    ])

    info_table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    colors.white
                ),
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.7,
                    border
                ),
                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.35,
                    border
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP"
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    7
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    7
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    6
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    6
                ),
            ]
        )
    )

    story.append(
        info_table
    )

    story.append(
        Spacer(1, 6 * mm)
    )

    # --------------------------------------------------------
    # Items table
    # --------------------------------------------------------

    item_header = [
        Paragraph(
            "<b>#</b>",
            center_style
        ),
        Paragraph(
            "<b>Item Description</b>",
            normal_style
        ),
        Paragraph(
            "<b>Qty</b>",
            center_style
        ),
        Paragraph(
            "<b>Units</b>",
            center_style
        ),
        Paragraph(
            "<b>Unit Price</b>",
            right_style
        ),
        Paragraph(
            "<b>Discount</b>",
            right_style
        ),
        Paragraph(
            "<b>After Discount</b>",
            right_style
        ),
        Paragraph(
            "<b>Total</b>",
            right_style
        ),
    ]

    table_data = [
        item_header
    ]

    for item in items:

        description = (
            safe_pdf_text(
                item.get(
                    "item_description",
                    ""
                )
            )
        )

        table_data.append(
            [
                Paragraph(
                    str(
                        item.get(
                            "serial_number",
                            ""
                        )
                    ),
                    center_style
                ),
                Paragraph(
                    description,
                    small_style
                ),
                Paragraph(
                    f"{money(item.get('quantity', 0)):g}",
                    center_style
                ),
                Paragraph(
                    safe_pdf_text(
                        item.get(
                            "units",
                            "pcs"
                        )
                    ),
                    center_style
                ),
                Paragraph(
                    f"£{money(item.get('unit_price', 0)):.2f}",
                    right_style
                ),
                Paragraph(
                    f"{money(item.get('discount_percent', 0)):.2f}%",
                    right_style
                ),
                Paragraph(
                    f"£{money(item.get('discounted_unit_price', 0)):.2f}",
                    right_style
                ),
                Paragraph(
                    f"£{money(item.get('total_price', 0)):.2f}",
                    right_style
                ),
            ]
        )

    items_table = Table(
        table_data,
        colWidths=[
            8 * mm,
            60 * mm,
            12 * mm,
            15 * mm,
            22 * mm,
            19 * mm,
            27 * mm,
            23 * mm,
        ],
        repeatRows=1,
    )

    items_table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    green
                ),
                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white
                ),
                (
                    "FONTNAME",
                    (0, 0),
                    (-1, -1),
                    font_name
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.35,
                    border
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),
                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [
                        colors.white,
                        pale_green,
                    ]
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    6
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    6
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),
            ]
        )
    )

    story.append(
        items_table
    )

    story.append(
        Spacer(1, 6 * mm)
    )

    # --------------------------------------------------------
    # Totals
    # --------------------------------------------------------

    totals_data = [
        [
            Paragraph(
                "Subtotal",
                normal_style
            ),
            Paragraph(
                f"£{money(invoice['subtotal']):.2f}",
                right_style
            ),
        ],
        [
            Paragraph(
                "Total Discount",
                normal_style
            ),
            Paragraph(
                f"£{money(invoice['total_discount']):.2f}",
                right_style
            ),
        ],
        [
            Paragraph(
                "Delivery Charges",
                normal_style
            ),
            Paragraph(
                f"£{money(invoice['delivery_charges']):.2f}",
                right_style
            ),
        ],
        [
            Paragraph(
                "<b>TOTAL</b>",
                normal_style
            ),
            Paragraph(
                f"<b>£{money(invoice['grand_total']):.2f}</b>",
                right_style
            ),
        ],
        [
            Paragraph(
                "Total Paid",
                normal_style
            ),
            Paragraph(
                f"£{money(invoice['total_paid']):.2f}",
                right_style
            ),
        ],
        [
            Paragraph(
                "Returned",
                normal_style
            ),
            Paragraph(
                f"£{money(invoice['returned_amount']):.2f}",
                right_style
            ),
        ],
        [
            Paragraph(
                "<b>Amount Due</b>",
                normal_style
            ),
            Paragraph(
                f"<b>£{money(invoice['total_due']):.2f}</b>",
                right_style
            ),
        ],
    ]

    totals_table = Table(
        totals_data,
        colWidths=[
            145 * mm,
            41 * mm,
        ],
    )

    totals_table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 3),
                    (-1, 3),
                    light_gold
                ),
                (
                    "BACKGROUND",
                    (0, 6),
                    (-1, 6),
                    pale_green
                ),
                (
                    "LINEABOVE",
                    (0, 3),
                    (-1, 3),
                    1.2,
                    gold
                ),
                (
                    "LINEABOVE",
                    (0, 6),
                    (-1, 6),
                    1,
                    green
                ),
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    border
                ),
                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.25,
                    border
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    7
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    7
                ),
            ]
        )
    )

    story.append(
        totals_table
    )

    story.append(
        Spacer(1, 6 * mm)
    )

    # --------------------------------------------------------
    # Payment
    # --------------------------------------------------------

    payment_data = [
        [
            Paragraph(
                "<b>Payment Details</b>",
                section_style
            ),
            "",
        ],
        [
            Paragraph(
                "Payment Method",
                normal_style
            ),
            Paragraph(
                safe_pdf_text(
                    invoice["payment_method"]
                    or "-"
                ),
                right_style
            ),
        ],
        [
            Paragraph(
                "Amount Received",
                normal_style
            ),
            Paragraph(
                f"£{money(invoice['cash_received']):.2f}",
                right_style
            ),
        ],
    ]

    payment_table = Table(
        payment_data,
        colWidths=[
            90 * mm,
            96 * mm
        ]
    )

    payment_table.setStyle(
        TableStyle(
            [
                (
                    "SPAN",
                    (0, 0),
                    (1, 0)
                ),
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    light_gold
                ),
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    border
                ),
                (
                    "INNERGRID",
                    (0, 1),
                    (-1, -1),
                    0.25,
                    border
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    5
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    7
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    7
                ),
            ]
        )
    )

    story.append(
        payment_table
    )

    story.append(
        Spacer(1, 6 * mm)
    )

    # --------------------------------------------------------
    # Bank details
    # --------------------------------------------------------

    bank_data = [
        [
            Paragraph(
                "<b>Bank Details</b>",
                section_style
            ),
            "",
        ],
        [
            Paragraph(
                "Bank",
                normal_style
            ),
            Paragraph(
                BANK_NAME,
                normal_style
            ),
        ],
        [
            Paragraph(
                "Account Name",
                normal_style
            ),
            Paragraph(
                BANK_ACCOUNT_NAME,
                normal_style
            ),
        ],
        [
            Paragraph(
                "Sort Code",
                normal_style
            ),
            Paragraph(
                SORT_CODE,
                normal_style
            ),
        ],
        [
            Paragraph(
                "Account Number",
                normal_style
            ),
            Paragraph(
                ACCOUNT_NUMBER,
                normal_style
            ),
        ],
    ]

    bank_table = Table(
        bank_data,
        colWidths=[
            50 * mm,
            70 * mm
        ]
    )

    bank_table.setStyle(
        TableStyle(
            [
                (
                    "SPAN",
                    (0, 0),
                    (1, 0)
                ),
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    pale_green
                ),
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    border
                ),
                (
                    "INNERGRID",
                    (0, 1),
                    (-1, -1),
                    0.25,
                    border
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    4
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    6
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    6
                ),
            ]
        )
    )

    story.append(
        bank_table
    )

    # --------------------------------------------------------
    # Remarks
    # --------------------------------------------------------

    if invoice.get("remarks"):

        story.append(
            Spacer(1, 5 * mm)
        )

        remarks_table = Table(
            [
                [
                    Paragraph(
                        "<b>Remarks</b><br/>"
                        + safe_pdf_text(
                            invoice["remarks"]
                        ),
                        normal_style
                    )
                ]
            ],
            colWidths=[
                186 * mm
            ]
        )

        remarks_table.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, -1),
                        colors.HexColor(
                            "#FAF8F0"
                        )
                    ),
                    (
                        "BOX",
                        (0, 0),
                        (-1, -1),
                        0.5,
                        border
                    ),
                    (
                        "LEFTPADDING",
                        (0, 0),
                        (-1, -1),
                        7
                    ),
                    (
                        "RIGHTPADDING",
                        (0, 0),
                        (-1, -1),
                        7
                    ),
                    (
                        "TOPPADDING",
                        (0, 0),
                        (-1, -1),
                        6
                    ),
                    (
                        "BOTTOMPADDING",
                        (0, 0),
                        (-1, -1),
                        6
                    ),
                ]
            )
        )

        story.append(
            remarks_table
        )

    story.append(
        Spacer(1, 8 * mm)
    )

    story.append(
        Paragraph(
            "Thank you for your business.",
            ParagraphStyle(
                "PDFFooter",
                parent=normal_style,
                alignment=TA_CENTER,
                fontSize=8.5,
                textColor=colors.HexColor(
                    "#66706B"
                ),
            )
        )
    )

    story.append(
        Spacer(1, 2 * mm)
    )

    story.append(
        Paragraph(
            "Dar Makkah International",
            ParagraphStyle(
                "PDFFooterShop",
                parent=normal_style,
                alignment=TA_CENTER,
                fontSize=7.5,
                textColor=green,
            )
        )
    )

    # --------------------------------------------------------
    # Page callback
    # --------------------------------------------------------

    def draw_page(canvas, doc):

        canvas.saveState()

        width, height = A4

        canvas.setStrokeColor(
            gold
        )

        canvas.setLineWidth(
            0.7
        )

        canvas.line(
            12 * mm,
            8 * mm,
            width - 12 * mm,
            8 * mm
        )

        canvas.setFont(
            font_name,
            7
        )

        canvas.setFillColor(
            colors.HexColor(
                "#777777"
            )
        )

        canvas.drawCentredString(
            width / 2,
            4.5 * mm,
            f"Page {doc.page}"
        )

        canvas.restoreState()

    document.build(
        story,
        onFirstPage=draw_page,
        onLaterPages=draw_page
    )

    return buffer.getvalue()


# ============================================================
# DATABASE - INVOICE SAVE
# ============================================================

def save_invoice(
    invoice,
    items,
    invoice_id=None
):

    conn = get_connection()
    cursor = conn.cursor()

    now = datetime.now().isoformat(
        timespec="seconds"
    )

    if invoice_id is None:

        cursor.execute(
            """
            INSERT INTO invoices (
                invoice_number,
                invoice_date,
                customer_name,
                customer_address,
                created_by,
                payment_method,
                cash_received,
                returned_amount,
                delivery_charges,
                total_discount,
                subtotal,
                grand_total,
                total_paid,
                total_due,
                remarks,
                created_at,
                updated_at
            )
            VALUES (
                ?, ?, ?, ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?, ?, ?
            )
            """,
            (
                invoice["invoice_number"],
                invoice["invoice_date"],
                invoice["customer_name"],
                invoice["customer_address"],
                invoice["created_by"],
                invoice["payment_method"],
                invoice["cash_received"],
                invoice["returned_amount"],
                invoice["delivery_charges"],
                invoice["total_discount"],
                invoice["subtotal"],
                invoice["grand_total"],
                invoice["total_paid"],
                invoice["total_due"],
                invoice["remarks"],
                now,
                now,
            )
        )

        invoice_id = cursor.lastrowid

    else:

        cursor.execute(
            """
            UPDATE invoices
            SET
                invoice_date = ?,
                customer_name = ?,
                customer_address = ?,
                created_by = ?,
                payment_method = ?,
                cash_received = ?,
                returned_amount = ?,
                delivery_charges = ?,
                total_discount = ?,
                subtotal = ?,
                grand_total = ?,
                total_paid = ?,
                total_due = ?,
                remarks = ?,
                updated_at = ?
            WHERE id = ?
            """,
            (
                invoice["invoice_date"],
                invoice["customer_name"],
                invoice["customer_address"],
                invoice["created_by"],
                invoice["payment_method"],
                invoice["cash_received"],
                invoice["returned_amount"],
                invoice["delivery_charges"],
                invoice["total_discount"],
                invoice["subtotal"],
                invoice["grand_total"],
                invoice["total_paid"],
                invoice["total_due"],
                invoice["remarks"],
                now,
                invoice_id,
            )
        )

        cursor.execute(
            """
            DELETE FROM invoice_items
            WHERE invoice_id = ?
            """,
            (invoice_id,)
        )

    for index, item in enumerate(items):

        cursor.execute(
            """
            INSERT INTO invoice_items (
                invoice_id,
                serial_number,
                product_id,
                sku,
                item_description,
                quantity,
                units,
                unit_price,
                discount_percent,
                discount_amount,
                discounted_unit_price,
                total_price
            )
            VALUES (
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
            )
            """,
            (
                invoice_id,
                index + 1,
                item.get("product_id", ""),
                item.get("sku", ""),
                item.get("item_description", ""),
                money(item.get("quantity", 1)),
                item.get("units", "pcs"),
                money(item.get("unit_price", 0)),
                money(item.get("discount_percent", 0)),
                money(item.get("discount_amount", 0)),
                money(
                    item.get(
                        "discounted_unit_price",
                        0
                    )
                ),
                money(
                    item.get(
                        "total_price",
                        0
                    )
                ),
            )
        )

    conn.commit()

    cursor.execute(
        """
        SELECT
            google_drive_id,
            google_drive_name
        FROM invoices
        WHERE id = ?
        """,
        (invoice_id,)
    )

    drive_row = cursor.fetchone()

    conn.close()

    return (
        invoice_id,
        dict(drive_row)
        if drive_row
        else {}
    )


def update_drive_information(
    invoice_id,
    drive_id,
    drive_name
):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        UPDATE invoices
        SET
            google_drive_id = ?,
            google_drive_name = ?
        WHERE id = ?
        """,
        (
            drive_id,
            drive_name,
            invoice_id,
        )
    )

    conn.commit()
    conn.close()


# ============================================================
# GET INVOICE
# ============================================================

def get_invoice(invoice_id):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT *
        FROM invoices
        WHERE id = ?
        """,
        (invoice_id,)
    )

    invoice = cursor.fetchone()

    cursor.execute(
        """
        SELECT *
        FROM invoice_items
        WHERE invoice_id = ?
        ORDER BY serial_number
        """,
        (invoice_id,)
    )

    items = cursor.fetchall()

    conn.close()

    return (
        dict(invoice)
        if invoice
        else None,
        [
            dict(item)
            for item in items
        ]
    )


# ============================================================
# INVOICE SEARCH
# ============================================================

def get_invoices(
    customer_search="",
    invoice_search="",
    date_from=None,
    date_to=None,
    created_by="",
):

    conn = get_connection()
    cursor = conn.cursor()

    query = """
        SELECT *
        FROM invoices
        WHERE 1 = 1
    """

    params = []

    if customer_search:

        normalized_customer = (
            normalize_text(
                customer_search
            )
        )

        # Normal database search first
        query += """
            AND LOWER(
                COALESCE(customer_name, '')
            ) LIKE ?
        """

        params.append(
            "%"
            + normalized_customer.lower()
            + "%"
        )

    if invoice_search:

        query += """
            AND invoice_number LIKE ?
        """

        params.append(
            "%"
            + invoice_search.strip()
            + "%"
        )

    if date_from:

        query += """
            AND invoice_date >= ?
        """

        params.append(
            str(date_from)
        )

    if date_to:

        query += """
            AND invoice_date <= ?
        """

        params.append(
            str(date_to)
        )

    if (
        created_by
        and created_by != "All"
    ):

        query += """
            AND created_by = ?
        """

        params.append(
            created_by
        )

    query += """
        ORDER BY id DESC
    """

    cursor.execute(
        query,
        params
    )

    rows = cursor.fetchall()

    conn.close()

    return [
        dict(row)
        for row in rows
    ]


# ============================================================
# DELETE
# ============================================================

def delete_invoice(
    invoice_id
):

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        DELETE FROM invoice_items
        WHERE invoice_id = ?
        """,
        (invoice_id,)
    )

    cursor.execute(
        """
        DELETE FROM invoices
        WHERE id = ?
        """,
        (invoice_id,)
    )

    conn.commit()
    conn.close()


# ============================================================
# SESSION STATE
# ============================================================

DEFAULT_STATE = {
    "invoice_items": [],
    "editing_invoice_id": None,
    "edit_loaded": False,
    "edit_invoice_data": {},
    "pending_navigation": None,
    "manual_item_counter": 0,
    "save_success_message": None,
}

for key, default in DEFAULT_STATE.items():

    if key not in st.session_state:

        st.session_state[key] = default


# ============================================================
# IMPORTANT NAVIGATION FIX
# ============================================================
#
# NEVER do:
#
# st.session_state.navigation_page = ...
#
# after the radio widget has been created.
#
# Instead:
#
# 1. Buttons set pending_navigation.
# 2. st.rerun().
# 3. At the TOP, BEFORE radio is created,
#    navigation_page is set.
#
# ============================================================

if st.session_state.pending_navigation:

    requested_page = (
        st.session_state.pending_navigation
    )

    st.session_state.navigation_page = (
        requested_page
    )

    st.session_state.pending_navigation = None


if "navigation_page" not in st.session_state:

    st.session_state.navigation_page = (
        NAV_CREATE
    )


# ============================================================
# CALLBACKS
# ============================================================

def request_create_invoice():

    st.session_state.editing_invoice_id = None
    st.session_state.edit_loaded = False
    st.session_state.edit_invoice_data = {}
    st.session_state.invoice_items = []

    st.session_state.pending_navigation = (
        NAV_CREATE
    )


def start_edit_invoice(
    invoice_id
):

    # IMPORTANT:
    # Do NOT modify navigation_page here.
    # That is the exact cause of the previous
    # StreamlitAPIException.

    st.session_state.editing_invoice_id = (
        invoice_id
    )

    st.session_state.edit_loaded = False
    st.session_state.edit_invoice_data = {}
    st.session_state.invoice_items = []

    st.session_state.pending_navigation = (
        NAV_CREATE
    )


# ============================================================
# ADD PRODUCT
# ============================================================

def add_product_to_invoice(
    product
):

    regular_price = (
        product.get(
            "regular_price"
        )
        or product.get(
            "price"
        )
        or 0
    )

    regular_price = money(
        regular_price
    )

    item = {
        "serial_number": (
            len(
                st.session_state.invoice_items
            ) + 1
        ),
        "product_id": str(
            product.get(
                "id",
                ""
            )
        ),
        "sku": product.get(
            "sku",
            ""
        ),
        "item_description": product.get(
            "name",
            ""
        ),
        "quantity": 1.0,
        "units": "pcs",
        "unit_price": regular_price,
        "discount_percent": 0.0,
        "discount_amount": 0.0,
        "discounted_unit_price": regular_price,
        "total_price": regular_price,
    }

    st.session_state.invoice_items.append(
        item
    )


def add_manual_item(
    description=""
):

    st.session_state.manual_item_counter += 1

    description = (
        description.strip()
        if description
        else "Manual Item"
    )

    st.session_state.invoice_items.append(
        {
            "serial_number": (
                len(
                    st.session_state.invoice_items
                ) + 1
            ),
            "product_id": "",
            "sku": "",
            "item_description": description,
            "quantity": 1.0,
            "units": "pcs",
            "unit_price": 0.0,
            "discount_percent": 0.0,
            "discount_amount": 0.0,
            "discounted_unit_price": 0.0,
            "total_price": 0.0,
        }
    )


# ============================================================
# SIDEBAR NAVIGATION
# ============================================================

with st.sidebar:

    st.markdown(
        """
        <div class="sidebar-brand">
            <div class="sidebar-icon">🕌</div>
            <div class="sidebar-title">
                Dar Makkah
            </div>
            <div style="font-size:12px;color:#d7d2c4;">
                International
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.divider()

    page = st.radio(
        "Navigation",
        [
            NAV_CREATE,
            NAV_INVOICES,
            NAV_EXCEL,
        ],
        key="navigation_page",
    )

    st.divider()

    st.caption(
        "WooCommerce products are searched live."
    )

    st.caption(
        "The complete product catalogue is not "
        "downloaded when the application starts."
    )


# ============================================================
# CREATE / EDIT INVOICE
# ============================================================

if page == NAV_CREATE:

    editing = (
        st.session_state.editing_invoice_id
        is not None
    )

    if editing:

        st.markdown(
            '<div class="page-heading">'
            "✏️ Edit Invoice"
            "</div>",
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="gold-line"></div>',
            unsafe_allow_html=True
        )

    else:

        st.markdown(
            '<div class="page-heading">'
            "🧾 Create New Invoice"
            "</div>",
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="gold-line"></div>',
            unsafe_allow_html=True
        )

    # --------------------------------------------------------
    # LOAD EDIT DATA
    # --------------------------------------------------------

    if (
        editing
        and not st.session_state.edit_loaded
    ):

        existing_invoice, existing_items = (
            get_invoice(
                st.session_state.editing_invoice_id
            )
        )

        if existing_invoice:

            st.session_state.edit_invoice_data = (
                existing_invoice
            )

            st.session_state.invoice_items = (
                existing_items
            )

            st.session_state.edit_loaded = True

        else:

            st.error(
                "The invoice could not be found."
            )

            st.session_state.editing_invoice_id = None
            st.session_state.edit_loaded = False
            st.rerun()

    edit_data = (
        st.session_state.edit_invoice_data
        if editing
        else {}
    )

    # --------------------------------------------------------
    # Invoice information
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-card">',
        unsafe_allow_html=True
    )

    c1, c2, c3 = st.columns(3)

    with c1:

        invoice_number = st.text_input(
            "Invoice Number",
            value=(
                edit_data.get(
                    "invoice_number",
                    get_next_invoice_number()
                )
            ),
            disabled=editing,
            key="invoice_number_field",
        )

    with c2:

        if edit_data.get(
            "invoice_date"
        ):

            try:

                default_date = (
                    datetime.strptime(
                        edit_data[
                            "invoice_date"
                        ],
                        "%Y-%m-%d"
                    ).date()
                )

            except Exception:

                default_date = date.today()

        else:

            default_date = date.today()

        invoice_date = st.date_input(
            "Invoice Date",
            value=default_date,
            key="invoice_date_field",
        )

    with c3:

        created_by = st.selectbox(
            "Added By",
            [
                "PC1",
                "PC2",
            ],
            index=(
                1
                if edit_data.get(
                    "created_by"
                ) == "PC2"
                else 0
            ),
            key="created_by_field",
        )

    c1, c2 = st.columns(2)

    with c1:

        customer_name = st.text_input(
            "Customer Name",
            value=edit_data.get(
                "customer_name",
                ""
            ),
            key="customer_name_field",
        )

    with c2:

        customer_address = st.text_input(
            "Customer Address",
            value=edit_data.get(
                "customer_address",
                ""
            ),
            key="customer_address_field",
        )

    st.markdown(
        "</div>",
        unsafe_allow_html=True
    )

    # ========================================================
    # PRODUCT SEARCH
    # ========================================================

    st.markdown(
        '<div class="section-card">',
        unsafe_allow_html=True
    )

    st.markdown(
        "### 🔎 Find Product"
    )

    st.caption(
        "Searches WooCommerce live. "
        "You can also add a manual item at any time."
    )

    search_text = st.text_input(
        "Product Search",
        placeholder=(
            "Try: sirah, seerah, quran, fiqh..."
        ),
        key="product_search_field",
    )

    # Manual button is ALWAYS available.
    manual_col1, manual_col2 = st.columns(
        [4, 1]
    )

    with manual_col1:

        manual_description = st.text_input(
            "Manual Item Description",
            placeholder=(
                "Enter a book/product name "
                "if it is not found"
            ),
            key="manual_description_field",
        )

    with manual_col2:

        st.write("")

        if st.button(
            "➕ Add Manual Item",
            key="always_manual_item",
            use_container_width=True,
        ):

            add_manual_item(
                manual_description
            )

            st.success(
                "Manual item added."
            )

            st.rerun()

    products = []

    if (
        search_text
        and len(
            search_text.strip()
        ) >= 2
    ):

        with st.spinner(
            "Searching WooCommerce..."
        ):

            result = search_products(
                search_text
            )

        if (
            isinstance(result, dict)
            and "error" in result
        ):

            st.error(
                result["error"]
            )

            st.code(
                result.get(
                    "message",
                    ""
                )
            )

        else:

            products = result

    if products:

        st.write(
            f"Found {len(products)} relevant product(s)"
        )

        for index, product in enumerate(
            products[:12]
        ):

            name = product.get(
                "name",
                ""
            )

            sku = product.get(
                "sku",
                ""
            )

            regular_price = (
                product.get(
                    "regular_price"
                )
                or product.get(
                    "price"
                )
                or "0"
            )

            sale_price = (
                product.get(
                    "sale_price"
                )
                or ""
            )

            score = fuzzy_score(
                search_text,
                name
            )

            col1, col2, col3, col4 = (
                st.columns(
                    [5.2, 1.3, 1.3, 1]
                )
            )

            with col1:

                st.markdown(
                    f"""
                    <div class="product-result">
                        <div class="product-name">
                            {html.escape(name)}
                        </div>
                        <div class="product-meta">
                            SKU:
                            {html.escape(sku or "-")}
                            &nbsp;&nbsp;|&nbsp;&nbsp;
                            Match: {score}%
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            with col2:

                st.write(
                    f"£{money(regular_price):.2f}"
                )

            with col3:

                if sale_price:

                    st.write(
                        f"Sale £{money(sale_price):.2f}"
                    )

                else:

                    st.caption(
                        "No sale"
                    )

            with col4:

                if st.button(
                    "Add",
                    key=(
                        f"add_product_"
                        f"{product.get('id')}_"
                        f"{index}"
                    ),
                    use_container_width=True,
                ):

                    add_product_to_invoice(
                        product
                    )

                    st.success(
                        "Product added."
                    )

                    st.rerun()

    elif (
        search_text
        and len(
            search_text.strip()
        ) >= 2
    ):

        st.warning(
            "No matching WooCommerce product found."
        )

        # Manual button still exists above.
        st.info(
            "You can use the Manual Item Description "
            "box above to add this product manually."
        )

    st.markdown(
        "</div>",
        unsafe_allow_html=True
    )

    # ========================================================
    # ITEMS
    # ========================================================

    st.markdown(
        '<div class="section-card">',
        unsafe_allow_html=True
    )

    st.markdown(
        "### 📚 Invoice Items"
    )

    if not st.session_state.invoice_items:

        st.info(
            "No items added yet. "
            "Search for a WooCommerce product or "
            "add a manual item."
        )

    # We intentionally iterate over a snapshot
    # so deleting an item does not corrupt the loop.
    for i in range(
        len(
            st.session_state.invoice_items
        )
    ):

        item = (
            st.session_state.invoice_items[i]
        )

        item["serial_number"] = i + 1

        st.markdown(
            f"**Item {i + 1}**"
        )

        c1, c2, c3 = st.columns(
            [6, 1.5, 0.7]
        )

        with c1:

            item["item_description"] = (
                st.text_input(
                    "Item Description",
                    value=item.get(
                        "item_description",
                        ""
                    ),
                    key=f"description_{i}",
                )
            )

        with c2:

            item["sku"] = st.text_input(
                "SKU",
                value=item.get(
                    "sku",
                    ""
                ),
                key=f"sku_{i}",
            )

        with c3:

            st.write("")

            if st.button(
                "🗑️",
                key=f"delete_item_{i}",
            ):

                st.session_state.invoice_items.pop(
                    i
                )

                st.rerun()

        c1, c2, c3, c4 = st.columns(4)

        with c1:

            item["quantity"] = (
                st.number_input(
                    "Quantity",
                    min_value=0.01,
                    value=float(
                        item.get(
                            "quantity",
                            1
                        )
                    ),
                    step=1.0,
                    key=f"quantity_{i}",
                )
            )

        with c2:

            item["units"] = st.text_input(
                "Units",
                value=item.get(
                    "units",
                    "pcs"
                ),
                key=f"units_{i}",
            )

        with c3:

            item["unit_price"] = (
                st.number_input(
                    "Unit Price (£)",
                    min_value=0.0,
                    value=float(
                        item.get(
                            "unit_price",
                            0
                        )
                    ),
                    step=0.01,
                    key=f"unit_price_{i}",
                )
            )

        with c4:

            item["discount_percent"] = (
                st.number_input(
                    "Discount %",
                    min_value=0.0,
                    max_value=100.0,
                    value=float(
                        item.get(
                            "discount_percent",
                            0
                        )
                    ),
                    step=1.0,
                    key=f"discount_{i}",
                )
            )

        (
            discount_amount,
            discounted_unit_price,
            total_price,
        ) = calculate_item(
            item["quantity"],
            item["unit_price"],
            item["discount_percent"],
        )

        item["discount_amount"] = (
            discount_amount
        )

        item["discounted_unit_price"] = (
            discounted_unit_price
        )

        item["total_price"] = (
            total_price
        )

        c1, c2, c3 = st.columns(3)

        with c1:

            st.metric(
                "Discount Amount",
                f"£{discount_amount:.2f}",
            )

        with c2:

            st.metric(
                "After Discount",
                f"£{discounted_unit_price:.2f}",
            )

        with c3:

            st.metric(
                "Item Total",
                f"£{total_price:.2f}",
            )

        if i < (
            len(
                st.session_state.invoice_items
            ) - 1
        ):

            st.divider()

    st.markdown(
        "</div>",
        unsafe_allow_html=True
    )

    # ========================================================
    # TOTALS
    # ========================================================

    subtotal = money(
        sum(
            money(
                item.get(
                    "quantity",
                    0
                )
            )
            * money(
                item.get(
                    "unit_price",
                    0
                )
            )
            for item in (
                st.session_state.invoice_items
            )
        )
    )

    total_after_item_discounts = money(
        sum(
            money(
                item.get(
                    "total_price",
                    0
                )
            )
            for item in (
                st.session_state.invoice_items
            )
        )
    )

    total_discount = money(
        subtotal
        - total_after_item_discounts
    )

    delivery_charges = st.number_input(
        "Delivery Charges (£)",
        min_value=0.0,
        value=float(
            edit_data.get(
                "delivery_charges",
                0
            )
        ),
        step=0.01,
        key="delivery_charges_field",
    )

    grand_total = money(
        total_after_item_discounts
        + delivery_charges
    )

    st.markdown(
        "### 💰 Invoice Summary"
    )

    c1, c2, c3, c4 = st.columns(4)

    with c1:

        st.metric(
            "Subtotal",
            f"£{subtotal:.2f}"
        )

    with c2:

        st.metric(
            "Discount",
            f"£{total_discount:.2f}"
        )

    with c3:

        st.metric(
            "Delivery",
            f"£{money(delivery_charges):.2f}"
        )

    with c4:

        st.metric(
            "TOTAL",
            f"£{grand_total:.2f}"
        )

    # ========================================================
    # PAYMENT
    # ========================================================

    st.markdown(
        "### 💳 Payment"
    )

    payment_options = [
        "Cash",
        "Bank Transfer",
        "Card",
    ]

    existing_payment = (
        edit_data.get(
            "payment_method",
            "Cash"
        )
    )

    payment_index = (
        payment_options.index(
            existing_payment
        )
        if existing_payment
        in payment_options
        else 0
    )

    c1, c2, c3 = st.columns(3)

    with c1:

        payment_method = st.selectbox(
            "Payment Method",
            payment_options,
            index=payment_index,
            key="payment_method_field",
        )

    with c2:

        cash_received = st.number_input(
            "Amount Received (£)",
            min_value=0.0,
            value=float(
                edit_data.get(
                    "cash_received",
                    0
                )
            ),
            step=0.01,
            key="cash_received_field",
        )

    with c3:

        returned_amount = st.number_input(
            "Returned (£)",
            min_value=0.0,
            value=float(
                edit_data.get(
                    "returned_amount",
                    0
                )
            ),
            step=0.01,
            key="returned_amount_field",
        )

    total_paid = money(
        max(
            0,
            cash_received
            - returned_amount
        )
    )

    total_due = money(
        max(
            0,
            grand_total
            - total_paid
        )
    )

    c1, c2 = st.columns(2)

    with c1:

        st.metric(
            "Total Paid",
            f"£{total_paid:.2f}"
        )

    with c2:

        st.metric(
            "Amount Due",
            f"£{total_due:.2f}"
        )

    remarks = st.text_area(
        "Remarks",
        value=edit_data.get(
            "remarks",
            ""
        ),
        key="remarks_field",
    )

    # ========================================================
    # SAVE
    # ========================================================

    st.divider()

    save_text = (
        "💾 Update Invoice"
        if editing
        else "💾 Save Invoice"
    )

    if st.button(
        save_text,
        type="primary",
        use_container_width=True,
        key="save_invoice_button",
    ):

        if not st.session_state.invoice_items:

            st.error(
                "Please add at least one item."
            )

        elif not invoice_number.strip():

            st.error(
                "Invoice number cannot be empty."
            )

        else:

            invoice_data = {
                "invoice_number": (
                    invoice_number.strip()
                ),
                "invoice_date": str(
                    invoice_date
                ),
                "customer_name": (
                    customer_name.strip()
                ),
                "customer_address": (
                    customer_address.strip()
                ),
                "created_by": created_by,
                "payment_method": (
                    payment_method
                ),
                "cash_received": money(
                    cash_received
                ),
                "returned_amount": money(
                    returned_amount
                ),
                "delivery_charges": money(
                    delivery_charges
                ),
                "total_discount": money(
                    total_discount
                ),
                "subtotal": money(
                    subtotal
                ),
                "grand_total": money(
                    grand_total
                ),
                "total_paid": money(
                    total_paid
                ),
                "total_due": money(
                    total_due
                ),
                "remarks": remarks,
            }

            try:

                invoice_id, drive_info = (
                    save_invoice(
                        invoice_data,
                        st.session_state.invoice_items,
                        st.session_state.editing_invoice_id,
                    )
                )

            except sqlite3.IntegrityError:

                st.error(
                    "That invoice number already exists. "
                    "Please use a different invoice number."
                )

                st.stop()

            invoice, saved_items = (
                get_invoice(
                    invoice_id
                )
            )

            pdf_bytes = build_invoice_pdf(
                invoice,
                saved_items
            )

            filename = (
                f"Invoice-{invoice_number}.pdf"
            )

            existing_drive_id = (
                invoice.get(
                    "google_drive_id"
                )
            )

            if existing_drive_id:

                with st.spinner(
                    "Updating the existing Google Drive PDF..."
                ):

                    success = update_pdf_on_drive(
                        pdf_bytes,
                        existing_drive_id,
                        filename,
                    )

                if success:

                    st.success(
                        "Invoice updated successfully. "
                        "The existing Google Drive PDF "
                        "has also been replaced."
                    )

                else:

                    st.warning(
                        "Invoice was updated locally, "
                        "but the Google Drive PDF could "
                        "not be updated."
                    )

            else:

                with st.spinner(
                    "Uploading invoice PDF to Google Drive..."
                ):

                    drive_id = (
                        upload_pdf_to_drive(
                            pdf_bytes,
                            filename
                        )
                    )

                if drive_id:

                    update_drive_information(
                        invoice_id,
                        drive_id,
                        filename
                    )

                    st.success(
                        "Invoice saved and PDF uploaded "
                        "to Google Drive."
                    )

                else:

                    st.warning(
                        "Invoice saved, but Google Drive "
                        "upload was unsuccessful."
                    )

            # Save a local copy too.
            try:

                with open(
                    filename,
                    "wb"
                ) as file:

                    file.write(
                        pdf_bytes
                    )

            except Exception:
                pass

            st.download_button(
                "⬇️ Download PDF",
                data=pdf_bytes,
                file_name=filename,
                mime="application/pdf",
                key=(
                    f"saved_download_"
                    f"{invoice_id}"
                ),
            )

            # Reset invoice editor.
            st.session_state.invoice_items = []
            st.session_state.editing_invoice_id = None
            st.session_state.edit_loaded = False
            st.session_state.edit_invoice_data = {}

            # Navigate safely.
            st.session_state.pending_navigation = (
                NAV_INVOICES
            )

            st.rerun()


# ============================================================
# INVOICES
# ============================================================

elif page == NAV_INVOICES:

    st.markdown(
        '<div class="page-heading">'
        "📋 Invoice Records"
        "</div>",
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="gold-line"></div>',
        unsafe_allow_html=True
    )

    st.caption(
        "Search by customer, invoice number, "
        "date range or PC."
    )

    # --------------------------------------------------------
    # Search
    # --------------------------------------------------------

    c1, c2, c3 = st.columns(3)

    with c1:

        customer_search = st.text_input(
            "Search Customer",
            key="invoice_customer_search",
            placeholder="Customer name..."
        )

    with c2:

        invoice_search = st.text_input(
            "Search Invoice Number",
            key="invoice_number_search",
            placeholder="e.g. 10001"
        )

    with c3:

        created_by_filter = st.selectbox(
            "Added By",
            [
                "All",
                "PC1",
                "PC2",
            ],
            key="invoice_created_by_filter",
        )

    c1, c2 = st.columns(2)

    with c1:

        date_from = st.date_input(
            "From Date",
            value=None,
            key="invoice_date_from",
        )

    with c2:

        date_to = st.date_input(
            "To Date",
            value=None,
            key="invoice_date_to",
        )

    # --------------------------------------------------------
    # Search database
    # --------------------------------------------------------

    rows = get_invoices(
        customer_search=customer_search,
        invoice_search=invoice_search,
        date_from=date_from,
        date_to=date_to,
        created_by=created_by_filter,
    )

    st.write(
        f"**Found {len(rows)} invoice(s)**"
    )

    if not rows:

        st.info(
            "No invoices found matching your search."
        )

    for row in rows:

        customer_display = (
            row["customer_name"]
            or "Walk-in Customer"
        )

        drive_status = (
            "✓ Stored"
            if row.get(
                "google_drive_id"
            )
            else "Not uploaded"
        )

        with st.expander(
            (
                f"Invoice #{row['invoice_number']} "
                f"— {customer_display} "
                f"— {row['invoice_date']} "
                f"— £{money(row['grand_total']):.2f}"
            )
        ):

            c1, c2, c3, c4 = st.columns(4)

            with c1:

                st.caption(
                    "Customer"
                )

                st.write(
                    customer_display
                )

            with c2:

                st.caption(
                    "Added By"
                )

                st.write(
                    row["created_by"]
                )

            with c3:

                st.caption(
                    "Payment"
                )

                st.write(
                    row["payment_method"]
                )

            with c4:

                st.caption(
                    "Amount Due"
                )

                st.write(
                    f"£{money(row['total_due']):.2f}"
                )

            st.caption(
                f"Google Drive: {drive_status}"
            )

            b1, b2, b3 = st.columns(3)

            with b1:

                if st.button(
                    "✏️ Edit Invoice",
                    key=f"edit_{row['id']}",
                    use_container_width=True,
                ):

                    # This calls the safe function.
                    # It does NOT modify the radio's
                    # state after the widget is created.
                    start_edit_invoice(
                        row["id"]
                    )

                    st.rerun()

            with b2:

                if st.button(
                    "📄 Generate PDF",
                    key=f"pdf_{row['id']}",
                    use_container_width=True,
                ):

                    invoice, items = (
                        get_invoice(
                            row["id"]
                        )
                    )

                    if invoice:

                        pdf_bytes = (
                            build_invoice_pdf(
                                invoice,
                                items
                            )
                        )

                        st.download_button(
                            "⬇️ Download PDF",
                            data=pdf_bytes,
                            file_name=(
                                "Invoice-"
                                f"{invoice['invoice_number']}.pdf"
                            ),
                            mime="application/pdf",
                            key=(
                                f"download_"
                                f"{row['id']}"
                            ),
                            use_container_width=True,
                        )

            with b3:

                if st.button(
                    "🗑️ Delete",
                    key=f"delete_{row['id']}",
                    use_container_width=True,
                ):

                    delete_invoice(
                        row["id"]
                    )

                    st.success(
                        "Invoice deleted."
                    )

                    st.rerun()


# ============================================================
# EXCEL EXPORT
# ============================================================

elif page == NAV_EXCEL:

    st.markdown(
        '<div class="page-heading">'
        "📊 Excel Export"
        "</div>",
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="gold-line"></div>',
        unsafe_allow_html=True
    )

    rows = get_invoices()

    if not rows:

        st.info(
            "There are no invoices to export."
        )

    else:

        workbook = openpyxl.Workbook()

        sheet = workbook.active

        sheet.title = "Invoices"

        headers = [
            "Invoice Number",
            "Date",
            "Customer",
            "Address",
            "Added By",
            "Payment Method",
            "Cash Received",
            "Returned",
            "Delivery Charges",
            "Discount",
            "Subtotal",
            "Grand Total",
            "Total Paid",
            "Total Due",
            "Remarks",
            "Google Drive ID",
            "Google Drive Name",
        ]

        sheet.append(
            headers
        )

        for row in rows:

            sheet.append(
                [
                    row.get(
                        "invoice_number",
                        ""
                    ),
                    row.get(
                        "invoice_date",
                        ""
                    ),
                    row.get(
                        "customer_name",
                        ""
                    ),
                    row.get(
                        "customer_address",
                        ""
                    ),
                    row.get(
                        "created_by",
                        ""
                    ),
                    row.get(
                        "payment_method",
                        ""
                    ),
                    row.get(
                        "cash_received",
                        0
                    ),
                    row.get(
                        "returned_amount",
                        0
                    ),
                    row.get(
                        "delivery_charges",
                        0
                    ),
                    row.get(
                        "total_discount",
                        0
                    ),
                    row.get(
                        "subtotal",
                        0
                    ),
                    row.get(
                        "grand_total",
                        0
                    ),
                    row.get(
                        "total_paid",
                        0
                    ),
                    row.get(
                        "total_due",
                        0
                    ),
                    row.get(
                        "remarks",
                        ""
                    ),
                    row.get(
                        "google_drive_id",
                        ""
                    ),
                    row.get(
                        "google_drive_name",
                        ""
                    ),
                ]
            )

        # ----------------------------------------------------
        # Excel formatting
        # ----------------------------------------------------

        header_fill = (
            openpyxl.styles.PatternFill(
                "solid",
                fgColor="164C3E"
            )
        )

        header_font = (
            openpyxl.styles.Font(
                bold=True,
                color="FFFFFF"
            )
        )

        for cell in sheet[1]:

            cell.fill = header_fill
            cell.font = header_font

        sheet.freeze_panes = "A2"

        sheet.auto_filter.ref = (
            sheet.dimensions
        )

        for column in sheet.columns:

            max_length = 0

            column_letter = (
                column[0].column_letter
            )

            for cell in column:

                try:

                    max_length = max(
                        max_length,
                        len(
                            str(
                                cell.value
                            )
                        )
                    )

                except Exception:
                    pass

            sheet.column_dimensions[
                column_letter
            ].width = min(
                max_length + 2,
                35
            )

        output = io.BytesIO()

        workbook.save(
            output
        )

        output.seek(0)

        st.download_button(
            "⬇️ Download Excel",
            data=output.getvalue(),
            file_name=(
                "Dar-Makkah-Invoices.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="professional-footer">
        <div style="color:#c3a45c;font-size:16px;">
            ۞
        </div>
        Dar Makkah International
        <br/>
        Invoice Management System
    </div>
    """,
    unsafe_allow_html=True
)
