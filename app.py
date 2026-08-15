import streamlit as st
import pdfplumber
import pandas as pd
import re
import fitz
import cv2
import numpy as np
from io import BytesIO

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="IAAS LC UNPAD – Registration Parser",
    page_icon="🌱",
    layout="wide"
)

# =========================================================
# CUSTOM CSS
# =========================================================

st.markdown("""
<style>

.stApp {
    background: linear-gradient(135deg, #a8e063, #56ab2f);
    color: white;
    font-family: "Poppins", sans-serif;
}

.main {
    padding-top: 20px;
}

.header-container {
    text-align: center;
    padding: 20px;
    margin-bottom: 20px;
}

.header-container img {
    width: 110px;
    margin-bottom: 10px;
}

.header-container h1 {
    font-size: 28px;
    font-weight: 700;
    margin: 5px 0;
}

.header-container p {
    font-size: 15px;
    opacity: 0.9;
}

.stFileUploader {
    background-color: rgba(255,255,255,0.10);
    border-radius: 15px;
    padding: 10px;
}

.stDownloadButton > button {
    background: linear-gradient(90deg, #56ab2f, #3c8d0d) !important;
    color: white !important;
    border-radius: 12px !important;
    padding: 10px 20px !important;
    font-size: 16px !important;
    font-weight: bold !important;
    border: none !important;
}

.stDownloadButton > button:hover {
    background: linear-gradient(90deg, #3c8d0d, #2f7010) !important;
}

div[data-testid="stDataFrame"] {
    border-radius: 12px;
    overflow: hidden;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# HEADER + LOGO
# =========================================================

st.markdown("""
<div class="header-container">
    <img src="https://raw.githubusercontent.com/faustusblack/ParsingIAAS/main/logo.png">
    <h1>IAAS LC UNPAD</h1>
    <p>Registration Form Data Parser</p>
</div>
""", unsafe_allow_html=True)


# =========================================================
# FIELD CONFIGURATION
# =========================================================

BIODATA_FIELDS = [
    "Full Name",
    "Preferred Name",
    "NPM",
    "Faculty",
    "Major",
    "Batch / Year of Entry",
    "Gender",
    "Current Address",
    "Original Address",
    "Current Living Arrangement",
    "WhatsApp Number",
    "Personal Email",
    "UNPAD Email",
    "Line ID"
]

INTEREST_FIELDS = [
    "Human development",
    "Social awareness/people (community) empowerment",
    "Design and public relations",
    "Exchange program and language",
    "Science and technology"
]

INTEREST_CHOICES = [
    "Most Interested",
    "Interested",
    "Quite Interested",
    "Less Interested",
    "Not Interested"
]


# =========================================================
# PDF TEXT EXTRACTION
# =========================================================

def extract_pdf(file_bytes):

    text = ""
    pages_data = []

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:

        for page in pdf.pages:

            page_text = page.extract_text() or ""

            text += page_text + "\n"

            pages_data.append({
                "page": page,
                "text": page_text,
                "words": page.extract_words() or [],
                "chars": page.chars or []
            })

    return text, pages_data


# =========================================================
# NORMALIZE TEXT
# =========================================================

def normalize_text(text):

    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)

    return text


# =========================================================
# GET VALUE FROM FIELD
# =========================================================

def get_field_value(text, field):

    """
    Mengambil nilai setelah:
    Field :
    """

    pattern = rf"{re.escape(field)}\s*:\s*(.*)"

    match = re.search(
        pattern,
        text,
        re.IGNORECASE
    )

    if match:

        value = match.group(1).strip()

        return value

    return ""


# =========================================================
# PARSE BIODATA
# =========================================================

def parse_biodata(text):

    text = normalize_text(text)

    data = {}

    # -----------------------------------------------------
    # Full Name
    # -----------------------------------------------------

    data["Full Name"] = get_field_value(
        text,
        "Full Name"
    )

    # -----------------------------------------------------
    # Preferred Name
    # -----------------------------------------------------

    data["Preferred Name"] = get_field_value(
        text,
        "Preferred Name"
    )

    # -----------------------------------------------------
    # NPM
    # -----------------------------------------------------

    data["NPM"] = get_field_value(
        text,
        "NPM"
    )

    # -----------------------------------------------------
    # Faculty
    # -----------------------------------------------------

    data["Faculty"] = get_field_value(
        text,
        "Faculty"
    )

    # -----------------------------------------------------
    # Major
    # -----------------------------------------------------

    data["Major"] = get_field_value(
        text,
        "Major"
    )

    # -----------------------------------------------------
    # Batch
    # -----------------------------------------------------

    batch_match = re.search(
        r"Batch\s*/\s*Year of Entry\s*:\s*(.*)",
        text,
        re.IGNORECASE
    )

    if batch_match:

        batch_text = batch_match.group(1).strip()

        # Kalau checkbox ikut terbaca
        years = re.findall(
            r"20\d{2}",
            batch_text
        )

        if years:
            data["Batch / Year of Entry"] = years[0]
        else:
            data["Batch / Year of Entry"] = batch_text

    else:

        data["Batch / Year of Entry"] = ""

    # -----------------------------------------------------
    # Gender
    # -----------------------------------------------------

    gender_match = re.search(
        r"Gender\s*:\s*(.*)",
        text,
        re.IGNORECASE
    )

    if gender_match:

        gender_text = gender_match.group(1).strip()

        # Coba cari checkbox yang tercentang
        if re.search(r"Female", gender_text, re.IGNORECASE):
            data["Gender"] = "Female"

        elif re.search(r"Male", gender_text, re.IGNORECASE):
            data["Gender"] = "Male"

        elif re.search(r"Other", gender_text, re.IGNORECASE):
            data["Gender"] = "Other"

        else:
            data["Gender"] = gender_text

    else:

        data["Gender"] = ""

    # -----------------------------------------------------
    # Address
    # -----------------------------------------------------

    data["Current Address"] = get_field_value(
        text,
        "Current address"
    )

    data["Original Address"] = get_field_value(
        text,
        "Original address"
    )

    # -----------------------------------------------------
    # Current Living Arrangement
    # -----------------------------------------------------

    living_match = re.search(
        r"Current Living\s+Arrangement\s*:\s*(.*)",
        text,
        re.IGNORECASE
    )

    if living_match:

        living_text = living_match.group(1).strip()

        data["Current Living Arrangement"] = living_text

    else:

        data["Current Living Arrangement"] = ""

    # -----------------------------------------------------
    # CONTACT
    # -----------------------------------------------------

    data["WhatsApp Number"] = get_field_value(
        text,
        "WhatsApp Number"
    )

    data["Personal Email"] = get_field_value(
        text,
        "Personal Email"
    )

    data["UNPAD Email"] = get_field_value(
        text,
        "UNPAD Email"
    )

    data["Line ID"] = get_field_value(
        text,
        "Line ID"
    )

    return data


# =========================================================
# INTEREST PARSER
# =========================================================

def find_interest_position(
    page_data,
    interest_keywords
):

    """
    Cari posisi horizontal field interest
    menggunakan pdfplumber.

    Karena checkbox pada formulir terbaru
    tidak selalu berada di baris teks yang sama,
    kita menggunakan koordinat PDF.
    """

    words = page_data["words"]

    matched_words = []

    for word in words:

        word_text = word["text"].lower()

        for keyword in interest_keywords:

            if keyword.lower() in word_text:

                matched_words.append(word)

    if not matched_words:

        return None

    # posisi Y rata-rata teks interest

    y_position = sum(
        (w["top"] + w["bottom"]) / 2
        for w in matched_words
    ) / len(matched_words)

    # posisi X awal teks
    x_position = min(
        w["x0"]
        for w in matched_words
    )

    return x_position, y_position


# =========================================================
# DETECT CHECKBOX
# =========================================================

def detect_checkboxes_near_y(
    page_data,
    target_y,
    tolerance=25
):

    checks = []

    for char in page_data["chars"]:

        char_text = char.get("text", "").strip()

        # karakter yang kemungkinan merupakan checkbox/checkmark
        if char_text in ["✓", "✔", "√", "V", "v", "☑"]:

            y = (
                char["top"] +
                char["bottom"]
            ) / 2

            if abs(y - target_y) <= tolerance:

                x = (
                    char["x0"] +
                    char["x1"]
                ) / 2

                checks.append(x)

    return sorted(checks)


# =========================================================
# CLUSTER CHECKBOX COLUMNS
# =========================================================

def detect_choice_columns(all_x):

    if not all_x:
        return []

    all_x = sorted(all_x)

    clusters = []

    current = [all_x[0]]

    for x in all_x[1:]:

        if abs(x - current[-1]) < 25:

            current.append(x)

        else:

            clusters.append(current)

            current = [x]

    clusters.append(current)

    centers = [
        sum(cluster) / len(cluster)
        for cluster in clusters
    ]

    return sorted(centers)


# =========================================================
# PARSE FIELD OF INTEREST
# =========================================================

def parse_interests(pages_data):

    result = {
        field: ""
        for field in INTEREST_FIELDS
    }

    # -----------------------------------------------------
    # Locate page containing Field of Interest
    # -----------------------------------------------------

    interest_page_indices = []

    for i, page in enumerate(pages_data):

        if (
            "FIELD OF INTEREST" in
            page["text"].upper()
        ):

            interest_page_indices.append(i)

        elif (
            "Human" in page["text"]
            and "development" in page["text"]
        ):

            interest_page_indices.append(i)

    if not interest_page_indices:

        # Form terbaru: field bisa terbagi halaman 5-6
        for i, page in enumerate(pages_data):

            page_text = page["text"].lower()

            if (
                "human" in page_text
                or "science" in page_text
                or "exchange" in page_text
            ):

                interest_page_indices.append(i)

    # -----------------------------------------------------
    # Find all checkbox positions
    # -----------------------------------------------------

    all_checkbox_x = []

    for i in interest_page_indices:

        page = pages_data[i]

        for char in page["chars"]:

            char_text = char.get(
                "text",
                ""
            ).strip()

            if char_text in [
                "✓",
                "✔",
                "√",
                "V",
                "v",
                "☑"
            ]:

                x = (
                    char["x0"] +
                    char["x1"]
                ) / 2

                all_checkbox_x.append(x)

    choice_columns = detect_choice_columns(
        all_checkbox_x
    )

    # -----------------------------------------------------
    # Fallback if columns not found
    # -----------------------------------------------------

    if len(choice_columns) < 2:

        return result

    # -----------------------------------------------------
    # Keywords per field
    # -----------------------------------------------------

    keywords = {

        "Human development":
            ["human", "development"],

        "Social awareness/people (community) empowerment":
            [
                "social",
                "awareness",
                "empowerment"
            ],

        "Design and public relations":
            [
                "design",
                "public",
                "relations"
            ],

        "Exchange program and language":
            [
                "exchange",
                "program",
                "language"
            ],

        "Science and technology":
            [
                "science",
                "technology"
            ]
    }

    # -----------------------------------------------------
    # Process every interest
    # -----------------------------------------------------

    for interest, keys in keywords.items():

        best_page = None
        best_y = None

        # Cari teks interest
        for page_index in interest_page_indices:

            page = pages_data[page_index]

            matches = []

            for word in page["words"]:

                word_text = word["text"].lower()

                if any(
                    key.lower() in word_text
                    for key in keys
                ):

                    matches.append(word)

            if matches:

                best_page = page_index

                best_y = sum(
                    (
                        w["top"] +
                        w["bottom"]
                    ) / 2
                    for w in matches
                ) / len(matches)

                break

        if best_page is None:
            continue

        page = pages_data[best_page]

        # Cari checkbox yang sejajar dengan interest
        checkbox_x = detect_checkboxes_near_y(
            page,
            best_y,
            tolerance=35
        )

        if not checkbox_x:
            continue

        # Ambil checkbox terdekat terhadap baris interest
        selected_x = checkbox_x[0]

        # Cari kolom pilihan terdekat
        distances = [
            abs(x - selected_x)
            for x in choice_columns
        ]

        choice_index = distances.index(
            min(distances)
        )

        if choice_index < len(
            INTEREST_CHOICES
        ):

            result[interest] = \
                INTEREST_CHOICES[
                    choice_index
                ]

    return result


# =========================================================
# EXCEL EXPORT
# =========================================================

def create_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Registration Data"
        )

        workbook = writer.book
        worksheet = writer.sheets[
            "Registration Data"
        ]

        # Freeze header
        worksheet.freeze_panes(
            1,
            0
        )

        # Header formatting
        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "top"
        })

        for col_num, value in enumerate(
            df.columns
        ):

            worksheet.write(
                0,
                col_num,
                value,
                header_format
            )

            worksheet.set_column(
                col_num,
                col_num,
                22
            )

    return output.getvalue()


# =========================================================
# UPLOAD
# =========================================================

uploaded_files = st.file_uploader(
    "📄 Upload satu atau beberapa Registration Form PDF",
    type=["pdf"],
    accept_multiple_files=True
)


# =========================================================
# MAIN PARSING
# =========================================================

if uploaded_files:

    all_data = []

    progress = st.progress(0)

    total_files = len(
        uploaded_files
    )

    for file_index, uploaded_file in enumerate(
        uploaded_files
    ):

        try:

            # ---------------------------------------------
            # Read PDF
            # ---------------------------------------------

            file_bytes = uploaded_file.read()

            text, pages_data = extract_pdf(
                file_bytes
            )

            # ---------------------------------------------
            # Biodata
            # ---------------------------------------------

            biodata = parse_biodata(
                text
            )

            # ---------------------------------------------
            # Interests
            # ---------------------------------------------

            interests = parse_interests(
                pages_data
            )

            # ---------------------------------------------
            # Combine
            # ---------------------------------------------

            row = {}

            row["File"] = uploaded_file.name

            # Biodata
            for field in BIODATA_FIELDS:

                row[field] = biodata.get(
                    field,
                    ""
                )

            # Interest
            for field in INTEREST_FIELDS:

                row[field] = interests.get(
                    field,
                    ""
                )

            all_data.append(row)

        except Exception as e:

            st.error(
                f"❌ Gagal memproses {uploaded_file.name}: {e}"
            )

        progress.progress(
            (file_index + 1) / total_files
        )

    # =====================================================
    # DATAFRAME
    # =====================================================

    if all_data:

        df = pd.DataFrame(
            all_data
        )

        st.success(
            f"✅ {len(all_data)} formulir berhasil diproses."
        )

        # -------------------------------------------------
        # Preview
        # -------------------------------------------------

        st.subheader(
            "📋 Hasil Rekap Data"
        )

        st.dataframe(
            df,
            use_container_width=True,
            height=500
        )

        # -------------------------------------------------
        # Interest summary
        # -------------------------------------------------

        st.subheader(
            "🎯 Field of Interest"
        )

        interest_preview = df[
            [
                "Full Name"
            ] + INTEREST_FIELDS
        ]

        st.dataframe(
            interest_preview,
            use_container_width=True
        )

        # -------------------------------------------------
        # Excel
        # -------------------------------------------------

        excel_data = create_excel(
            df
        )

        st.download_button(
            label="⬇️ Download Excel",
            data=excel_data,
            file_name=(
                "IAAS_Registration_Rekap.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            )
        )
