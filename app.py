import streamlit as st
import pdfplumber
import pandas as pd
import re

# ====== CSS Custom UI Hijau ======
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #a8e063, #56ab2f);
    color: #ffffff;
    font-family: "Poppins", sans-serif;
}
header {
    text-align: center;
    padding: 20px;
    margin-bottom: 20px;
}
header img {
    width: 100px;
    border-radius: 8px;
}
header h1 {
    margin: 10px 0 0;
    font-size: 28px;
    font-weight: 700;
}
.css-1fcdlh1 {
    border: 2px dashed #145A32 !important;
    border-radius: 15px !important;
    background-color: rgba(255, 255, 255, 0.1) !important;
    padding: 20px;
}
.stDownloadButton > button, button[kind="primary"] {
    background: linear-gradient(90deg, #56ab2f, #3c8d0d) !important;
    color: white !important;
    border-radius: 12px !important;
    padding: 10px 20px !important;
    font-size: 16px !important;
    font-weight: bold;
    border: none !important;
    transition: 0.3s;
}
.stDownloadButton > button:hover {
    background: linear-gradient(90deg, #3c8d0d, #2f7010) !important;
}
.dataframe {
    border: 2px solid #2f7010;
    border-radius: 12px;
    overflow: hidden;
    background-color: white;
    color: black;
}
</style>
""", unsafe_allow_html=True)

# ====== Header dengan Logo IAAS ======
st.markdown(f"""
<header>
    <img src="https://raw.githubusercontent.com/faustusblack/ParsingIAAS/main/logo.png">
    <h1>IAAS LC UNPAD – Rekap Data Formulir PDF</h1>
</header>
""", unsafe_allow_html=True)

# ====== Upload PDF ======
uploaded_files = st.file_uploader(
    "Upload satu atau beberapa PDF", type="pdf", accept_multiple_files=True
)

# ====== Fields yang di-parse ======
fields = [
    "Full name :", "Nickname :", "NIM :", "Faculty/Major/Batch :",
    "Place/date of birth :", "Gender :", "Current address :",
    "Original Address :", "Address :", "Phone number :",
    "ID Line/WA/etc :", "Email address :"
]

# ====== Parsing PDF ======
if uploaded_files:
    all_data = []
    
    # Definisikan kolom untuk Field of Interest
    interest_columns = ["Most Interested", "Interested", "Quite Interested", "Less Interested", "Not Interested"]
    interest_fields = {
        "Human development": "",
        "Social awareness/people empowerment": "",
        "Design and public relations": ""
    }

    for uploaded_file in uploaded_files:
        form_data = {f.replace(" :", ""): "" for f in fields}
        form_data["File"] = uploaded_file.name

        # Tambahkan kolom untuk 'Field of Interest'
        form_data.update(interest_fields)

        with pdfplumber.open(uploaded_file) as pdf:
            text_all = "\n".join(page.extract_text() or "" for page in pdf.pages)
            
            # --- Bagian baru untuk parsing field of interest ---
            
            # Mendapatkan data teks dari tabel "FIELD OF INTEREST"
            for page in pdf.pages:
                tables = page.find_tables()
                for table in tables:
                    table_text = table.extract()
                    if table_text and len(table_text) > 2 and "FIELD OF INTEREST" in table_text[0][0]:
                        # Melewati header tabel
                        for row in table_text[2:]: # Mulai dari baris ke-3 (setelah header)
                            if not row[0]: # Lewati baris kosong
                                continue
                            field_name = row[0].replace("\n", " ").strip()
                            
                            # Cari tanda centang (✓) di setiap kolom
                            for i, column_name in enumerate(interest_columns):
                                # Pastikan indeks kolom tidak melebihi batas
                                if i + 1 < len(row):
                                    if "✓" in (row[i + 1] or ""):
                                        # Simpan nama kolom yang memiliki centang
                                        form_data[field_name] = column_name
            
            # --- Akhir bagian baru ---

            # Parsing data personal seperti sebelumnya
            for f in fields:
                for line in text_all.split("\n"):
                    if line.strip().startswith(f):
                        form_data[f.replace(" :", "")] = line.replace(f, "").strip()

        all_data.append(form_data)

    df = pd.DataFrame(all_data)

    # Mengubah urutan kolom agar 'Field of Interest' muncul setelah 'Email address'
    new_column_order = [
        "File", "Full name", "Nickname", "NIM", "Faculty/Major/Batch",
        "Place/date of birth", "Gender", "Current address",
        "Original Address", "Address", "Phone number",
        "ID Line/WA/etc", "Email address",
        "Human development", "Social awareness/people empowerment",
        "Design and public relations"
    ]
    df = df[new_column_order]

    # ====== Tampilkan Hasil ======
    st.subheader("📋 Hasil Rekap Data")
    st.dataframe(df, use_container_width=True)

    # ====== Download Excel ======
    output_file = "rekap_semua_formulir.xlsx"
    df.to_excel(output_file, index=False)
    with open(output_file, "rb") as f:
        st.download_button(
            "⬇️ Download Excel",
            f,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
