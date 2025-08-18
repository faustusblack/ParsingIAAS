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
    
    # Definisikan Field of Interest dan opsi-opsinya
    interest_options_map = {
        "Human development": ["Most Interested", "Interested", "Quite Interested", "Less Interested", "Not Interested"],
        "Social awareness/people empowerment": ["Most Interested", "Interested", "Quite Interested", "Less Interested", "Not Interested"],
        "Design and public relations": ["Most Interested", "Interested", "Quite Interested", "Less Interested", "Not Interested"],
        "Exchange program and language": ["Most Interested", "Interested", "Quite Interested", "Less Interested", "Not Interested"],
        "Science and technology": ["Most Interested", "Interested", "Quite Interested", "Less Interested", "Not Interested"]
    }
    
    # Ambil daftar bidang dan opsi sebagai list
    interest_fields = list(interest_options_map.keys())
    interest_options_header = interest_options_map[interest_fields[0]]
    
    for uploaded_file in uploaded_files:
        form_data = {f.replace(" :", ""): "" for f in fields}
        form_data["File"] = uploaded_file.name
        
        # Tambahkan kolom untuk Field of Interest
        form_data.update({k: "" for k in interest_fields})

        with pdfplumber.open(uploaded_file) as pdf:
            text_all = "\n".join(page.extract_text() or "" for page in pdf.pages)
            
            # --- Logika baru: Deteksi berdasarkan posisi (koordinat) karakter dan string yang dibangun dari char ---
            
            # Mendapatkan semua karakter dan posisinya dari setiap halaman
            all_chars = []
            for page in pdf.pages:
                all_chars.extend(page.chars)

            # Temukan koordinat (y-axis) dari setiap bidang minat
            field_coords_map = {}
            for field in interest_fields:
                field_chars = [c for c in all_chars if c['text'].strip() and c['text'] in field]
                if field_chars:
                    # Temukan posisi y0 dari karakter pertama dari bidang
                    field_coords_map[field] = field_chars[0]['y0']

            # Temukan koordinat (x-axis) dari setiap opsi header
            option_coords_map = {}
            for option in interest_options_header:
                option_chars = [c for c in all_chars if c['text'].strip() and c['text'] in option]
                if option_chars:
                    # Temukan posisi x0 dari karakter pertama dari opsi
                    option_coords_map[option] = option_chars[0]['x0']

            # Cari tanda centang (✓) dan cocokkan dengan koordinat
            for char in all_chars:
                if 'text' in char and char['text'] == '✓':
                    checkmark_x = char['x0']
                    checkmark_y = char['y0']

                    # Cari baris (bidang minat) yang cocok dengan y-koordinat checkmark
                    for field, y_coord in field_coords_map.items():
                        if abs(checkmark_y - y_coord) < 10: # Toleransi vertikal
                            
                            # Cari kolom (opsi) yang cocok dengan x-koordinat checkmark
                            for option, x_coord in option_coords_map.items():
                                if abs(checkmark_x - x_coord) < 20: # Toleransi horizontal
                                    form_data[field] = option
                                    break # Hentikan jika sudah menemukan opsi yang cocok
                            break # Hentikan jika sudah menemukan bidang yang cocok
            
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
        "Design and public relations", "Exchange program and language", "Science and technology"
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
