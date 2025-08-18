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
    interest_options = {
        "Human development": ["Most Interested", "Interested", "Quite Interested", "Less Interested", "Not Interested"],
        "Social awareness/people empowerment": ["Most Interested", "Interested", "Quite Interested", "Less Interested", "Not Interested"],
        "Design and public relations": ["Most Interested", "Interested", "Quite Interested", "Less Interested", "Not Interested"]
    }
    
    # Kunci untuk data yang akan diekstrak
    extraction_keys = list(fields) + list(interest_options.keys())

    for uploaded_file in uploaded_files:
        form_data = {f.replace(" :", ""): "" for f in fields}
        form_data["File"] = uploaded_file.name
        
        # Tambahkan kolom untuk Field of Interest
        form_data.update({k: "" for k in interest_options.keys()})

        with pdfplumber.open(uploaded_file) as pdf:
            text_all = ""
            for page in pdf.pages:
                text_all += page.extract_text() or ""
                
                # --- Logika baru untuk deteksi checkbox ---
                # Coba deteksi tanda centang (✓) di setiap halaman
                for key, options in interest_options.items():
                    # Mencari pola "Nama Bidang ✓"
                    for option in options:
                        pattern = re.compile(re.escape(key) + r'\s+' + re.escape(option) + r'\s*✓', re.IGNORECASE)
                        if pattern.search(text_all):
                            form_data[key] = option
                            break # Hentikan pencarian jika sudah ditemukan
            
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
