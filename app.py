import streamlit as st
import pdfplumber
import pandas as pd

st.title("Rekap Data Formulir PDF ke Excel")

uploaded_files = st.file_uploader(
    "Upload satu atau beberapa PDF", type="pdf", accept_multiple_files=True
)

fields = [
    "Full name :",
    "Nickname :",
    "NIM :",
    "Faculty/Major/Batch :",
    "Place/date of birth :",
    "Gender :",
    "Current address :",
    "Original Address :",
    "Address :",
    "Phone number :",
    "ID Line/WA/etc :",
    "Email address :"
]

if uploaded_files:
    all_data = []

    for uploaded_file in uploaded_files:
        form_data = {field.replace(" :", ""): "" for field in fields}
        form_data["File"] = uploaded_file.name

        with pdfplumber.open(uploaded_file) as pdf:
            text_all = ""
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                text_all += page_text + "\n"

            for field in fields:
                for line in text_all.split("\n"):
                    if line.strip().startswith(field):
                        value = line.replace(field, "").strip()
                        form_data[field.replace(" :", "")] = value

        all_data.append(form_data)

    df = pd.DataFrame(all_data)
    st.subheader("Hasil Rekap Data")
    st.dataframe(df)

    # Tombol download Excel
    output_file = "rekap_semua_formulir.xlsx"
    df.to_excel(output_file, index=False)
    with open(output_file, "rb") as f:
        st.download_button(
            "Download Excel", f, file_name=output_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
