import streamlit as st
import pandas as pd
from io import BytesIO
import random
import os

st.set_page_config(page_title="Gabung File Excel", layout="centered")

tab1, tab2 = st.tabs(["🗂️ Gabung File", "📖 Panduan"])

with tab1:
    st.markdown("<div style='text-align:center'><h2>🗂️ Gabung File Excel</h2></div>",unsafe_allow_html=True)

    if "uploaded_files" not in st.session_state:
        st.session_state["uploaded_files"] = []
    if "gabung_log" not in st.session_state:
        st.session_state["gabung_log"] = ""
    if "reset_flag" not in st.session_state:
        st.session_state["reset_flag"] = False
    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = "uploader_1"

    st.markdown("#### 1️⃣ Upload File Excel (.xlsx/.xls)")

    uploaded_files = st.file_uploader(
        "Drag & Drop file Excel dari satu folder sekaligus ke sini ⬇️",
        accept_multiple_files=True,
        type=["xlsx", "xls"],
        key=st.session_state["uploader_key"],
        label_visibility="visible"
    )

    if uploaded_files:
        st.session_state["uploaded_files"] = uploaded_files
        st.session_state["reset_flag"] = False

    st.markdown("")
    col1, col2 = st.columns([1,1])
    with col1:
        hapus_semua = st.button("🗑️ Hapus Semua File", use_container_width=True)
    with col2:
        output_filename1 = st.text_input("Nama file output (tanpa ekstensi)", value="GabunganExcel")

    if hapus_semua or st.session_state["reset_flag"]:
        st.session_state["uploaded_files"] = []
        st.session_state["gabung_log"] = ""
        st.session_state["uploader_key"] = f"uploader_{random.randint(1000,9999)}"
        st.session_state["reset_flag"] = False
        st.success("✅ Semua file telah dihapus.")

    gabung = st.button("🚀 Gabungkan Semua File", type="primary", use_container_width=True)

    preview_combined = None
    error_gabungan = None

    if gabung:
        st.session_state["gabung_log"] = ""
        preview_combined = None
        error_gabungan = None

        if not st.session_state["uploaded_files"]:
            st.warning("⚠️ Harap upload minimal satu file Excel.")
        else:
            all_data = []
            mapping_rows = []
            current_row = 1

            for file in st.session_state["uploaded_files"]:
                filename = file.name.lower()
                ext = os.path.splitext(filename)[-1]
                try:
                    if ext == ".xls":
                        try:
                            xls = pd.ExcelFile(file, engine="xlrd")
                        except ImportError:
                            st.session_state["gabung_log"] += (
                                f"❌ {file.name} - Modul 'xlrd' belum terinstall. "
                                "Install dengan perintah: pip install xlrd\n"
                            )
                            continue
                    else:
                        xls = pd.ExcelFile(file, engine="openpyxl")
                except Exception as e:
                    st.session_state["gabung_log"] += f"❌ {file.name} - Gagal membaca file: {type(e).__name__}: {e}\n"
                    continue

                for sheet in xls.sheet_names:
                    try:
                        if ext == ".xls":
                            try:
                                df = pd.read_excel(xls, sheet, engine="xlrd")
                            except ImportError:
                                st.session_state["gabung_log"] += (
                                    f"❌ {file.name} - Modul 'xlrd' belum terinstall. "
                                    "Install dengan perintah: pip install xlrd\n"
                                )
                                continue
                        else:
                            df = pd.read_excel(xls, sheet, engine="openpyxl")
                    except Exception as e:
                        st.session_state["gabung_log"] += f"❌ {file.name} - Sheet '{sheet}' gagal dibaca: {type(e).__name__}: {e}\n"
                        continue
                    if df.empty:
                        st.session_state["gabung_log"] += f"⚠️ {file.name} - Sheet '{sheet}' kosong, dilewati.\n"
                        continue
                    jumlah_baris = len(df)
                    mapping_rows.append({
                        "Nama File": file.name,
                        "Sheet": sheet,
                        "Baris Awal": current_row,
                        "Baris Akhir": current_row + jumlah_baris - 1
                    })
                    current_row += jumlah_baris
                    all_data.append(df)
                    st.session_state["gabung_log"] += f"✅ {file.name} - Sheet '{sheet}' ({jumlah_baris} baris)\n"

            if not all_data:
                st.session_state["gabung_log"] += "\n❌ Tidak ada data yang berhasil digabung."
                error_gabungan = "Tidak ada data yang berhasil digabung."
            else:
                combined = pd.concat(all_data, ignore_index=True)
                total = len(combined)
                st.session_state["gabung_log"] += f"\n📊 Total baris gabungan: {total}"

                preview_combined = combined.head(10)

                # Sheet keterangan mapping
                df_keterangan = pd.DataFrame(mapping_rows)

                # Export hanya XLSX
                max_rows = 65000
                num_sheets = (total // max_rows) + 1
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Sheet data (bisa lebih dari satu jika > 65000 baris)
                    for i in range(num_sheets):
                        part = combined.iloc[i*max_rows:(i+1)*max_rows]
                        part.to_excel(writer, index=False, sheet_name=f"Data_{i+1}")
                    # Sheet keterangan mapping
                    df_keterangan.to_excel(writer, index=False, sheet_name="Keterangan")
                output.seek(0)
                filename = f"{output_filename1.strip() or 'GabunganExcel'}.xlsx"
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                st.download_button(
                    f"📥 Download File Gabungan (XLSX)",
                    data=output,
                    file_name=filename,
                    mime=mime_type,
                    on_click=lambda: st.session_state.update({"reset_flag": True}),
                    use_container_width=True
                )

                # Preview keterangan
                st.markdown("#### Info Baris Gabungan (Sheet Keterangan):")
                st.dataframe(df_keterangan, height=200)

    if st.session_state["gabung_log"]:
        st.text_area("Log Proses", st.session_state["gabung_log"], height=180)

    if preview_combined is not None:
        st.markdown("#### Preview Data Gabungan (10 baris pertama):")
        st.dataframe(preview_combined, height=300)
    elif error_gabungan:
        st.error(error_gabungan)

    st.markdown("---")
    st.markdown("""
    **Tips:**
    - Untuk upload banyak file, buka folder di komputer, block semua file, lalu drag ke area upload.
    - Untuk file .xls, install: `pip install xlrd`
    - Output aplikasi hanya dalam format XLSX (Excel 2007+) yang kompatibel dengan semua versi Excel modern.
    - Jika ada error, cek log yang muncul.
    """)

with tab2:
    st.markdown("<div style='text-align:center'><h2>📖 Panduan</h2></div>",unsafe_allow_html=True)
    st.markdown("""
    <div style='max-width:700px;margin:auto;font-size:18px;'>
    <b>Aplikasi ini untuk menggabungkan banyak file Excel (.xlsx, .xls) jadi satu file rekap otomatis.</b>
    <br><b>Cara penggunaan:</b>
    <ol>
    <li>Kumpulkan file Excel yang ingin digabung dalam satu folder.</li>
    <li>Drag & Drop semua file ke area upload di Tab <b>Gabung File</b>.</li>
    <li>Klik tombol <b>Gabungkan</b> untuk proses.</li>
    <li>Preview hasil dan file rekap siap diunduh.</li>
    </ol>
    <br><b>Penjelasan Kolom & Hasil Gabungan:</b>
    <ul>
    <li><b>Kolom Awal :</b> Semua kolom dari file yang diupload digabung, perbedaan kolom antar file otomatis diberi kolom baru dan data lain dikosongkan (NaN).</li>
    <li><b>Hasil Akhir :</b> File gabungan berisi semua data dari seluruh file, format tabel sudah disesuaikan. Hasil bisa diunduh sebagai <b>XLSX</b> (Excel 2007+).</li>
    <li><b>Sheet Keterangan :</b> Sheet tambahan yang menampilkan mapping baris dari tiap file dan sheet asalnya.</li>
    </ul>
    Jika ada file dengan nama kolom berbeda, aplikasi tetap menampilkan semua kolom agar data tidak hilang.
    </div>
    """, unsafe_allow_html=True)
