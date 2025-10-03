import streamlit as st
import pandas as pd
import io

# --- PENGATURAN TAMPILAN HALAMAN ---
st.set_page_config(
    page_title="Analisis Absensi Karyawan",
    page_icon="📊",
    layout="centered"
)

# --- FUNGSI UTAMA UNTUK ANALISIS (diambil dari skrip sebelumnya) ---
def format_dengan_spasi(df, group_by_col='Department'):
    if df is None or df.empty:
        return df
    df = df.sort_values(by=[group_by_col, 'Name', 'Date' if 'Date' in df.columns else 'Tanggal'])
    all_dfs = []
    empty_row = pd.DataFrame([[pd.NA] * len(df.columns)], columns=df.columns)
    for i, group_name in enumerate(df[group_by_col].unique()):
        df_group = df[df[group_by_col] == group_name]
        all_dfs.append(df_group)
        if i < len(df[group_by_col].unique()) - 1:
            all_dfs.append(empty_row)
    return pd.concat(all_dfs, ignore_index=True)

def analisis_absensi_lanjutan(file_obj, min_durasi_menit):
    try:
        if file_obj.name.endswith('.xlsx'):
            df = pd.read_excel(file_obj)
        elif file_obj.name.endswith('.csv'):
            df = pd.read_csv(file_obj)
        else:
            st.error("Format file tidak didukung. Harap gunakan .xlsx atau .csv")
            return None, None, None
        
        df.columns = df.columns.str.strip()
        df_original = df.copy()
        df['Date/Time'] = pd.to_datetime(df['Date/Time'])
        df = df.sort_values(by=['Name', 'Date/Time']).reset_index(drop=True)
        
        valid_shifts = []
        anomalies = []
        for name, group in df.groupby('Name'):
            last_check_in_data = None
            for index, row in group.iterrows():
                department = row['Department']
                if row['Status'] == 'C/In':
                    if last_check_in_data is not None:
                        anomalies.append({'Department': last_check_in_data['Department'],'Name': name,'Tanggal': last_check_in_data['Date/Time'].date(),'Jenis Anomali': 'Absen Masuk Ganda','Detail': f"C/In pada {last_check_in_data['Date/Time']} tidak memiliki pasangan C/Out."})
                    last_check_in_data = row.to_dict()
                elif row['Status'] == 'C/Out':
                    if last_check_in_data is not None:
                        valid_shifts.append({'Department': department,'Name': name,'Check_In': last_check_in_data['Date/Time'],'Check_Out': row['Date/Time']})
                        last_check_in_data = None
                    else:
                        anomalies.append({'Department': department,'Name': name,'Tanggal': row['Date/Time'].date(),'Jenis Anomali': 'Absen Pulang tanpa Absen Masuk','Detail': f"C/Out pada {row['Date/Time']} tidak memiliki pasangan C/In."})
            if last_check_in_data is not None:
                anomalies.append({'Department': last_check_in_data['Department'],'Name': name,'Tanggal': last_check_in_data['Date/Time'].date(),'Jenis Anomali': 'Tidak ada jam pulang','Detail': f"C/In pada {last_check_in_data['Date/Time']} tidak memiliki pasangan C/Out."})

        rekap_final = pd.DataFrame(valid_shifts)
        anomali_awal = pd.DataFrame(anomalies)
        rekap_valid = pd.DataFrame()
        if not rekap_final.empty:
            rekap_final['Duration_Seconds'] = (rekap_final['Check_Out'] - rekap_final['Check_In']).dt.total_seconds()
            is_valid_duration = rekap_final['Duration_Seconds'] >= (min_durasi_menit * 60)
            rekap_valid = rekap_final[is_valid_duration].copy()
            shift_pendek = rekap_final[~is_valid_duration].copy()
            if not shift_pendek.empty:
                shift_pendek['Tanggal'] = shift_pendek['Check_In'].dt.date
                shift_pendek['Jenis Anomali'] = 'Durasi Shift Kurang dari Standar'
                shift_pendek['Detail'] = shift_pendek.apply(lambda row: f"Total jam kerja hanya {int(row['Duration_Seconds'] // 3600):02}:{int((row['Duration_Seconds'] % 3600) // 60):02}", axis=1)
                anomali_tambahan = shift_pendek[['Department', 'Name', 'Tanggal', 'Jenis Anomali', 'Detail']]
                anomali_final = pd.concat([anomali_awal, anomali_tambahan], ignore_index=True)
            else:
                anomali_final = anomali_awal
        else:
            anomali_final = anomali_awal

        if not rekap_valid.empty:
            rekap_valid['Total_Hours'] = rekap_valid.apply(lambda row: f"{int(row['Duration_Seconds'] // 3600):02}:{int((row['Duration_Seconds'] % 3600) // 60):02}", axis=1)
            rekap_valid['Date'] = pd.to_datetime(rekap_valid['Check_In']).dt.strftime('%Y-%m-%d')
            rekap_valid['Check_In'] = rekap_valid['Check_In'].dt.strftime('%Y-%m-%d %H:%M:%S')
            rekap_valid['Check_Out'] = rekap_valid['Check_Out'].dt.strftime('%Y-%m-%d %H:%M:%S')
            rekap_valid = rekap_valid[['Department', 'Name', 'Date', 'Check_In', 'Check_Out', 'Total_Hours']]
        if not anomali_final.empty:
            anomali_final['Tanggal'] = pd.to_datetime(anomali_final['Tanggal']).dt.strftime('%Y-%m-%d')
        return df_original, rekap_valid, anomali_final
    except Exception as e:
        st.error(f"Terjadi error saat memproses file: {e}")
        return None, None, None

# --- TAMPILAN APLIKASI WEB (UI) ---
st.title("📊 Aplikasi Analisis Data Absensi")
st.write("Unggah file absensi mentah Anda untuk mendapatkan rekap shift yang valid, laporan anomali, dan data mentah dalam satu file Excel.")

# Opsi di sidebar
with st.sidebar:
    st.header("⚙️ Pengaturan")
    min_durasi = st.number_input(
        "Durasi shift valid minimum (menit)", 
        min_value=0, 
        value=480,  # Default 8 jam
        step=15,
        help="Shift yang lebih pendek dari ini akan dianggap anomali."
    )
    
# File uploader
uploaded_file = st.file_uploader(
    "Pilih file absensi (.xlsx atau .csv)", 
    type=['xlsx', 'csv']
)

if uploaded_file is not None:
    st.success(f"File **{uploaded_file.name}** berhasil diunggah.")
    
    if st.button("🚀 Proses Sekarang!"):
        with st.spinner("Mohon tunggu, sedang menganalisis data..."):
            data_mentah, rekap_final, anomali_final = analisis_absensi_lanjutan(uploaded_file, min_durasi)

            if data_mentah is not None:
                # Format dengan spasi
                rekap_final_spaced = format_dengan_spasi(rekap_final, 'Department')
                anomali_final_spaced = format_dengan_spasi(anomali_final, 'Department')
                
                # Simpan ke memori dalam format Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    if rekap_final_spaced is not None and not rekap_final_spaced.empty:
                        rekap_final_spaced.to_excel(writer, sheet_name='Rekap Shift Valid', index=False)
                    if anomali_final_spaced is not None and not anomali_final_spaced.empty:
                        anomali_final_spaced.to_excel(writer, sheet_name='Laporan Anomali', index=False)
                    else:
                        pd.DataFrame([{'Status': 'Tidak ada anomali yang ditemukan'}]).to_excel(writer, sheet_name='Laporan Anomali', index=False)
                    data_mentah.to_excel(writer, sheet_name='Data Mentah', index=False)
                
                output.seek(0)
                st.balloons()
                st.header("✅ Analisis Selesai!")
                
                # Tombol download
                st.download_button(
                    label="📥 Unduh Hasil Analisis (File Excel)",
                    data=output,
                    file_name='Hasil_Analisis_Absensi.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )