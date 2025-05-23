import streamlit as st
import pandas as pd
import io
import base64
import json
import os
import numpy as np


def parse_kode_mesin_Kamboja(file): 
    try:
        import pandas as pd
        import streamlit as st
        import json
        
        # Baca file Excel
        df = pd.read_excel(file, header=None)
        
        # Simpan dataframe asli untuk JSON
        original_df = df.copy()
        
        # Hapus baris yang berisi "Kalibrasi Ulang" di kolom D (index 3)
        filtered_display_df = df[~df[3].astype(str).str.contains("Kalibrasi Ulang", na=False)]
        
        # Dictionary untuk menyimpan kode mesin dan batch yang terkait
        mesin_map = {}
        current_mesin = None
        
        # Loop melalui seluruh baris data ORIGINAL untuk JSON (bukan yang difilter)
        for i in range(len(original_df)):
            # Periksa apakah sel di kolom D (index 3) berisi "Kode Mesin"
            if str(original_df.iloc[i, 3]).strip() == "Kode Mesin":
                # Ambil kode mesin dari kolom F (index 5) di baris yang sama
                current_mesin = str(original_df.iloc[i, 5]).strip()
                
                # Inisialisasi list untuk mesin ini jika belum ada
                if current_mesin not in mesin_map:
                    mesin_map[current_mesin] = []
                    
            # Jika sudah ada mesin yang aktif, periksa nomor batch
            elif current_mesin is not None:
                # Ambil nomor batch dari kolom A (index 0)
                batch = str(original_df.iloc[i, 0]).strip()
                
                # Tambahkan batch yang valid ke list mesin saat ini
                if batch and batch.upper() != "NAN" and not pd.isna(batch):
                    mesin_map[current_mesin].append(batch)
        
        # Buat ringkasan data kode mesin dalam bentuk tabel
        summary_data = []
        for mesin, batches in mesin_map.items():
            summary_data.append({
                "Kode Mesin": mesin,
                "Jumlah Batch": len(batches)
            })
        
        summary_df = pd.DataFrame(summary_data)
        
        # Tampilkan dalam bentuk tabel ringkas
        st.write("🔍 Ringkasan Kode Mesin yang Ditemukan:")
        st.dataframe(summary_df)
        
        # Ubah dict menjadi DataFrame untuk tampilan batch
        # Menentukan panjang maksimum untuk padding list yang lebih pendek
        max_length = max([len(v) for v in mesin_map.values()]) if mesin_map else 0
        
        # Pad list yang lebih pendek dengan NaN
        padded_mesin_map = {}
        for k in mesin_map:
            padded_list = mesin_map[k].copy()
            if len(padded_list) < max_length:
                padded_list.extend([None] * (max_length - len(padded_list)))
            padded_mesin_map[k] = padded_list
        
        # Buat DataFrame untuk tampilan
        result_df = pd.DataFrame(padded_mesin_map)
        
        st.write("📊 Detail Batch Berdasarkan Kode Mesin:")
        st.dataframe(result_df)
        
        # Tampilkan informasi jumlah baris
        st.write(f"Jumlah baris asli: {len(df)}")
        st.write(f"Jumlah baris setelah menghapus 'Kalibrasi Ulang': {len(filtered_display_df)}")
        
        # Simpan data original ke session state untuk digunakan di tab lain
        st.session_state.original_tab1_json = json.dumps(mesin_map)
        st.session_state.tab1_json = json.dumps(mesin_map)
        
        # --- BAGIAN FILTERING DAN PENGHAPUSAN --- #
        st.write("### Filter Batch Berdasarkan Kode Mesin")
        
        # Pilih mesin yang ingin disimpan batchnya
        mesin_to_keep = st.multiselect(
            "Pilih kode mesin yang batchnya ingin disimpan:",
            options=list(mesin_map.keys())
        )
        
        # Kumpulkan batch yang akan disimpan
        batches_to_keep = []
        for mesin in mesin_to_keep:
            batches_to_keep.extend([b for b in mesin_map[mesin] if b is not None])
            
        # Terapkan filter jika tombol diklik dan ada mesin yang dipilih
        if mesin_to_keep and st.button("Terapkan Filter"):
            if not batches_to_keep:
                st.warning("Tidak ada batch yang dapat disimpan dari mesin yang dipilih.")
            else:
                # Gunakan filtered_display_df untuk menampilkan hasil filter (tanpa Kalibrasi Ulang)
                filtered_rows = []
                for i in range(len(filtered_display_df)):
                    batch = str(filtered_display_df.iloc[i, 0]).strip()
                    if batch in batches_to_keep or batch == "" or pd.isna(batch) or batch.upper() == "NAN":
                        filtered_rows.append(filtered_display_df.iloc[i])
            
                # Buat DataFrame hasil filter
                filtered_df = pd.DataFrame(filtered_rows)
                
                # Tampilkan hasil
                st.write(f"### Hasil Filter (Menyimpan {len(batches_to_keep)} batch)")
                st.write(f"Jumlah baris sebelum filter: {len(filtered_display_df)}")
                st.write(f"Jumlah baris setelah filter: {len(filtered_df)}")
                st.dataframe(filtered_df)
                
                # Update session state dengan data yang telah difilter
                st.session_state.filtered_df = filtered_df
                
                # Buat dictionary baru hanya dengan mesin yang dipilih user (tanpa pengelompokan)
                filtered_mesin_map = {}
                for mesin in mesin_to_keep:
                    filtered_mesin_map[mesin] = [b for b in mesin_map[mesin] if b is not None]

                # Simpan langsung hasil user ke session untuk Tab2
                st.session_state.filtered_tab1_json = json.dumps(filtered_mesin_map)
                
                return filtered_df
            
        # Return filtered DataFrame untuk tampilan (tanpa Kalibrasi Ulang)
        return filtered_display_df

    except Exception as e:
        st.error(f"Gagal parsing file: {str(e)}")
        st.exception(e)  # Tampilkan detail error untuk debugging
        return None
    
def parse_kode_mesin_Vietnam(file): 

    try:

        import pandas as pd
        import streamlit as st
        import json        
        # Baca file Excel
        df = pd.read_excel(file, header=None)
        
        # Simpan dataframe asli untuk JSON
        original_df = df.copy()

        # Dictionary untuk menyimpan batch dari kolom A dengan label tetap "Olsa Mames"
        vietnam_batches = [] 
        # Loop melalui baris data mulai dari indeks 1 (baris kedua)
        # untuk mengabaikan baris pertama (indeks 0)
        for i in range(1, len(original_df)):
            # Ambil nomor batch dari kolom A (index 0)
            batch = str(original_df.iloc[i, 0]).strip()
            # Tambahkan batch yang valid ke list
            if batch and batch.upper() != "NAN" and not pd.isna(batch):
                vietnam_batches.append(batch)
            # Buat dictionary dengan satu kunci "Olsa Mames" dan semua batch
        mesin_map = {"Olsa Mames": vietnam_batches}
        

        # Buat ringkasan data
        summary_data = [{"Kode Mesin": "Olsa Mames", "Jumlah Batch": len(vietnam_batches)}]
        summary_df = pd.DataFrame(summary_data)


        # Tampilkan dalam bentuk tabel ringkas
        st.write("🔍 Ringkasan Vietnam:")
        st.dataframe(summary_df)
        # Buat DataFrame untuk tampilan batch
        result_df = pd.DataFrame({"Olsa Mames": pd.Series(vietnam_batches)})


        st.write("📊 Detail Batch Vietnam:")
        st.dataframe(result_df)

    
        # Tampilkan informasi jumlah baris
        st.write(f"Jumlah baris asli: {len(df)}")
        st.write(f"Jumlah batch unik: {len(vietnam_batches)}")

        
        # Simpan data ke session state untuk digunakan di tab lain
        st.session_state.original_tab1_json = json.dumps(mesin_map)
        st.session_state.tab1_json = json.dumps(mesin_map)

        

        # Return DataFrame untuk tampilan
        return result_df
    except Exception as e:
        st.error(f"Error saat memproses file Vietnam: {str(e)}")
        return None
    except Exception as e:
        st.error(f"Gagal parsing file Vietnam: {str(e)}")
        st.exception(e)  # Tampilkan detail error untuk debugging
        return None
def parse_nama_mesin_tab2(file):
    try:
        from difflib import SequenceMatcher
        import pandas as pd
        import streamlit as st
        import json
        import re
        import datetime
        import numpy as np
        import os
        import base64
        import io
        
        # Fungsi untuk membuat tombol download Excel
        def export_dataframe(df, filename="data_export", sheet_name="Sheet1"):
            """
            Membuat tombol download untuk DataFrame
            """
            # Konversi df ke Excel dalam memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Konversi ke bytes dan encode ke base64
            b64 = base64.b64encode(output.getvalue()).decode()
            
            # Membuat link download
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}.xlsx" class="btn" style="background-color:#4CAF50;color:white;padding:8px 12px;text-decoration:none;border-radius:4px;">📥 Download {filename}.xlsx</a>'
            
            return href

        df = pd.read_excel(file, header=None)

        def similarity_score(str1, str2):
            str1_norm = ' '.join(str1.lower().split()) if isinstance(str1, str) else ""
            str2_norm = ' '.join(str2.lower().split()) if isinstance(str2, str) else ""
            if not str1_norm or not str2_norm:
                return 0
            score = SequenceMatcher(None, str1_norm, str2_norm).ratio()
            substring_bonus = 0.2 if str1_norm in str2_norm or str2_norm in str1_norm else 0
            return score + substring_bonus

        # Initialize variables
        all_machine_names = []
        mesin_batch_groups = {}
        mesin_original = {}
        batch_machine_mapping = {}
        
        # Get valid batches from tab1 (works for both Vietnam and Kamboja data)
        valid_filter_batches = []
        if 'tab1_json' in st.session_state:
            tab1_data = json.loads(st.session_state.tab1_json)
            for key, batches in tab1_data.items():
                if batches:
                    valid_filter_batches.extend([str(b).strip() for b in batches if b is not None])
        
        # Machine keywords
        machine_keywords = ["hassia", "sacklok", "redatron", "packaging", "machine", "vietnam"]
        
        # Scan the file to identify potential machine names
        for i in range(len(df)):
            for col in range(df.shape[1]):
                cell_value = str(df.iloc[i, col]).strip()
                if cell_value and cell_value.upper() != "NAN" and not pd.isna(cell_value):
                    cell_lower = cell_value.lower()
                    
                    # Look for machine name patterns
                    if any(keyword in cell_lower for keyword in machine_keywords):
                        machine_candidate = cell_value
                        
                        # Add surrounding context if possible
                        if col > 0 and not pd.isna(df.iloc[i, col-1]):
                            prev_text = str(df.iloc[i, col-1]).strip()
                            if prev_text and len(prev_text) < 30:
                                machine_candidate = f"{prev_text} {machine_candidate}"
                        
                        if col < df.shape[1]-1 and not pd.isna(df.iloc[i, col+1]):
                            next_text = str(df.iloc[i, col+1]).strip()
                            if next_text and len(next_text) < 30:
                                machine_candidate = f"{machine_candidate} {next_text}"
                        
                        machine_candidate = " ".join(machine_candidate.split())
                        if len(machine_candidate) > 5 and machine_candidate not in all_machine_names:
                            all_machine_names.append(machine_candidate)
        
        # Add default machine types if not found
        if not all_machine_names:
            all_machine_names.extend(["HASSIA REDATRON", "SACKLOK 00001"])
        
        # Group machine names
        machine_groups = {}
        for name in all_machine_names:
            name_lower = name.lower()
            if "hassia" in name_lower or "redatron" in name_lower:
                machine_groups[name] = "HASSIA REDATRON"
            elif "sacklok" in name_lower:
                machine_groups[name] = "SACKLOK 00001"
            else:
                machine_groups[name] = name
        
        # Process batches and assign to machines
        for i in range(len(df)):
            batch = str(df.iloc[i, 0]).strip()
            if batch and batch.upper() != "NAN" and not pd.isna(batch) and batch != "-":
                if batch not in valid_filter_batches:
                    continue
                
                # Look for machine indicators in the entire row
                row_str = ' '.join([str(df.iloc[i, j]).lower() for j in range(df.shape[1]) if not pd.isna(df.iloc[i, j])])
                
                # Try to find the machine from the row text
                found_machine = False
                for machine_name in machine_groups.keys():
                    if machine_name.lower() in row_str:
                        canonical_machine = machine_groups[machine_name]
                        found_machine = True
                        
                        if canonical_machine not in mesin_batch_groups:
                            mesin_batch_groups[canonical_machine] = []
                        if canonical_machine not in mesin_original:
                            mesin_original[canonical_machine] = []
                        
                        if machine_name not in mesin_original[canonical_machine]:
                            mesin_original[canonical_machine].append(machine_name)
                        
                        mesin_batch_groups[canonical_machine].append(batch)
                        
                        if batch not in batch_machine_mapping:
                            batch_machine_mapping[batch] = []
                        if canonical_machine not in batch_machine_mapping[batch]:
                            batch_machine_mapping[batch].append(canonical_machine)
                        
                        break

        # Clean up empty machine groups
        mesin_batch_groups = {k: v for k, v in mesin_batch_groups.items() if v}

        # Display machine groups
        st.write("### Detail Grup Mesin")
        for canonical, originals in mesin_original.items():
            if canonical in mesin_batch_groups and mesin_batch_groups[canonical]:
                st.write(f"- **{canonical}** (dari: {', '.join(originals)})")

        # Create summary
        summary_data = []
        for mesin, batches in mesin_batch_groups.items():
            summary_data.append({
                "Nama Mesin": mesin,
                "Jumlah Batch": len(batches),
                "Kode/Nama Asli": ", ".join(mesin_original.get(mesin, [mesin]))
            })

        summary_df = pd.DataFrame(summary_data)
        st.write("🔍 Ringkasan Nama Mesin yang Ditemukan:")
        st.dataframe(summary_df)

        # Create display table
        table_data = []
        for mesin, batches in mesin_batch_groups.items():
            if len(batches) > 0:
                batch_sample = ", ".join(batches[:5]) + ("..." if len(batches) > 5 else "")
                table_data.append({
                    "Nama Mesin": mesin,
                    "Jumlah Batch": len(batches),
                    "Contoh Batch": batch_sample,
                    "Kode/Nama Asli": ", ".join(mesin_original.get(mesin, [mesin]))
                })
        
        result_table = pd.DataFrame(table_data)
        st.write("### Pengelompokan Batch Berdasarkan Nama Mesin:")
        st.dataframe(result_table)

        # Store filtered data
        st.session_state['filtered_nama_mesin_map'] = mesin_batch_groups

        # Create uniform dataframe for full batch details
        result_df = None
        if mesin_batch_groups:
            max_length = max([len(v) for v in mesin_batch_groups.values()])
            uniform_data = {}
            for k in mesin_batch_groups:
                if len(mesin_batch_groups[k]) > 0:
                    uniform_data[k] = mesin_batch_groups[k] + [None] * (max_length - len(mesin_batch_groups[k]))
            
            if uniform_data:
                result_df = pd.DataFrame(uniform_data)
                st.write("📊 Detail Lengkap Batch Per Mesin:")
                st.dataframe(result_df)
                
                # Add download buttons
                st.write("### Download Excel per Kategori Mesin")
                for mesin_name, batch_list in mesin_batch_groups.items():
                    if batch_list:
                        mesin_df = pd.DataFrame({
                            "Batch": batch_list,
                            "Mesin": [mesin_name] * len(batch_list)
                        })
                        
                        filename = f"batch_{mesin_name.lower().replace(' ', '_')}"
                        download_button = export_dataframe(mesin_df, filename)
                        
                        col1, col2 = st.columns([1, 3])
                        with col1:
                            st.markdown(f"**{mesin_name}**")
                        with col2:
                            st.markdown(download_button, unsafe_allow_html=True)
                        st.caption(f"{len(batch_list)} batch teridentifikasi")

                # Download all batches
                all_batches = []
                all_mesins = []
                for mesin_name, batch_list in mesin_batch_groups.items():
                    all_batches.extend(batch_list)
                    all_mesins.extend([mesin_name] * len(batch_list))
                
                if all_batches:
                    st.write("### Download Semua Batch")
                    all_df = pd.DataFrame({
                        "Batch": all_batches,
                        "Mesin": all_mesins
                    })
                    download_all = export_dataframe(all_df, "semua_batch_mesin")
                    st.markdown(download_all, unsafe_allow_html=True)
                    st.caption(f"Total {len(all_batches)} batch dari {len(mesin_batch_groups)} mesin")

        # Save reference
        if st.button("Simpan Referensi Nama Mesin ke JSON"):
            try:
                reference_data = {mesin: list(set([b for b in batch_list if b]))
                                for mesin, batch_list in mesin_batch_groups.items()}
                with open("mesin_batch_reference.json", "w") as f:
                    json.dump(reference_data, f)
                st.success("Referensi batch per nama mesin berhasil disimpan ke mesin_batch_reference.json")
            except Exception as e:
                st.error(f"Gagal menyimpan referensi: {str(e)}")

        return result_df

    except Exception as e:
        st.error(f"Gagal parsing file: {str(e)}")
        st.exception(e)
        return None
def save_kode_mesin_batch_reference(mesin_map, filename="kode_mesin_batch_reference.json"):
    """
    Menyimpan referensi batch-kode mesin ke file JSON
    """
    try:
        # Convert dictionary values (lists) to set for unique batch entries
        reference_data = {}
        for mesin, batches in mesin_map.items():
            # Filter out None values and ensure unique entries
            reference_data[mesin] = list(set([b for b in batches if b is not None]))
        
        # Save reference data to JSON file
        with open(filename, 'w') as f:
            json.dump(reference_data, f)
        
        st.success(f"Referensi batch-kode mesin berhasil disimpan ke {filename}")
        return True
    except Exception as e:
        st.error(f"Gagal menyimpan referensi batch-kode mesin: {str(e)}")
        return False

def load_mesin_batch_reference(filename="mesin_batch_reference.json"):
    """
    Memuat referensi batch-mesin dari file JSON
    """
    try:
        if not os.path.exists(filename):
            st.warning(f"File referensi {filename} tidak ditemukan")
            return {}
        
        with open(filename, 'r') as f:
            reference_data = json.load(f)
        
        st.success(f"Referensi batch-mesin berhasil dimuat dari {filename}")
        return reference_data
    except Exception as e:
        st.error(f"Gagal memuat referensi batch-mesin: {str(e)}")
        return {}

def parse_batch_only_file(file):
    """
    Parsing file yang hanya berisi batch tanpa informasi mesin
    """
    try:
        import pandas as pd
        import streamlit as st
        
        df = pd.read_excel(file, header=None)
        
        st.write("Preview 5 baris pertama data:")
        st.dataframe(df.head())
        
        # Kumpulkan semua batch dari kolom A
        batch_list = []
        for i in range(len(df)):
            batch = str(df.iloc[i, 0]).strip()
            if batch and batch.upper() != "NAN" and not pd.isna(batch):
                batch_list.append(batch)
        
        st.write(f"Jumlah batch yang ditemukan: {len(batch_list)}")
        return batch_list
    
    except Exception as e:
        st.error(f"Gagal parsing file batch: {str(e)}")
        st.exception(e)
        return None
        
def pisahkan_data_grinding_berdasarkan_mesin(file_grinding, reference_data):
    """
    Memisahkan file grinding menjadi beberapa bagian berdasarkan batch yang sudah diklasifikasi dengan nama mesin.
    Header "Nomor Batch" selalu di A1, tapi data bisa mulai dari baris 2, 3, dst (skip baris kosong).
    """
    try:
        # Baca file Excel dengan header di baris 1, tapi tetap baca semua baris (termasuk yang kosong)
        df = pd.read_excel(file_grinding, header=0, keep_default_na=False)
        
        # Pastikan kolom "Nomor Batch" ada
        if "Nomor Batch" not in df.columns:
            raise ValueError("Kolom 'Nomor Batch' tidak ditemukan dalam file grinding.")
        
        # Konversi ke string dan bersihkan dari spasi, tapi jangan hilangkan baris kosong dulu
        df["Nomor Batch"] = df["Nomor Batch"].astype(str).str.strip()
        
        # Sekarang hapus baris yang benar-benar kosong di kolom Nomor Batch
        df_clean = df[
            (df["Nomor Batch"] != '') & 
            (df["Nomor Batch"] != 'nan') &
            (df["Nomor Batch"] != 'NaN') &
            (df["Nomor Batch"] != 'None')
        ].copy()
        
        # Reset index setelah filtering
        df_clean.reset_index(drop=True, inplace=True)
        
        # Debug: print untuk melihat data yang terbaca
        print(f"Total baris setelah cleaning: {len(df_clean)}")
        if len(df_clean) > 0:
            print(f"5 batch pertama: {df_clean['Nomor Batch'].head().tolist()}")
        
        # Inisialisasi hasil per mesin
        hasil_per_mesin = {}
        
        # Inisialisasi list untuk menyimpan batch yang sudah diklasifikasi
        batch_terklasifikasi = []
        
        # Untuk setiap mesin dalam referensi, filter data
        for mesin, daftar_batch in reference_data.items():
            # Bersihkan daftar batch dari reference_data
            daftar_batch_bersih = [str(b).strip() for b in daftar_batch if str(b).strip() != '']
            
            # Filter data berdasarkan nomor batch
            df_filtered = df_clean[df_clean["Nomor Batch"].isin(daftar_batch_bersih)]
            
            # Simpan hasil filter
            hasil_per_mesin[mesin] = df_filtered
            
            # Tambahkan batch yang terklasifikasi ke dalam list
            batch_terklasifikasi.extend(df_filtered["Nomor Batch"].tolist())
        
        # Filter data yang tidak terklasifikasi
        df_unclassified = df_clean[~df_clean["Nomor Batch"].isin(batch_terklasifikasi)]
        hasil_per_mesin["Unclassified"] = df_unclassified
        
        return hasil_per_mesin

    except Exception as e:
        st.error(f"Gagal memisahkan file grinding: {str(e)}")
        return {}


# Fungsi untuk mengeksport DataFrame ke Excel
def export_dataframe(df, filename="data_export"):
    """
    Fungsi untuk mengekspor DataFrame ke file Excel yang dapat diunduh
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    b64 = base64.b64encode(output.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}.xlsx">📥 Download Excel File</a>'
    return href

# Fungsi untuk mengeksport beberapa DataFrame ke Excel dalam file yang sama
def export_multiple_dataframes(df_dict, filename="data_export_multi"):
    """
    Fungsi untuk mengekspor beberapa DataFrame ke file Excel yang dapat diunduh
    Setiap DataFrame akan ditempatkan dalam sheet terpisah
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            # Bersihkan nama sheet dari karakter yang tidak valid
            valid_sheet_name = str(sheet_name)[:31].replace('/', '_').replace('\\', '_').replace('?', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_')
            df.to_excel(writer, sheet_name=valid_sheet_name, index=False)
    output.seek(0)
    b64 = base64.b64encode(output.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}.xlsx">📥 Download Excel File (Semua Mesin)</a>'
    return href

# Tambahkan kode ini ke dalam fungsi tampilkan_obat() setelah dengan tab1 dan sebelum tab3

def tampilkan_obat():
    import streamlit as st
    import json
    import io
    import os
    
    st.title("Halaman Produk Obat")
    st.write("Ini adalah halaman CPP OBAT")
    
    # Tampilkan tab untuk memilih mode operasi
    tab1, tab2, tab3 = st.tabs(["Pengelompokan Batch dengan Kode Mesin", "Pengelompokan Batch dengan Nama Mesin", "Pisahkan Hasil Proses per Mesin"])
    
    with tab1:
        # Kode tab1 yang sudah ada...
        st.header("Pengelompokan Batch dengan Kode Mesin 1.2")
        st.write("Upload file Excel yang berisi informasi batch dan kode mesin.")
        
        # Menambahkan radio button untuk memilih opsi
        selected_option = st.radio(
            "Pilih jenis filter batch:",
            [
                "Kamboja",
                "Vietnam"
            ],
            horizontal=True
        )

        # Menampilkan informasi template yang harus digunakan
        template_info = {
            "Kamboja": "Template Excel untuk pengujian Kamboja",
            "Vietnam": "Template Excel untuk pengecekan Vietnam"
        }
        
        st.info(f"Upload file Excel dengan format: {template_info[selected_option]}")

        # File uploader untuk file dengan informasi kode mesin
        uploaded_file = st.file_uploader("Upload file Excel sesuai template", 
                                        type=["xlsx","ods"], 
                                        key="uploader_kode_mesin")
        
        if uploaded_file:
            # Simpan salinan file untuk diproses
            file_copy = io.BytesIO(uploaded_file.getvalue())
            
            st.success(f"File berhasil diupload")
            st.subheader("Hasil Pengelompokan")

            # Parsing file berdasarkan jenis pengujian yang dipilih
            if selected_option == "Vietnam":
                df = parse_kode_mesin_Vietnam(file_copy)
                if df is not None:
                    # Data Vietnam sudah diproses dalam fungsi, dengan satu kunci "Olsa Mames"
                    kode_mesin_batch_dict = {"Olsa Mames": df["Olsa Mames"].dropna().tolist()}
            
            elif selected_option == "Kamboja":
                df = parse_kode_mesin_Kamboja(file_copy)
                if df is not None:
                    # Ekstrak dictionary kode-mesin-batch dari DataFrame
                    kode_mesin_batch_dict = {}
                    for kode_mesin in df.columns:
                        kode_mesin_batch_dict[kode_mesin] = df[kode_mesin].dropna().tolist()
                
            # Tombol untuk menyimpan data referensi
            if st.button("Simpan Referensi Pengelompokan Kode Mesin"):
                # Gunakan hasil filter jika tersedia
                if 'filtered_tab1_json' in st.session_state and selected_option == "Kamboja":
                    st.session_state.tab1_json = st.session_state.filtered_tab1_json
                    save_kode_mesin_batch_reference(json.loads(st.session_state.filtered_tab1_json))
                else:
                    st.session_state.tab1_json = json.dumps(kode_mesin_batch_dict)
                    save_kode_mesin_batch_reference(kode_mesin_batch_dict)
                
                # Tampilkan tombol download jika df tersedia
                if df is not None:
                    filename = "data_batch_" + selected_option.lower()
                    st.markdown(export_dataframe(df, filename), unsafe_allow_html=True)
                    st.success(f"Data siap diunduh. Klik tombol di atas untuk mengunduh file Excel.")

            # Tambahkan tombol untuk menghapus cache JSON
            if st.button("🧹 Hapus Cache JSON Mesin"):
                try:
                    if os.path.exists("kode_mesin_batch_reference.json"):
                        os.remove("kode_mesin_batch_reference.json")
                        st.success("File kode_mesin_batch_reference.json berhasil dihapus")
                    else:
                        st.info("File kode_batch_reference.json tidak ditemukan")
                except Exception as e:
                    st.error(f"Gagal menghapus file: {str(e)}")
                    
    with tab2:
        st.header("Pengelompokan Batch dengan Nama Mesin")
        st.write("Upload file Excel yang berisi daftar batch dan filter berdasarkan data dari Tab 1.")
        
        # Periksa apakah data tab1 tersedia
        tab1_data_available = 'tab1_json' in st.session_state
        
        if not tab1_data_available:
            st.warning("Silakan proses data di Tab 1 terlebih dahulu sebelum menggunakan fitur ini.")
        else:
            # Tampilkan informasi batch yang tersedia dari tab1
            mesin_batch_map = json.loads(st.session_state.tab1_json)
            total_batches = sum(len([b for b in batches if b is not None]) for batches in mesin_batch_map.values())
            
            st.info(f"ℹ️ Data dari Tab 1 tersedia: {len(mesin_batch_map)} kode mesin dengan total {total_batches} batch.")
            
            # File uploader untuk file Excel tab2
            uploaded_file = st.file_uploader("Upload file Excel dengan kolom Nomor Batch", 
                                            type=["xlsx","ods"], 
                                            key="uploader_tab2")
            
            if uploaded_file:
                # Simpan salinan file untuk diproses
                file_copy = io.BytesIO(uploaded_file.getvalue())
                
                st.success(f"File berhasil diupload")
                
                # Parse file - USING THE CORRECT ARGUMENT COUNT
                filtered_df = parse_nama_mesin_tab2(file_copy)
                

    with tab3:
        # Kode tab3 yang sudah ada...
        st.header("Pisahkan Data Proses Berdasarkan Mesin")
        st.write("Upload file proses grinding dan gunakan referensi batch–mesin yang sudah ada.")
        
        ref_file = "mesin_batch_reference.json"
        if not os.path.exists(ref_file):
            st.warning("Referensi nama mesin-batch belum ditemukan.")
        else:
            reference_data = load_mesin_batch_reference(ref_file)
            grinding_file = st.file_uploader("Upload file Proses Grinding", type=["xlsx"])

            if grinding_file:
                grinding_copy = io.BytesIO(grinding_file.getvalue())
                hasil_split = pisahkan_data_grinding_berdasarkan_mesin(grinding_copy, reference_data)

                if hasil_split:
                    st.subheader("Hasil Sortir Proses berdasarkan Nama Mesin")
                    
                    # Buat tabel ringkasan hasil pemisahan
                    data_ringkasan = []
                    for mesin, df_mesin in hasil_split.items():
                        jumlah_data = len(df_mesin) if not df_mesin.empty else 0
                        status = "✅ Ada data" if jumlah_data > 0 else "❌ Tidak ada data"
                        download_link = export_dataframe(df_mesin, f"grinding_{mesin.replace(' ', '_')}") if jumlah_data > 0 else ""
                        
                        data_ringkasan.append({
                            "Nama Mesin": mesin,
                            "Jumlah Data": jumlah_data,
                            "Status": status,
                            "Download": download_link if jumlah_data > 0 else "-"
                        })
                    
                    # Tampilkan dalam bentuk tabel
                    ringkasan_df = pd.DataFrame(data_ringkasan)
                    st.table(ringkasan_df[["Nama Mesin", "Jumlah Data", "Status"]])
                    
                    # Tampilkan tombol download untuk setiap mesin yang memiliki data
                    st.subheader("Download Data per Nama Mesin")
                    for row in data_ringkasan:
                        if row["Jumlah Data"] > 0:
                            st.markdown(f"**{row['Nama Mesin']}** ({row['Jumlah Data']} data: {row['Download']}", unsafe_allow_html=True)
