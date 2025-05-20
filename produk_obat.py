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
                
                # Buat dictionary baru hanya dengan mesin yang dipilih
                filtered_mesin_map = {}
                for mesin in mesin_to_keep:
                    filtered_mesin_map[mesin] = [b for b in mesin_map[mesin] if b is not None]
                    
                # Update session state untuk tab2
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
                # Gunakan result_df sebagai sumber untuk filtering
                filtered_rows = []
                for i in range(len(original_df)):
                    if i == 0:  # Pertahankan header
                        filtered_rows.append(original_df.iloc[i])
                    else:
                        batch = str(original_df.iloc[i, 0]).strip()
                        if batch in batches_to_keep or batch == "" or pd.isna(batch) or batch.upper() == "NAN":
                            filtered_rows.append(original_df.iloc[i])
            
                # Buat DataFrame hasil filter
                filtered_df = pd.DataFrame(filtered_rows)
                
                # Tampilkan hasil
                st.write(f"### Hasil Filter (Menyimpan {len(batches_to_keep)} batch)")
                st.write(f"Jumlah baris sebelum filter: {len(original_df)}")
                st.write(f"Jumlah baris setelah filter: {len(filtered_df)}")
                st.dataframe(filtered_df)
                
                # Update session state dengan data yang telah difilter
                st.session_state.filtered_df = filtered_df
                
                # Buat dictionary baru hanya dengan mesin yang dipilih
                filtered_mesin_map = {}
                for mesin in mesin_to_keep:
                    filtered_mesin_map[mesin] = [b for b in mesin_map[mesin] if b is not None]
                    
                # Update session state untuk tab Vietnam
                st.session_state.filtered_tab1_json = json.dumps(filtered_mesin_map)
                
                return filtered_df
            
        # Return DataFrame asli jika tidak ada filtering
        return result_df
    except Exception as e:
        st.error(f"Error saat memproses file Vietnam: {str(e)}")
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
        import os  # Added missing import

        df = pd.read_excel(file, header=None)

        st.write("📋 Informasi File:")
        st.write(f"- Jumlah baris: {len(df)}")

        def similarity_score(str1, str2):
            str1_norm = ' '.join(str1.lower().split()) if isinstance(str1, str) else ""
            str2_norm = ' '.join(str2.lower().split()) if isinstance(str2, str) else ""
            if not str1_norm or not str2_norm:
                return 0
            score = SequenceMatcher(None, str1_norm, str2_norm).ratio()
            substring_bonus = 0.2 if str1_norm in str2_norm or str2_norm in str1_norm else 0
            return score + substring_bonus

        all_machine_names = []
        for i in range(len(df)):
            if str(df.iloc[i, 3]).strip() == "Nama Mesin":
                machine_name = str(df.iloc[i, 5]).strip()
                if machine_name and machine_name.upper() != "NAN" and not pd.isna(machine_name):
                    all_machine_names.append(machine_name)

        all_machine_names.sort(key=len, reverse=True)

        machine_groups = {}
        processed_machines = set()
        threshold = 0.6

        for name in all_machine_names:
            if name in processed_machines:
                continue
            similar_machines = [name]
            processed_machines.add(name)
            for other_name in all_machine_names:
                if other_name != name and other_name not in processed_machines:
                    score = similarity_score(name, other_name)
                    if score >= threshold:
                        similar_machines.append(other_name)
                        processed_machines.add(other_name)
            canonical_name = max(similar_machines, key=len)
            for m in similar_machines:
                machine_groups[m] = canonical_name

        for name in list(machine_groups.keys()):
            if "hassia" in name.lower() or "redatron" in name.lower():
                machine_groups[name] = "HASSIA REDATRON"
            elif "sacklok" in name.lower():
                machine_groups[name] = "SACKLOK 00001"

        # Store batch numbers with their machine names
        batch_machine_mapping = {}
        mesin_batch_groups = {}
        current_mesin = None
        mesin_original = {}

        for i in range(len(df)):
            if str(df.iloc[i, 3]).strip() == "Nama Mesin":
                original_mesin = str(df.iloc[i, 5]).strip()
                if not original_mesin or original_mesin.upper() in ["NAN", "-", ""] or pd.isna(original_mesin):
                    current_mesin = None
                    continue
                if original_mesin in machine_groups:
                    current_mesin = machine_groups[original_mesin]
                else:
                    current_mesin = original_mesin
                if current_mesin not in mesin_original:
                    mesin_original[current_mesin] = []
                if original_mesin not in mesin_original[current_mesin]:
                    mesin_original[current_mesin].append(original_mesin)
                if current_mesin and current_mesin not in mesin_batch_groups:
                    mesin_batch_groups[current_mesin] = []
            elif str(df.iloc[i, 3]).strip() == "Tanggal Kalibrasi":
                continue
            elif current_mesin is not None:
                batch = str(df.iloc[i, 0]).strip()
                if batch and batch.upper() != "NAN" and not pd.isna(batch) and batch != "-":
                    # Store batch with its machine name information
                    if batch not in batch_machine_mapping:
                        batch_machine_mapping[batch] = []
                    
                    # Store the current machine as a valid machine for this batch
                    if current_mesin not in batch_machine_mapping[batch]:
                        batch_machine_mapping[batch].append(current_mesin)
                    
                    mesin_batch_groups[current_mesin].append(batch)

        st.write("### Detail Grup Mesin")
        for canonical, originals in mesin_original.items():
            st.write(f"- **{canonical}** mencakup: {', '.join(originals)}")

        summary_data = []
        for mesin, batches in mesin_batch_groups.items():
            summary_data.append({
                "Nama Mesin": mesin,
                "Jumlah Batch": len(batches),
                "Nama Asli": ", ".join(mesin_original.get(mesin, [mesin]))
            })

        summary_df = pd.DataFrame(summary_data)
        st.write("🔍 Ringkasan Nama Mesin yang Ditemukan:")
        st.dataframe(summary_df)

        # Initialize result_df to None at the beginning
        result_df = None
        
        if 'tab1_json' not in st.session_state:
            st.warning("Data batch dari Tab 1 belum tersedia. Silakan proses data di Tab 1 terlebih dahulu.")
            st.session_state['filtered_nama_mesin_map'] = mesin_batch_groups
            return None

        mesin_batch_map = json.loads(st.session_state.tab1_json)
        
        valid_batches = []
        mesin_terpilih = []
        for mesin, batches in mesin_batch_map.items():
            valid = [str(b).strip() for b in batches if b is not None]
            if len(valid) > 0:
                mesin_terpilih.append(mesin)
            valid_batches.extend(valid)

        valid_batches = list(set(valid_batches))
        
        # Gunakan hanya batch dari mesin pertama yang dipilih (diasumsikan)
        mesin_dipilih = mesin_terpilih[0] if mesin_terpilih else None
        if mesin_dipilih:
            valid_batches = mesin_batch_map[mesin_dipilih]
            st.info(f"📌 Menggunakan batch dari mesin: {mesin_dipilih} ({len(valid_batches)} batch)")
        
        # Process with proper duplicate handling - keep ALL valid batch instances
        filtered_mesin_batch = {}
        for mesin, batches in mesin_batch_groups.items():
            # Initialize list for this machine if it doesn't exist
            if mesin not in filtered_mesin_batch:
                filtered_mesin_batch[mesin] = []
            
            # Process each batch, keeping only valid ones
            for batch in batches:
                if batch in valid_batches:
                    filtered_mesin_batch[mesin].append(batch)

        # Prepare data for the table display
        table_data = []
        for mesin, batches in filtered_mesin_batch.items():
            if len(batches) > 0:
                batch_sample = ", ".join(batches[:5]) + ("..." if len(batches) > 5 else "")
                table_data.append({
                    "Nama Mesin": mesin,
                    "Jumlah Batch": len(batches),
                    "Contoh Batch": batch_sample,
                    "Nama Asli": ", ".join(mesin_original.get(mesin, [mesin]))
                })
        
        # Create and display the table
        result_table = pd.DataFrame(table_data)
        st.write("### Pengelompokan Batch Berdasarkan Nama Mesin (Setelah Dicocokkan dengan Tab 1):")
        st.dataframe(result_table)

        st.session_state['filtered_nama_mesin_map'] = filtered_mesin_batch

        # Create uniform dataframe for full batch details
        if filtered_mesin_batch:
            max_length = max([len(v) for v in filtered_mesin_batch.values()]) if filtered_mesin_batch else 0
            uniform_data = {}
            for k in filtered_mesin_batch:
                if len(filtered_mesin_batch[k]) > 0:  # Only include machines with batches
                    uniform_data[k] = filtered_mesin_batch[k] + [None] * (max_length - len(filtered_mesin_batch[k]))
            
            if uniform_data:
                result_df = pd.DataFrame(uniform_data)
                st.write("📊 Detail Lengkap Batch Per Mesin:")
                st.dataframe(result_df)

        if st.button("Simpan Referensi Nama Mesin ke JSON"):
            try:
                reference_data = {mesin: list(set([b for b in batch_list if b]))
                                for mesin, batch_list in filtered_mesin_batch.items()}
                with open("mesin_batch_reference.json", "w") as f:
                    json.dump(reference_data, f)
                st.success("Referensi batch per nama mesin berhasil disimpan ke mesin_batch_reference.json")
            except Exception as e:
                st.error(f"Gagal menyimpan referensi: {str(e)}")

        # Return result_df (which may still be None in some cases, but is now initialized)
        return result_df

    except Exception as e:
        st.error(f"Gagal parsing file: {str(e)}")
        st.exception(e)
        return None  # Added explicit return None for exception case
        
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
    """
    try:
        # Baca file Excel dengan pandas
        df = pd.read_excel(file_grinding)
        
        # Pastikan kolom "Nomor Batch" ada
        if "Nomor Batch" not in df.columns:
            raise ValueError("Kolom 'Nomor Batch' tidak ditemukan dalam file grinding.")
        
        # Bersihkan kolom Nomor Batch dari spasi
        df["Nomor Batch"] = df["Nomor Batch"].astype(str).str.strip()
        
        # Inisialisasi hasil per mesin
        hasil_per_mesin = {}
        
        # Inisialisasi list untuk menyimpan batch yang sudah diklasifikasi
        batch_terklasifikasi = []
        
        # Untuk setiap mesin dalam referensi, filter data
        for mesin, daftar_batch in reference_data.items():
            # Bersihkan daftar batch dari reference_data
            daftar_batch_bersih = [str(b).strip() for b in daftar_batch]
            
            # Filter data berdasarkan nomor batch
            df_filtered = df[df["Nomor Batch"].isin(daftar_batch_bersih)]
            
            # Simpan hasil filter
            hasil_per_mesin[mesin] = df_filtered
            
            # Tambahkan batch yang terklasifikasi ke dalam list
            batch_terklasifikasi.extend(df_filtered["Nomor Batch"].tolist())
        
        # Filter data yang tidak terklasifikasi
        df_unclassified = df[~df["Nomor Batch"].isin(batch_terklasifikasi)]
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
    tab1, tab2, tab3 = st.tabs(["Pengelompokan Batch dengan Kode Mesin", "Pengelompokan Batch dengan Nama Mesin", "Pisahkan Data Grinding per Mesin"])
    
    with tab1:
        # Kode tab1 yang sudah ada...
        st.header("Pengelompokan Batch dengan Kode Mesin")
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
        st.header("Pisahkan Data Grinding Berdasarkan Mesin")
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
                    st.subheader("Hasil Pemisahan Data Grinding berdasarkan Nama Mesin")
                    
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
