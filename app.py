import streamlit as st
import pandas as pd
import cv2
import numpy as np
import pytesseract
from PIL import Image
import difflib

# WAJIB UNTUK WINDOWS LOCALHOST (Sesuaikan jika di server Linux)
#pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

st.set_page_config(page_title="Sparepart Admin Panel", layout="wide")
st.title("🛠️ Sparepart Database Admin Panel")

# Membuat 2 Tab Menu di Website
tab1, tab2 = st.tabs(["📁 1. Generate Data (Insert & Reset)", "🎯 2. Auto-Pin Koordinat (OCR)"])

# =====================================================================
# TAB 1: GENERATE DATA (RESET & INSERT)
# =====================================================================
with tab1:
    st.header("Generate Query SQL untuk Data Part")
    st.markdown("Fitur ini akan membaca Excel, menghapus data lama di database (Cascade Delete), lalu membuat query Insert untuk Part Names dan Part Numbers baru.")
    
    col_t1_1, col_t1_2 = st.columns(2)
    with col_t1_1:
        prod_name_tab1 = st.text_input("Nama Produk (Cth: Sigra Gen 3)", value="Sigra Gen 3", key="prod1")
        excel_file_tab1 = st.file_uploader("Upload File Excel (.xlsx)", type=["xlsx", "csv"], key="exc1")
        
    if excel_file_tab1 and st.button("Generate SQL Data", key="btn1"):
        with st.spinner("Memproses data Excel..."):
            if excel_file_tab1.name.endswith('.csv'): df = pd.read_csv(excel_file_tab1)
            else: df = pd.read_excel(excel_file_tab1)
            
            sql_stmts = ["-- SCRIPT RESET & INSERT PART NAMES & PART NUMBERS\n"]
            
            # 0. HAPUS DATA LAMA
            unique_figures = df['Part Figure Index'].dropna().unique()
            fig_list_sql = ", ".join([f"'{str(f).strip()}'" for f in unique_figures])
            
            sql_stmts.append("-- 0. HAPUS DATA LAMA (CASCADE)")
            sql_stmts.append(f"DELETE pcm FROM part_numbers_compatible_models pcm JOIN part_numbers pnum ON pcm.partnumber_id = pnum.id JOIN part_names pn ON pnum.part_name_id = pn.id JOIN part_figures pf ON pn.part_figure_id = pf.id JOIN part_groups pg ON pf.part_group_id = pg.id JOIN products pr ON pg.product_id = pr.id WHERE pr.name = '{prod_name_tab1}' AND pf.number IN ({fig_list_sql});")
            sql_stmts.append(f"DELETE pnum FROM part_numbers pnum JOIN part_names pn ON pnum.part_name_id = pn.id JOIN part_figures pf ON pn.part_figure_id = pf.id JOIN part_groups pg ON pf.part_group_id = pg.id JOIN products pr ON pg.product_id = pr.id WHERE pr.name = '{prod_name_tab1}' AND pf.number IN ({fig_list_sql});")
            sql_stmts.append(f"DELETE pn FROM part_names pn JOIN part_figures pf ON pn.part_figure_id = pf.id JOIN part_groups pg ON pf.part_group_id = pg.id JOIN products pr ON pg.product_id = pr.id WHERE pr.name = '{prod_name_tab1}' AND pf.number IN ({fig_list_sql});\n")

            # 1. INSERT PART NAMES
            sql_stmts.append("-- 1. INSERT PART NAMES (PNC)")
            unique_pncs = df[['Figure No', 'Part Name', 'Part Figure Index']].dropna().drop_duplicates()
            for _, row in unique_pncs.iterrows():
                pnc_num = str(row['Figure No']).strip().replace("'", "\\'")
                p_name = str(row['Part Name']).strip().replace("'", "\\'")
                f_idx = str(row['Part Figure Index']).strip().replace("'", "\\'")
                
                sql = f"INSERT INTO part_names (id, created_at, updated_at, is_deleted, number, name, part_figure_id, status, x_position, y_position) SELECT REPLACE(UUID(), '-', ''), NOW(), NOW(), 0, '{pnc_num}', '{p_name}', pf.id, 'active', 0, 0 FROM part_figures pf JOIN part_groups pg ON pf.part_group_id = pg.id JOIN products pr ON pg.product_id = pr.id WHERE pr.name = '{prod_name_tab1}' AND pf.number = '{f_idx}' AND NOT EXISTS (SELECT 1 FROM part_names pn WHERE pn.number = '{pnc_num}' AND pn.part_figure_id = pf.id) LIMIT 1;"
                sql_stmts.append(sql)

            # 2. INSERT PART NUMBERS
            sql_stmts.append("\n-- 2. INSERT PART NUMBERS")
            for _, row in df.iterrows():
                pnc_num = str(row['Figure No']).strip().replace("'", "\\'")
                part_num = str(row['Part Number']).strip().replace("'", "\\'")
                f_idx = str(row['Part Figure Index']).strip().replace("'", "\\'")
                desc = str(row['Description']).strip().replace("'", "\\'")
                if desc.lower() == 'nan': desc = ''
                model = str(row['Model']).strip().replace("'", "\\'")
                try: qty = int(row['Qty'])
                except: qty = 1
                
                prod_date = str(row['Prod Date']).strip()
                start_yr, end_yr = "NULL", "NULL"
                if prod_date.lower() != 'nan' and '-' in prod_date:
                    pts = prod_date.split('-')
                    if len(pts[0].strip()) >= 2: start_yr = f"20{pts[0].strip()[:2]}"
                    if len(pts) > 1 and len(pts[1].strip()) >= 2: end_yr = f"20{pts[1].strip()[:2]}"
                elif prod_date.lower() != 'nan' and len(prod_date) >= 2:
                    start_yr = f"20{prod_date[:2]}"
                    
                spec = str(row['Spec Code']).strip().replace("'", "\\'")
                spec_val = "NULL" if spec.lower() == 'nan' or spec == '' else f"'{spec}'"

                sql = f"INSERT INTO part_numbers (id, created_at, updated_at, is_deleted, number, description, qty, model, production_date, spec_code, status, production_start_year, production_end_year, part_name_id) SELECT REPLACE(UUID(), '-', ''), NOW(), NOW(), 0, '{part_num}', '{desc}', {qty}, '{model}', '{prod_date}', {spec_val}, 'active', {start_yr}, {end_yr}, pn.id FROM part_names pn JOIN part_figures pf ON pn.part_figure_id = pf.id JOIN part_groups pg ON pf.part_group_id = pg.id JOIN products pr ON pg.product_id = pr.id WHERE pr.name = '{prod_name_tab1}' AND pf.number = '{f_idx}' AND pn.number = '{pnc_num}' AND NOT EXISTS (SELECT 1 FROM part_numbers pnum WHERE pnum.number = '{part_num}' AND pnum.part_name_id = pn.id AND pnum.model = '{model}') LIMIT 1;"
                sql_stmts.append(sql)

            # Hasil
            final_sql = "\n".join(sql_stmts)
            st.success("✅ Query SQL berhasil di-generate!")
            st.download_button(label="⬇️ Download Query Insert/Reset (.sql)", data=final_sql, file_name=f"insert_data_{prod_name_tab1}.sql", mime="text/plain")


# =====================================================================
# TAB 2: AUTO-PIN KOORDINAT (OCR)
# =====================================================================
with tab2:
    st.header("Auto-Pin Koordinat PNC")
    st.markdown("Fitur ini akan men-scan gambar diagram, mencari teks PNC menggunakan Excel sebagai validasi, dan otomatis membuat query Update Koordinat X & Y.")
    
    col_t2_1, col_t2_2 = st.columns([1, 2])
    with col_t2_1:
        prod_name_tab2 = st.text_input("Nama Produk", value="Sigra Gen 3", key="prod2")
        fig_index_tab2 = st.text_input("Figure Index (Cth: 85-03)", value="85-03")
        excel_file_tab2 = st.file_uploader("Upload File Excel Validasi", type=["xlsx", "csv"], key="exc2")
        image_file = st.file_uploader("Upload Gambar Diagram", type=["png", "jpg", "jpeg"])
        pin_gap = 5

    if excel_file_tab2 and image_file and st.button("Mulai Scan Gambar", key="btn2"):
        with st.spinner('Membaca Excel dan menganalisa gambar dengan AI...'):
            if excel_file_tab2.name.endswith('.csv'): df2 = pd.read_csv(excel_file_tab2)
            else: df2 = pd.read_excel(excel_file_tab2)
                
            df_filt = df2[df2['Part Figure Index'].astype(str).str.strip() == fig_index_tab2]
            if df_filt.empty:
                st.error(f"❌ Tidak ditemukan data untuk Figure {fig_index_tab2} di Excel!")
                st.stop()
                
            valid_pncs = df_filt['Figure No'].dropna().astype(str).str.strip().unique().tolist()

            # Proses Gambar
            img_bytes = np.asarray(bytearray(image_file.read()), dtype=np.uint8)
            img = cv2.imdecode(img_bytes, 1)
            annotated_img = img.copy()
            img_h, img_w, _ = img.shape
            
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            _, bin_line = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)
            
            SCALE = 2.0
            gray_scaled = cv2.resize(gray, None, fx=SCALE, fy=SCALE, interpolation=cv2.INTER_CUBIC)
            _, gray_scaled = cv2.threshold(gray_scaled, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
            
            cfg = r'--oem 3 --psm 11 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'
            data = pytesseract.image_to_data(gray_scaled, output_type=pytesseract.Output.DICT, config=cfg)

            found_pncs = {}
            sql_upd = [f"-- UPDATE KOORDINAT {prod_name_tab2} FIGURE {fig_index_tab2}\n"]
            
            for i in range(len(data['text'])):
                txt = data['text'][i].upper().strip().replace('O', '0').replace('I', '1').replace('L', '1')
                txt_clean = ''.join(e for e in txt if e.isalnum())
                if len(txt_clean) < 4: continue
                
                matches = difflib.get_close_matches(txt_clean, valid_pncs, n=1, cutoff=0.7)
                if matches:
                    true_pnc = matches[0]
                    if true_pnc not in found_pncs:
                        x, y = int(data['left'][i]/SCALE), int(data['top'][i]/SCALE)
                        w, h = int(data['width'][i]/SCALE), int(data['height'][i]/SCALE)
                        if w == 0 or h == 0: continue
                        
                        mar, thk = 3, 4
                        ty1, ty2 = max(0, y-mar-thk), max(0, y-mar)
                        by1, by2 = min(img_h, y+h+mar), min(img_h, y+h+mar+thk)
                        lx1, lx2 = max(0, x-mar-thk), max(0, x-mar)
                        rx1, rx2 = min(img_w, x+w+mar), min(img_w, x+w+mar+thk)
                        
                        scores = {'atas': np.sum(bin_line[ty1:ty2, x:x+w]), 'bawah': np.sum(bin_line[by1:by2, x:x+w]), 'kiri': np.sum(bin_line[y:y+h, lx1:lx2]), 'kanan': np.sum(bin_line[y:y+h, rx1:rx2])}
                        best = max(scores, key=scores.get)
                        
                        tx, ty = x + (w/2.0), y + (h/2.0)
                        if max(scores.values()) == 0: tx = x + w + pin_gap
                        else:
                            if best == 'atas': ty = y - pin_gap
                            elif best == 'bawah': ty = y + h + pin_gap
                            elif best == 'kiri': tx = x - pin_gap
                            elif best == 'kanan': tx = x + w + pin_gap

                        px = max(0.0, min(100.0, round((tx/img_w)*100, 2)))
                        py = max(0.0, min(100.0, round((ty/img_h)*100, 2)))
                        found_pncs[true_pnc] = True
                        
                        cv2.rectangle(annotated_img, (x, y), (x+w, y+h), (0, 255, 0), 2)
                        cv2.putText(annotated_img, true_pnc, (x, max(10, y-5)), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 0, 0), 2)
                        cv2.circle(annotated_img, (int(tx), int(ty)), 6, (0, 0, 255), -1)
                        
                        sql = f"UPDATE part_names pn JOIN part_figures pf ON pn.part_figure_id = pf.id JOIN part_groups pg ON pf.part_group_id = pg.id JOIN products pr ON pg.product_id = pr.id SET pn.x_position = {px}, pn.y_position = {py} WHERE pr.name = '{prod_name_tab2}' AND pf.number = '{fig_index_tab2}' AND pn.number = '{true_pnc}';"
                        sql_upd.append(sql)

            with col_t2_2:
                st.success(f"🎯 Berhasil pin {len(found_pncs)} dari total {len(valid_pncs)} PNC di Excel!")
                missed = [p for p in valid_pncs if p not in found_pncs]
                if missed: st.warning(f"Teks ini tidak ketemu di gambar: {', '.join(missed)}")
                
                rgb_img = cv2.cvtColor(annotated_img, cv2.COLOR_BGR2RGB)
                st.image(Image.fromarray(rgb_img), caption="Preview Hasil", use_column_width=True)
                
                sql_text = "\n".join(sql_upd)
                st.download_button(label="⬇️ Download Query Update Koordinat (.sql)", data=sql_text, file_name=f"koordinat_{fig_index_tab2}.sql", mime="text/plain")