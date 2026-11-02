import os
import re
from flask import Flask, render_template, request, send_file
import pandas as pd
from pypdf import PdfReader
from docx import Document
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_EXCEL = 'hasil_ekstraksi.xlsx'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def bersihkan_karakter_ilegal(teks):
    """
    Fungsi ketat untuk memastikan hanya karakter yang kompatibel dengan XML/Excel 
    yang lolos, sehingga mencegah error ValueError pada openpyxl.
    """
    if not isinstance(teks, str):
        teks = str(teks)
    
    # Hanya izinkan karakter yang didukung oleh format XML Excel
    cleaned = re.sub(r'[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD]', '', teks)
    return cleaned.strip()

def format_excel_dengan_garis(file_excel):
    """
    Fungsi untuk memberikan garis (border), warna header, dan merapikan tabel di Excel.
    """
    import openpyxl
    wb = openpyxl.load_workbook(file_excel)
    ws = wb.active

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            if cell.row > 1:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 15)

    wb.save(file_excel)

@app.route('/', methods=['GET', 'POST'])
def index():
    pesan = None
    if request.method == 'POST':
        files = request.files.getlist('file_dokumen')
        
        if files:
            daftar_baris = []
            
            for file in files:
                if file and file.filename != '':
                    filename = file.filename
                    file_path = os.path.join(UPLOAD_FOLDER, filename)
                    file.save(file_path)
                    
                    ext = os.path.splitext(filename)[1].lower()
                    
                    try:
                        # Proses Ekstraksi PDF
                        if ext == '.pdf':
                            reader = PdfReader(file_path)
                            for i, page in enumerate(reader.pages):
                                teks_halaman = page.extract_text() or ""
                                for baris in teks_halaman.split('\n'):
                                    clean_text = bersihkan_karakter_ilegal(baris)
                                    if clean_text:
                                        daftar_baris.append({
                                            "No": len(daftar_baris) + 1,
                                            "Nama File": filename,
                                            "Halaman": i + 1,
                                            "Isi Teks / Kata": clean_text
                                        })
                                        
                        # Proses Ekstraksi Word (.docx)
                        elif ext == '.docx':
                            doc = Document(file_path)
                            for para in doc.paragraphs:
                                clean_text = bersihkan_karakter_ilegal(para.text)
                                if clean_text:
                                    daftar_baris.append({
                                        "No": len(daftar_baris) + 1,
                                        "Nama File": filename,
                                        "Halaman": 1,
                                        "Isi Teks / Kata": clean_text
                                    })
                                    
                        # Untuk file format teks lainnya atau .doc lama
                        else:
                            try:
                                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                                    isi = f.read()
                                    for baris in isi.split('\n'):
                                        clean_text = bersihkan_karakter_ilegal(baris)
                                        if clean_text:
                                            daftar_baris.append({
                                                "No": len(daftar_baris) + 1,
                                                "Nama File": filename,
                                                "Halaman": 1,
                                                "Isi Teks / Kata": clean_text
                                            })
                            except:
                                clean_name = bersihkan_karakter_ilegal(filename)
                                daftar_baris.append({
                                    "No": len(daftar_baris) + 1,
                                    "Nama File": clean_name,
                                    "Halaman": 1,
                                    "Isi Teks / Kata": f"[File berformat {ext} berhasil diupload]"
                                })
                    except Exception as e:
                        daftar_baris.append({
                            "No": len(daftar_baris) + 1,
                            "Nama File": filename,
                            "Halaman": 1,
                            "Isi Teks / Kata": f"Error membaca file: {str(e)}"
                        })

            if daftar_baris:
                df_baru = pd.DataFrame(daftar_baris)
                df_baru.to_excel(OUTPUT_EXCEL, index=False)
                format_excel_dengan_garis(OUTPUT_EXCEL)
                
                pesan = f"Berhasil mengekstrak {len(files)} file ke dalam Excel dengan bersih!"
            
    return render_template('index.html', pesan=pesan)

@app.route('/download')
def download():
    if os.path.exists(OUTPUT_EXCEL):
        return send_file(OUTPUT_EXCEL, as_attachment=True)
    return "File Excel belum tersedia."

if __name__ == '__main__':
    # Konfigurasi port otomatis agar mendukung server lokal maupun cloud (Render)
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)