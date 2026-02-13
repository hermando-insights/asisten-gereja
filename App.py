from flask import Flask, request, send_file
from flask_cors import CORS
from pptx import Presentation
import io

app = Flask(__name__)

# Konfigurasi CORS khusus agar mengizinkan port Vite kamu (5173)
# Ini solusi agar tidak kena blokir browser saat klik "Buat PPT"
CORS(app, resources={r"/*": {"origins": "http://localhost:5173"}})

@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    data = request.json  # Mengambil data slides dari React
    template_file = 'Template PowerPoint.pptx'
    
    try:
        # Memastikan file template ada di folder yang sama dengan App.py
        prs = Presentation(template_file)
    except Exception as e:
        print(f"Kesalahan: {e}")
        return {"error": "Template PowerPoint tidak ditemukan di folder server"}, 500

    # Melakukan looping untuk setiap slide yang dikirim dari form web
    for item in data['slides']:
        try:
            # Mengambil index layout dan validasi
            layout_idx = int(item['layout_idx'])
            if layout_idx >= len(prs.slide_layouts):
                continue
                
            layout_dipilih = prs.slide_layouts[layout_idx]
            slide = prs.slides.add_slide(layout_dipilih)
            
            # Mengurutkan placeholder agar pengisian judul dan isi tidak tertukar
            shapes = sorted(slide.placeholders, key=lambda p: p.placeholder_format.idx)

            # Mengisi Judul (Placeholder pertama)
            if len(shapes) > 0 and item.get('judul'):
                shapes[0].text = item['judul']
            
            # Mengisi Isi/Konten (Placeholder kedua)
            if len(shapes) > 1 and item.get('isi'):
                tf = shapes[1].text_frame
                # Menangani baris baru agar tidak berantakan
                tf.text = item['isi'].replace('\\n', '\n')
                
        except Exception as e:
            print(f"Gagal memproses slide: {e}")
            continue

    # Menggunakan BytesIO agar file disimpan di RAM (lebih cepat & bersih)
    target_stream = io.BytesIO()
    prs.save(target_stream)
    target_stream.seek(0)

    return send_file(
        target_stream,
        as_attachment=True,
        download_name="Ibadah_Minggu.pptx",
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

if __name__ == '__main__':
    import os
    # Mengambil port dari environment variable server, default ke 5000 jika lokal
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)