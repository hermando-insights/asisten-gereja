from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from pptx import Presentation
import io
import os

app = Flask(__name__)

# Izinkan semua domain agar GitHub Pages kamu bisa akses tanpa blokir
CORS(app)

# Rute Utama agar tidak muncul "Not Found" di link Render
@app.route('/')
def home():
    return "<h1>Server Asisten Gereja Aktif!</h1><p>Siap memproses data PowerPoint dari Hermando Insights.</p>"

@app.route('/generate-ppt', methods=['POST', 'OPTIONS'])
def generate_ppt():
    # Menangani preflight request dari browser
    if request.method == 'OPTIONS':
        return jsonify({"status": "ok"}), 200

    data = request.json
    template_file = 'Template PowerPoint.pptx'
    
    # Gunakan path absolut agar file template selalu terbaca di server Render
    base_path = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_path, template_file)
    
    try:
        prs = Presentation(template_path)
    except Exception as e:
        print(f"Error Template: {e}")
        return jsonify({"error": "File template tidak ditemukan di server"}), 500

    # Proses pembuatan slide
    for item in data.get('slides', []):
        try:
            layout_idx = int(item.get('layout_idx', 0))
            if layout_idx >= len(prs.slide_layouts):
                layout_idx = 0
                
            layout_dipilih = prs.slide_layouts[layout_idx]
            slide = prs.slides.add_slide(layout_dipilih)
            
            shapes = sorted(slide.placeholders, key=lambda p: p.placeholder_format.idx)

            if len(shapes) > 0 and item.get('judul'):
                shapes[0].text = item['judul']
            
            if len(shapes) > 1 and item.get('isi'):
                tf = shapes[1].text_frame
                tf.text = item['isi']
                
        except Exception as e:
            print(f"Gagal proses slide: {e}")
            continue

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
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
