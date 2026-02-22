from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
import io
import os
import uuid

app = Flask(__name__)
CORS(app)

# Fungsi bantuan untuk menambahkan Section ke dalam XML PowerPoint
def add_section(prs, name, slide_id_list):
    # Pastikan elemen p14 terdaftar di namespace
    ns_p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
    
    # Mencari atau membuat extLst di tempat yang benar
    try:
        # Kita cari extLst yang punya URI khusus untuk Sections
        ext_lst = prs.element.find(qn('p:extLst'))
        if ext_lst is None:
            ext_lst = prs.element.add_extLst()
    except Exception:
        ext_lst = prs.element.add_extLst()

    section_id = f"{{{str(uuid.uuid4()).upper()}}}"
    sld_id_xml = "".join([f'<p14:sldId id="{sid}"/>' for sid in slide_id_list])
    
    # Gunakan XML yang lebih eksplisit untuk p14 namespace
    xml = f'''
    <p:ext xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
           uri="{{521415D9-36F0-43E3-9257-96A12269D11F}}">
        <p14:section xmlns:p14="{ns_p14}" 
                     name="{name}" 
                     id="{section_id}">
            <p14:sldIdLst>
                {sld_id_xml}
            </p14:sldIdLst>
        </p14:section>
    </p:ext>
    '''
    ext = parse_xml(xml)
    ext_lst.append(ext)

@app.route('/')
def home():
    return "<h1>Server Asisten Gereja Aktif!</h1><p>Siap memproses data PowerPoint dengan fitur Section.</p>"

@app.route('/generate-ppt', methods=['POST', 'OPTIONS'])
def generate_ppt():
    if request.method == 'OPTIONS':
        return jsonify({"status": "ok"}), 200

    data = request.json
    template_file = 'Source_PowerPoint.pptx'
    
    base_path = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_path, template_file)
    
    try:
        prs = Presentation(template_path)
    except Exception as e:
        return jsonify({"error": f"File template tidak ditemukan: {e}"}), 500

    # Dictionary untuk menampung slide berdasarkan nama section
    # Format: {"Nama Section": [slide_id1, slide_id2, ...]}
    sections_map = {}
    current_section = "Default"

    for item in data.get('slides', []):
        try:
            # Jika ada field 'section' baru dari React, kita ganti section aktif
            if item.get('section'):
                current_section = item['section']
            
            if current_section not in sections_map:
                sections_map[current_section] = []

            layout_idx = int(item.get('layout_idx', 0))
            if layout_idx >= len(prs.slide_layouts):
                layout_idx = 0
                
            layout_dipilih = prs.slide_layouts[layout_idx]
            slide = prs.slides.add_slide(layout_dipilih)
            
            # Simpan slide_id untuk didaftarkan ke section nanti
            sections_map[current_section].append(slide.slide_id)
            
            shapes = sorted(slide.placeholders, key=lambda p: p.placeholder_format.idx)

            if len(shapes) > 0 and item.get('judul'):
                shapes[0].text = item['judul']
            
            if len(shapes) > 1 and item.get('isi'):
                tf = shapes[1].text_frame
                tf.text = item['isi']
                
        except Exception as e:
            print(f"Gagal proses slide: {e}")
            continue

    # Daftarkan semua section yang terkumpul ke dalam XML PPT
    for sec_name, sld_ids in sections_map.items():
        if sld_ids:
            add_section(prs, sec_name, sld_ids)

    target_stream = io.BytesIO()
    prs.save(target_stream)
    target_stream.seek(0)

    return send_file(
        target_stream,
        as_attachment=True,
        download_name="Ibadah_Minggu_Sectioned.pptx",
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
