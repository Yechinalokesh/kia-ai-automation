from flask import Flask, request, send_file, jsonify
import io, sys, os
from werkzeug.utils import secure_filename
sys.path.insert(0, os.path.dirname(__file__))
from ppt_generator import generate_ppt

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB

HTML = open(os.path.join(os.path.dirname(__file__), 'index.html'), encoding='utf-8').read()
@app.route('/')
def index():
    return HTML

@app.route('/generate', methods=['POST'])
def generate():
    if 'file' not in request.files:
        return jsonify(error="No file uploaded"), 400
    f = request.files['file']
    safe_name = secure_filename(f.filename)
    if not safe_name.endswith('.xlsx'):
        return jsonify(error="Please upload an .xlsx file"), 400
    try:
        ppt_bytes = generate_ppt(f.read())
        fname = safe_name.replace('.xlsx', '_report.pptx')
        return send_file(
            io.BytesIO(ppt_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=fname
        )
    except Exception as e:
        import traceback
        return jsonify(error=str(e), detail=traceback.format_exc()), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=os.environ.get('FLASK_DEBUG', '0') == '1')
