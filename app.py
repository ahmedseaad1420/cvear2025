from flask import Flask, render_template, request, send_from_directory
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

def save_to_excel(data_dict):
    df = pd.DataFrame([data_dict])
    folder = 'saved_submissions'
    os.makedirs(folder, exist_ok=True)
    filename = f"submission_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    filepath = os.path.join(folder, filename)
    df.to_excel(filepath, index=False)

@app.route('/', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        data = request.form.to_dict(flat=False)
        processed = {}
        for k, v in data.items():
            processed[k] = ', '.join(v) if isinstance(v, list) else v
        save_to_excel(processed)
        return "✅ تم حفظ البيانات في Excel بنجاح!"
    return render_template('form.html')

@app.route('/admin')
def admin():
    folder = 'saved_submissions'
    try:
        files = os.listdir(folder)
        file_links = [f"<li><a href='/download/{fname}' target='_blank'>{fname}</a></li>" for fname in files]
        return "<h2>📄 الطلبات المستلمة:</h2><ul>" + "".join(file_links) + "</ul>"
    except FileNotFoundError:
        return "🚫 لا يوجد ملفات حالياً."

@app.route('/download/<filename>')
def download_file(filename):
    folder = 'saved_submissions'
    return send_from_directory(folder, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
