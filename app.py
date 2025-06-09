
from flask import Flask, render_template, request
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

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
