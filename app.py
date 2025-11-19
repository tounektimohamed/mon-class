from flask import Flask, render_template_string, request, send_file
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import io
import os

app = Flask(__name__)

HTML = """<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<title>إنشاء جدول DOCX</title>
<style>
body { font-family: Arial, sans-serif; text-align: right; padding: 20px; }
input, select { width: 100%; margin: 5px 0; padding: 5px; }
button { padding: 10px 20px; font-size: 16px; }
</style>
</head>
<body>
<h2>إنشاء جدول DOCX ديناميكي</h2>
<form method="POST">
    <label>القسم:</label>
    <input type="text" name="classe" required>
    
    <label>المادة:</label>
    <input type="text" name="matiere" required>
    
    <label>المعايير (افصل بين كل معيار بفاصلة):</label>
    <input type="text" name="criteria" placeholder="مع1, مع2, مع3..." required>
    
    <label>اختر مجموعة التلاميذ:</label>
    <select name="group_choice" required>
        <option value="1">المجموعة السابقة</option>
        <option value="2">المجموعة الجديدة</option>
    </select>
    
    <button type="submit">إنشاء الملف</button>
</form>
</body>
</html>
"""

group_old = ["أمنه عبد اللطيف","أروى يقين طنيش","اسامه بنضو","أنس الخطيب","إسراء بنمفتاح"]  # et ainsi de suite
group_new = ["احلام الغليظ","أحمد التايب","أحمد الحمزي","أيمن حلموس","إدريس القرسان"]  # etc.

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        classe = request.form.get("classe")
        matiere = request.form.get("matiere")
        criteria_input = request.form.get("criteria", "")
        criteria = [c.strip() for c in criteria_input.split(",") if c.strip()]
        if not criteria:
            criteria = ["مع 1", "مع 2", "مع 3"]

        group_choice = request.form.get("group_choice")
        names = group_new if group_choice=="2" else group_old

        doc = Document()
        doc.add_heading(f"جداول إسناد إعداد {matiere}", level=0).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        doc.add_paragraph(f"القسم: {classe}    -    مدرسة الحبيب بورقيبة تطاوين").alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        cols = 1 + len(criteria)
        table = doc.add_table(rows=1, cols=cols)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        hdr[0].text = "الاسم"
        for i, c in enumerate(criteria):
            hdr[i+1].text = c

        for name in names:
            row = table.add_row().cells
            row[0].text = name
            for j in range(len(criteria)):
                row[j+1].text = ""

        f = io.BytesIO()
        doc.save(f)
        f.seek(0)
        return send_file(f,
            as_attachment=True,
            download_name="table_RTL_web.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    return render_template_string(HTML)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
