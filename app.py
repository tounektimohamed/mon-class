from flask import Flask, render_template_string, request, send_file
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement, ns
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import os
import json

app = Flask(__name__)

# HTML avec interface améliorée
HTML = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>إنشاء جدول DOCX</title>
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            text-align: right; 
            padding: 20px; 
            background-color: #f5f5f5;
            margin: 0;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h2 {
            color: #2c3e50;
            text-align: center;
            margin-bottom: 30px;
        }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            margin-bottom: 8px;
            font-weight: bold;
            color: #34495e;
        }
        input, select {
            width: 100%;
            margin: 5px 0;
            padding: 12px;
            border: 1px solid #ddd;
            border-radius: 5px;
            font-size: 16px;
        }
        .criteria-container {
            border: 2px dashed #3498db;
            padding: 15px;
            border-radius: 5px;
            margin: 10px 0;
            min-height: 100px;
            background-color: #f8f9fa;
        }
        .criteria-item {
            background: #3498db;
            color: white;
            padding: 8px 15px;
            margin: 5px;
            border-radius: 20px;
            display: inline-block;
            cursor: move;
            position: relative;
        }
        .criteria-item:hover {
            background: #2980b9;
        }
        .criteria-item.editing {
            background: #e74c3c;
        }
        .edit-input {
            background: transparent;
            border: none;
            color: white;
            font-size: 14px;
            text-align: center;
            width: 120px;
            outline: none;
        }
        .edit-input::placeholder {
            color: rgba(255,255,255,0.7);
        }
        .criteria-input {
            display: flex;
            gap: 10px;
            margin-bottom: 15px;
        }
        .criteria-input input {
            flex: 1;
        }
        .btn {
            background: #3498db;
            color: white;
            border: none;
            padding: 15px 30px;
            border-radius: 5px;
            font-size: 18px;
            cursor: pointer;
            width: 100%;
            margin-top: 20px;
        }
        .btn:hover {
            background: #2980b9;
        }
        .btn-secondary {
            background: #95a5a6;
            padding: 10px 20px;
            font-size: 14px;
            width: auto;
        }
        .btn-secondary:hover {
            background: #7f8c8d;
        }
        .drag-info {
            text-align: center;
            color: #7f8c8d;
            font-style: italic;
            margin: 10px 0;
        }
        .predefined-criteria {
            margin: 15px 0;
        }
        .predefined-item {
            background: #e74c3c;
            color: white;
            padding: 5px 10px;
            margin: 3px;
            border-radius: 15px;
            display: inline-block;
            cursor: pointer;
            font-size: 14px;
        }
        .predefined-item:hover {
            background: #c0392b;
        }
        .edit-icon {
            margin-right: 5px;
            cursor: pointer;
            font-size: 12px;
            opacity: 0.7;
        }
        .edit-icon:hover {
            opacity: 1;
        }
        .instructions {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 10px;
            border-radius: 5px;
            margin: 10px 0;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>إنشاء جدول DOCX ديناميكي</h2>
        <form method="POST" id="docxForm">
            <div class="form-group">
                <label>القسم:</label>
                <input type="text" name="classe" required placeholder="أدخل اسم القسم">
            </div>
            
            <div class="form-group">
                <label>المادة:</label>
                <input type="text" name="matiere" required placeholder="أدخل اسم المادة">
            </div>
            
            <div class="form-group">
                <label>المعايير:</label>
                
                <div class="instructions">
                    💡 <strong>تعليمات:</strong> انقر نقراً مزدوجاً على أي معيار لتعديل اسمه. اسحب المعايير لإعادة ترتيبها.
                </div>
                
                <div class="predefined-criteria">
                    <strong>معايير جاهزة:</strong><br>
                    <span class="predefined-item" onclick="addPredefined('المشاركة في الحصة')">المشاركة في الحصة</span>
                    <span class="predefined-item" onclick="addPredefined('إنجاز الواجبات')">إنجاز الواجبات</span>
                    <span class="predefined-item" onclick="addPredefined('الاختبار التحريري')">الاختبار التحريري</span>
                    <span class="predefined-item" onclick="addPredefined('التطبيق العملي')">التطبيق العملي</span>
                    <span class="predefined-item" onclick="addPredefined('المشروع الجماعي')">المشروع الجماعي</span>
                </div>
                
                <div class="criteria-input">
                    <input type="text" id="newCriteria" placeholder="أدخل معيار جديد">
                    <button type="button" class="btn-secondary" onclick="addCriteria()">إضافة معيار</button>
                </div>
                
                <div class="criteria-container" id="criteriaContainer" ondragover="allowDrop(event)">
                    <div class="drag-info">اسحب المعايير لإعادة ترتيبها - انقر نقراً مزدوجاً للتعديل</div>
                </div>
                <input type="hidden" name="criteria" id="criteriaInput" required>
            </div>
            
            <div class="form-group">
                <label>اختر مجموعة التلاميذ:</label>
                <select name="group_choice" required>
                    <option value="1">المجموعة السابقة</option>
                    <option value="2">المجموعة الجديدة</option>
                </select>
            </div>
            
            <button type="submit" class="btn">إنشاء الملف</button>
        </form>
    </div>

    <script>
        let criteriaList = [];
        
        function updateCriteriaInput() {
            document.getElementById('criteriaInput').value = JSON.stringify(criteriaList);
        }
        
        function addCriteria() {
            const input = document.getElementById('newCriteria');
            const value = input.value.trim();
            if (value && !criteriaList.includes(value)) {
                criteriaList.push(value);
                renderCriteria();
                input.value = '';
            }
        }
        
        function addPredefined(criteria) {
            if (!criteriaList.includes(criteria)) {
                criteriaList.push(criteria);
                renderCriteria();
            }
        }
        
        function removeCriteria(index) {
            criteriaList.splice(index, 1);
            renderCriteria();
        }
        
        function editCriteria(index) {
            const container = document.getElementById('criteriaContainer');
            const items = container.querySelectorAll('.criteria-item');
            const item = items[index];
            
            if (item.classList.contains('editing')) {
                return; // Déjà en mode édition
            }
            
            item.classList.add('editing');
            const currentText = criteriaList[index];
            
            item.innerHTML = `
                <input type="text" 
                       class="edit-input" 
                       value="${currentText}" 
                       onblur="saveEdit(${index}, this.value)"
                       onkeypress="handleEditKeypress(event, ${index}, this)"
                       placeholder="أدخل اسم المعيار">
            `;
            
            const input = item.querySelector('.edit-input');
            input.focus();
            input.select();
        }
        
        function saveEdit(index, newValue) {
            const trimmedValue = newValue.trim();
            if (trimmedValue && !criteriaList.includes(trimmedValue)) {
                criteriaList[index] = trimmedValue;
            }
            renderCriteria();
        }
        
        function handleEditKeypress(event, index, input) {
            if (event.key === 'Enter') {
                saveEdit(index, input.value);
            } else if (event.key === 'Escape') {
                renderCriteria();
            }
        }
        
        function renderCriteria() {
            const container = document.getElementById('criteriaContainer');
            container.innerHTML = '<div class="drag-info">اسحب المعايير لإعادة ترتيبها - انقر نقراً مزدوجاً للتعديل</div>';
            
            if (criteriaList.length === 0) {
                container.innerHTML += '<div class="drag-info">لا توجد معايير مضافة</div>';
            }
            
            criteriaList.forEach((criteria, index) => {
                const item = document.createElement('div');
                item.className = 'criteria-item';
                item.innerHTML = `
                    <span class="edit-icon" onclick="editCriteria(${index})">✏️</span>
                    ${criteria}
                `;
                item.draggable = true;
                item.ondragstart = (e) => dragStart(e, index);
                item.ondblclick = () => editCriteria(index);
                container.appendChild(item);
            });
            
            updateCriteriaInput();
        }
        
        function allowDrop(ev) {
            ev.preventDefault();
        }
        
        function dragStart(ev, index) {
            ev.dataTransfer.setData("text/plain", index);
        }
        
        document.getElementById('criteriaContainer').addEventListener('dragover', allowDrop);
        
        document.getElementById('criteriaContainer').addEventListener('drop', (ev) => {
            ev.preventDefault();
            const fromIndex = parseInt(ev.dataTransfer.getData("text/plain"));
            const items = Array.from(document.querySelectorAll('.criteria-item'));
            const toIndex = items.indexOf(ev.target.closest('.criteria-item'));
            
            if (toIndex !== -1 && fromIndex !== toIndex) {
                const [removed] = criteriaList.splice(fromIndex, 1);
                criteriaList.splice(toIndex, 0, removed);
                renderCriteria();
            }
        });
        
        document.getElementById('newCriteria').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                addCriteria();
            }
        });
        
        // إضافة المعايير الافتراضية عند التحميل
        document.addEventListener('DOMContentLoaded', function() {
            // إضافة المعايير الافتراضية من 1 إلى 5
            setTimeout(() => {
                const defaultCriteria = ["مع 1", "مع 2", "مع 3", "مع 4", "مع 5"];
                defaultCriteria.forEach(criteria => {
                    if (!criteriaList.includes(criteria)) {
                        criteriaList.push(criteria);
                    }
                });
                renderCriteria();
            }, 100);
        });
    </script>
</body>
</html>
"""

# Groupes complets (inchangés)
group_old = [
    "أمنه عبد اللطيف","أروى يقين طنيش","اسامه بنضو","أنس الخطيب","إسراء بنمفتاح",
    "اياد بوحريه","إياد منصور عمار","المختار عبد الواحد","بادیس دقنيش","جاهد السياري",
    "رنيم العزلوك","ريتاج الطالب","رحمة الونيسي","زينب طنيش","زينب عبد الواحد",
    "سلمان الشبلي","فادي القلعاوي","الجين الزردابي","ليان الطالبي","مؤمن بنمبارك",
    "محمد أمير الحمدي","محمد الطاهر مشيري","محمد زكرياء حلاوط","مريم الذكار",
    "ملاك عبد اللطيف","منال بوحربه","هديل بن حامد","ياسمين الحاجي","ياسمين المستيسر",
    "ياسين جويد","يقين بوروحه","يوسف الشيباني","يوسف بن يحي","يونس بوصفة"
]

group_new = [
    "احلام الغليظ","أحمد التايب","أحمد الحمزي","أيمن حلموس","إدريس القرسان",
    "إسراء المرزوقي","باديس سكيب","بتول الفيتوري","تسنيم الطالب","خليل الشلاخ",
    "رضوان عبدالستار","رمزي المقدميني","رنیم خلفه","رنیم عازق","رياض لهول",
    "سيرين العربي","شيماء المورو","عبد الرحمان الوذان","عبد الرحمان بومروة",
    "الجين زهمول","محمد الطاهر بوطالب","محمد جاسم العطوي","محمد ياسين الجليدي",
    "مريم الذكار","مريم حسين","میار حسن","ميس بنصميده","ميار دباغي",
    "نزار عکار","نضال ابن غنيه","نادين مراحي","همام الغرياني","أميمة ذكار"
]

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        classe = request.form.get("classe")
        matiere = request.form.get("matiere")
        
        # Récupération des critères depuis JSON
        criteria_json = request.form.get("criteria", "[]")
        criteria = json.loads(criteria_json)
        if not criteria:
            criteria = ["مع 1", "مع 2", "مع 3", "مع 4", "مع 5"]

        group_choice = request.form.get("group_choice")
        names = group_new if group_choice == "2" else group_old

        # Création du document amélioré
        doc = Document()
        
        # Configuration de la page
        section = doc.sections[0]
        section.page_height = Cm(29.7)
        section.page_width = Cm(21.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        
        # Titre principal
        title = doc.add_heading(f"جداول إسناد إعداد {matiere}", level=1)
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        title_run = title.runs[0]
        title_run.font.size = Pt(16)
        title_run.font.bold = True
        title_run.font.name = 'Arial'

        # Sous-titre
        subtitle = doc.add_paragraph()
        subtitle.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        subtitle_run = subtitle.add_run(f"القسم: {classe} - مدرسة الحبيب بورقيبة تطاوين")
        subtitle_run.font.size = Pt(12)
        subtitle_run.font.name = 'Arial'
        
        # Date
        from datetime import datetime
        date_para = doc.add_paragraph()
        date_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        date_run = date_para.add_run(f"تاريخ الإنشاء: {datetime.now().strftime('%Y-%m-%d')}")
        date_run.font.size = Pt(10)
        date_run.font.name = 'Arial'
        date_run.font.italic = True

        doc.add_paragraph().add_run().add_break()  # Ligne vide

        # Tableau amélioré
        cols = 1 + len(criteria)
        table = doc.add_table(rows=1, cols=cols)
        table.style = 'Table Grid'
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Configuration RTL pour le tableau
        tbl = table._tbl
        tblPr = tbl.tblPr
        bidi = OxmlElement('w:bidiVisual')
        tblPr.append(bidi)

        # En-têtes du tableau
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "الاسم واللقب"
        hdr_cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        hdr_cells[0].paragraphs[0].runs[0].font.size = Pt(12)
        hdr_cells[0].paragraphs[0].runs[0].font.bold = True
        hdr_cells[0].paragraphs[0].runs[0].font.name = 'Arial'

        for i, criterion in enumerate(criteria):
            hdr_cells[i+1].text = criterion
            hdr_cells[i+1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            hdr_cells[i+1].paragraphs[0].runs[0].font.size = Pt(11)
            hdr_cells[i+1].paragraphs[0].runs[0].font.bold = True
            hdr_cells[i+1].paragraphs[0].runs[0].font.name = 'Arial'

        # Lignes des étudiants
        for name in names:
            row_cells = table.add_row().cells
            row_cells[0].text = name
            row_cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            row_cells[0].paragraphs[0].runs[0].font.size = Pt(10)
            row_cells[0].paragraphs[0].runs[0].font.name = 'Arial'
            
            for j in range(len(criteria)):
                row_cells[j+1].text = ""
                row_cells[j+1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                row_cells[j+1].paragraphs[0].runs[0].font.size = Pt(10)
                row_cells[j+1].paragraphs[0].runs[0].font.name = 'Arial'

        # Ajustement des largeurs de colonnes
        for i, column in enumerate(table.columns):
            for cell in column.cells:
                if i == 0:  # Colonne des noms
                    cell.width = Cm(5.0)
                else:  # Colonnes des critères
                    cell.width = Cm(3.0)

        # Pied de page
        doc.add_paragraph().add_run().add_break()
        footer = doc.add_paragraph()
        footer.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        footer_run = footer.add_run("تم إنشاء هذا الجدول آلياً - جميع الحقوق محفوظة")
        footer_run.font.size = Pt(9)
        footer_run.font.italic = True
        footer_run.font.name = 'Arial'

        # Sauvegarde et retour du fichier
        f = io.BytesIO()
        doc.save(f)
        f.seek(0)
        
        filename = f"جدول_{matiere}_{classe}_{datetime.now().strftime('%Y%m%d')}.docx"
        return send_file(
            f,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    return render_template_string(HTML)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)