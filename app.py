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
            max-width: 1200px;
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
        .criteria-section {
            display: flex;
            gap: 20px;
            margin-top: 20px;
        }
        .suggested-criteria, .selected-criteria {
            flex: 1;
        }
        .suggested-container, .selected-container {
            border: 2px dashed #3498db;
            padding: 15px;
            border-radius: 5px;
            margin: 10px 0;
            min-height: 200px;
            background-color: #f8f9fa;
        }
        .suggested-item {
            background: #2ecc71;
            color: white;
            padding: 10px 15px;
            margin: 5px;
            border-radius: 20px;
            display: block;
            cursor: move;
            text-align: center;
        }
        .suggested-item:hover {
            background: #27ae60;
        }
        .criteria-item {
            background: #3498db;
            color: white;
            padding: 10px 15px;
            margin: 5px;
            border-radius: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            cursor: move;
        }
        .criteria-item:hover {
            background: #2980b9;
        }
        .criteria-actions {
            display: flex;
            gap: 5px;
        }
        .action-btn {
            background: rgba(255,255,255,0.2);
            border: none;
            color: white;
            padding: 5px 8px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 12px;
        }
        .action-btn:hover {
            background: rgba(255,255,255,0.3);
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
        .btn-danger {
            background: #e74c3c;
            padding: 8px 15px;
            font-size: 12px;
            width: auto;
        }
        .btn-danger:hover {
            background: #c0392b;
        }
        .drag-info {
            text-align: center;
            color: #7f8c8d;
            font-style: italic;
            margin: 10px 0;
        }
        .instructions {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 10px;
            border-radius: 5px;
            margin: 10px 0;
            font-size: 14px;
        }
        .section-title {
            background: #34495e;
            color: white;
            padding: 10px;
            border-radius: 5px;
            text-align: center;
            margin-bottom: 10px;
        }
        .empty-message {
            text-align: center;
            color: #7f8c8d;
            font-style: italic;
            padding: 20px;
        }
        .table-preview {
            margin-top: 20px;
            border: 2px solid #3498db;
            border-radius: 5px;
            padding: 15px;
            background: white;
        }
        .preview-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
        }
        .preview-table th, .preview-table td {
            border: 1px solid #ddd;
            padding: 5px;
            text-align: center;
        }
        .preview-table th {
            background-color: #f8f9fa;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>إنشاء جدول التقييم DOCX</h2>
        <form method="POST" id="docxForm">
            <div class="form-group">
                <label>القسم:</label>
                <input type="text" name="classe" required placeholder="أدخل اسم القسم">
            </div>
            
            <div class="form-group">
                <label>نوع التقييم:</label>
                <select id="matiere" name="matiere" required onchange="updateSuggestedCriteria()">
                    <option value="">اختر نوع التقييم</option>
                    <option value="التواصل الشفوي">التواصل الشفوي</option>
                    <option value="القراءة">القراءة</option>
                    <option value="الإنتاج الكتابي">الإنتاج الكتابي</option>
                    <option value="قواعد اللغة">قواعد اللغة</option>
                    <option value="أخرى">أخرى</option>
                </select>
            </div>
            
            <div class="form-group">
                <label>إعداد المعايير:</label>
                
                <div class="instructions">
                    💡 <strong>تعليمات:</strong> 
                    <br>• اختر نوع التقييم أولاً
                    <br>• اسحب المعايير من القائمة المقترحة إلى قائمة المعايير المختارة
                    <br>• كل معيار سيكون له 3 خانات للمؤشرات في الجدول النهائي
                </div>

                <div class="criteria-section">
                    <!-- القائمة المقترحة -->
                    <div class="suggested-criteria">
                        <div class="section-title">المعايير المقترحة</div>
                        <div class="suggested-container" id="suggestedContainer" ondragover="allowDrop(event)" ondrop="dropInSuggested(event)">
                            <div class="drag-info">اسحب المعايير إلى القائمة المختارة</div>
                            <div id="suggestedList"></div>
                        </div>
                    </div>
                    
                    <!-- القائمة المختارة -->
                    <div class="selected-criteria">
                        <div class="section-title">المعايير المختارة</div>
                        <div class="selected-container" id="selectedContainer" ondragover="allowDrop(event)" ondrop="dropInSelected(event)">
                            <div class="drag-info">اسحب المعايير هنا</div>
                            <div id="selectedList"></div>
                        </div>
                    </div>
                </div>
                
                <input type="hidden" name="criteria" id="criteriaInput" required>
                
                <div class="criteria-actions" style="justify-content: center; margin-top: 20px;">
                    <button type="button" class="btn-danger" onclick="clearAllCriteria()">حذف الكل</button>
                </div>
            </div>

            <!-- معاينة الجدول -->
            <div class="form-group">
                <div class="table-preview">
                    <div class="section-title">معاينة الجدول</div>
                    <div id="tablePreview">
                        <div class="empty-message">سيظهر معاينة الجدول هنا بعد اختيار المعايير</div>
                    </div>
                </div>
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
        let selectedCriteria = [];
        let suggestedCriteria = [];
        const subjectCriteria = {
            "التواصل الشفوي": [
                "الملائمة", "التغنيم", "الانسجام", "الاتساق", "الثراء"
            ],
            "القراءة": [
                "القراءة الجهرية", "معالجة النص", "التصرف في النص", "إبداء الرأي"
            ],
            "الإنتاج الكتابي": [
                "الملائمة", "سلامة بناء النص", "المقروئية", "ثراء اللغة والطرافة"
            ],
            "قواعد اللغة": [
                "التعرف على الظاهرة اللغوية", "توظيف الظاهرة اللغوية"
            ],
            "أخرى": [
                "معيار 1", "معيار 2", "معيار 3"
            ]
        };
        
        function updateCriteriaInput() {
            document.getElementById('criteriaInput').value = JSON.stringify(selectedCriteria);
            updateTablePreview();
        }
        
        function updateSuggestedCriteria() {
            const subject = document.getElementById('matiere').value;
            const suggestedList = document.getElementById('suggestedList');
            
            suggestedCriteria = subjectCriteria[subject] || [];
            suggestedList.innerHTML = '';
            
            if (suggestedCriteria.length === 0) {
                suggestedList.innerHTML = '<div class="empty-message">لا توجد معايير مقترحة</div>';
                return;
            }
            
            suggestedCriteria.forEach(criteria => {
                if (!selectedCriteria.includes(criteria)) {
                    const item = document.createElement('div');
                    item.className = 'suggested-item';
                    item.textContent = criteria;
                    item.draggable = true;
                    item.ondragstart = (e) => dragStart(e, criteria, 'suggested');
                    suggestedList.appendChild(item);
                }
            });
            
            if (suggestedList.children.length === 0) {
                suggestedList.innerHTML = '<div class="empty-message">جميع المعايير مضافة</div>';
            }
        }
        
        function addToSelected(criteria) {
            if (!selectedCriteria.includes(criteria)) {
                selectedCriteria.push(criteria);
                renderSelectedCriteria();
                updateSuggestedCriteria();
            }
        }
        
        function removeFromSelected(criteria) {
            const index = selectedCriteria.indexOf(criteria);
            if (index > -1) {
                selectedCriteria.splice(index, 1);
                renderSelectedCriteria();
                updateSuggestedCriteria();
            }
        }
        
        function renderSelectedCriteria() {
            const selectedList = document.getElementById('selectedList');
            selectedList.innerHTML = '';
            
            if (selectedCriteria.length === 0) {
                selectedList.innerHTML = '<div class="empty-message">لم يتم اختيار أي معايير</div>';
                updateCriteriaInput();
                return;
            }
            
            selectedCriteria.forEach(criteria => {
                const item = document.createElement('div');
                item.className = 'criteria-item';
                item.draggable = true;
                item.ondragstart = (e) => dragStart(e, criteria, 'selected');
                
                const criteriaText = document.createElement('span');
                criteriaText.textContent = criteria;
                
                const actions = document.createElement('div');
                actions.className = 'criteria-actions';
                
                const deleteBtn = document.createElement('button');
                deleteBtn.className = 'action-btn';
                deleteBtn.innerHTML = '🗑️';
                deleteBtn.title = 'حذف';
                deleteBtn.onclick = () => removeFromSelected(criteria);
                actions.appendChild(deleteBtn);
                
                item.appendChild(criteriaText);
                item.appendChild(actions);
                selectedList.appendChild(item);
            });
            
            updateCriteriaInput();
        }
        
        function updateTablePreview() {
            const preview = document.getElementById('tablePreview');
            
            if (selectedCriteria.length === 0) {
                preview.innerHTML = '<div class="empty-message">سيظهر معاينة الجدول هنا بعد اختيار المعايير</div>';
                return;
            }
            
            let html = '<table class="preview-table">';
            
            // رأس الجدول - الصف الأول
            html += '<tr>';
            html += '<th rowspan="2">اسم التلميذ</th>';
            selectedCriteria.forEach(criteria => {
                html += `<th colspan="3">${criteria}</th>`;
            });
            html += '</tr>';
            
            // رأس الجدول - الصف الثاني (المؤشرات)
            html += '<tr>';
            selectedCriteria.forEach(() => {
                html += '<th>مؤشر 1</th><th>مؤشر 2</th><th>مؤشر 3</th>';
            });
            html += '</tr>';
            
            // صفوف التلاميذ (3 صفوف كمثال)
            for (let i = 1; i <= 3; i++) {
                html += '<tr>';
                html += `<td>التلميذ ${i}</td>`;
                selectedCriteria.forEach(() => {
                    html += '<td></td><td></td><td></td>';
                });
                html += '</tr>';
            }
            
            html += '</table>';
            html += '<div style="text-align: center; margin-top: 10px; color: #7f8c8d; font-size: 12px;">';
            html += 'هذه معاينة مبسطة للجدول. الملف النهائي سيحتوي على جميع التلاميذ';
            html += '</div>';
            
            preview.innerHTML = html;
        }
        
        function clearAllCriteria() {
            if (confirm('هل أنت متأكد من حذف جميع المعايير المختارة؟')) {
                selectedCriteria = [];
                renderSelectedCriteria();
                updateSuggestedCriteria();
            }
        }
        
        function allowDrop(ev) {
            ev.preventDefault();
        }
        
        function dragStart(ev, criteria, source) {
            ev.dataTransfer.setData("text/plain", JSON.stringify({
                criteria: criteria,
                source: source
            }));
        }
        
        function dropInSuggested(ev) {
            ev.preventDefault();
            const data = JSON.parse(ev.dataTransfer.getData("text/plain"));
            if (data.source === 'selected') {
                removeFromSelected(data.criteria);
            }
        }
        
        function dropInSelected(ev) {
            ev.preventDefault();
            const data = JSON.parse(ev.dataTransfer.getData("text/plain"));
            if (data.source === 'suggested') {
                addToSelected(data.criteria);
            }
        }
        
        // التهيئة الأولية
        document.addEventListener('DOMContentLoaded', function() {
            updateSuggestedCriteria();
        });
    </script>
</body>
</html>
"""

# Groupes complets
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
        
        # Récupération des données
        criteria_json = request.form.get("criteria", "[]")
        criteria = json.loads(criteria_json)
        
        if not criteria:
            criteria = ["معيار 1", "معيار 2", "معيار 3"]

        group_choice = request.form.get("group_choice")
        names = group_new if group_choice == "2" else group_old

        # Création du document
        doc = Document()
        
        # Configuration de la page
        section = doc.sections[0]
        section.page_height = Cm(29.7)
        section.page_width = Cm(21.0)
        section.left_margin = Cm(1.0)
        section.right_margin = Cm(1.0)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        
        # Titre principal
        title = doc.add_heading(f"جدول التقييم - {matiere}", level=1)
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        title_run = title.runs[0]
        title_run.font.size = Pt(14)
        title_run.font.bold = True
        title_run.font.name = 'Arial'

        # Sous-titre
        subtitle = doc.add_paragraph()
        subtitle.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        subtitle_run = subtitle.add_run(f"القسم: {classe}")
        subtitle_run.font.size = Pt(12)
        subtitle_run.font.name = 'Arial'

        doc.add_paragraph().add_run().add_break()

        # Création du tableau avec la structure demandée
        total_cols = 1 + (len(criteria) * 3)  # اسم + 3 خانات لكل معيار
        
        table = doc.add_table(rows=2, cols=total_cols)  # صفين للرأس
        table.style = 'Table Grid'
        table.autofit = False
        
        # الصف الأول من الرأس (دمج الخلايا للمعايير)
        hdr_row1 = table.rows[0]
        hdr_row1.cells[0].text = "اسم التلميذ"
        hdr_row1.cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        col_index = 1
        for criterion in criteria:
            # دمج 3 خانات لكل معيار
            if col_index + 2 < total_cols:
                hdr_row1.cells[col_index].merge(hdr_row1.cells[col_index + 2])
            
            hdr_row1.cells[col_index].text = criterion
            hdr_row1.cells[col_index].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            col_index += 3

        # الصف الثاني من الرأس (المؤشرات)
        hdr_row2 = table.rows[1]
        hdr_row2.cells[0].text = ""  # الخلية الأولى فارغة
        
        col_index = 1
        for criterion in criteria:
            for i in range(3):
                hdr_row2.cells[col_index + i].text = f"مؤشر {i+1}"
                hdr_row2.cells[col_index + i].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            col_index += 3

        # إضافة صفوف التلاميذ
        for name in names:
            row_cells = table.add_row().cells
            row_cells[0].text = name
            row_cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            
            for j in range(total_cols - 1):
                row_cells[j + 1].text = ""
                row_cells[j + 1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # تطبيق التنسيق على جميع الخلايا
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(8)
                        run.font.name = 'Arial'

        # جعل الرأس عريض
        for i in range(2):  # الصفين الأولين
            for cell in table.rows[i].cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

        # ضبط عرض الأعمدة
        for i, column in enumerate(table.columns):
            for cell in column.cells:
                if i == 0:  # عمود الأسماء
                    cell.width = Cm(3.5)
                else:  # أعمدة المؤشرات
                    cell.width = Cm(1.8)

        # تكبير الخط في الرأس قليلاً
        for cell in table.rows[0].cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)

        # إعداد RTL للجدول
        tbl = table._tbl
        tblPr = tbl.tblPr
        bidi = OxmlElement('w:bidiVisual')
        tblPr.append(bidi)

        # محاذاة الجدول
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Sauvegarde
        f = io.BytesIO()
        doc.save(f)
        f.seek(0)
        
        from datetime import datetime
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