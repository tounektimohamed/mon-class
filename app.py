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
        .indicators-container {
            margin-top: 10px;
            padding: 10px;
            background: #ecf0f1;
            border-radius: 5px;
            display: none;
        }
        .indicator-input {
            display: flex;
            align-items: center;
            gap: 10px;
            margin: 5px 0;
        }
        .indicator-input input {
            flex: 1;
            padding: 8px;
            font-size: 14px;
        }
        .indicator-label {
            font-size: 12px;
            color: #7f8c8d;
            min-width: 80px;
        }
        .indicator-option {
            margin-top: 10px;
            padding: 10px;
            background: #e8f6f3;
            border-radius: 5px;
            border: 1px solid #27ae60;
        }
        .checkbox-group {
            display: flex;
            align-items: center;
            gap: 10px;
            margin: 5px 0;
        }
        .empty-message {
            text-align: center;
            color: #7f8c8d;
            font-style: italic;
            padding: 20px;
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
                    <br>• يمكنك إضافة مؤشرات لكل معيار إذا رغبت
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
                <input type="hidden" name="indicators" id="indicatorsInput" required>
                <input type="hidden" name="use_indicators" id="useIndicatorsInput" value="false">
                
                <div class="criteria-actions" style="justify-content: center; margin-top: 20px;">
                    <button type="button" class="btn-danger" onclick="clearAllCriteria()">حذف الكل</button>
                </div>
            </div>

            <!-- خيار المؤشرات -->
            <div class="form-group">
                <div class="indicator-option">
                    <div class="checkbox-group">
                        <input type="checkbox" id="useIndicators" onchange="toggleIndicatorsOption()">
                        <label for="useIndicators" style="margin: 0; font-weight: normal;">
                            إضافة مؤشرات للتقييم (3 مؤشرات لكل معيار)
                        </label>
                    </div>
                    <div id="indicatorsPreview" style="font-size: 12px; color: #7f8c8d; margin-top: 5px;">
                        سيتم إضافة 3 أعمدة لكل معيار في الجدول النهائي
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
        let indicatorsData = {};
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
            document.getElementById('indicatorsInput').value = JSON.stringify(indicatorsData);
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
                // إضافة مؤشرات افتراضية إذا كان الخيار مفعل
                if (document.getElementById('useIndicators').checked) {
                    indicatorsData[criteria] = ["مؤشر 1", "مؤشر 2", "مؤشر 3"];
                }
                renderSelectedCriteria();
                updateSuggestedCriteria();
            }
        }
        
        function removeFromSelected(criteria) {
            const index = selectedCriteria.indexOf(criteria);
            if (index > -1) {
                selectedCriteria.splice(index, 1);
                delete indicatorsData[criteria];
                renderSelectedCriteria();
                updateSuggestedCriteria();
            }
        }
        
        function editIndicators(criteria) {
            if (!document.getElementById('useIndicators').checked) {
                alert('يجب تفعيل خيار المؤشرات أولاً');
                return;
            }
            
            const indicators = indicatorsData[criteria] || ["مؤشر 1", "مؤشر 2", "مؤشر 3"];
            const newIndicators = [];
            
            for (let i = 0; i < 3; i++) {
                const newName = prompt(`أدخل اسم المؤشر ${i + 1} لـ "${criteria}":`, indicators[i]);
                if (newName === null) return; // User cancelled
                newIndicators.push(newName.trim() || indicators[i]);
            }
            
            indicatorsData[criteria] = newIndicators;
            renderSelectedCriteria();
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
                
                if (document.getElementById('useIndicators').checked) {
                    const indicatorsBtn = document.createElement('button');
                    indicatorsBtn.className = 'action-btn';
                    indicatorsBtn.innerHTML = '📊';
                    indicatorsBtn.title = 'تعديل المؤشرات';
                    indicatorsBtn.onclick = () => editIndicators(criteria);
                    actions.appendChild(indicatorsBtn);
                }
                
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
        
        function toggleIndicatorsOption() {
            const useIndicators = document.getElementById('useIndicators').checked;
            document.getElementById('useIndicatorsInput').value = useIndicators;
            
            if (useIndicators) {
                // إضافة مؤشرات افتراضية للمعايير المختارة
                selectedCriteria.forEach(criteria => {
                    if (!indicatorsData[criteria]) {
                        indicatorsData[criteria] = ["مؤشر 1", "مؤشر 2", "مؤشر 3"];
                    }
                });
            } else {
                // إزالة جميع المؤشرات
                indicatorsData = {};
            }
            
            renderSelectedCriteria();
        }
        
        function clearAllCriteria() {
            if (confirm('هل أنت متأكد من حذف جميع المعايير المختارة؟')) {
                selectedCriteria = [];
                indicatorsData = {};
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
        indicators_json = request.form.get("indicators", "{}")
        use_indicators = request.form.get("use_indicators") == "true"
        
        criteria = json.loads(criteria_json)
        indicators = json.loads(indicators_json)
        
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
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        
        # Titre principal
        title = doc.add_heading(f"جداول التقييم - {matiere}", level=1)
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

        doc.add_paragraph().add_run().add_break()

        # Création du tableau
        if use_indicators:
            # Tableau avec indicateurs
            total_cols = 1  # Colonne des noms
            
            for criterion in criteria:
                total_cols += 3  # 3 colonnes pour chaque critère (المؤشرات)
            
            table = doc.add_table(rows=1, cols=total_cols)
            
            # En-têtes du tableau
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = "الاسم واللقب"
            hdr_cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            col_index = 1
            for criterion in criteria:
                # Fusionner les cellules pour le critère
                if col_index + 2 < total_cols:
                    hdr_cells[col_index].merge(hdr_cells[col_index + 2])
                
                hdr_cells[col_index].text = criterion
                hdr_cells[col_index].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                
                # Ajouter les indicateurs
                indicator_names = indicators.get(criterion, ["مؤشر 1", "مؤشر 2", "مؤشر 3"])
                for i in range(3):
                    if col_index + i < total_cols:
                        indicator_cell = table.rows[0].cells[col_index + i]
                        indicator_cell.text = indicator_names[i] if i < len(indicator_names) else f"مؤشر {i+1}"
                        indicator_cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        indicator_cell.paragraphs[0].runs[0].font.size = Pt(9)
                
                col_index += 3
        else:
            # Tableau simple بدون مؤشرات
            total_cols = 1 + len(criteria)
            table = doc.add_table(rows=1, cols=total_cols)
            
            # En-têtes du tableau
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = "الاسم واللقب"
            hdr_cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            for i, criterion in enumerate(criteria):
                hdr_cells[i + 1].text = criterion
                hdr_cells[i + 1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Style du tableau
        table.style = 'Table Grid'
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Configuration RTL
        tbl = table._tbl
        tblPr = tbl.tblPr
        bidi = OxmlElement('w:bidiVisual')
        tblPr.append(bidi)

        # Appliquer le style aux en-têtes
        for i in range(len(table.rows[0].cells)):
            cell = table.rows[0].cells[i]
            cell.paragraphs[0].runs[0].font.size = Pt(10)
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].runs[0].font.name = 'Arial'

        # Lignes des étudiants
        for name in names:
            row_cells = table.add_row().cells
            row_cells[0].text = name
            row_cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            row_cells[0].paragraphs[0].runs[0].font.size = Pt(9)
            row_cells[0].paragraphs[0].runs[0].font.name = 'Arial'
            
            for j in range(len(row_cells) - 1):
                row_cells[j + 1].text = ""
                row_cells[j + 1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                row_cells[j + 1].paragraphs[0].runs[0].font.size = Pt(9)
                row_cells[j + 1].paragraphs[0].runs[0].font.name = 'Arial'

        # Ajustement des largeurs
        for i, column in enumerate(table.columns):
            for cell in column.cells:
                if i == 0:  # Colonne des noms
                    cell.width = Cm(4.0)
                else:  # Colonnes des critères/المؤشرات
                    cell.width = Cm(2.5)

        # Pied de page
        doc.add_paragraph().add_run().add_break()
        footer = doc.add_paragraph()
        footer.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        footer_text = "تم إنشاء هذا الجدول آلياً - نظام التقييم بالمؤشرات" if use_indicators else "تم إنشاء هذا الجدول آلياً"
        footer_run = footer.add_run(footer_text)
        footer_run.font.size = Pt(9)
        footer_run.font.italic = True
        footer_run.font.name = 'Arial'

        # Sauvegarde
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