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
            max-width: 1000px;
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
        .criteria-item.to-delete {
            background: #e74c3c;
            animation: shake 0.5s;
        }
        @keyframes shake {
            0%, 100% { transform: translateX(0); }
            25% { transform: translateX(-5px); }
            75% { transform: translateX(5px); }
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
        .instructions {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 10px;
            border-radius: 5px;
            margin: 10px 0;
            font-size: 14px;
        }
        .criteria-actions {
            display: flex;
            gap: 5px;
            margin-top: 10px;
            justify-content: center;
        }
        .delete-icon {
            margin-right: 5px;
            cursor: pointer;
            font-size: 12px;
            opacity: 0.7;
        }
        .delete-icon:hover {
            opacity: 1;
            color: #ffeb3b;
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
        .delete-confirm {
            background: #e74c3c;
            color: white;
            padding: 10px;
            border-radius: 5px;
            margin: 5px 0;
            text-align: center;
        }
        .delete-confirm-buttons {
            display: flex;
            gap: 10px;
            justify-content: center;
            margin-top: 5px;
        }
        .indicators-container {
            margin-top: 10px;
            padding: 10px;
            background: #ecf0f1;
            border-radius: 5px;
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
        .toggle-indicators {
            background: #9b59b6;
            color: white;
            border: none;
            padding: 5px 10px;
            border-radius: 3px;
            cursor: pointer;
            font-size: 12px;
            margin-top: 5px;
        }
        .toggle-indicators:hover {
            background: #8e44ad;
        }
        .subject-criteria {
            background: #d5edf7;
            padding: 10px;
            border-radius: 5px;
            margin: 10px 0;
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
                <label>نوع التقييم:</label>
                <select id="matiere" name="matiere" required onchange="updateCriteriaBySubject()">
                    <option value="">اختر نوع التقييم</option>
                    <option value="التواصل الشفوي">التواصل الشفوي</option>
                    <option value="القراءة">القراءة</option>
                    <option value="الإنتاج الكتابي">الإنتاج الكتابي</option>
                    <option value="قواعد اللغة">قواعد اللغة</option>
                    <option value="أخرى">أخرى</option>
                </select>
            </div>
            
            <div class="form-group">
                <label>المعايير:</label>
                
                <div class="instructions">
                    💡 <strong>تعليمات:</strong> 
                    <br>• اختر نوع التقييم لتحميل المعايير المقترحة
                    <br>• انقر على ✏️ لتعديل اسم المعيار
                    <br>• انقر على 🗑️ لحذف المعيار  
                    <br>• انقر على 📊 لإظهار/إخفاء المؤشرات
                    <br>• اسحب المعايير لإعادة ترتيبها
                </div>

                <div id="subjectCriteria" class="subject-criteria" style="display: none;">
                    <strong>المعايير المقترحة:</strong>
                    <div id="suggestedCriteria"></div>
                </div>
                
                <div class="criteria-container" id="criteriaContainer" ondragover="allowDrop(event)">
                    <div class="drag-info">اسحب المعايير لإعادة ترتيبها</div>
                </div>
                <input type="hidden" name="criteria" id="criteriaInput" required>
                <input type="hidden" name="indicators" id="indicatorsInput" required>
                
                <div class="criteria-actions">
                    <button type="button" class="btn-danger" onclick="clearAllCriteria()">حذف الكل</button>
                    <button type="button" class="btn-secondary" onclick="addDefaultCriteria()">إضافة معايير افتراضية</button>
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
        let criteriaList = [];
        let indicatorsData = {};
        let itemToDelete = null;
        const subjectCriteria = {
            "التواصل الشفوي": [
                {name: "الملائمة", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]},
                {name: "التغنيم", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]},
                {name: "الانسجام", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]},
                {name: "الاتساق", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]},
                {name: "الثراء", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]}
            ],
            "القراءة": [
                {name: "القراءة الجهرية", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]},
                {name: "معالجة النص", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]},
                {name: "التصرف في النص", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]},
                {name: "إبداء الرأي", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]}
            ],
            "الإنتاج الكتابي": [
                {name: "الملائمة", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]},
                {name: "سلامة بناء النص", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]},
                {name: "المقروئية", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]},
                {name: "ثراء اللغة والطرافة", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]}
            ],
            "قواعد اللغة": [
                {name: "التعرف على الظاهرة اللغوية", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]},
                {name: "توظيف الظاهرة اللغوية", indicators: ["مؤشر 1", "مؤشر 2", "مؤشر 3"]}
            ]
        };
        
        function updateCriteriaInput() {
            document.getElementById('criteriaInput').value = JSON.stringify(criteriaList);
            document.getElementById('indicatorsInput').value = JSON.stringify(indicatorsData);
        }
        
        function updateCriteriaBySubject() {
            const subject = document.getElementById('matiere').value;
            const subjectDiv = document.getElementById('subjectCriteria');
            const suggestedDiv = document.getElementById('suggestedCriteria');
            
            if (subjectCriteria[subject]) {
                subjectDiv.style.display = 'block';
                suggestedDiv.innerHTML = '';
                
                subjectCriteria[subject].forEach(criteria => {
                    const item = document.createElement('span');
                    item.className = 'predefined-item';
                    item.textContent = criteria.name;
                    item.onclick = () => addCriteriaWithIndicators(criteria.name, criteria.indicators);
                    suggestedDiv.appendChild(item);
                });
            } else {
                subjectDiv.style.display = 'none';
            }
        }
        
        function addCriteriaWithIndicators(criteriaName, defaultIndicators) {
            if (!criteriaList.includes(criteriaName)) {
                criteriaList.push(criteriaName);
                indicatorsData[criteriaName] = [...defaultIndicators];
                renderCriteria();
            }
        }
        
        function addDefaultCriteria() {
            const defaultCriteria = ["مع 1", "مع 2", "مع 3"];
            defaultCriteria.forEach(criteria => {
                if (!criteriaList.includes(criteria)) {
                    criteriaList.push(criteria);
                    indicatorsData[criteria] = ["مؤشر 1", "مؤشر 2", "مؤشر 3"];
                }
            });
            renderCriteria();
        }
        
        function removeCriteria(index) {
            const criteriaName = criteriaList[index];
            delete indicatorsData[criteriaName];
            criteriaList.splice(index, 1);
            renderCriteria();
            itemToDelete = null;
        }
        
        function confirmDelete(index) {
            itemToDelete = index;
            renderCriteria();
        }
        
        function cancelDelete() {
            itemToDelete = null;
            renderCriteria();
        }
        
        function clearAllCriteria() {
            if (confirm('هل أنت متأكد من حذف جميع المعايير؟')) {
                criteriaList = [];
                indicatorsData = {};
                renderCriteria();
            }
        }
        
        function editCriteria(index) {
            const container = document.getElementById('criteriaContainer');
            const items = container.querySelectorAll('.criteria-item');
            const item = items[index];
            
            if (item.classList.contains('editing')) {
                return;
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
                const oldName = criteriaList[index];
                // نقل المؤشرات إلى الاسم الجديد
                if (indicatorsData[oldName]) {
                    indicatorsData[trimmedValue] = indicatorsData[oldName];
                    delete indicatorsData[oldName];
                }
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
        
        function toggleIndicators(criteriaName) {
            const indicatorsDiv = document.getElementById(`indicators-${btoa(criteriaName)}`);
            if (indicatorsDiv.style.display === 'none') {
                indicatorsDiv.style.display = 'block';
            } else {
                indicatorsDiv.style.display = 'none';
            }
        }
        
        function updateIndicator(criteriaName, index, value) {
            if (!indicatorsData[criteriaName]) {
                indicatorsData[criteriaName] = ["مؤشر 1", "مؤشر 2", "مؤشر 3"];
            }
            indicatorsData[criteriaName][index] = value;
            updateCriteriaInput();
        }
        
        function renderCriteria() {
            const container = document.getElementById('criteriaContainer');
            container.innerHTML = '<div class="drag-info">اسحب المعايير لإعادة ترتيبها</div>';
            
            if (criteriaList.length === 0) {
                container.innerHTML += '<div class="drag-info">لا توجد معايير مضافة</div>';
                updateCriteriaInput();
                return;
            }
            
            criteriaList.forEach((criteria, index) => {
                if (itemToDelete === index) {
                    const deleteConfirm = document.createElement('div');
                    deleteConfirm.className = 'delete-confirm';
                    deleteConfirm.innerHTML = `
                        هل تريد حذف "${criteria}"؟
                        <div class="delete-confirm-buttons">
                            <button class="btn-danger" onclick="removeCriteria(${index})">نعم، احذف</button>
                            <button class="btn-secondary" onclick="cancelDelete()">إلغاء</button>
                        </div>
                    `;
                    container.appendChild(deleteConfirm);
                } else {
                    const item = document.createElement('div');
                    item.className = 'criteria-item';
                    if (itemToDelete === index) {
                        item.classList.add('to-delete');
                    }
                    item.innerHTML = `
                        <span class="edit-icon" onclick="editCriteria(${index})" title="تعديل">✏️</span>
                        <span class="delete-icon" onclick="confirmDelete(${index})" title="حذف">🗑️</span>
                        <span style="margin-right: 5px; cursor: pointer;" onclick="toggleIndicators('${criteria}')" title="المؤشرات">📊</span>
                        ${criteria}
                    `;
                    item.draggable = true;
                    item.ondragstart = (e) => dragStart(e, index);
                    container.appendChild(item);
                    
                    // إضافة المؤشرات
                    const indicatorsDiv = document.createElement('div');
                    indicatorsDiv.id = `indicators-${btoa(criteria)}`;
                    indicatorsDiv.className = 'indicators-container';
                    indicatorsDiv.style.display = 'none';
                    
                    if (!indicatorsData[criteria]) {
                        indicatorsData[criteria] = ["مؤشر 1", "مؤشر 2", "مؤشر 3"];
                    }
                    
                    indicatorsData[criteria].forEach((indicator, i) => {
                        const indicatorDiv = document.createElement('div');
                        indicatorDiv.className = 'indicator-input';
                        indicatorDiv.innerHTML = `
                            <span class="indicator-label">مؤشر ${i + 1}:</span>
                            <input type="text" 
                                   value="${indicator}" 
                                   onchange="updateIndicator('${criteria}', ${i}, this.value)"
                                   placeholder="أدخل اسم المؤشر">
                        `;
                        indicatorsDiv.appendChild(indicatorDiv);
                    });
                    
                    container.appendChild(indicatorsDiv);
                }
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
        
        // إضافة المعايير الافتراضية عند التحميل
        document.addEventListener('DOMContentLoaded', function() {
            addDefaultCriteria();
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
        
        # Récupération des critères et indicateurs depuis JSON
        criteria_json = request.form.get("criteria", "[]")
        indicators_json = request.form.get("indicators", "{}")
        criteria = json.loads(criteria_json)
        indicators = json.loads(indicators_json)
        
        if not criteria:
            criteria = ["مع 1", "مع 2", "مع 3"]

        group_choice = request.form.get("group_choice")
        names = group_new if group_choice == "2" else group_old

        # Création du document amélioré
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

        # Tableau avec indicateurs
        total_cols = 1  # Colonne des noms
        for criterion in criteria:
            # Pour chaque critère, on ajoute 3 colonnes pour les indicateurs
            total_cols += 3
        
        table = doc.add_table(rows=1, cols=total_cols)
        table.style = 'Table Grid'
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Configuration RTL
        tbl = table._tbl
        tblPr = tbl.tblPr
        bidi = OxmlElement('w:bidiVisual')
        tblPr.append(bidi)

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
            hdr_cells[col_index].paragraphs[0].runs[0].font.bold = True
            
            # Ajouter les indicateurs
            indicator_names = indicators.get(criterion, ["مؤشر 1", "مؤشر 2", "مؤشر 3"])
            for i in range(3):
                if col_index + i < total_cols:
                    indicator_cell = table.rows[0].cells[col_index + i] if i > 0 else hdr_cells[col_index + i]
                    indicator_cell.text = indicator_names[i] if i < len(indicator_names) else f"مؤشر {i+1}"
                    indicator_cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    indicator_cell.paragraphs[0].runs[0].font.size = Pt(9)
            
            col_index += 3

        # Appliquer le style aux en-têtes
        for i in range(total_cols):
            cell = table.rows[0].cells[i]
            cell.paragraphs[0].runs[0].font.size = Pt(10)
            cell.paragraphs[0].runs[0].font.name = 'Arial'

        # Lignes des étudiants
        for name in names:
            row_cells = table.add_row().cells
            row_cells[0].text = name
            row_cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            row_cells[0].paragraphs[0].runs[0].font.size = Pt(9)
            row_cells[0].paragraphs[0].runs[0].font.name = 'Arial'
            
            for j in range(total_cols - 1):
                row_cells[j + 1].text = ""
                row_cells[j + 1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                row_cells[j + 1].paragraphs[0].runs[0].font.size = Pt(9)
                row_cells[j + 1].paragraphs[0].runs[0].font.name = 'Arial'

        # Ajustement des largeurs
        for i, column in enumerate(table.columns):
            for cell in column.cells:
                if i == 0:  # Colonne des noms
                    cell.width = Cm(4.0)
                else:  # Colonnes des indicateurs
                    cell.width = Cm(2.0)

        # Pied de page
        doc.add_paragraph().add_run().add_break()
        footer = doc.add_paragraph()
        footer.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        footer_run = footer.add_run("تم إنشاء هذا الجدول آلياً - نظام التقييم بالمؤشرات")
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