# /xlsx-safe-edit - ปลอดภัยแก้ไข Excel

> **ปัญหา**: openpyxl save() ทำให้ drawing/relationships พัง
> **วิธีแก้**: Extract → Edit XML → Re-zip (SAFE METHOD)

---

## 🎯 เมื่อใช้

- แก้ไข Excel ที่มี **chart, image, shape** มากมาย
- ต้อง **ไม่เสียหาย** ไฟล์เดิม
- ต้อง **เปลี่ยนข้อมูล** (ค่า, สูตร, format)

---

## ⚠️ ห้ามทำ

```python
# ❌ WRONG - ทำให้ drawing พัง
from openpyxl import load_workbook
wb = load_workbook('file.xlsx')
sheet = wb.active
sheet['A1'] = 'new value'
wb.save('file.xlsx')  # 💥 DANGER!
```

---

## ✅ วิธีที่ถูก: Extract-XML Method

### Step 1: แตก XLSX (zip)

```python
import zipfile
import shutil

file_path = 'file.xlsx'
extract_dir = 'xlsx_extract'

with zipfile.ZipFile(file_path, 'r') as zip_ref:
    zip_ref.extractall(extract_dir)
```

### Step 2: แก้ไข XML โดยตรง

```python
from lxml import etree

# เปิด sheet XML
sheet_path = f'{extract_dir}/xl/worksheets/sheet1.xml'
parser = etree.XMLParser(remove_blank_text=False)
tree = etree.parse(sheet_path, parser)
root = tree.getroot()

# namespace
ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

# แก้ไข cell values
rows = root.findall('.//x:row', ns)
for row in rows:
    cells = row.findall('.//x:c', ns)
    for cell in cells:
        v_elem = cell.find('.//x:v', ns)
        if v_elem is not None:
            # เปลี่ยนค่า
            v_elem.text = 'new_value'

# บันทึก
tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
```

### Step 3: บีบอัดกลับ

```python
def zipdir(path, ziph):
    for root_dir, dirs, files in os.walk(path):
        for file in files:
            file_full = os.path.join(root_dir, file)
            arcname = os.path.relpath(file_full, path)
            ziph.write(file_full, arcname)

with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
    zipdir(extract_dir, zipf)

# ล้าง temp
shutil.rmtree(extract_dir)
```

---

## 🔑 Key Points

### Column Letter to Number
```python
def col_letter_to_num(letter):
    """A=1, B=2, Z=26, AA=27, AZ=52, BA=53, BK=63, BO=67"""
    num = 0
    for char in letter:
        num = num * 26 + (ord(char) - ord('A') + 1)
    return num

def num_to_col_letter(num):
    """1=A, 2=B, 26=Z, 27=AA, ..."""
    result = ""
    while num > 0:
        num -= 1
        result = chr(num % 26 + ord('A')) + result
        num //= 26
    return result
```

### ค้นหา Sheet Number
- sheet1.xml = sheet ที่ 1 (หน้าหลัก)
- sheet2.xml = sheet ที่ 2 (ข้อมูลนักเรียน)
- sheet6.xml = sheet ที่ 6 (คะแนน1)

### Cell Reference Format
- "A1", "B5", "AA10" = row number at end
- cell.get('r') = get cell reference
- col = ''.join([c for c in cell_ref if c.isalpha()]) = get column letter

### Value Element
```python
v_elem = cell.find('.//x:v', ns)  # value element
if v_elem is not None:
    v_elem.text = 'new_value'  # change value
```

---

## 🧪 ทดสอบ

**Verification Steps:**
1. ✅ ตรวจสอบขนาดไฟล์ (ไม่เปลี่ยน ~5%)
2. ✅ เปิด Excel - ไม่มี Error dialog
3. ✅ ค่าเปลี่ยนแปลงตามที่ต้องการ
4. ✅ Chart/Image/Shape ยังอยู่

---

## 📋 Full Template

```python
import zipfile
import shutil
import os
from lxml import etree

def safe_edit_xlsx(file_path, sheet_number, edits):
    """
    Safe Excel editing without corrupting drawings

    Args:
        file_path: Path to xlsx file
        sheet_number: Sheet number (1-indexed)
        edits: List of (cell_ref, new_value) tuples
    """
    extract_dir = 'xlsx_temp'

    # Step 1: Extract
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)

    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    # Step 2: Edit XML
    sheet_path = f'{extract_dir}/xl/worksheets/sheet{sheet_number}.xml'
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(sheet_path, parser)
    root = tree.getroot()

    ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    for cell_ref, new_val in edits:
        # หา cell ตาม reference
        for elem in root.iter():
            if elem.get('r') == cell_ref:
                v_elem = elem.find('.//x:v', ns)
                if v_elem is not None:
                    v_elem.text = str(new_val)

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)

    # Step 3: Repack
    def zipdir(path, ziph):
        for root_dir, dirs, files in os.walk(path):
            for file in files:
                file_full = os.path.join(root_dir, file)
                arcname = os.path.relpath(file_full, path)
                ziph.write(file_full, arcname)

    with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipdir(extract_dir, zipf)

    shutil.rmtree(extract_dir)
    print(f"✅ Safe edit complete: {file_path}")
```

---

## ⚡ สำหรับ Sheet ที่มี Formula

**ปัญหา**: Formula cells อาจไม่อยู่ใน rows list
**วิธีแก้**:
- แก้เฉพาะ data cells (ค่า hardcoded)
- ปล่อย formula cells ให้ Excel คำนวณเอง
- หรือหา formula cell แล้วแก้ argument ข้างในสูตร

---

## 🎓 เรียนรู้เพิ่ม

- XLSX = ZIP archive ของ XML files
- `/xl/worksheets/sheet*.xml` = ข้อมูล sheet
- `/xl/drawings/` = รูป, chart, shape
- `/_rels/` = relationships (สำคัญ!)
- `[Content_Types].xml` = ประเภท file ทั้งหมด

---

---

## 📚 Sheet "คะแนน1" — โครงสร้างและวิธีแก้เกรด

> ไฟล์: `ปพ.5 ม.3 วิทยาการคำนวณ.xlsx` | Sheet: `sheet6.xml`

### 🗂️ Column Map (pandas index → Excel column)

| pandas | Excel | ชื่อ | หมายเหตุ |
|--------|-------|------|---------|
| 8–14 | I–O | ตัวชี้วัด 1–7 | **INPUT** (hardcoded) max=10 |
| 58 | BG | คะแนนกลางภาค | **INPUT** max=10 |
| 59 | BH | รวมระหว่างเรียน | FORMULA: `SUMIF(I:BG,"<>-1")` |
| 60 | BI | คะแนนปลายภาค (raw) | **INPUT** max=40 |
| 61 | BJ | รวมปลายภาค | FORMULA: `ROUND(BI*BJ$7/BI$7,0)` |
| 65 | BN | รวมคะแนน | FORMULA: `SUM(BH,BM)` |
| 66 | BO | สรุปคะแนน | FORMULA: `ROUND(BN/BN$7*BO$7,0)` |
| 67 | BP | ผลการประเมิน | FORMULA: `VLOOKUP(BO, หน้าหลัก!Q:U, 4, TRUE)` |

### 🔢 Grading Scale (จาก Sheet หน้าหลัก)

| ช่วงคะแนน | เกรด |
|----------|------|
| 0 – 49.49 | 1 |
| 49.5 – 54.49 | 1.5 |
| 54.5 – 59.49 | 2 |
| 59.5 – 64.49 | 2.5 |
| 64.5 – 69.49 | 2.5 |
| **69.5 – 74.49** | **3** |
| 74.5 – 79.49 | 3.5 |
| 80+ | 4 |

### 🧮 สูตรคำนวณ Final Score

```
between   = sum(I:O) + BG         ← รวมระหว่างเรียน (max 80)
exam      = ROUND(BI * 20/40, 0)  ← รวมปลายภาค (max 20)
total     = between + exam         ← สรุปคะแนน (max 100)
```

**เป้าหมาย Grade 3 (total = 70–74)**:
```
between ≈ 56–60  →  sum(I:O) เฉลี่ย 6.5–7.5 + BG = 7–9
exam    ≈ 12–16  →  BI = 24–32
```

### ✅ วิธีแก้เกรดให้เป็น 3 (ถูกต้อง)

```python
import zipfile, shutil, os, random
from lxml import etree

def fix_grade3(file_path):
    extract_dir = 'xlsx_tmp'
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)

    with zipfile.ZipFile(file_path, 'r') as z:
        z.extractall(extract_dir)

    sheet_path = f'{extract_dir}/xl/worksheets/sheet6.xml'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    rows_list = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows_list}

    # ⚠️ CRITICAL: Restore Row 7 (คะแนนเต็ม header) ก่อนเสมอ!
    # ห้ามแก้ row 7 เพราะ formula ทุกตัวอ้างอิง row 7 เป็น max score
    header_restore = {
        'I':'10','J':'10','K':'10','L':'10','M':'10','N':'10','O':'10',
        'BG':'10', 'BI':'40', 'BJ':'20', 'BM':'20', 'BN':'100', 'BO':'100'
    }
    row7 = row_by_num.get(7)
    if row7:
        for cell in row7.findall('.//x:c', ns):
            col = ''.join([c for c in cell.get('r','') if c.isalpha()])
            if col in header_restore:
                v = cell.find('.//x:v', ns)
                if v is not None:
                    v.text = header_restore[col]

    # หา student rows (ตรวจจาก column B = formula จาก เวลาเรียน1)
    student_rows = {}
    for excel_row in range(8, 60):
        row = row_by_num.get(excel_row)
        if not row:
            continue
        for cell in row.findall('.//x:c', ns):
            col = ''.join([c for c in cell.get('r','') if c.isalpha()])
            if col == 'B':
                # ตรวจว่ามี formula (เป็นชื่อนักเรียน ไม่ใช่ header)
                f_elem = cell.find('.//x:f', ns)
                v = cell.find('.//x:v', ns)
                if f_elem is not None and v is not None and v.text:
                    try:
                        num = int(float(v.text))
                        if 1 <= num <= 50:
                            student_rows[excel_row] = num
                    except:
                        pass

    # สร้างคะแนน Grade 3 (ไม่ซ้ำกัน)
    def gen_grade3():
        for _ in range(300):
            scores = [random.choice([6, 7, 7, 8]) for _ in range(7)]
            midterm = random.choice([7, 8, 9])
            exam_raw = random.randint(24, 32)
            exam_portion = round(exam_raw / 40 * 20)
            total = sum(scores) + midterm + exam_portion
            if 70 <= total <= 74:
                return scores, midterm, exam_raw
        return [7,7,6,7,7,6,7], 8, 28

    score_cols = {'I':0,'J':1,'K':2,'L':3,'M':4,'N':5,'O':6}
    for excel_row in sorted(student_rows):
        scores, midterm, exam_raw = gen_grade3()
        row = row_by_num[excel_row]
        for cell in row.findall('.//x:c', ns):
            col = ''.join([c for c in cell.get('r','') if c.isalpha()])
            v = cell.find('.//x:v', ns)
            if v is None:
                continue
            if col in score_cols:
                v.text = str(scores[score_cols[col]])
            elif col == 'BG':
                v.text = str(midterm)
            elif col == 'BI':
                v.text = str(exam_raw)

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)

    def zipdir(path, ziph):
        for rr, dd, ff in os.walk(path):
            for f in ff:
                fp = os.path.join(rr, f)
                ziph.write(fp, os.path.relpath(fp, path))

    with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipdir(extract_dir, zipf)
    shutil.rmtree(extract_dir)
```

### ⚠️ บทเรียนสำคัญ (เรียนรู้จากการทำงานจริง)

1. **ห้ามแก้ Row 7** — เป็น header คะแนนเต็ม ทุก formula อ้างอิง `$7`
2. **ระวัง blank rows** — XML ไม่เก็บ blank rows → ต้องใช้ `row.get('r')` แทน index
3. **ตรวจ cell type ก่อน** — formula cells มี `<f>` element, input cells ไม่มี
4. **ปิด Excel ก่อน save** — ไม่งั้น PermissionError
5. **อย่าแก้ cached value ของ formula** — Excel จะ recalculate ทับทันที
6. **SUMIF(I:BG,"<>-1")** รวมทุก cell ที่ไม่ใช่ -1 → เปลี่ยน I–O และ BG เพียงพอ

---

## 📚 Primary `คะแนน1` in `sheet8.xml` (Two-Semester Layout)

Use this when the workbook is primary level and `คะแนน1` stores both Semester 1 and Semester 2 in the same sheet.

### Semester split

- Semester 1 inputs usually live at `I:BI`
- Semester 2 inputs usually live at `BJ:DV`
- Typical primary exam inputs:
  - `BH` = Semester 1 exam
  - `DI` = Semester 2 exam
  - `DQ` = year total
  - `DR` = 100-point score
  - `DS` = grade result
  - `DU` = grade override

### Row 7 normalization for common primary templates

If row 7 still contains placeholder values for Semester 2, normalize it before editing student rows:

```text
BJ7:BQ7 = 10
DH7 = 80
DI7 = 20
DJ7 = 20
DK7 = 0
DL7 = 0
DM7 = 20
DN7 = 100
DO7 = 100
DP7 = 100
DQ7 = 200
DR7 = 100
```

### Distribution heuristics for target grades

1. Prefer real score inputs before `DU` override.
2. If the user says "distribute across both semesters", keep Semester 1 and Semester 2 totals reasonably close.
3. For grade targets around `3` or `4`, use indicator values mostly in the `6-9` range.
4. Do not give every student full exam scores; keep exams consistent with the indicator pattern.
5. Aim the final `DR/DS` near the lower edge of the requested grade band unless the user asks for a bigger safety margin.
6. Inspect nearby students first so edited rows match the class scoring style.

---

## 📚 Sheet "คุณลักษณะ", "อ่านคิด", "สมรรถนะ" — โครงสร้างและวิธีใส่คะแนน

> ไฟล์: `ปพ.5 ม.3 วิทยาการคำนวณ.xlsx`

### 🗂️ Sheet File Mapping

| Sheet Name | File | rId |
|------------|------|-----|
| คุณลักษณะ | sheet7.xml | rId7 |
| อ่านคิด | sheet8.xml | rId8 |
| สมรรถนะ | sheet9.xml | rId9 |

> ดู mapping จริงใน `xl/_rels/workbook.xml.rels`

---

### 📋 คุณลักษณะ (sheet7.xml)

| Excel | ชื่อ | ประเภท | หมายเหตุ |
|-------|------|--------|---------|
| H–O | คุณลักษณะ 1–8 | **INPUT** | max=10 ต่อด้าน |
| R | รวมคะแนน | FORMULA | `SUM(H:Q)` |
| S | ร้อยละ | FORMULA | `ROUND(R/80*100, 0)` |
| T | ระดับคุณภาพ | FORMULA | `VLOOKUP(S, kun, 5, TRUE)` |

**เกณฑ์ระดับ (kun table):**
| ช่วงร้อยละ | ระดับ |
|-----------|-------|
| 0–49.49 | ไม่ผ่าน |
| 49.5–64.49 | ผ่าน |
| 64.5–79.49 | ดี |
| 79.5–100 | ดีเยี่ยม |

**เป้าหมาย "ดี":** sum(H:O) = 52–63 จาก 80

---

### 📋 อ่านคิด (sheet8.xml)

| Excel | ชื่อ | ประเภท | หมายเหตุ |
|-------|------|--------|---------|
| H–L | ข้อ 1–5 | **INPUT** | max=5 ต่อข้อ |
| M | รวม | FORMULA | `SUM(H:L)` max=25 |
| N | ร้อยละ | FORMULA | `ROUND(M/25*100, 0)` |
| O,P | ผลการประเมิน | FORMULA | `VLOOKUP(N, Arn, 4/5, TRUE)` |

**เป้าหมาย "ดี":** sum(H:L) = 17–19 จาก 25

---

### 📋 สมรรถนะ (sheet9.xml)

| Excel | ชื่อ | ประเภท | หมายเหตุ |
|-------|------|--------|---------|
| H | สมรรถนะ 1 (การสื่อสาร) | **INPUT** | score 0–100 |
| J | สมรรถนะ 2 (การคิด) | **INPUT** | score 0–100 |
| L | สมรรถนะ 3 (การแก้ปัญหา) | **INPUT** | score 0–100 |
| N | สมรรถนะ 4 (ทักษะชีวิต) | **INPUT** | score 0–100 |
| P | สมรรถนะ 5 (เทคโนโลยี) | **INPUT** | score 0–100 |
| I,K,M,O,Q | สรุปรายด้าน | FORMULA | `VLOOKUP(input, capacity, 4)` → 0–3 |
| R | เฉลี่ย | FORMULA | `SUM(H,J,L,N,P)/5` |
| S | ระดับรวม | FORMULA | `VLOOKUP(R, capacity, 4)` → 0–3 |
| T | ผลการประเมิน | FORMULA | `VLOOKUP(S, capacity_level, 2)` |

**เกณฑ์ระดับ (capacity table) — เหมือน kun:**
- 0–49.49 → 0 (ปรับปรุง) | 49.5–64.49 → 1 (พอใช้)
- 64.5–79.49 → 2 (ดี) | 79.5–100 → 3 (ดีเยี่ยม)

**เป้าหมาย "ดี":** แต่ละ H,J,L,N,P = 65–79

---

### ✅ Helper Functions (ใช้ซ้ำสำหรับทุก sheet)

```python
import zipfile, shutil, os, random
from lxml import etree

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def col_to_num(col):
    num = 0
    for c in col:
        num = num * 26 + (ord(c) - ord('A') + 1)
    return num

def ensure_cell(row_elem, col_letter, row_num):
    """หา cell หรือสร้างใหม่ถ้าไม่มี (insert ในตำแหน่งที่ถูกต้อง)"""
    target_ref = f'{col_letter}{row_num}'
    cells = row_elem.findall(f'{{{NS}}}c')
    for c in cells:
        if c.get('r') == target_ref:
            return c
    new_c = etree.Element(f'{{{NS}}}c')
    new_c.set('r', target_ref)
    col_num = col_to_num(col_letter)
    insert_pos = len(cells)
    for i, c in enumerate(cells):
        existing_col = ''.join([ch for ch in c.get('r', '') if ch.isalpha()])
        if col_to_num(existing_col) > col_num:
            insert_pos = i
            break
    row_elem.insert(insert_pos, new_c)
    return new_c

def set_val(row_elem, col_letter, row_num, value):
    """Set numeric value ใน cell (สร้าง <v> ถ้าไม่มี)"""
    c = ensure_cell(row_elem, col_letter, row_num)
    v = c.find(f'{{{NS}}}v')
    if v is None:
        v = etree.SubElement(c, f'{{{NS}}}v')
    v.text = str(value)
```

> **สำคัญ**: ใช้ `ensure_cell` + `set_val` แทนการหา cell ตรงๆ เพราะ INPUT cells ที่ยังไม่มีค่า อาจไม่มี `<v>` element หรืออาจไม่มี `<c>` element เลย

---

### ✅ วิธีใส่คะแนน ทั้ง 3 Sheet (Template)

```python
def fill_assessment_sheets(file_path, kun_scores, read_scores, cap_scores):
    """
    Args:
        kun_scores:  list of 12 lists, each = [s1,s2,...,s8] (0-10 per criterion)
        read_scores: list of 12 lists, each = [s1,s2,s3,s4,s5] (0-5 per item)
        cap_scores:  list of 12 lists, each = [s1,s2,s3,s4,s5] (0-100 per competency)
    """
    extract_dir = 'xlsx_tmp'
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z:
        z.extractall(extract_dir)

    # คุณลักษณะ — sheet7.xml, INPUT: H-O
    _fill_sheet(extract_dir, 'sheet7.xml', ['H','I','J','K','L','M','N','O'], kun_scores)

    # อ่านคิด — sheet8.xml, INPUT: H-L
    _fill_sheet(extract_dir, 'sheet8.xml', ['H','I','J','K','L'], read_scores)

    # สมรรถนะ — sheet9.xml, INPUT: H,J,L,N,P (สลับกับ FORMULA)
    _fill_sheet(extract_dir, 'sheet9.xml', ['H','J','L','N','P'], cap_scores)

    def zipdir(path, ziph):
        for root_dir, dirs, files in os.walk(path):
            for file in files:
                fp = os.path.join(root_dir, file)
                ziph.write(fp, os.path.relpath(fp, path))
    with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipdir(extract_dir, zipf)
    shutil.rmtree(extract_dir)
    print(f'✅ บันทึกสำเร็จ: {file_path}')

def _fill_sheet(extract_dir, sheet_file, input_cols, scores_list):
    sheet_path = f'{extract_dir}/xl/worksheets/{sheet_file}'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    rows = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}
    for idx, excel_row in enumerate(range(8, 8 + len(scores_list))):
        row = row_by_num.get(excel_row)
        if not row:
            continue
        for ci, col in enumerate(input_cols):
            set_val(row, col, excel_row, scores_list[idx][ci])
    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
```

---

### 🎯 ตัวอย่าง: ให้ผลระดับ "ดี" ทุกคน

```python
import random

def gen_kun_di():     # คุณลักษณะ ดี: sum 52-63 จาก 80
    scores = [random.choice([6,7,7,8]) for _ in range(8)]
    while not (52 <= sum(scores) <= 63):
        scores = [random.choice([6,7,7,8]) for _ in range(8)]
    return scores

def gen_read_di():    # อ่านคิด ดี: sum 17-19 จาก 25
    scores = [random.choice([3,4,4]) for _ in range(5)]
    while not (17 <= sum(scores) <= 19):
        scores = [random.choice([3,4,4]) for _ in range(5)]
    return scores

def gen_cap_di():     # สมรรถนะ ดี: 65-79 ต่อด้าน
    return [random.randint(65, 79) for _ in range(5)]

n = 12  # จำนวนนักเรียน
fill_assessment_sheets(
    'ปพ.5 ม.3 วิทยาการคำนวณ.xlsx',
    kun_scores  = [gen_kun_di()  for _ in range(n)],
    read_scores = [gen_read_di() for _ in range(n)],
    cap_scores  = [gen_cap_di()  for _ in range(n)],
)
```

---

### ⚠️ บทเรียนสำคัญ

1. **INPUT cells อาจไม่มี `<c>` element** — ถ้าไม่เคยกรอกค่า XML จะไม่บันทึก cell ว่างไว้ → ต้องสร้างใหม่ด้วย `ensure_cell()`
2. **INPUT cells อาจมี `<c>` แต่ไม่มี `<v>`** — ต้องตรวจและ SubElement ก่อน set
3. **Student rows เริ่มที่ row 8** — rows 8–19 = นักเรียน 12 คน (ปพ.5 ม.3)
4. **ห้ามแก้ row 7** — เป็น header คะแนนเต็ม formula ทุกตัวอ้างอิง `$7`
5. **สมรรถนะ INPUT เป็นคอลัมน์เว้น** — H,J,L,N,P (ไม่ใช่ H-L ต่อเนื่อง)

---

---

## 📚 Sheet "เวลาเรียน1" — โครงสร้างและวิธีอัพเดตตารางเวลาเรียน

> ไฟล์: `ปพ.5 ม.3 วิทยาการคำนวณ.xlsx` | Sheet: `sheet5.xml` (rId5)

### 🗂️ โครงสร้าง Header (Rows 3–7)

| Row | ชื่อ | ตัวอย่าง |
|-----|------|---------|
| 3 | สัปดาห์ | 1, 2, 3... (ที่ col แรกของแต่ละสัปดาห์) |
| 4 | เดือน | ตุลาคม, พฤศจิกายน, ธ.ค.-ม.ค. |
| 5 | วัน | จ, อ, พ, พฤ, ศ (5 คอลัมน์ต่อสัปดาห์) |
| 6 | วันที่ | เลขวันที่ของแต่ละวัน |
| 7 | ชั่วโมงที่ | (ว่าง หรือใส่เลข period) |

### 🗂️ โครงสร้างคอลัมน์

```
H → EU  = ตารางเวลาเรียน (24 สัปดาห์ × 6 คอลัมน์)
            แต่ละสัปดาห์: [จ, อ, พ, พฤ, ศ, separator]
EW      = มาเรียน  FORMULA: COUNTIF($H:$EU, "/")
EX      = ป่วย     FORMULA: COUNTIF($H:$EU, "ป")
EY      = ลา       FORMULA: COUNTIF($H:$EU, "ล")
EZ      = ขาด      FORMULA: COUNTIF($H:$EU, "ข")
EY5     = เวลาเรียนเต็ม (INPUT - ต้องปรับตามจำนวนครั้งที่สอน)
FA      = ร้อยละที่มาเรียน = EW/EY5*100
```

### 🔢 Column Mapping (สูตรหาคอลัมน์ของแต่ละสัปดาห์)

```python
H_COL = 8  # column H = index 8

def week_cols(week_pos):  # 1-indexed
    """คืนค่า [จ, อ, พ, พฤ, ศ] ของสัปดาห์ที่ week_pos"""
    start = H_COL + (week_pos - 1) * 6
    return [col_num_to_letter(start + i) for i in range(5)]

# ตัวอย่าง:
# week_cols(1) → ['H','I','J','K','L']   (สัปดาห์ 1)
# week_cols(2) → ['N','O','P','Q','R']   (สัปดาห์ 2)
# week_cols(3) → ['T','U','V','W','X']   (สัปดาห์ 3)
# Tuesday = cols[1] (index 1)
```

**Sheet มีทั้งหมด 24 สัปดาห์** (H col 8 → EU col 151 = 144 cols / 6 = 24 weeks)

### 📅 วิธีคำนวณวันสอน (ตัวอย่าง เทอม 2 ปี 2568)

```python
from datetime import date, timedelta

start = date(2025, 10, 1)  # วันเริ่มสอน
end   = date(2026, 3, 27)  # วันสิ้นสุด
teach_weekday = 1           # 0=จ, 1=อ, 2=พ, 3=พฤ, 4=ศ

# วันหยุดราชการที่กระทบ
holidays = {
    date(2026, 3, 3): 'วันมาฆบูชา',
    # เพิ่มตามปีที่ใช้
}

# สร้าง teaching_weeks (Monday ของแต่ละสัปดาห์ที่สอน)
# ข้ามสัปดาห์ที่วันสอนตรงกับวันหยุด
teaching_weeks = []
week_start = start - timedelta(days=start.weekday())  # Monday ของสัปดาห์แรก
while week_start <= end:
    teach_date = week_start + timedelta(days=teach_weekday)
    if start <= teach_date <= end and teach_date not in holidays:
        teaching_weeks.append(week_start)
    elif teach_date not in holidays:
        pass  # นอกช่วง semester
    week_start += timedelta(days=7)

# ⚠️ ถ้าได้มากกว่า 24 สัปดาห์ → ข้ามสัปดาห์หยุด (เช่น มาฆบูชา)
# เพื่อให้ fit ใน 24 slots ของ sheet
```

### ✅ Template: อัพเดต Sheet เวลาเรียน (สมบูรณ์)

```python
import zipfile, shutil, os
from lxml import etree
from datetime import date, timedelta

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
H_COL = 8
EU_COL = 151  # col_letter_to_num('EU')

mfull = {1:'มกราคม',2:'กุมภาพันธ์',3:'มีนาคม',4:'เมษายน',
         5:'พฤษภาคม',6:'มิถุนายน',7:'กรกฎาคม',8:'สิงหาคม',
         9:'กันยายน',10:'ตุลาคม',11:'พฤศจิกายน',12:'ธันวาคม'}
mabbr = {1:'ม.ค.',2:'ก.พ.',3:'มี.ค.',4:'เม.ย.',
         5:'พ.ค.',6:'มิ.ย.',7:'ก.ค.',8:'ส.ค.',
         9:'ก.ย.',10:'ต.ค.',11:'พ.ย.',12:'ธ.ค.'}

def month_label(monday):
    fri = monday + timedelta(days=4)
    return mfull[monday.month] if monday.month == fri.month \
           else f'{mabbr[monday.month]}-{mabbr[fri.month]}'

def update_attendance(file_path, teaching_weeks, teach_weekday,
                      sem_start, sem_end, total_hours):
    """
    Args:
        teaching_weeks:  list of date (Monday of each teaching week, max 24)
        teach_weekday:   0=จ, 1=อ, 2=พ, 3=พฤ, 4=ศ
        sem_start/end:   date range of semester
        total_hours:     int — ใส่ใน EY5 (เวลาเรียนเต็ม)
    """
    extract_dir = 'xlsx_tmp_att'
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z: z.extractall(extract_dir)

    sheet_path = f'{extract_dir}/xl/worksheets/sheet5.xml'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    rows = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}

    def col_num_to_letter(n):
        result = ''
        while n > 0:
            n -= 1; result = chr(n % 26 + ord('A')) + result; n //= 26
        return result

    def col_letter_to_num(col):
        num = 0
        for c in col: num = num * 26 + (ord(c) - ord('A') + 1)
        return num

    def week_cols(wp):
        s = H_COL + (wp - 1) * 6
        return [col_num_to_letter(s + i) for i in range(5)]

    def find_or_make(row_elem, col_letter, row_num):
        ref = f'{col_letter}{row_num}'; cn = col_letter_to_num(col_letter)
        for c in row_elem.findall(f'{{{NS}}}c'):
            if c.get('r') == ref: return c
        new_c = etree.Element(f'{{{NS}}}c'); new_c.set('r', ref)
        cells = row_elem.findall(f'{{{NS}}}c'); pos = len(cells)
        for i, c in enumerate(cells):
            if col_letter_to_num(''.join([ch for ch in c.get('r','') if ch.isalpha()])) > cn:
                pos = i; break
        row_elem.insert(pos, new_c); return new_c

    def set_str(row_elem, col, rn, text):
        c = find_or_make(row_elem, col, rn)
        if 't' in c.attrib: del c.attrib['t']
        for ch in list(c):
            if ch.tag.split('}')[-1] in ['v','is']: c.remove(ch)
        c.set('t', 'str')
        v = etree.SubElement(c, f'{{{NS}}}v'); v.text = text

    def set_num(row_elem, col, rn, number):
        c = find_or_make(row_elem, col, rn)
        if 't' in c.attrib: del c.attrib['t']
        for ch in list(c):
            if ch.tag.split('}')[-1] in ['v','is']: c.remove(ch)
        v = etree.SubElement(c, f'{{{NS}}}v'); v.text = str(number)

    # หา shared string index สำหรับ '/'
    ss_path = f'{extract_dir}/xl/sharedStrings.xml'
    ss_tree = etree.parse(ss_path, etree.XMLParser(remove_blank_text=False))
    ss_root = ss_tree.getroot()
    slash_idx = None
    for i, si in enumerate(ss_root.findall('.//x:si', ns)):
        t_list = si.findall('.//x:t', ns)
        if ''.join([t.text or '' for t in t_list]) == '/':
            slash_idx = i; break

    # --- Step 1: Clear H-EU in rows 4, 6, 8-19 ---
    for rn in [4, 6] + list(range(8, 20)):
        row = row_by_num.get(rn)
        if not row: continue
        for c in row.findall(f'{{{NS}}}c'):
            col_l = ''.join([ch for ch in c.get('r','') if ch.isalpha()])
            cn = col_letter_to_num(col_l)
            if H_COL <= cn <= EU_COL:
                if rn in [4, 6] or c.find(f'{{{NS}}}f') is None:
                    if 't' in c.attrib: del c.attrib['t']
                    v = c.find(f'{{{NS}}}v')
                    if v is not None: c.remove(v)

    # --- Step 2: Fill months, dates, and '/' ---
    row4 = row_by_num.get(4)
    row6 = row_by_num.get(6)
    for wp, monday in enumerate(teaching_weeks, 1):
        cols = week_cols(wp)
        set_str(row4, cols[0], 4, month_label(monday))
        for di, col in enumerate(cols):
            d = monday + timedelta(days=di)
            if sem_start <= d <= sem_end:
                set_num(row6, col, 6, d.day)
        teach_date = monday + timedelta(days=teach_weekday)
        if sem_start <= teach_date <= sem_end:
            for rn in range(8, 20):
                row = row_by_num.get(rn)
                if not row: continue
                c = find_or_make(row, cols[teach_weekday], rn)
                for ch in list(c):
                    if ch.tag.split('}')[-1] in ['v','is']: c.remove(ch)
                if slash_idx is not None:
                    c.set('t', 's')
                    v = etree.SubElement(c, f'{{{NS}}}v'); v.text = str(slash_idx)
                else:
                    c.set('t', 'str')
                    v = etree.SubElement(c, f'{{{NS}}}v'); v.text = '/'

    # --- Step 3: Update EY5 (เวลาเรียนเต็ม) ---
    row5 = row_by_num.get(5)
    for c in row5.findall(f'{{{NS}}}c'):
        if c.get('r') == 'EY5':
            v = c.find(f'{{{NS}}}v')
            if v is not None: v.text = str(total_hours)

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    def zipdir(path, ziph):
        for rd, dd, ff in os.walk(path):
            for f in ff:
                fp = os.path.join(rd, f)
                ziph.write(fp, os.path.relpath(fp, path))
    with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipdir(extract_dir, zipf)
    shutil.rmtree(extract_dir)
    print(f'✅ บันทึกสำเร็จ: {file_path}')
```

### 🎯 ตัวอย่างการใช้งาน (เทอม 2 ปี 2568 สอนวันอังคาร)

```python
from datetime import date

# teaching_weeks ที่สร้างจากการคำนวณ (ข้าม Makha Bucha week)
teaching_weeks = [
    date(2025,10,6), date(2025,10,13), date(2025,10,20), date(2025,10,27),
    date(2025,11,3), date(2025,11,10), date(2025,11,17), date(2025,11,24),
    date(2025,12,1), date(2025,12,8), date(2025,12,15), date(2025,12,22),
    date(2025,12,29), date(2026,1,5), date(2026,1,12), date(2026,1,19),
    date(2026,1,26), date(2026,2,2), date(2026,2,9), date(2026,2,16),
    date(2026,2,23), date(2026,3,9), date(2026,3,16), date(2026,3,23),
]

update_attendance(
    file_path      = 'ปพ.5 ม.3 วิทยาการคำนวณ.xlsx',
    teaching_weeks = teaching_weeks,
    teach_weekday  = 1,              # 1 = วันอังคาร
    sem_start      = date(2025,10,1),
    sem_end        = date(2026,3,27),
    total_hours    = 24,
)
```

### ⚠️ บทเรียนสำคัญ

1. **Sheet มี 24 สัปดาห์พอดี** (H→EU = 144 cols / 6 = 24) → ต้องจัด teaching_weeks ให้ ≤ 24
2. **ข้ามสัปดาห์หยุด** — ถ้าวันสอนตรงวันหยุด ให้ข้ามทั้งสัปดาห์นั้นออกจาก list (วันที่จะกระโดด เช่น ก.พ. → มี.ค. ข้าม 1 สัปดาห์)
3. **'/' คือ shared string** — ต้องหา index ใน sharedStrings.xml แล้วใช้ t="s" (อย่าใช้ t="str" เพราะ COUNTIF อาจนับไม่ถูก)
4. **เดือนข้ามเดือน** — ถ้า Monday–Friday คนละเดือน ให้ใช้ abbr เช่น "ธ.ค.-ม.ค."
5. **EY5 = เวลาเรียนเต็ม** — ตั้งให้ตรงกับ len(teaching_weeks) เพื่อให้ % ถูกต้อง
6. **ระวัง FF column** — มีข้อมูลเก่าอยู่นอก H-EU ไม่ต้องสนใจ (ไม่ถูกนับใน COUNTIF)

---

**ส้มต้อนรับ! ใช้ skill นี้สำหรับแก้ Excel ปลอดภัย** 🐱✨
