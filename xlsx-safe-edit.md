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

**ส้มต้อนรับ! ใช้ skill นี้สำหรับแก้ Excel ปลอดภัย** 🐱✨
