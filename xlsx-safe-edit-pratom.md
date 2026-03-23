# /xlsx-safe-edit-pratom - แก้ไข Excel ปพ.5 ประถม (ปลอดภัย)

> **ไฟล์**: `ปพ.5 ป.X -วิชา.xlsx`
> **วิธี**: Extract → Edit XML → Re-zip (SAFE METHOD เหมือนมัธยม)
> **ความต่างจากมัธยม**: คะแนน1 มี 2 ภาคเรียนในชีตเดียว, สมรรถนะแยกภาค, มีเวลาเรียน2

---

## 🗂️ Sheet Mapping (rId → XML file)

| Sheet Name | rId | XML file |
|------------|-----|---------|
| หน้าหลัก | rId1 | sheet1.xml |
| ข้อมูลนักเรียน | rId2 | sheet2.xml |
| ปก | rId3 | sheet3.xml |
| ประกาศผล | rId4 | sheet4.xml |
| เวลาเรียน1 | rId5 | sheet5.xml |
| เวลาเรียน2 | rId6 | sheet6.xml |
| สรุปเวลาเรียน | rId7 | sheet7.xml |
| **คะแนน1** | rId8 | **sheet8.xml** |
| **คุณลักษณะ** | rId9 | **sheet9.xml** |
| **อ่านคิด** | rId10 | **sheet10.xml** |
| **สมรรถนะ** | rId11 | **sheet11.xml** |

> ⚠️ ยืนยัน mapping จาก `xl/_rels/workbook.xml.rels` ก่อนเสมอ

---

## 🔄 Workflow (ทำตามลำดับ)

```
Step 0: BACKUP ไฟล์ → copy เป็น _backup.xlsx ก่อนเสมอ
Step 1: อ่านไฟล์ → auto-detect ประถม/มัธยม + จำนวนนักเรียน + ข้อมูลหน้าหลัก
Step 2: ถามภาค (ประถม: 1 หรือ 2? / มัธยม: ไม่ต้องถาม)
Step 3: ถามเป้าหมายระดับ (ดี / ดีเยี่ยม / ผ่าน)
Step 4: หน้าหลัก — แสดงค่า + ขอยืนยัน + แก้ไข
Step 5: เวลาเรียน (ประถม: ภาค1=sheet5 / ภาค2=sheet6)
Step 6: คะแนน1
Step 7: คุณลักษณะ
Step 8: อ่านคิด
Step 9: สมรรถนะ
→ แต่ละ step: แจ้งข้อมูลที่จะเปลี่ยน → ขอยืนยัน → ทำ → verify XML
```

**Auto-detect ประถม/มัธยม**: ดูจาก sheet names — มี "เวลาเรียน2" = ประถม

---

## 👥 โครงสร้างนักเรียน — Auto-detect

- **ห้ามใช้ค่า hardcode** — แต่ละห้องจำนวนไม่เท่ากัน
- **ตรวจจากคอลัมน์ C** (เลขประจำตัว) ที่เป็น FORMULA (`<f>` element)
- Row 7 = header คะแนนเต็ม (ห้ามแก้!)

```python
def detect_students(z, sheet_file='sheet8.xml'):
    """Auto-detect student rows จาก col C ที่เป็น FORMULA"""
    data = z.read(f'xl/worksheets/{sheet_file}')
    tree = etree.fromstring(data)
    ns = {'x': NS}
    rows = tree.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}
    student_rows = []
    for rn in range(8, 100):
        row = row_by_num.get(rn)
        if not row: continue
        for c in row.findall('.//x:c', ns):
            col = ''.join([ch for ch in c.get('r', '') if ch.isalpha()])
            if col == 'C':
                f = c.find(f'{{{NS}}}f')
                v = c.find(f'{{{NS}}}v')
                if f is not None and v is not None and v.text:
                    student_rows.append(rn)
    return student_rows
    # ป.1 กอท → [8,9,...,20] = 13 คน
    # ห้องอื่นอาจได้ผลต่างกัน
```

---

## 📋 หน้าหลัก (sheet1.xml) — ตรวจสอบก่อนทำทุกครั้ง

### 🗂️ Fields ที่ต้องตรวจ/อัพเดท

| Cell | ชื่อ | ตัวอย่างค่า | ต้องตรวจ |
|------|------|------------|---------|
| **E13** | ภาคเรียนที่ | `1` | ✅ เปลี่ยนให้ตรงภาคที่ทำ |
| **E12** | รหัสวิชา | `ง 11101` | ✅ ตรวจตามวิชา |
| **I12** | ชื่อวิชา | `การงานอาชีพฯ` | ✅ ตรวจตามวิชา |
| **E15** | ครูประจำวิชา | `นายอุดร มะโนเนือง` | ✅ ตรวจตามวิชา |
| **L15** | ตำแหน่งครู | `ครู ชำนาญการพิเศษ` | ✅ |
| **E16** | **หัวหน้ากลุ่มสาระ** | `นายอุดร มะโนเนือง` | ✅ **ตรวจเสมอ อาจต่างจากครูวิชา** |
| **L16** | ตำแหน่งหัวหน้า | `ครู ชำนาญการพิเศษ` | ✅ |
| **L14** | วันที่อนุมัติ | `29` | ✅ อัพเดทวันจริง |
| **M14** | เดือนอนุมัติ | `มีนาคม` | ✅ อัพเดท |
| **O14** | ปี (พ.ศ. ย่อ) | `68` | ✅ อัพเดท |

> ⚠️ **หัวหน้ากลุ่มสาระ** ≠ ครูประจำวิชาเสมอไป — ต้องถามยืนยันทุกครั้ง

### 💬 Script ยืนยัน (แสดงครบทุก field ก่อนแก้)

```
📋 Sheet 1/7 — หน้าหลัก

ข้อมูลสถานศึกษา (ไม่น่าเปลี่ยน — แจ้งถ้าผิด):
  โรงเรียน         [E5]  : ...
  สังกัด           [E6]  : ...
  ที่อยู่           [E7]  : ...
  หัวหน้างานวัดผล  [E8]  : ...  ตำแหน่ง [L8] : ...
  หัวหน้าฝ่ายวิชาการ [E9] : ...  ตำแหน่ง [L9] : ...
  ผู้อำนวยการ      [E10] : ...  ตำแหน่ง [L10]: ...
  ครูประจำชั้น     [E20] : ...  ตำแหน่ง [K20]: ...

ข้อมูลรายวิชา (ตรวจสอบ):
  รหัสวิชา    [E12] : ...
  ชื่อวิชา    [I12] : ...
  กลุ่มสาระ   [E14] : ...
  ภาคเรียนที่  [E13] : ...  ← เปลี่ยน?
  ปีการศึกษา  [I13] : ...
  เวลาเรียน/ปี [N13] : ... ชม.

บุคลากร (ตรวจสอบ):
  ครูประจำวิชา    [E15] : ...  ตำแหน่ง [L15]: ...
  หัวหน้ากลุ่มสาระ [E16] : ...  ตำแหน่ง [L16]: ...  ← อาจต่างจากครูวิชา

ข้อมูลชั้นเรียน:
  ระดับชั้น  [E19] : ...
  ห้อง       [M19] : ...

สัดส่วนคะแนน:
  ระหว่างเรียน [M17] : ...  ปลายภาค [O17] : ...  ← ตรวจสอบ

วันที่อนุมัติ:
  [L14] วัน / [M14] เดือน / [O14] ปี : ... ... ...  ← เปลี่ยน?

→ อะไรต้องแก้? บอกเป็นข้อ หรือพิมพ์ OK ถ้าถูกต้องทั้งหมด
```

### ✅ Template: อ่านและแสดงหน้าหลัก

```python
def read_main_sheet(file_path):
    """อ่านและแสดงข้อมูลทุก field ใน หน้าหลัก"""
    with zipfile.ZipFile(file_path, 'r') as z:
        ss = z.read('xl/sharedStrings.xml')
        ss_tree = etree.fromstring(ss)
        ns = {'x': NS}
        shared = []
        for si in ss_tree.findall('.//x:si', ns):
            texts = si.findall('.//x:t', ns)
            shared.append(''.join([t.text or '' for t in texts]))

        data = z.read('xl/worksheets/sheet1.xml')
        tree = etree.fromstring(data)
        rows = tree.findall('.//x:row', ns)
        row_by_num = {int(r.get('r',0)): r for r in rows}

        def get(ref):
            col = ''.join([c for c in ref if c.isalpha()])
            rn = int(''.join([c for c in ref if c.isdigit()]))
            row = row_by_num.get(rn)
            if not row: return ''
            for c in row.findall('.//x:c', ns):
                if c.get('r') == ref:
                    v = c.find(f'{{{NS}}}v')
                    t = c.get('t','')
                    val = v.text if v is not None else ''
                    if t == 's' and val and val.isdigit():
                        idx = int(val)
                        return shared[idx] if idx < len(shared) else val
                    return val
            return ''

        print(f'\n📋 Sheet 1/7 — หน้าหลัก\n')
        print('ข้อมูลสถานศึกษา (ไม่น่าเปลี่ยน — แจ้งถ้าผิด):')
        print(f'  โรงเรียน          [E5]  : {get("E5")}')
        print(f'  สังกัด            [E6]  : {get("E6")}')
        print(f'  ที่อยู่            [E7]  : {get("E7")}')
        print(f'  หัวหน้างานวัดผล   [E8]  : {get("E8")}  ตำแหน่ง [L8] : {get("L8")}')
        print(f'  หัวหน้าฝ่ายวิชาการ [E9]  : {get("E9")}  ตำแหน่ง [L9] : {get("L9")}')
        print(f'  ผู้อำนวยการ        [E10] : {get("E10")}  ตำแหน่ง [L10]: {get("L10")}')
        print(f'  ครูประจำชั้น      [E20] : {get("E20")}  ตำแหน่ง [K20]: {get("K20")}')
        print()
        print('ข้อมูลรายวิชา (ตรวจสอบ):')
        print(f'  รหัสวิชา    [E12] : {get("E12")}')
        print(f'  ชื่อวิชา    [I12] : {get("I12")}')
        print(f'  กลุ่มสาระ   [E14] : {get("E14")}')
        print(f'  ภาคเรียนที่  [E13] : {get("E13")}  ← เปลี่ยน?')
        print(f'  ปีการศึกษา  [I13] : {get("I13")}')
        print(f'  เวลาเรียน/ปี [N13] : {get("N13")} ชม.')
        print()
        print('บุคลากร (ตรวจสอบ):')
        print(f'  ครูประจำวิชา     [E15] : {get("E15")}  ตำแหน่ง [L15]: {get("L15")}')
        print(f'  หัวหน้ากลุ่มสาระ  [E16] : {get("E16")}  ตำแหน่ง [L16]: {get("L16")}  ← อาจต่างจากครูวิชา')
        print()
        print('ข้อมูลชั้นเรียน:')
        print(f'  ระดับชั้น  [E19] : {get("E19")}')
        print(f'  ห้อง       [M19] : {get("M19")}')
        print()
        print('สัดส่วนคะแนน:')
        print(f'  ระหว่างเรียน [M17] : {get("M17")}  ปลายภาค [O17] : {get("O17")}  ← ตรวจสอบ')
        print()
        print('วันที่อนุมัติ:')
        print(f'  [L14/M14/O14] : {get("L14")} {get("M14")} {get("O14")}  ← เปลี่ยน?')
        print()
        print('→ อะไรต้องแก้? บอกเป็นข้อ หรือพิมพ์ OK ถ้าถูกต้องทั้งหมด')
```

### ✅ Template: อัพเดทหน้าหลัก

```python
def update_main_sheet(file_path, updates):
    """
    Args:
        updates: dict เช่น {
            'E13': '2',           # ภาคเรียนที่
            'L14': '29',          # วันที่
            'M14': 'มีนาคม',      # เดือน
            'O14': '68',          # ปี
            'E16': 'ชื่อหัวหน้า',  # หัวหน้ากลุ่มสาระ (ถ้าต้องเปลี่ยน)
        }
    """
    extract_dir = 'xlsx_tmp'
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z:
        z.extractall(extract_dir)

    sheet_path = f'{extract_dir}/xl/worksheets/sheet1.xml'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    rows = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}

    def is_number(val):
        try: float(str(val)); return True
        except: return False

    for cell_ref, new_val in updates.items():
        rn = int(''.join([c for c in cell_ref if c.isdigit()]))
        row = row_by_num.get(rn)
        if not row: continue
        for c in row.findall('.//x:c', ns):
            if c.get('r') == cell_ref:
                # ⚠️ text ต้องใส่ t="str" / ตัวเลขไม่ใส่
                if 't' in c.attrib:
                    del c.attrib['t']
                if not is_number(new_val):
                    c.set('t', 'str')   # ← สำคัญ! text ใน <v> ต้องมี t="str"
                v = c.find(f'{{{NS}}}v')
                if v is None:
                    v = etree.SubElement(c, f'{{{NS}}}v')
                v.text = str(new_val)

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    _repack(file_path, extract_dir)
    print(f'✅ หน้าหลัก อัพเดทสำเร็จ')
```

---

## 📚 Sheet คะแนน1 (sheet8.xml) — 2 ภาคเรียนในชีตเดียว

### ⚠️ ความแตกต่างจากมัธยม
ประถมมี **ภาคเรียนที่ 1 และ 2 อยู่ในชีตเดียวกัน** — ต้องแยกให้ถูก

### 🚫 กฎเด็ดขาด — ห้ามแก้ข้อมูลเหล่านี้

| ช่วงคอลัมน์ | ชื่อ | สถานะ |
|------------|------|-------|
| I–P | ตัวชี้วัด 1–8 ภาค1 | 🚫 ห้ามแก้ |
| BG | รวมระหว่างเรียน ภาค1 | 🚫 FORMULA ห้ามแก้ |
| BH | กลางภาค ภาค1 | ⚠️ แก้ได้ถ้าครูสั่ง (เพื่อให้เกรดสอดคล้อง) |
| BI | รวมทั้งหมด ภาค1 | 🚫 FORMULA ห้ามแก้ |
| **BJ–BQ** | **ตัวชี้วัด 1–8 ภาค2** | ✅ INPUT (≥1 ห้ามเป็น 0) |
| **DI** | **ปลายภาค ภาค2** | ✅ INPUT |

---

### 🎯 การกระจายคะแนน BH และ DI ให้สอดคล้องกัน

**หลักการ:** BH (กลางภาค) และ DI (ปลายภาค) ต้องไปทิศทางเดียวกันและใกล้เคียงกัน
- เกรดสูง → BH สูง + DI สูง
- เกรดต่ำ → BH ต่ำ + DI ต่ำ
- BH และ DI ควรต่างกันไม่เกิน 2-3 คะแนน

**สูตรคำนวณ:**
```
DQ_target = กลางช่วงของ grade range
BH + DI = DQ_target - BG - DH
แบ่งเกือบเท่ากัน: BH ≈ DI ≈ (BH+DI) / 2
```

**ตาราง grade → DQ_target (DQ7=200 หลัง BJ7:BQ7=10):**

| เกรด | DR | DQ range | DQ_aim | BJ-BQ per item |
|------|-----|---------|--------|----------------|
| 1 | 50–54 | 100–108 | 104 | 5–6 (DH≈40–44) |
| 1.5 | 55–59 | 110–118 | 114 | 5–7 |
| 2 | 60–64 | 120–128 | 124 | 5–7 (DH≈44–52) |
| 2.5 | 65–69 | 130–138 | 134 | 6–7 (DH≈48–56) |
| 3 | 70–74 | 140–148 | 144 | 6–8 (DH≈52–60) |
| 3.5 | 75–79 | 149–158 | 154 | 7–8 (DH≈56–64) |
| 4 | 80+ | 160+ | 164 | 8–10 |

**สูตร BH และ DI:**
```
BH + DI = DQ_aim - BG - DH
แบ่งเกือบเท่ากัน: BH ≈ DI ≈ (BH+DI)/2  (ต่างกันไม่เกิน 2)
```

> ⚠️ **BG สูง + grade ต่ำ** (เช่น BG=56, grade 1) → DH min=40, BH+DI ≈ 8 → BH=4, DI=4 (บังคับ)
> ⚠️ **BG ต่ำ + grade สูง** (เช่น BG=40, grade 3.5) → DH ต้องสูง ≈ 74 → BJ-BQ items=9–10
> ⚠️ **ตรวจสอบ**: `DQ = BG+BH+DH+DI`, `DR = round(DQ/200*100)` ต้องอยู่ใน grade range
> ⚠️ **BJ-BQ ≥ 5** ทุกช่อง (ห้ามน้อยกว่า)

**ตัวอย่างจริง (ปพ.5 ป.1 กอท 2568 ภาค2, DQ7=200):**

| ชื่อ | เกรด | BG | BH | DH | BJ-BQ | DI | DQ | DR |
|------|------|----|----|----|----|----|----|-----|
| ก้องเกียรติ | 1 | 40 | 10 | 44 | 5,6,5,6,5,5,6,6 | 10 | 104 | 52 |
| สัณห์พิชญ์ | 1 | 56 | 4 | 40 | 5,5,5,5,5,5,5,5 | 4 | 104 | 52 |
| กรวิชญ์ | 2 | 45 | 15 | 48 | 5,7,6,6,5,7,6,6 | 16 | 124 | 62 |
| เขตภากร | 2 | 50 | 13 | 48 | 6,5,6,7,5,6,7,6 | 13 | 124 | 62 |
| กชพร | 2.5 | 61 | 12 | 48 | 6,5,7,5,6,7,6,6 | 13 | 134 | 67 |
| ธนากร | 3 | 59 | 18 | 48 | 6,5,6,7,6,5,7,6 | 19 | 144 | 72 |
| ภูตะวัน | 3 | 54 | 17 | 56 | 7,6,7,8,6,7,8,7 | 17 | 144 | 72 |
| กรฤต | 3 | 53 | 17 | 54 | 7,6,7,7,7,6,8,6 | 18 | 142 | 71 |
| ณัฐพสิษฐ์ | 3.5 | 57 | 18 | 60 | 8,7,7,8,7,8,8,7 | 19 | 154 | 77 |
| ธัญพิชชา | 3.5 | 62 | 18 | 56 | 7,8,6,7,8,6,7,7 | 18 | 154 | 77 |
| พัชรธิดา | 3.5 | 61 | 18 | 56 | 7,6,8,6,7,8,7,7 | 19 | 154 | 77 |
| ณัฐกานต์ | 3.5 | 40 | 20 | 74 | 9,9,10,9,9,10,9,9 | 20 | 154 | 77 |

---

### 🔧 การปรับเกรดนักเรียนรายคน (Single Student Grade Adjustment)

**หลักการ**: ปรับ **BJ-BQ (คะแนนเก็บ)** ก่อนเสมอ — ไม่แตะ BH/DI ถ้าไม่จำเป็น
- เหตุผล: "เก็บดี สอบแย่" สมเหตุสมผลกว่า "สอบดีแต่เก็บแย่" สำหรับเกรดต่ำ-กลาง
- BH/DI ควรสะท้อนระดับเกรดจริง ไม่ควรสูงกว่าคะแนนเก็บมากเกินไป

**ขั้นตอน:**
```
1. คำนวณ DQ_target (กลาง grade range ใหม่)
2. DH_need = DQ_target - BG - BH - DI
3. ออกแบบ BJ-BQ ใหม่ให้ sum = DH_need (ค่าแต่ละช่อง 5-8)
4. อัพเดทไฟล์ + cached values (DH, DJ, DM, DN, DO, DP, DQ, DR)
5. บันทึกลงไฟล์จริงทันที (ไม่ใช่แค่ simulate)
```

**ตัวอย่าง: ปรับ ก้องเกียรติ grade 1 → 1.5**
- BG=40, BH=10, DI=10 (คงเดิม)
- DQ_target=114 (DR=57, grade 1.5)
- DH_need = 114-40-10-10 = **54**
- BJ-BQ ใหม่: [6,7,7,6,7,6,7,8] sum=54 (จากเดิม [5,6,5,6,5,5,6,6] sum=44)
- DQ=114, DR=57 → grade 1.5 ✓

```python
# ปรับเกรดนักเรียนรายคน — แก้ BJ-BQ ให้ DH ตรงเป้า
ROW = 10  # row ของนักเรียนที่ต้องปรับ
new_inds = [6,7,7,6,7,6,7,8]  # BJ-BQ ใหม่ sum=54
BH = 10; DI = 10               # คงเดิม (หรือปรับถ้าจำเป็น)

dh = sum(new_inds)
bg = <อ่านจาก BG{ROW} cached value>
dj = DI; dm = dj; dn = dh + dm
do_ = bg + BH; dq = do_ + dn; dr = round(dq/200*100)

# set BJ-BQ + update cached values ทั้งหมด
for i, col in enumerate(bj_cols):
    set_val(row_el, f'{col}{ROW}', new_inds[i])
set_val(row_el, f'DH{ROW}', dh); set_val(row_el, f'DJ{ROW}', dj)
set_val(row_el, f'DM{ROW}', dm); set_val(row_el, f'DN{ROW}', dn)
set_val(row_el, f'DO{ROW}', do_); set_val(row_el, f'DP{ROW}', dn)
set_val(row_el, f'DQ{ROW}', dq); set_val(row_el, f'DR{ROW}', dr)
# บันทึกไฟล์จริงทันที ไม่ใช่แค่ preview
```

### 🗂️ ภาคเรียนที่ 1 (cols I–BI)

| Excel | Col# | ชื่อ | ประเภท |
|-------|------|------|--------|
| I–P | 9–16 | ตัวชี้วัด 1–8 | **INPUT** max=10 |
| BG | 59 | รวมระหว่างเรียน | FORMULA: `SUMIF(I:BF,"<>-1")` |
| BH | 60 | คะแนนปลายภาค | **INPUT** max=20 |
| BI | 61 | รวมทั้งหมด | FORMULA: `SUMIF(BG:BH,"<>-1")` |

### 🗂️ ภาคเรียนที่ 2 (cols BJ–DV)

| Excel | Col# | ชื่อ | ประเภท |
|-------|------|------|--------|
| BJ–BQ | 62–69 | ตัวชี้วัด 1–8 | **INPUT** max=10 (ต้องสร้าง cell ใหม่) |
| DH | 112 | รวมระหว่างเรียน | FORMULA: `SUMIF(BJ:DG,"<>-1")` |
| DI | 113 | คะแนนปลายภาค | **INPUT** max=20 (ต้องสร้าง cell ใหม่) |
| DK | 115 | คะแนนพิเศษ | **INPUT** (ปกติ=0) |
| DJ | 114 | สัดส่วนปลายภาค | FORMULA: `ROUND(DI8*DJ$7/DI$7,0)` |
| DL | 116 | สัดส่วนพิเศษ | FORMULA |
| DM | 117 | รวมปลาย+พิเศษ | FORMULA: `DJ8+DL8` |
| DN | 118 | รวมภาค2 | FORMULA: `DH8+DM8` |
| DO | 119 | (copy sem1) | FORMULA: `BI8` |
| DP | 120 | (copy sem2 total) | FORMULA: `DN8` |
| DQ | 121 | รวมทั้งปี | FORMULA: `DO8+DP8` |
| DR | 122 | สรุปคะแนน (100) | FORMULA: `ROUND(DQ8/DQ$7*DR$7,0)` |
| DS | 123 | เกรด/ระดับ | FORMULA: VLOOKUP |

### 🔑 Row 7 (header) — ต้องตั้ง BJ7:BQ7=10

```
I-P = 10 (max ตัวชี้วัด ภาค1) — ห้ามแก้
BG=80(F), BH=20(V), BI=100(F) — ห้ามแก้
BJ7:BQ7 = 10 (V) each ← ✅ ต้องตั้งค่านี้ (max ตัวชี้วัด ภาค2)
  → DH7 = 80 (F: SUM(BJ7:DG7))
  → DN7 = 100, DP7 = 100
  → DQ7 = 200 (เปลี่ยนจาก 120)
DI7=20(V), DK7=0(V), DM7=20(V), DR7=100(V)
```

> ⚠️ **CRITICAL — Cached Value Bug**: XML เก็บ cached `<v>` แยกจาก formula
> ถ้าไม่อัพเดท cache: Excel เห็น DQ7=120 → DR=DQ/120×100 → **เกรด 4 ทุกคน**
> ต้องอัพเดท cached value ของ DH7=80, DN7=100, DP7=100, DQ7=200
> **และ** cached value ของทุก formula cell ใน student rows (DH, DJ, DM, DN, DO, DP, DQ, DR)

### ✅ Template: กรอกคะแนน ภาคเรียนที่ 2

```python
import zipfile, shutil, os, random
from lxml import etree

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def col_to_num(col):
    num = 0
    for c in col: num = num * 26 + (ord(c) - ord('A') + 1)
    return num

def ensure_cell(row_elem, col_letter, row_num):
    """หา cell หรือสร้างใหม่"""
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
    """Set numeric value"""
    c = ensure_cell(row_elem, col_letter, row_num)
    v = c.find(f'{{{NS}}}v')
    if v is None:
        v = etree.SubElement(c, f'{{{NS}}}v')
    v.text = str(value)

def fill_score_sem2(file_path, scores_list):
    """
    กรอกคะแนนภาคเรียนที่ 2 ในชีต คะแนน1
    *** ต้องอัพเดท cached values ด้วย ไม่งั้น Excel เห็น DQ7=120 → เกรด 4 ทุกคน ***

    Args:
        scores_list: list of dict per student, เช่น:
            [{'bh': 15, 'indicators': [5,7,6,6,5,7,6,6], 'di': 16}, ...]
            bh = กลางภาค (INPUT, max=20)
            indicators = ตัวชี้วัด BJ-BQ 8 ตัว (≥5 ต่อช่อง)
            di = ปลายภาค (INPUT, max=20)

    Constants (row 7): DJ7=20, DI7=20, DK7=0, DM7=20
    Formula chain: DH=sum(BJ:BQ), DJ=DI, DM=DI, DN=DH+DI, DO=BG+BH, DQ=DO+DN, DR=round(DQ/200*100)
    """
    extract_dir = 'xlsx_tmp'
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z:
        z.extractall(extract_dir)

    sheet_path = f'{extract_dir}/xl/worksheets/sheet8.xml'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    rows = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}

    indicator_cols = ['BJ','BK','BL','BM','BN','BO','BP','BQ']

    # Step 1: Row 7 — ตั้ง BJ7:BQ7=10 → DQ7=200
    row7 = row_by_num.get(7)
    if row7:
        for col in indicator_cols:
            set_val(row7, col, 7, 10)
        # อัพเดท cached values row 7 (formula chain)
        set_val(row7, 'DH', 7, 80)   # SUM(BJ7:DG7) = 80
        set_val(row7, 'DN', 7, 100)  # DH7+DM7 = 80+20
        set_val(row7, 'DP', 7, 100)  # DN7
        set_val(row7, 'DQ', 7, 200)  # DO7+DP7 = 100+100

    # Step 2: Student rows — ใส่คะแนนและอัพเดท cached formula values
    for idx, excel_row in enumerate(range(8, 8 + len(scores_list))):
        row = row_by_num.get(excel_row)
        if not row or idx >= len(scores_list): continue
        data = scores_list[idx]
        bh = data['bh']
        di = data['di']
        inds = data['indicators']
        dh = sum(inds)

        # อ่าน BG (formula cached value)
        bg = 0
        for c in row.findall(f'{{{NS}}}c'):
            if c.get('r') == f'BG{excel_row}':
                v = c.find(f'{{{NS}}}v')
                if v is not None: bg = int(float(v.text))

        # คำนวณ formula chain
        dj = di          # ROUND(DI*DJ7/DI7,0) = ROUND(di*20/20,0) = di
        dm = dj          # DJ+DL = DJ+0
        dn = dh + dm     # DH+DM
        do_ = bg + bh    # BI = BG+BH
        dq = do_ + dn    # DO+DP
        dr = round(dq / 200 * 100)

        # INPUT values
        set_val(row, 'BH', excel_row, bh)
        for ci, col in enumerate(indicator_cols):
            set_val(row, col, excel_row, inds[ci])
        set_val(row, 'DI', excel_row, di)

        # Cached formula values (CRITICAL)
        set_val(row, 'DH', excel_row, dh)
        set_val(row, 'DJ', excel_row, dj)
        set_val(row, 'DM', excel_row, dm)
        set_val(row, 'DN', excel_row, dn)
        set_val(row, 'DO', excel_row, do_)
        set_val(row, 'DP', excel_row, dn)
        set_val(row, 'DQ', excel_row, dq)
        set_val(row, 'DR', excel_row, dr)

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    _repack(file_path, extract_dir)
    print(f'✅ บันทึกคะแนน ภาค2 สำเร็จ: {file_path}')

def _repack(file_path, extract_dir):
    # ลบ calcChain.xml ก่อน repack เสมอ
    # เหตุผล: chain เก่าทำให้ Excel recalculate ตอนเปิด → ถามบันทึกตอนปิดทั้งที่ไม่ได้แก้อะไร
    calc = os.path.join(extract_dir, 'xl', 'calcChain.xml')
    if os.path.exists(calc):
        os.remove(calc)
        ct_path = os.path.join(extract_dir, '[Content_Types].xml')
        with open(ct_path, 'r', encoding='utf-8') as f: ct = f.read()
        ct = ct.replace('<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>', '')
        with open(ct_path, 'w', encoding='utf-8') as f: f.write(ct)

    def zipdir(path, ziph):
        for rd, dd, ff in os.walk(path):
            for f in ff:
                fp = os.path.join(rd, f)
                ziph.write(fp, os.path.relpath(fp, path))
    with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipdir(extract_dir, zipf)
    shutil.rmtree(extract_dir)
```

---

## 📚 Sheet คุณลักษณะ (sheet9.xml)

### 🗂️ โครงสร้าง

| Excel | ชื่อ | ประเภท | หมายเหตุ |
|-------|------|--------|---------|
| H–O | คุณลักษณะ 1–8 | **INPUT** | max=10 ต่อด้าน |
| T | รวมคะแนน | FORMULA | `SUM(H:O)` max=80 |
| U | ร้อยละ | FORMULA | |

**เกณฑ์ระดับ:**
| ช่วงร้อยละ | ระดับ |
|-----------|-------|
| 0–49.49 | ไม่ผ่าน |
| 49.5–64.49 | ผ่าน |
| 64.5–79.49 | ดี |
| 79.5–100 | ดีเยี่ยม |

**เป้าหมาย "ดี":** sum(H:O) = 52–63 จาก 80

> ⚠️ คุณลักษณะ **ไม่แยกภาค** — กรอกครั้งเดียวต่อปี

### ✅ Template: กรอกคุณลักษณะ

```python
def fill_kun_sheet(file_path, kun_scores):
    """
    Args:
        kun_scores: list of 13 lists, each = [s1..s8] (0-10)
    """
    extract_dir = 'xlsx_tmp'
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z:
        z.extractall(extract_dir)

    sheet_path = f'{extract_dir}/xl/worksheets/sheet9.xml'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    rows = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}

    kun_cols = ['H','I','J','K','L','M','N','O']  # 8 คุณลักษณะ
    for idx, excel_row in enumerate(range(8, 8 + len(kun_scores))):
        row = row_by_num.get(excel_row)
        if not row: continue
        for ci, col in enumerate(kun_cols):
            set_val(row, col, excel_row, kun_scores[idx][ci])

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    _repack(file_path, extract_dir)
    print(f'✅ คุณลักษณะ บันทึกสำเร็จ')
```

---

## 📚 Sheet อ่านคิด (sheet10.xml)

### 🗂️ โครงสร้าง

| Excel | ชื่อ | ประเภท | หมายเหตุ |
|-------|------|--------|---------|
| H, I | การอ่าน ข้อ 1–2 | **INPUT** | max=5 ต่อข้อ |
| J, K | การคิดฯ ข้อ 3–4 | **INPUT** | max=5 ต่อข้อ |
| L | เขียน ข้อ 5 | **INPUT** | max=5 |
| M | รวมคะแนน | FORMULA | max=25 |
| N | รวมร้อยละ | FORMULA | |

**เป้าหมาย "ดี":** sum(H:L) = 17–19 จาก 25

> ⚠️ อ่านคิด **ไม่แยกภาค** — กรอกครั้งเดียวต่อปี

### ✅ Template: กรอกอ่านคิด

```python
def fill_read_sheet(file_path, read_scores):
    """
    Args:
        read_scores: list of 13 lists, each = [s1..s5] (0-5 per item)
    """
    extract_dir = 'xlsx_tmp'
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z:
        z.extractall(extract_dir)

    sheet_path = f'{extract_dir}/xl/worksheets/sheet10.xml'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    rows = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}

    read_cols = ['H','I','J','K','L']  # 5 ข้อ
    for idx, excel_row in enumerate(range(8, 8 + len(read_scores))):
        row = row_by_num.get(excel_row)
        if not row: continue
        for ci, col in enumerate(read_cols):
            set_val(row, col, excel_row, read_scores[idx][ci])

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    _repack(file_path, extract_dir)
    print(f'✅ อ่านคิด บันทึกสำเร็จ')
```

---

## 📚 Sheet สมรรถนะ (sheet11.xml)

### 🗂️ โครงสร้าง — แยกภาค!

| สมรรถนะ | ภาค1 INPUT | ภาค2 INPUT | รวม(F) | สรุป(F) |
|---------|-----------|-----------|--------|---------|
| 1 การสื่อสาร | **H** | **I** | J | K |
| 2 การคิด | **L** | **M** | N | O |
| 3 การแก้ปัญหา | **P** | **Q** | R | S |
| 4 ทักษะชีวิต | **T** | **U** | V | W |
| 5 เทคโนโลยี | **X** | **Y** | Z | AA |
| รวมภาค1 | AB(F) | | | |
| รวมภาค2 | AC(F) | | | |
| รวมทั้งปี | AD(F) | | | |
| สรุปรวม | AE(F) | | | |

**เกณฑ์ (capacity table):**
- 0–49 → 0 (ปรับปรุง) | 50–64 → 1 (พอใช้)
- 65–79 → 2 (ดี) | 80–100 → 3 (ดีเยี่ยม)

**เป้าหมาย "ดี" ต่อด้าน:** 65–79 (ต่อภาค)

### ✅ Template: กรอกสมรรถนะ

```python
def fill_cap_sheet(file_path, cap_sem1, cap_sem2):
    """
    Args:
        cap_sem1: list of 13 lists, each = [s1,s2,s3,s4,s5] (0-100 ต่อด้าน) ภาค1
        cap_sem2: list of 13 lists, each = [s1,s2,s3,s4,s5] (0-100 ต่อด้าน) ภาค2
    """
    extract_dir = 'xlsx_tmp'
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z:
        z.extractall(extract_dir)

    sheet_path = f'{extract_dir}/xl/worksheets/sheet11.xml'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    rows = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}

    # ภาค1 = H,L,P,T,X | ภาค2 = I,M,Q,U,Y
    sem1_cols = ['H','L','P','T','X']
    sem2_cols = ['I','M','Q','U','Y']

    for idx, excel_row in enumerate(range(8, 8 + len(cap_sem1))):
        row = row_by_num.get(excel_row)
        if not row: continue
        for ci, col in enumerate(sem1_cols):
            set_val(row, col, excel_row, cap_sem1[idx][ci])
        for ci, col in enumerate(sem2_cols):
            set_val(row, col, excel_row, cap_sem2[idx][ci])

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    _repack(file_path, extract_dir)
    print(f'✅ สมรรถนะ บันทึกสำเร็จ')
```

---

## 🎯 ตัวอย่าง: กรอกทุกชีตครั้งเดียว (ภาคเรียนที่ 2)

```python
import random

FILE = 'ปพ.5 ป.1 -กอท.xlsx'
N = 13  # จำนวนนักเรียน

# --- Generator functions ---

def gen_indicator_sem2():
    """ตัวชี้วัด 8 ข้อ ภาค2 — ระดับดี (ไม่เหมือนกันทุกคน)"""
    for _ in range(200):
        scores = [random.choice([6,7,7,8]) for _ in range(8)]
        exam = random.randint(14, 18)
        total_sem2 = sum(scores) + exam  # ~70-82
        if 65 <= sum(scores) <= 76 and 14 <= exam <= 18:
            return scores, exam
    return [7,7,7,6,7,7,6,7], 15

def gen_kun_di():
    scores = [random.choice([6,7,7,8]) for _ in range(8)]
    while not (52 <= sum(scores) <= 63):
        scores = [random.choice([6,7,7,8]) for _ in range(8)]
    return scores

def gen_read_di():
    scores = [random.choice([3,4,4]) for _ in range(5)]
    while not (17 <= sum(scores) <= 19):
        scores = [random.choice([3,4,4]) for _ in range(5)]
    return scores

def gen_cap_di():
    return [random.randint(65, 79) for _ in range(5)]

# --- Generate scores ---
sem2_data = []
for _ in range(N):
    indicators, exam = gen_indicator_sem2()
    sem2_data.append({'indicators': indicators, 'exam': exam})

kun_scores  = [gen_kun_di()  for _ in range(N)]
read_scores = [gen_read_di() for _ in range(N)]
cap_sem1    = [gen_cap_di()  for _ in range(N)]  # ถ้ายังไม่ได้กรอก ภาค1
cap_sem2    = [gen_cap_di()  for _ in range(N)]

# --- Fill all sheets ---
fill_score_sem2(FILE, sem2_data)
fill_kun_sheet(FILE, kun_scores)
fill_read_sheet(FILE, read_scores)
fill_cap_sheet(FILE, cap_sem1, cap_sem2)
```

---

## 📅 Sheet เวลาเรียน2 (sheet6.xml) — บันทึกการเช็คชื่อ ภาค2

### 🗂️ โครงสร้าง

| Row | ชื่อ | รายละเอียด |
|-----|------|-----------|
| 3 | สัปดาห์ | 1–24 (slot H, N, T, Z ... แต่ละ 6 col) |
| 4 | เดือน | ชื่อเดือน/ช่วงเดือน ต่อ week slot |
| 5 | วัน | จ อ พ พฤ ศ (5 วันต่อ week) |
| 6 | วันที่ | ตัวเลขวันที่ (number) |
| 7 | ชั่วโมงที่ | ตัวเลขคาบ (1,2,3...) เฉพาะวันที่สอน |
| 8–N | นักเรียน | เช็คชื่อ `/` `ป` `ล` `ข` ต่อวัน |
| EW–FA | สรุป | มาเรียน / ป่วย / ลา / ขาด / ร้อยละ (FORMULA) |

**Column layout ต่อสัปดาห์:**
- Week n (0-indexed): base_col = 8 + n×6 (H=col8)
- จันทร์=+0, อังคาร=+1, **พุธ=+2**, พฤหัส=+3, ศุกร์=+4, gap=+5
- `EY5 = COUNTIF(H7:EU7,">=0")` → เวลาเรียนเต็ม (นับจาก row 7)

### ✅ Template: กรอกเวลาเรียน2 (วันสอนเดียว + stdlib ET)

```python
import zipfile, shutil, os, sys, re
from datetime import date, timedelta
import xml.etree.ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', NS)

MONTH_TH   = {1:'มกราคม',2:'กุมภาพันธ์',3:'มีนาคม',4:'เมษายน',
              5:'พฤษภาคม',6:'มิถุนายน',7:'กรกฎาคม',8:'สิงหาคม',
              9:'กันยายน',10:'ตุลาคม',11:'พฤศจิกายน',12:'ธันวาคม'}
MONTH_ABBR = {1:'ม.ค.',2:'ก.พ.',3:'มี.ค.',4:'เม.ย.',
              5:'พ.ค.',6:'มิ.ย.',7:'ก.ค.',8:'ส.ค.',
              9:'ก.ย.',10:'ต.ค.',11:'พ.ย.',12:'ธ.ค.'}

def col_letter(n):   # 1-indexed → letter e.g. 8→H
    r=''
    while n>0:
        n,m=divmod(n-1,26); r=chr(65+m)+r
    return r

def col_index(s):    # letter → 1-indexed e.g. H→8
    r=0
    for c in s: r=r*26+ord(c)-64
    return r

def fill_attendance_sem2(file_path, sem_start, sem_end,
                         teach_weekday=2,   # 0=จ 1=อ 2=พ 3=พฤ 4=ศ
                         period=1,          # คาบที่
                         mark='/',          # เครื่องหมาย (/ ป ล ข)
                         holidays=None,     # set of date
                         student_rows=None  # list of row ints e.g. [8..20]
                         ):
    """
    กรอกเวลาเรียน2 ทั้งหมด
    - วนทุก week จาก sem_start ถึง sem_end
    - กรอก row4(เดือน), row6(วันที่), row7(คาบ), rows8-N(เช็คชื่อ)
    """
    if holidays is None: holidays = set()
    if student_rows is None: student_rows = list(range(8, 21))

    # --- หา Monday แรก ---
    mon = sem_start
    while mon.weekday() != 0:
        mon += timedelta(days=1)

    weeks = []
    while mon <= sem_end:
        weeks.append([mon+timedelta(days=i) for i in range(5)])
        mon += timedelta(days=7)

    # --- build updates {(row, col_str): (value, type)} ---
    updates = {}
    for n, wd in enumerate(weeks):
        base = 8 + n*6
        # row 4: month header
        m0, m4 = wd[0], wd[4]
        mlabel = MONTH_TH[m0.month] if m0.month==m4.month \
                 else f"{MONTH_ABBR[m0.month]}-{MONTH_ABBR[m4.month]}"
        updates[(4, col_letter(base))] = (mlabel, 'str')
        # row 6: dates
        for d in range(5):
            dd = wd[d]
            if dd <= sem_end:
                updates[(6, col_letter(base+d))] = (str(dd.day), 'n')
        # teaching day
        tday = wd[teach_weekday]
        tc   = col_letter(base+teach_weekday)
        if tday <= sem_end and tday not in holidays:
            updates[(7, tc)] = (str(period), 'n')
            for sr in student_rows:
                updates[(sr, tc)] = (mark, 'str')

    # --- patch XML in-memory (no extract) ---
    with zipfile.ZipFile(file_path) as z:
        names = z.namelist()
        files = {n2: z.read(n2) for n2 in names}

    root = ET.fromstring(files['xl/worksheets/sheet6.xml'].decode('utf-8'))
    sd   = root.find(f'{{{NS}}}sheetData')
    row_map = {int(r.get('r')): r for r in sd.findall(f'{{{NS}}}row')}

    def get_row(rn):
        if rn not in row_map:
            nr=ET.SubElement(sd,f'{{{NS}}}row'); nr.set('r',str(rn))
            row_map[rn]=nr
        return row_map[rn]

    def set_cell(row_el, cref, value, vtype):
        col_n = col_index(''.join(filter(str.isalpha, cref)))
        c = next((x for x in row_el.findall(f'{{{NS}}}c') if x.get('r')==cref), None)
        if c is None:
            c = ET.Element(f'{{{NS}}}c'); c.set('r', cref)
            idx = sum(1 for x in row_el.findall(f'{{{NS}}}c')
                      if col_index(''.join(filter(str.isalpha,x.get('r')))) < col_n)
            row_el.insert(idx, c)
        for tag in [f'{{{NS}}}v', f'{{{NS}}}f']:
            el=c.find(tag)
            if el is not None: c.remove(el)
        if vtype=='str': c.set('t','str')
        elif 't' in c.attrib: del c.attrib['t']
        v=ET.SubElement(c,f'{{{NS}}}v'); v.text=value

    for (rn, cl),(val,vtype) in sorted(updates.items()):
        set_cell(get_row(rn), f'{cl}{rn}', val, vtype)

    # re-sort rows
    rows_sorted = sorted(sd.findall(f'{{{NS}}}row'), key=lambda r:int(r.get('r')))
    for r in list(sd): sd.remove(r)
    for r in rows_sorted: sd.append(r)

    files['xl/worksheets/sheet6.xml'] = ET.tostring(root,'utf-8',xml_declaration=True)

    # ลบ calcChain
    files.pop('xl/calcChain.xml', None)
    ct = files['[Content_Types].xml'].decode('utf-8')
    ct = re.sub(r'<Override[^>]+calcChain[^>]+/>', '', ct)
    files['[Content_Types].xml'] = ct.encode('utf-8')

    tmp = file_path+'.tmp'
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
        for name,data in files.items(): zout.writestr(name,data)
    os.replace(tmp, file_path)
    print(f'✅ เวลาเรียน2 บันทึกสำเร็จ: {len([k for k in updates if k[0]>=8])} cells เช็คชื่อ')


# --- ตัวอย่างใช้งาน ---
# วันหยุดราชการ พ.ย.68 – มี.ค.69
holidays_sem2_68 = {
    date(2025,12, 5),   # วันพ่อแห่งชาติ (ศุกร์)
    date(2025,12,10),   # วันรัฐธรรมนูญ (พุธ) ← กระทบ!
    date(2026, 1, 1),   # วันขึ้นปีใหม่ (พฤหัส)
    date(2026, 2,12),   # วันมาฆบูชา (พฤหัส)
}

fill_attendance_sem2(
    'ปพ.5 ป.1 -กอท.xlsx',
    sem_start  = date(2025,11,1),
    sem_end    = date(2026,3,31),
    teach_weekday = 2,      # วันพุธ
    period        = 1,      # คาบที่ 1
    mark          = '/',    # มาทุกคน
    holidays      = holidays_sem2_68,
    student_rows  = list(range(8,21)),  # 13 คน
)
```

### 📌 หมายเหตุ

- `teach_weekday`: 0=จ 1=อ 2=พ 3=พฤ 4=ศ
- วันที่สอนหลายวัน: เรียก `fill_attendance_sem2` หลายครั้ง (ต่างค่า `teach_weekday`)
- วันหยุดราชการต้องตรวจทุกปี — วันที่กระทบคือวันตรงกับ `teach_weekday` เท่านั้น
- ถ้าบางนักเรียนขาด/ลา: สร้าง updates แยกแล้ว patch ทับหลังจาก fill ครั้งแรก

---

## 📅 Sheet เวลาเรียน1 (sheet5.xml) — บันทึกการเช็คชื่อ ภาค1

### ⚠️ ต่างจาก เวลาเรียน2 ตรงนี้

| | sheet5 (ภาค1) | sheet6 (ภาค2) |
|--|--|--|
| Row 4 เดือน | **มีแล้ว** (pre-filled) | ต้องสร้าง |
| Row 6 วันที่ | **มีแล้ว** (pre-filled) | ต้องสร้าง |
| Row 7 คาบ | ต้องเติม | ต้องสร้าง |
| Attendance 8-N | ต้องเติม | ต้องสร้าง |

→ sheet5 ไม่ต้องคำนวณ calendar — loop col offset +2 ต่อ week slot ได้เลย

### 📌 วิธีตรวจวันหยุด

ตรวจ holidays จาก **dates จริงใน row 6 ของ "พ" column** (ไม่ใช่จาก day-of-week จริง เพราะ template อาจไม่ตรงกับ calendar จริง)

```python
# อ่าน dates จาก row 6 col "พ" (offset +2) แล้วเช็ค holidays
wed_dates = []
for n in range(20):  # 20 week slots
    wed_col = col_letter(8 + n*6 + 2)
    day_str = get_val(row6, wed_col)
    month   = get_month_from_row4(n)   # เดือนจาก row 4
    year    = 2025                     # ปี ค.ศ.
    d = date(year, month, int(day_str))
    if d not in holidays:
        wed_dates.append((d, wed_col))
```

### ✅ Template: กรอกเวลาเรียน1 (sheet5 — dates pre-filled)

```python
import zipfile, os, re
import xml.etree.ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', NS)

def col_letter(n):
    r=''
    while n>0: n,m=divmod(n-1,26); r=chr(65+m)+r
    return r
def col_index(s):
    r=0
    for c in s: r=r*26+ord(c)-64
    return r

def fill_attendance_sem1(file_path,
                         teach_offset=2,      # 0=จ 1=อ 2=พ(default) 3=พฤ 4=ศ
                         period=1,
                         mark='/',
                         holidays=None,       # set of date — ตรวจจาก row6 col ที่สอน
                         student_rows=None,
                         num_weeks=20):       # จำนวน week slots ในชีต
    """
    กรอกเวลาเรียน1 (sheet5)
    - rows 4,6 มี dates/months อยู่แล้ว — ไม่ต้องสร้าง
    - เติมเฉพาะ row 7 (คาบ) + attendance rows
    - ตรวจ holidays จาก dates จริงใน row 6
    """
    if holidays is None: holidays = set()
    if student_rows is None: student_rows = list(range(8, 21))

    with zipfile.ZipFile(file_path) as z:
        with z.open('xl/sharedStrings.xml') as f:
            ss_root = ET.parse(f).getroot()
        files = {n: z.read(n) for n in z.namelist()}

    shared = [''.join(t.text or '' for t in si.findall(f'.//{{{NS}}}t'))
              for si in ss_root.findall(f'{{{NS}}}si')]

    root = ET.fromstring(files['xl/worksheets/sheet5.xml'].decode('utf-8'))
    sd   = root.find(f'{{{NS}}}sheetData')
    row_map = {int(r.get('r')): r for r in sd.findall(f'{{{NS}}}row')}

    MONTH_NUM = {
        'มกราคม':1,'กุมภาพันธ์':2,'มีนาคม':3,'เมษายน':4,
        'พฤษภาคม':5,'มิถุนายน':6,'กรกฎาคม':7,'สิงหาคม':8,
        'กันยายน':9,'ตุลาคม':10,'พฤศจิกายน':11,'ธันวาคม':12,
        'พ.ค.-มิ.ย.':5,'มิ.ย.-ก.ค.':6,'ก.ค.-ส.ค.':7,'ส.ค.-ก.ย.':8,
        'ก.ย.-ต.ค.':9,'ต.ค.-พ.ย.':10,'พ.ย.-ธ.ค.':11,'ธ.ค.-ม.ค.':12,
        'ม.ค.-ก.พ.':1,'ก.พ.-มี.ค.':2,
    }

    def cell_val(row_el, col_str):
        if row_el is None: return ''
        cref = f"{col_str}{row_el.get('r')}"
        for c in row_el.findall(f'{{{NS}}}c'):
            if c.get('r') == cref:
                v = c.find(f'{{{NS}}}v')
                t = c.get('t','')
                if v is None: return ''
                return shared[int(v.text)] if t=='s' else (v.text or '')
        return ''

    def get_row(rn):
        if rn not in row_map:
            nr=ET.SubElement(sd,f'{{{NS}}}row'); nr.set('r',str(rn))
            row_map[rn]=nr
        return row_map[rn]

    def set_cell(row_el, cref, value, vtype):
        col_n = col_index(''.join(filter(str.isalpha, cref)))
        c = next((x for x in row_el.findall(f'{{{NS}}}c') if x.get('r')==cref), None)
        if c is None:
            c = ET.Element(f'{{{NS}}}c'); c.set('r', cref)
            idx = sum(1 for x in row_el.findall(f'{{{NS}}}c')
                      if col_index(''.join(filter(str.isalpha, x.get('r')))) < col_n)
            row_el.insert(idx, c)
        for tag in [f'{{{NS}}}v', f'{{{NS}}}f']:
            el=c.find(tag)
            if el is not None: c.remove(el)
        if vtype=='str': c.set('t','str')
        elif 't' in c.attrib: del c.attrib['t']
        v=ET.SubElement(c,f'{{{NS}}}v'); v.text=value

    r4 = row_map.get(4)
    r6 = row_map.get(6)
    filled = 0
    skipped = []

    for n in range(num_weeks):
        base     = 8 + n*6
        teach_col = col_letter(base + teach_offset)

        # อ่าน month จาก row4 (base col)
        month_label = cell_val(r4, col_letter(base))
        month_num   = MONTH_NUM.get(month_label, 0)

        # อ่าน date จาก row6 (teach col)
        day_str = cell_val(r6, teach_col)
        if not day_str or not month_num:
            continue

        # ตรวจ holiday
        if holidays:
            from datetime import date as _date
            year = 2025  # ภาค1 = 2025
            if month_num < 5:  # ถ้าน้อยกว่า พ.ค. แสดงว่าข้ามปีไม่ได้
                year = 2026
            try:
                d = _date(year, month_num, int(day_str))
                if d in holidays:
                    skipped.append(f"{teach_col} ({d})")
                    continue
            except:
                pass

        # เติม row7 + attendance
        set_cell(get_row(7), f'{teach_col}7', str(period), 'n')
        for sr in student_rows:
            set_cell(get_row(sr), f'{teach_col}{sr}', mark, 'str')
            filled += 1

    # re-sort rows
    rows_sorted = sorted(sd.findall(f'{{{NS}}}row'), key=lambda r:int(r.get('r')))
    for r in list(sd): sd.remove(r)
    for r in rows_sorted: sd.append(r)

    files['xl/worksheets/sheet5.xml'] = ET.tostring(root,'utf-8',xml_declaration=True)
    files.pop('xl/calcChain.xml', None)
    ct = files['[Content_Types].xml'].decode('utf-8')
    ct = re.sub(r'<Override[^>]+calcChain[^>]+/>', '', ct)
    files['[Content_Types].xml'] = ct.encode('utf-8')

    tmp = file_path+'.tmp'
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
        for name,data in files.items(): zout.writestr(name,data)
    os.replace(tmp, file_path)

    weeks_taught = filled // len(student_rows)
    print(f'✅ เวลาเรียน1 บันทึกสำเร็จ: {weeks_taught} สัปดาห์, {filled} cells')
    if skipped:
        print(f'  ข้าม (วันหยุด): {skipped}')


# --- ตัวอย่างใช้งาน ---
from datetime import date

holidays_sem1_68 = {
    date(2025, 5, 1),   # วันแรงงาน
    date(2025, 5, 5),   # วันฉัตรมงคล
    date(2025, 5,12),   # วันวิสาขบูชา
    date(2025, 6, 3),   # วันเฉลิมพระชนมพรรษา ร.10
    date(2025, 7,28),   # วันเฉลิมพระชนมพรรษา ร.10
    date(2025, 8,12),   # วันแม่แห่งชาติ
}

fill_attendance_sem1(
    'ปพ.5 ป.1 -กอท.xlsx',
    teach_offset = 2,     # วันพุธ
    period       = 1,     # คาบที่ 1
    mark         = '/',   # มาทุกคน
    holidays     = holidays_sem1_68,
    student_rows = list(range(8, 21)),  # 13 คน
    num_weeks    = 20,
)
```

---

## ⚠️ บทเรียนสำคัญ (ต่างจากมัธยม + สิ่งที่มักลืม)

1. **คะแนน1 มี 2 ภาคในชีตเดียว** — ภาค1 = I–BI, ภาค2 = BJ–DV ห้ามสับสน
2. **ภาค2 INPUT cells ยังไม่มีใน XML** — ต้องใช้ `ensure_cell()` สร้างใหม่
3. **สมรรถนะแยกภาค** — H,L,P,T,X = ภาค1 / I,M,Q,U,Y = ภาค2
4. **ห้ามแก้ Row 7** — เป็น header คะแนนเต็ม formula ทุกตัวอ้างอิง `$7`
5. **จำนวนนักเรียนไม่เท่ากันทุกห้อง** — ใช้ `detect_students()` เสมอ ห้าม hardcode
6. **มีเวลาเรียน1 + เวลาเรียน2** — คนละ sheet (sheet5, sheet6) สำหรับ 2 ภาค
7. **คุณลักษณะและอ่านคิดไม่แยกภาค** — กรอกครั้งเดียวต่อปี
8. **หน้าหลัก E13** — ภาคเรียนที่ต้องอัพเดทให้ตรงก่อนบันทึก
9. **หัวหน้ากลุ่มสาระ (E16)** — อาจ ≠ ครูประจำวิชา ต้องถามยืนยันทุกครั้ง
10. **Backup ก่อนเสมอ** — copy ไฟล์เป็น _backup.xlsx ก่อน step แรก
11. **Text ใน `<v>` ต้องมี `t="str"`** — ถ้าใส่ข้อความในแท็ก `<v>` โดยไม่มี `t="str"` Excel จะ repair/พัง (ตัวเลขไม่ต้องใส่)
12. **ลบ calcChain.xml ทุกครั้งก่อน repack** — ถ้าไม่ลบ Excel จะ recalculate ตอนเปิด → ถามบันทึกตอนปิดทั้งที่ไม่ได้แก้อะไร → ต้องลบออกจาก zip และ Content_Types.xml ด้วย

---

## 🔧 Helper Functions (ใช้ซ้ำทุกที่)

```python
import zipfile, shutil, os
from lxml import etree

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def col_to_num(col):
    num = 0
    for c in col: num = num * 26 + (ord(c) - ord('A') + 1)
    return num

def ensure_cell(row_elem, col_letter, row_num):
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
    c = ensure_cell(row_elem, col_letter, row_num)
    v = c.find(f'{{{NS}}}v')
    if v is None:
        v = etree.SubElement(c, f'{{{NS}}}v')
    v.text = str(value)

def _repack(file_path, extract_dir):
    # ลบ calcChain.xml ก่อน repack เสมอ
    # เหตุผล: chain เก่าทำให้ Excel recalculate ตอนเปิด → ถามบันทึกตอนปิดทั้งที่ไม่ได้แก้อะไร
    calc = os.path.join(extract_dir, 'xl', 'calcChain.xml')
    if os.path.exists(calc):
        os.remove(calc)
        ct_path = os.path.join(extract_dir, '[Content_Types].xml')
        with open(ct_path, 'r', encoding='utf-8') as f: ct = f.read()
        ct = ct.replace('<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>', '')
        with open(ct_path, 'w', encoding='utf-8') as f: f.write(ct)

    def zipdir(path, ziph):
        for rd, dd, ff in os.walk(path):
            for f in ff:
                fp = os.path.join(rd, f)
                ziph.write(fp, os.path.relpath(fp, path))
    with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipdir(extract_dir, zipf)
    shutil.rmtree(extract_dir)
```
