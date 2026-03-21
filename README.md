# 📊 Skill: Excel Grade Adjustment (ปพ.5)

> Claude Code Skill สำหรับปรับเกรดนักเรียนในไฟล์ Excel ปพ.5 อย่างปลอดภัย
> พัฒนาโดย Krudony | โรงเรียนบ้านแม่ทราย

---

## ✨ ความสามารถ

- ✅ แก้ไขคะแนนนักเรียนรายคน หรือทั้งหมด
- ✅ ปรับ Grade 1, 1.5, 2, 2.5, 3, 3.5, 4
- ✅ ไม่ทำให้ไฟล์พัง (drawing/chart/image ยังอยู่ครบ)
- ✅ คะแนนไม่ซ้ำกัน ต่างกันแต่ละคน
- ✅ เข้าใจ formula chain ของ ปพ.5 ระดับมัธยม

---

## 🔧 ติดตั้ง Skill

```bash
# copy skill ไปไว้ใน Claude skills folder
cp xlsx-safe-edit.md ~/.claude/skills/xlsx-safe-edit.md
```

**Windows:**
```
copy xlsx-safe-edit.md C:\Users\{ชื่อUser}\.claude\skills\xlsx-safe-edit.md
```

---

## 📚 วิธีใช้งาน

พิมพ์ใน Claude Code:

```
ปรับคะแนนทุกคนให้ได้ Grade 3
```

```
นักเรียนคนที่ 1-5 ให้ Grade 3 / คนที่ 6-12 ให้ Grade 3.5
```

```
นักเรียนคนที่ 3 ปรับเป็น Grade 4
```

---

## 📐 โครงสร้างไฟล์ ปพ.5 ม.3

### Sheet "คะแนน1" (sheet6.xml)

| Excel Column | ชื่อ | ประเภท | หมายเหตุ |
|-------------|------|--------|---------|
| I–O | ตัวชี้วัด 1–7 | **INPUT** | max = 10 ต่อข้อ |
| BG | คะแนนกลางภาค | **INPUT** | max = 10 |
| BH | รวมระหว่างเรียน | FORMULA | `SUMIF(I:BG,"<>-1")` |
| BI | คะแนนปลายภาค | **INPUT** | max = 40 |
| BJ | รวมปลายภาค | FORMULA | `ROUND(BI×20/40, 0)` |
| BN | รวมคะแนน | FORMULA | `SUM(BH, BM)` |
| BO | สรุปคะแนน | FORMULA | `ROUND(BN/100×100, 0)` |
| BP | ผลการประเมิน | FORMULA | `VLOOKUP(BO, หน้าหลัก!, 4)` |

### เกณฑ์การให้เกรด

| ช่วงคะแนน | เกรด |
|----------|------|
| 80–100 | 4 |
| 74.5–79.49 | 3.5 |
| **69.5–74.49** | **3** |
| 64.5–69.49 | 2.5 |
| 54.5–64.49 | 2 |
| 49.5–54.49 | 1.5 |
| 0–49.49 | 1 |

### สูตรคำนวณ

```
between   = sum(ตัวชี้วัด 1-7) + คะแนนกลางภาค   (max 80)
exam      = round(คะแนนปลายภาค × 20/40, 0)      (max 20)
total     = between + exam                        (max 100)
```

**ตัวอย่าง Grade 3 (total 70–74):**
```
ตัวชี้วัด 1-7: เฉลี่ย 6-8 ต่อข้อ
คะแนนกลางภาค: 7-9
คะแนนปลายภาค: 24-32
```

---

## ⚠️ ข้อควรระวัง (เรียนรู้จากการใช้งานจริง)

1. **ห้ามแก้ Row 7** — เป็น header คะแนนเต็ม ทุก formula อ้างอิง `$7`
2. **ปิด Excel ก่อน** — ไม่งั้นบันทึกไม่ได้ (PermissionError)
3. **อย่าใช้ openpyxl save()** — ทำให้ drawing/chart พัง
4. **ใช้ Extract-XML** — แตก zip → แก้ XML → บีบอัดคืน
5. **ใช้ row attribute `r`** — อย่า index เพราะ blank rows ไม่ถูก save ใน XML

---

## 🔄 วิธีการทำงาน (Safe Method)

```
xlsx (zip) → แตก → แก้ XML → บีบอัด → xlsx
```

```python
# ✅ SAFE
import zipfile
from lxml import etree

# 1. Extract
with zipfile.ZipFile('file.xlsx', 'r') as z:
    z.extractall('tmp')

# 2. Edit XML (ไม่ใช้ openpyxl!)
tree = etree.parse('tmp/xl/worksheets/sheet6.xml')
# ... แก้ v_elem.text ...

# 3. Repack
with zipfile.ZipFile('file.xlsx', 'w', zipfile.ZIP_DEFLATED) as z:
    # zipdir...
```

---

## 📁 Files

| ไฟล์ | คำอธิบาย |
|------|---------|
| `xlsx-safe-edit.md` | Claude Code skill พร้อม code template |
| `README.md` | คู่มือนี้ |

---

*พัฒนาสำหรับโรงเรียนบ้านแม่ทราย (คุรุราษฎร์เจริญวิทย์) — 2026*
