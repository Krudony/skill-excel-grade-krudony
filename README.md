# Skill: Excel Grade Adjustment (ปพ.5)

> Claude Code Skill สำหรับแก้ไขไฟล์ Excel ปพ.5 อย่างปลอดภัย
> พัฒนาโดย Krudony | โรงเรียนบ้านแม่ทราย

---

## Files

| ไฟล์ | ระดับ | Slash Command |
|------|-------|--------------|
| `xlsx-safe-edit.md` | มัธยม | `/xlsx-safe-edit` |
| `xlsx-safe-edit-pratom.md` | ประถม | `/xlsx-safe-edit-pratom` |

---

## ติดตั้ง

```bash
# Linux / macOS
cp xlsx-safe-edit.md ~/.claude/skills/
cp xlsx-safe-edit-pratom.md ~/.claude/skills/
```

```bat
rem Windows
copy xlsx-safe-edit.md %USERPROFILE%\.claude\skills\
copy xlsx-safe-edit-pratom.md %USERPROFILE%\.claude\skills\
```

---

## ความสามารถ

- แก้ไข XML ตรง ไม่ใช้ openpyxl.save() (drawing/chart/image ยังอยู่ครบ)
- Auto-detect จำนวนนักเรียนแต่ละห้อง
- Backup ก่อนทำงานทุกครั้ง
- ลบ calcChain.xml ก่อน repack (ป้องกัน Excel ถามบันทึกตอนปิด)
- ทำทีละ sheet พร้อม verify

---

## ความต่าง มัธยม vs ประถม

| รายการ | มัธยม | ประถม |
|--------|-------|-------|
| คะแนน1 | ภาคเดียวต่อ sheet | 2 ภาคในชีตเดียว (ภาค1=I-BI / ภาค2=BJ-DV) |
| เวลาเรียน | ไม่มี | sheet5 (ภาค1) + sheet6 (ภาค2) |
| สมรรถนะ | รวมปี | แยกภาค (H,L,P,T,X=ภาค1 / I,M,Q,U,Y=ภาค2) |
| คุณลักษณะ/อ่านคิด | - | ไม่แยกภาค กรอกครั้งเดียว/ปี |
| Sheet ทั้งหมด | 6 sheets | 11 sheets |

---

## Sheet Mapping — ประถม

| Sheet Name | XML file | หมายเหตุ |
|------------|----------|---------|
| หน้าหลัก | sheet1.xml | ข้อมูลวิชา/ครู/วันอนุมัติ |
| ข้อมูลนักเรียน | sheet2.xml | source ชื่อ-นามสกุล |
| เวลาเรียน1 | sheet5.xml | เช็คชื่อ ภาค1 |
| เวลาเรียน2 | sheet6.xml | เช็คชื่อ ภาค2 |
| คะแนน1 | sheet8.xml | 2 ภาคในชีตเดียว |
| คุณลักษณะ | sheet9.xml | H-O |
| อ่านคิด | sheet10.xml | H-L |
| สมรรถนะ | sheet11.xml | แยกภาค |

## Sheet Mapping — มัธยม

| Sheet Name | XML file |
|------------|----------|
| หน้าหลัก | sheet1.xml |
| คะแนน1 | sheet6.xml |
| คุณลักษณะ | sheet7.xml |
| อ่านคิด | sheet8.xml |
| สมรรถนะ | sheet9.xml |

---

## บทเรียนสำคัญ

1. ห้ามใช้ openpyxl.save() — ทำลาย drawing/chart/image
2. ปิด Excel ก่อนทุกครั้ง — ไม่งั้นเกิด PermissionError
3. Text ใน `<v>` ต้องมี `t="str"` — ไม่งั้น Excel repair
4. ลบ calcChain.xml ทุกครั้งก่อน repack — ไม่งั้น Excel ถามบันทึกตอนปิด
5. ห้ามแก้ Row 7 — header คะแนนเต็ม ทุก formula อ้างอิง `$7`
6. ห้าม hardcode จำนวนนักเรียน — ใช้ auto-detect จาก formula ใน col C

---

## Safe Method (หลักการ)

```
xlsx (zip) -> อ่าน XML in-memory -> แก้ -> repack -> xlsx
```

```python
import zipfile, xml.etree.ElementTree as ET

with zipfile.ZipFile(fname) as z:
    files = {n: z.read(n) for n in z.namelist()}

root = ET.fromstring(files['xl/worksheets/sheet6.xml'])
# ... แก้ XML ...
files['xl/worksheets/sheet6.xml'] = ET.tostring(root, 'utf-8', xml_declaration=True)

# ลบ calcChain ก่อน repack เสมอ
files.pop('xl/calcChain.xml', None)

with zipfile.ZipFile(fname, 'w', zipfile.ZIP_DEFLATED) as z:
    for name, data in files.items():
        z.writestr(name, data)
```

---

*พัฒนาสำหรับโรงเรียนบ้านแม่ทราย (คุรุราษฎร์เจริญวิทย์) — 2026*
