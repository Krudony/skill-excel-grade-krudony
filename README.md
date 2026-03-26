# Skill: Excel Grade Adjustment (ปพ.5)

> Claude Code skill for safe `ปพ.5` Excel editing without breaking drawings, charts, or workbook relationships.
> Maintained by Krudony.

---

## Files

| File | Level | Slash Command |
|------|-------|---------------|
| `xlsx-safe-edit.md` | มัธยม / shared safe-edit reference | `/xlsx-safe-edit` |
| `xlsx-safe-edit-pratom.md` | ประถม | `/xlsx-safe-edit-pratom` |

---

## Install

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

## Capabilities

- Edit workbook XML directly instead of using `openpyxl.save()`.
- Preserve drawings, charts, images, and relationships.
- Auto-detect student rows from formula-driven sheets.
- Back up the workbook before editing.
- Remove `calcChain.xml` before repacking so Excel recalculates cleanly.
- Support primary layouts where `คะแนน1` is `sheet8.xml` and contains both semesters.

---

## Secondary vs Primary

| Item | มัธยม | ประถม |
|------|-------|--------|
| `คะแนน1` | One term per sheet | Two terms in one sheet (`I:BI` and `BJ:DV`) |
| Attendance | Not used in this repo's main workflow | `sheet5` + `sheet6` |
| Competencies | Year summary | Split by term |
| Total sheets | Smaller layout | Larger layout |

---

## Field Update: Primary Two-Semester Case

- Many primary `ปพ.5` templates keep both Semester 1 and Semester 2 in `คะแนน1 = sheet8.xml`.
- Before filling Semester 2, normalize row 7 if the template still has placeholder totals:
  `BJ7:BQ7=10`, `DH7=80`, `DI7=20`, `DJ7=20`, `DK7=0`, `DL7=0`, `DM7=20`, `DN7=100`, `DO7=100`, `DP7=100`, `DQ7=200`, `DR7=100`.
- When the user gives target grades but not exact raw scores, fill input score cells first and avoid grade override cells unless explicitly requested.
- If asked to distribute scores across both terms, keep the two term totals reasonably close and avoid unnatural all-10 or perfectly repeated patterns.

### Practical scoring heuristics

1. Prefer real scores before `DU` override.
2. For grade targets around `3` or `4`, keep indicator scores mostly in the `6-9` range.
3. Do not give every student full exam marks.
4. Aim `DR/DS` near the lower edge of the requested band unless the user asks for extra margin.
5. Inspect nearby students first so edited rows match the class scoring pattern.

---

## Sheet Mapping

### Primary

| Sheet Name | XML file | Note |
|------------|----------|------|
| หน้าหลัก | `sheet1.xml` | Subject / teacher / approval data |
| ข้อมูลนักเรียน | `sheet2.xml` | Student source data |
| เวลาเรียน1 | `sheet5.xml` | Attendance term 1 |
| เวลาเรียน2 | `sheet6.xml` | Attendance term 2 |
| คะแนน1 | `sheet8.xml` | Two semesters in one sheet |
| คุณลักษณะ | `sheet9.xml` | `H:O` |
| อ่านคิด | `sheet10.xml` | `H:L` |
| สมรรถนะ | `sheet11.xml` | Split by term |

### Secondary

| Sheet Name | XML file |
|------------|----------|
| หน้าหลัก | `sheet1.xml` |
| คะแนน1 | `sheet6.xml` |
| คุณลักษณะ | `sheet7.xml` |
| อ่านคิด | `sheet8.xml` |
| สมรรถนะ | `sheet9.xml` |

---

## Core Lessons

1. Never use `openpyxl.save()` on these workbooks.
2. Close Excel before editing to avoid file locks.
3. Use `t="str"` for plain text cell values when needed.
4. Remove `calcChain.xml` before repacking.
5. Respect row 7 because formulas depend on it as the scoring base.
6. Do not hardcode student counts; detect active rows from formulas or IDs.

---

## Safe Method

```text
xlsx (zip) -> read XML -> edit -> remove calcChain -> repack -> xlsx
```

```python
import zipfile
import xml.etree.ElementTree as ET

with zipfile.ZipFile(fname) as z:
    files = {name: z.read(name) for name in z.namelist()}

root = ET.fromstring(files['xl/worksheets/sheet6.xml'])
# edit XML here
files['xl/worksheets/sheet6.xml'] = ET.tostring(root, 'utf-8', xml_declaration=True)

files.pop('xl/calcChain.xml', None)

with zipfile.ZipFile(fname, 'w', zipfile.ZIP_DEFLATED) as z:
    for name, data in files.items():
        z.writestr(name, data)
```

---

*Built for school grading workflows that must preserve workbook integrity.*
