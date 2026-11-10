r"""第11批构建器离线自测：只用临时假数据，不碰 0.xlsx / 文献 / 准 / 待。"""
import importlib.util
import os
import shutil
import sys
import tempfile

ROOT = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
spec = importlib.util.spec_from_file_location("e11", os.path.join(ROOT, "scripts", "entry11_20261108.py"))
e11 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(e11)

tmp = tempfile.mkdtemp(prefix="entry11_selftest_", dir=os.path.join(ROOT, "temp"))
try:
    fake_backup = os.path.join(tmp, "base.xlsx")
    fake_stage = os.path.join(tmp, "stage.xlsx")
    shutil.copy2(e11.BACKUP, fake_backup)

    FIELDS = e11.FIELDS
    data = []
    for no in range(e11.FIRST_NO, e11.LAST_NO + 1):
        nrows = 2 if no in (160, 165, 177) else 1
        rows = []
        for i in range(nrows):
            r = {k: "未提及" for k in FIELDS}
            r["sample"] = f"SELFT11-{no}-{i}"
            r["Fe"] = "余量"
            r["process"] = " SELFTEST "
            r["WE"] = "试样"
            r["RE"] = "SCE"
            r["CE"] = "Pt"
            r["temp"] = "25℃"
            r["Epit"] = "N/A（待读图）" if no % 2 else "0.25"
            r["stable"] = "30 min"
            rows.append(r)
        data.append({
            "no": no, "idx": f"{no:03d}",
            "file": f"fake_{no}.pdf", "verdict": "selftest",
            "citation": f"Selftest {no}. Journal. 2026.",
            "notes": "数据取自模拟表。", "flag": "", "rows": rows,
        })

    m, base_rows, anchor = e11.patched()
    m.BACKUP = fake_backup
    m.STAGE = fake_stage
    blocks = m.build(data)
    problems = m.verify(data, blocks, fake_stage)

    import openpyxl
    ws = openpyxl.load_workbook(fake_stage)["Sheet1"]
    total = sum(len(p["rows"]) for p in data)
    print(f"selftest11: {len(data)} 篇 / {total} 条 / 多行块 {sum(1 for p in data if len(p['rows'])>1)} / "
          f"底本 {base_rows} 行 / 模板块首 r{anchor}")
    print(f"stage dims {ws.max_row} x {ws.max_column} / merges "
          f"{len(ws.merged_cells.ranges)} (base "
          f"{len(openpyxl.load_workbook(fake_backup)['Sheet1'].merged_cells.ranges)})")
    print(f"首块 r{base_rows+1}: A={ws.cell(base_rows+1,1).value} B={ws.cell(base_rows+1,2).value!r} "
          f"Y={ws.cell(base_rows+1,25).value!r} Z={ws.cell(base_rows+1,26).value!r}")
    print(f"末块 r{ws.max_row}: A={ws.cell(ws.max_row,1).value} "
          f"编号连续={[ws.cell(r,1).value for r in range(base_rows+1, ws.max_row+1) if ws.cell(r,1).value] == list(range(160,178))}")
    print("problems =", len(problems))
    if problems:
        for x in problems[:10]:
            print(" -", x)
        sys.exit(2)
    print("selftest11: OK")
finally:
    shutil.rmtree(tmp, ignore_errors=True)
