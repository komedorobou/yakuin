# 歴代役員経歴名簿.xlsx から index.html にデータを注入する
# 使い方: python build.py  → git commit & push で公開反映
import json, re, sys
from pathlib import Path

import openpyxl

XLSX = Path.home() / 'OneDrive/デスクトップ/歴代役員経歴名簿.xlsx'
HTML = Path(__file__).parent / 'index.html'

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb.worksheets[0]
r1 = [c.value for c in ws[1]]
r2 = [c.value for c in ws[2]]

# 1行目のグループ名（中央執行部・各支部・青年部・女性部）を前方埋め
grp = ''
groups = []
for i in range(len(r2)):
    if r1[i]:
        grp = str(r1[i]).strip()
    groups.append(grp)

# 役職列 = E列(4)〜女性部常任幹事まで。集計列(中央四役〜)以降は除外
roles = []
for i in range(4, len(r2)):
    name = str(r2[i]).replace('\n', '').replace('　', '').strip() if r2[i] else ''
    if name in ('中央四役', '支部四役', '青年部四役', '女性部四役', '通算総数'):
        break
    roles.append((i, groups[i], name))
LAST = roles[-1][0]
TOTAL, ACT, MEMO = LAST + 9, LAST + 10, LAST + 11  # 集計8列の後: 通算総数, 2026現役, 備考

people = []
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row[1]:
        continue
    rl = []
    for i, g, rn in roles:
        v = row[i]
        if v not in (None, ''):
            try:
                t = int(v)
            except (TypeError, ValueError):
                t = str(v)
            rl.append({'g': g, 'r': rn, 't': t})
    people.append({
        'k': str(row[0] or ''), 's': str(row[1]).strip(), 'm': str(row[2] or '').strip(),
        'b': str(row[3] or ''), 'roles': rl,
        'total': int(row[TOTAL] or 0),
        'act': bool(row[ACT]), 'memo': str(row[MEMO] or ''),
    })

data = json.dumps(people, ensure_ascii=False, separators=(',', ':'))
html = HTML.read_text(encoding='utf-8')
new = re.sub(
    r'(<script id="data" type="application/json">).*?(</script>)',
    lambda m: m.group(1) + data + m.group(2),
    html, count=1, flags=re.S,
)
HTML.write_text(new, encoding='utf-8')
print(f'{len(people)}名を index.html に注入した')
