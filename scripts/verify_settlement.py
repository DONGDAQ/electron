"""核对脚本：从报价单/结算单/飞书表读取真实数据，与 settlement_tracker 对比"""
import openpyxl, json, os, sys, re
from datetime import datetime, timedelta
sys.stdout.reconfigure(encoding='utf-8')

BASE = r'\\192.168.110.111\【管理者专用】\04合同\000-DQ业务对接用\报价结算系统'

DIR_TO_KEY = {
    '马娘': 'maniang', 'HBR': 'hbr', 'bang2': 'bang2',
    '幻塔': 'huanta', '异环发行': 'yihuan_faxing', '异环游戏内': 'yihuan_nei',
    '战双发行': 'zhan_shuang_faxing', '战双版更': 'zhan_shuang',
    '恋与深空': 'liandishenkong', '闪暖': 'shining_nikki', 'ニキ新作': 'niki_xinzuo',
}

def read_quote_totals(path):
    """尝试读取报价单的合计行数据"""
    result = {'word_count': 0, 'total_price': 0}
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                val = str(cell.value).strip() if cell.value else ''
                if val in ('合计', '总计', '合计/Total'):
                    r = cell.row
                    for c in range(1, ws.max_column + 1):
                        v = ws.cell(r, c).value
                        if isinstance(v, (int, float)):
                            if v > 100 and v != int(v):
                                result['total_price'] = max(result['total_price'], v)
                            elif v > 100:
                                result['word_count'] = max(result['word_count'], int(v))
                    break
        if result['total_price'] == 0 and result['word_count'] == 0:
            last_row = ws.max_row
            for c in range(1, min(ws.max_column + 1, 15)):
                v = ws.cell(last_row, c).value
                if isinstance(v, (int, float)):
                    if v > 100 and v != int(v):
                        result['total_price'] = max(result['total_price'], v)
                    elif v > 100:
                        result['word_count'] = max(result['word_count'], int(v))
        wb.close()
    except Exception as e:
        result['error'] = str(e)
    return result

def read_settlement_total(path):
    """读取结算单的总金额"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        max_val = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                v = cell.value
                if isinstance(v, (int, float)) and v > max_val and v > 100:
                    max_val = v
        wb.close()
        return max_val
    except:
        return 0

def infer_key_from_filename(f):
    if 'X3' in f or '恋与深空' in f: return 'liandishenkong'
    elif '闪暖' in f: return 'shining_nikki'
    elif 'X6' in f or 'ニキ' in f: return 'niki_xinzuo'
    elif '幻塔' in f: return 'huanta'
    elif '异环' in f and '发行' in f: return 'yihuan_faxing'
    elif '异环' in f and '游戏内' in f: return 'yihuan_nei'
    elif '战双' in f and '发行' in f: return 'zhan_shuang_faxing'
    elif '战双' in f and ('版更' in f or 'LQA' in f): return 'zhan_shuang'
    elif '代号PD' in f or '马娘' in f: return 'maniang'
    elif 'Trickcal' in f or 'TK' in f: return 'tk'
    elif '4399' in f or '冒险' in f or '指尖' in f or '主宰' in f or '波奇' in f: return '4399'
    elif '祖龙' in f or '以闪亮之名' in f or '龙族' in f: return 'zulong'
    return None

# ===== 1. 扫描报价 =====
truth = {}
quote_dir = os.path.join(BASE, '报价')
for company in os.listdir(quote_dir):
    comp_path = os.path.join(quote_dir, company)
    if not os.path.isdir(comp_path):
        continue
    for proj in os.listdir(comp_path):
        proj_path = os.path.join(comp_path, proj)
        if not os.path.isdir(proj_path):
            continue
        key = DIR_TO_KEY.get(proj)
        if not key:
            continue
        if key not in truth:
            truth[key] = {'settled': [], 'unsettled': []}
        for root, dirs, files in os.walk(proj_path):
            is_settled = '已结算' in root
            for f in files:
                if not f.endswith('.xlsx'):
                    continue
                if any(k in f for k in ['结算', 'Invoice', '账单', '汇总', 'zip']):
                    continue
                fpath = os.path.join(root, f)
                totals = read_quote_totals(fpath)
                entry = {
                    'filename': f,
                    'path': fpath,
                    'word_count': totals.get('word_count', 0),
                    'total_price': totals.get('total_price', 0),
                }
                if 'error' in totals:
                    entry['error'] = totals['error']
                if is_settled:
                    truth[key]['settled'].append(entry)
                else:
                    truth[key]['unsettled'].append(entry)

# ===== 2. 扫描结算单 =====
settle_202605 = {}
settle_dir = os.path.join(BASE, '结算')
if os.path.isdir(settle_dir):
    month_path = os.path.join(settle_dir, '2026年5月')
    if os.path.isdir(month_path):
        for root, dirs, files in os.walk(month_path):
            for f in files:
                if not f.endswith('.xlsx'):
                    continue
                if '结算' not in f and '账单' not in f:
                    continue
                if f.startswith('~$'):
                    continue
                fpath = os.path.join(root, f)
                parts = root.replace(month_path, '').split(os.sep)
                proj_name = parts[-1] if parts else ''
                key = DIR_TO_KEY.get(proj_name) or infer_key_from_filename(f)
                if not key:
                    continue
                amount = read_settlement_total(fpath)
                if key not in settle_202605:
                    settle_202605[key] = []
                settle_202605[key].append({'filename': f, 'amount': amount})

# ===== 3. 输出文件系统真数据 =====
print('=' * 60)
print('文件系统真实数据（报价单 + 结算单）')
print('=' * 60)

all_keys = sorted(set(list(truth.keys()) + list(settle_202605.keys())))
for key in all_keys:
    d = truth.get(key, {'settled': [], 'unsettled': []})
    s = settle_202605.get(key, [])
    print(f'\n【{key}】')
    if d['settled']:
        print(f'  已结算报价单: {len(d["settled"])}')
        for f in d['settled']:
            print(f'    ✅ {f["filename"]}  words={f["word_count"]}  amount={f["total_price"]}')
    if d['unsettled']:
        print(f'  未结算报价单: {len(d["unsettled"])}')
        for f in d['unsettled']:
            print(f'    ⏳ {f["filename"]}  words={f.get("word_count",0)}  amount={f.get("total_price",0)}')
    if s:
        total_settle = sum(x['amount'] for x in s)
        print(f'  5月结算单: {len(s)} 份, 总额≈{total_settle:.2f}')
        for x in s:
            print(f'    💰 {x["filename"]}  amount≈{x["amount"]:.2f}')
    if not d['settled'] and not d['unsettled'] and not s:
        print(f'  (无数据)')

# ===== 4. 对比 settlement_tracker =====
print('\n' + '=' * 60)
print('差异对比：文件系统 vs settlement_tracker')
print('=' * 60)

tracker_dir = os.path.join('outputs', 'settlement_tracker')
for key in all_keys:
    rec_path = os.path.join(tracker_dir, key, 'records.json')
    if not os.path.isfile(rec_path):
        print(f'\n【{key}】⚠️ tracker 无记录')
        continue
    with open(rec_path, 'r', encoding='utf-8') as f:
        records = json.load(f)

    tk_settled = [r for r in records if r.get('settled')]
    tk_unsettled = [r for r in records if not r.get('settled')]

    fs_data = truth.get(key, {'settled': [], 'unsettled': []})
    fs_settled_count = len(fs_data['settled'])
    fs_unsettled_count = len(fs_data['unsettled'])
    s_data = settle_202605.get(key, [])

    issues = []

    if len(tk_settled) != fs_settled_count and fs_settled_count > 0:
        issues.append(f'已结算数: 文件={fs_settled_count}, tracker={len(tk_settled)}')

    if len(tk_unsettled) != fs_unsettled_count and fs_unsettled_count > 0:
        issues.append(f'未结算数: 文件={fs_unsettled_count}, tracker={len(tk_unsettled)}')

    # 5月结算单存在但 tracker 仍有5月未结算记录
    if s_data and tk_unsettled:
        may_unsettled = [r for r in tk_unsettled if (r.get('delivery_date') or '').startswith('2026-05')]
        if may_unsettled:
            issues.append(f'5月有结算单但tracker仍有{len(may_unsettled)}条5月交付的未结算记录')

    # 金额差异
    fs_total_unsettled = sum(f['total_price'] for f in fs_data['unsettled'])
    tk_total_unsettled = sum(r.get('total_price', 0) for r in tk_unsettled)
    if abs(fs_total_unsettled - tk_total_unsettled) > 1 and fs_total_unsettled > 0:
        issues.append(f'未结算金额差异: 文件≈{fs_total_unsettled:.2f}, tracker={tk_total_unsettled:.2f}')

    if issues:
        print(f'\n【{key}】❌ 差异:')
        for iss in issues:
            print(f'    ⚠️ {iss}')
    else:
        print(f'\n【{key}】✅ 一致')
