# -*- coding: utf-8 -*-
"""
下单费用核算工具
选择 账面毛利 / 付款信息 / 累计下单费用 三个文件（均只读第一个 sheet），点击"核算"生成《下单费用.xlsx》。
核算规则见《服务销售事业部 下单费用核算流程需求说明书》。
"""

import collections
import datetime
import json
import os
import queue
import re
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# (结果表表头, 行数据键)，键为 None 的列固定留空
COLUMNS = [
    ('是否已算过费用', 'counted'),
    ('索引1', 'index'),
    ('下单合同号（外购不算下单成本）', 'contract'),
    ('项目名称', 'project'),
    ('客户名称', 'customer'),
    ('下单时间', None),
    ('性质', 'nature'),
    ('平台', 'platform'),
    ('区域', 'region'),
    ('事业部', 'dept'),
    ('销售员', 'salesman'),
    ('销售员编码', 'salesman_code'),
    ('开单成本（未税）', 'cost'),
    ('税率', 'tax'),
    ('开单金额（含税）', 'amount'),
    ('付款日期', 'pay_date'),
    ('付款金额', 'pay_amount'),
    ('开单日期', 'invoice_date'),
    ('付款金额（不含税）', 'pay_amount_ex'),
    ('贷款利率', 'rate'),
    ('总财务费用', 'fee'),
    ('扣款时间', 'deduct_period'),
    ('扣款月份', 'deduct_month'),
    ('备注', 'remark'),
    ('运输方式', None),
    ('核查', 'check'),
    ('使用激励金额', 'incentive'),
    ('发货款天数差', 'days'),
    ('产品类别', None),
    ('产品线', None),
    ('服务产品线', 'sp_line'),
    ('类型', 'sp_type'),
    ('特殊备注', None),
    ('核查-使用激励', 'check_inc'),
]
FILE_LABELS = ('账面毛利', '付款信息', '累计下单费用')

# 要从各过程表取用的列名，按表头名定位而不依赖列序，表头改序不影响核算
GM_COLS = ('下单合同号', '项目名称', '客户名称', '一级部门', '二级部门', '三级部门', '实际销售员名称',
           '行项目销售员', '成本总价', '税率', '出具发票日', '备注', '服务产品线', '服务产品线类别')
PAY_COLS = ('合同号', '付款金额', '付款时间', '华为激励', '超聚变激励')
ACC_COLS = ('索引1',)

# 外购、价外费用不算下单成本；暂未下单不是订单号，一并剔除
EXCLUDE_CONTRACTS = ('外购', '价外费用', '暂未下单')
# 不计费用：培训区域且项目名称含关键字（与累计表已算过、开单成本≤0 任一成立即标）
NO_FEE_REGION = '培训'
NO_FEE_KEYWORDS = ('考试券', '囤券', '半导体客户培训项目')
REMARK_NO_FEE = '不计费用'
REMARK_CHECK = '人工检查'
# HYX 合同不拆行，改为金额比对
HYX_PREFIX = 'HYX'
TOLERANCE = 1.0        # HYX 开单金额与付款金额的比较容差（元）
DAYS_BASE = 8         # 发货款免息天数
CHECK_LIMIT = 5       # |核查-使用激励| 超过该金额需人工检查（元）

# 写结果时的单元格格式：日期、金额、百分比，其余保持常规
DATE_KEYS = {'pay_date', 'invoice_date'}
MONEY_KEYS = {'cost', 'amount', 'pay_amount', 'pay_amount_ex', 'fee', 'check', 'incentive', 'check_inc'}
PERCENT_KEYS = {'rate'}
WIDTHS = {'contract': 26, 'project': 36, 'customer': 32, 'remark': 24}

DEFAULT_RATE = 0.0828       # 超贷利率默认值，配置文件缺失或非法时使用
OUTPUT_NAME = '下单费用.xlsx'


class AppError(Exception):
    """需要直接展示给用户的错误"""


class Stats:
    def __init__(self):
        self.count = collections.Counter()


def program_dir():
    # 配置文件跟随程序目录（打包后为 exe 所在目录）
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def config_path():
    return os.path.join(program_dir(), '下单费用配置.json')


def load_rate_config():
    try:
        with open(config_path(), 'r', encoding='utf-8') as f:
            rate = float(json.load(f).get('超贷利率', DEFAULT_RATE))
        return rate if 0 <= rate < 1 else DEFAULT_RATE
    except (OSError, ValueError, TypeError):
        return DEFAULT_RATE


def save_rate_config(rate):
    try:
        with open(config_path(), 'w', encoding='utf-8') as f:
            json.dump({'超贷利率': rate}, f, ensure_ascii=False, indent=2)
    except OSError as e:
        raise AppError(f'利率配置文件写入失败：{config_path()}\n{e}')


def parse_rate(text):
    s = text.strip()
    is_percent = s.endswith('%')
    try:
        rate = float(s.rstrip('%').strip()) / (100 if is_percent else 1)
    except ValueError:
        raise AppError(f'超贷利率格式错误："{s}"，应为数字，例如 0.0828 或 8.28%')
    if not 0 <= rate < 1:
        raise AppError(f'超贷利率超出范围（0~100%）："{s}"')
    return rate


def cell(row, idx):
    # 只读模式下行尾空单元格会被省略，行可能比表头短
    return row[idx] if idx < len(row) else None


def text(value):
    return str(value).strip() if value is not None else ''


def to_number(value):
    # 付款金额等可能是带千分位的文本，统一转数字，转不了的一律当空
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(',', ''))
    except ValueError:
        return None


def to_date(value):
    # 付款时间/开单日期存在文本、日期、序列号多种形态，统一转 date 才能比较和拆行
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    s = text(value)
    if not s:
        return None
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y%m%d', '%Y年%m月%d日', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    serial = to_number(s)
    # Excel 日期序列号（1900 日期系统）
    if serial is not None and 20000 < serial < 80000:
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=serial)).date()
    return None


def excel_str(value):
    # 按 Excel"常规"格式转文本，保证索引1 与累计表拼接结果一致（0.06→"0.06"，0→"0"）
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def split_contracts(contract):
    # 一个格子里可能填多个合同号，分隔符有半角逗号和全角逗号两种
    return [c for c in (p.strip() for p in re.split(r'[,，]+', contract)) if c]


def first_sheet_rows(path, label):
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
        wb.close()
        return rows
    except Exception as e:
        raise AppError(f'{label}读取失败：{path}\n{e}')


def find_header(rows, required, label):
    # 账面毛利第 1 行是分组标题，表头可能下移，因此在首 10 行内找命中列名最多的一行
    best_idx, best_hit = -1, {}
    for i, row in enumerate(rows[:10]):
        cells = [text(c) for c in row]
        hit = {name: cells.index(name) for name in required if name in cells}
        if len(hit) > len(best_hit):
            best_idx, best_hit = i, hit
    missing = [name for name in required if name not in best_hit]
    if best_idx < 0 or missing:
        raise AppError(f'{label}缺少必需的列：{"、".join(missing)}')
    return best_idx, best_hit


def build_pay_index(rows, col):
    # 同一合同号可能有多笔付款（同日多笔或异日多笔），按合同号聚成列表供后续合并/拆行
    index = {}
    for row in rows:
        contract = text(cell(row, col['合同号']))
        if not contract:
            continue
        # (付款日期, 付款金额, 华为激励+超聚变激励)
        index.setdefault(contract, []).append((
            to_date(cell(row, col['付款时间'])),
            to_number(cell(row, col['付款金额'])) or 0.0,
            (to_number(cell(row, col['华为激励'])) or 0.0) + (to_number(cell(row, col['超聚变激励'])) or 0.0),
        ))
    return index


def build_acc_index(rows, col):
    # 累计表已算过的索引集合，结果表据此标"不计费用"
    return {text(cell(r, col['索引1'])) for r in rows if cell(r, col['索引1']) is not None}


def with_tax(amount, tax):
    # 税率异常时 1+税率 可能为 0，避免除零
    return amount / (1 + tax) if 1 + tax else None


def make_base(row, col, contract):
    # 结果表基础行：开单金额由 成本*(1+税率) 推算，下单时间/产品类别/产品线等列按文档留空
    tax = to_number(cell(row, col['税率']))
    cost = to_number(cell(row, col['成本总价'])) or 0.0
    code = cell(row, col['行项目销售员'])
    if isinstance(code, float) and code.is_integer():
        code = int(code)   # 销售员编码按文本存，避免出现 2003185.0
    return {
        'counted': None, 'index': '', 'contract': contract, 'contracts': split_contracts(contract),
        'project': text(cell(row, col['项目名称'])), 'customer': text(cell(row, col['客户名称'])),
        'nature': '服务',
        'platform': text(cell(row, col['三级部门'])), 'region': text(cell(row, col['二级部门'])),
        'dept': text(cell(row, col['一级部门'])),
        'salesman': text(cell(row, col['实际销售员名称'])), 'salesman_code': text(code),
        'cost': cost, 'tax': 0.0 if tax is None else tax,
        'amount': cost * (1 + (0.0 if tax is None else tax)),
        'invoice_date': to_date(cell(row, col['出具发票日'])),
        'sp_line': text(cell(row, col['服务产品线'])), 'sp_type': text(cell(row, col['服务产品线类别'])),
        'pay_date': None, 'pay_amount': None, 'incentive': None, 'remarks': [],
    }


def copy_row(row):
    # remarks 是列表，浅拷贝会让拆出的行共用同一个备注列表
    return dict(row, remarks=list(row['remarks']))


def apply_pays(row, pays):
    # 多笔付款合并规则：金额、激励求和，付款日期取最早
    row['pay_amount'] = sum(p[1] for p in pays)
    row['incentive'] = sum(p[2] for p in pays)
    dates = [p[0] for p in pays if p[0] is not None]
    row['pay_date'] = min(dates) if dates else None


def match_hyx(base, pays, stats):
    # HYX：开单金额依次与 单条金额→同日合计→全部合计 比较，均不符则填开单金额并人工检查；不拆行
    if not pays:
        stats.count['no_pay'] += 1
        return [base]
    row = copy_row(base)
    by_day = collections.defaultdict(list)
    for p in pays:
        by_day[p[0]].append(p)
    candidates = [(p[1], [p]) for p in pays]
    candidates += [(sum(p[1] for p in g), g) for g in by_day.values() if len(g) > 1]
    candidates.append((sum(p[1] for p in pays), pays))
    for amount, group in candidates:
        if abs(amount - base['amount']) <= TOLERANCE:
            apply_pays(row, group)
            return [row]
    apply_pays(row, pays)
    row['pay_amount'] = base['amount']
    row['remarks'].append(REMARK_CHECK)
    stats.count['hyx_mismatch'] += 1
    return [row]


def split_by_pay_date(base, pay_index, stats):
    # 先按"合同号 → 付款日期"分组，同一天的多笔付款只有一组时才合并
    per = {c: collections.defaultdict(list) for c in base['contracts'] if c in pay_index}
    for c in per:
        for p in pay_index[c]:
            per[c][p[0]].append(p)
    if not per:
        stats.count['no_pay'] += 1
        return [base]   # 付款表未匹配，付款相关列留空
    dates = sorted({d for g in per.values() for d in g if d is not None})
    if len(dates) <= 1:
        # 付款日期一致（或有多个合同号但同日付款），合并成一行
        apply_pays(base, [p for g in per.values() for ps in g.values() for p in ps])
        return [base]

    # 付款日期不同：最早日期留在原行，其余日期各拆一行；无法解析付款日期的记录归原行
    first = dates[0]
    orig = copy_row(base)
    orig['contracts'] = [c for c in base['contracts']
                         if c not in per or not [d for d in per[c] if d is not None]
                         or min(d for d in per[c] if d is not None) == first]
    orig['contract'] = ','.join(orig['contracts'])   # 拆出的合同号从原行去除并重排
    orig_pays = [p for c in orig['contracts'] if c in per
                 for d, ps in per[c].items() if d is None or d == first for p in ps]
    apply_pays(orig, orig_pays)
    orig['pay_date'] = first

    new_rows = []
    for d in dates[1:]:
        pays = [p for g in per.values() for p in g.get(d, [])]
        new = copy_row(base)
        new['contracts'] = [c for c in base['contracts'] if c in per and d in per[c]]
        new['contract'] = ','.join(new['contracts'])
        new['amount'] = sum(p[1] for p in pays)                     # 拆分行开单金额取该日付款金额
        new['cost'] = with_tax(new['amount'], base['tax'])          # 成本按含税金额反算
        apply_pays(new, pays)
        new['pay_date'] = d
        new_rows.append(new)
    # 原行开单金额扣减拆分行的付款金额，差额由"核查"体现
    orig['amount'] = base['amount'] - sum(r['amount'] for r in new_rows)
    orig['cost'] = with_tax(orig['amount'], base['tax'])
    if orig['amount'] < 0:
        orig['remarks'].append(REMARK_CHECK)
    stats.count['split'] += len(new_rows)
    return [orig] + new_rows


def fill_derived(row, acc_index, rate, stats):
    # 索引1 = 下单合同号 & 税率，与累计表同格式拼接才能比对
    tax, invoice, pay_amount = row['tax'], row['invoice_date'], row['pay_amount']
    row['index'] = row['contract'] + excel_str(tax)
    counted = row['index'] in acc_index
    if counted:
        row['counted'] = row['index']
    if pay_amount is not None:
        row['pay_amount_ex'] = with_tax(pay_amount, tax)
        row['check'] = row['amount'] - pay_amount
        row['check_inc'] = row['check'] - (row['incentive'] or 0.0)
        if invoice and row['pay_date']:
            days = (invoice - row['pay_date']).days
            row['days'] = days
            # 天数差超过免息天数才计财务费用，否则利率、费用记 0
            if days > DAYS_BASE:
                row['rate'] = rate
                row['fee'] = (days - DAYS_BASE) * pay_amount * rate / 365
            else:
                row['rate'], row['fee'] = 0, 0
    if invoice:
        # 扣款期间跟随开单日期所在季度/月份，而非核算当前日期
        row['deduct_period'] = f'{invoice.year}Q{(invoice.month - 1) // 3 + 1}'
        row['deduct_month'] = f'{invoice.month}月'
    # 不计费用：累计表已算过 / 开单成本≤0 / 培训区域的囤券类项目，满足其一
    no_fee = (counted or (row['cost'] or 0) <= 0
              or (row['region'] == NO_FEE_REGION and any(k in row['project'] for k in NO_FEE_KEYWORDS)))
    remarks = row['remarks']
    check_inc = row.get('check_inc')
    if no_fee:
        remarks = [REMARK_NO_FEE] + remarks
        stats.count['no_fee'] += 1
        if counted:
            stats.count['counted'] += 1
    elif isinstance(check_inc, float) and abs(check_inc) > CHECK_LIMIT:
        remarks.append(REMARK_CHECK)   # 已标不计费用的行不再追加人工检查
    if REMARK_CHECK in remarks:
        stats.count['manual'] += 1
    row['remark'] = '；'.join(dict.fromkeys(remarks))   # 多个备注去重后用分号连接
    return row


def build_output_rows(gm_rows, col, pay_index, acc_index, rate, stats):
    outputs = []
    for row in gm_rows:
        contract = text(cell(row, col['下单合同号']))
        if not contract:
            if any(v is not None for v in row):
                stats.count['no_contract'] += 1
            continue
        # 剔除口径：外购/价外费用/暂未下单，以及备注为预提的行
        if contract in EXCLUDE_CONTRACTS:
            stats.count['excluded'] += 1
            continue
        if '预提' in text(cell(row, col['备注'])):
            stats.count['excluded_yuti'] += 1
            continue
        base = make_base(row, col, contract)
        # HYX 合同走金额比对、不拆行；其余按付款日期合并/拆行
        if len(base['contracts']) == 1 and base['contracts'][0].startswith(HYX_PREFIX):
            matched = match_hyx(base, pay_index.get(base['contracts'][0]), stats)
        else:
            matched = split_by_pay_date(base, pay_index, stats)
        outputs.extend(fill_derived(r, acc_index, rate, stats) for r in matched)
    return outputs


def write_result(rows, path):
    # 表头加粗并冻结首行，方便人工核对时横向对比
    wb = Workbook()
    ws = wb.active
    ws.title = '下单费用'
    header_font = Font(bold=True, name='微软雅黑', size=10)
    header_fill = PatternFill(start_color='D9E1F2', fill_type='solid')
    body_font = Font(name='微软雅黑', size=10)
    for c, (title, _) in enumerate(COLUMNS, start=1):
        h = ws.cell(row=1, column=c, value=title)
        h.font, h.fill = header_font, header_fill
        h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = WIDTHS.get(COLUMNS[c - 1][1], 14)
    for r, row in enumerate(rows, start=2):
        for c, (_, key) in enumerate(COLUMNS, start=1):
            value = row.get(key)
            target = ws.cell(row=r, column=c, value=value)
            target.font = body_font
            if value is None:
                continue
            # 按列设格式，避免日期/金额在 Excel 里显示成小数
            if key in DATE_KEYS:
                target.number_format = 'yyyy-mm-dd'
            elif key in MONEY_KEYS:
                target.number_format = '#,##0.00'
            elif key in PERCENT_KEYS:
                target.number_format = '0.00%'
            elif key == 'tax':
                target.number_format = '0.00'
    ws.freeze_panes = 'A2'
    try:
        wb.save(path)
    except PermissionError:
        raise AppError(f'结果文件被占用，请先关闭后再核算：{path}')
    except OSError as e:
        raise AppError(f'结果文件保存失败：{path}\n{e}')


def compute(paths, output, rate, log, progress):
    # 三个过程表都只读第一个 sheet，表头行由 find_header 自动定位
    log(f'开始核算：{datetime.datetime.now():%Y-%m-%d %H:%M:%S}')
    progress(10)
    gm_rows = first_sheet_rows(paths['账面毛利'], '账面毛利')
    gm_idx, gm_col = find_header(gm_rows, GM_COLS, '账面毛利')
    pay_rows = first_sheet_rows(paths['付款信息'], '付款信息')
    pay_idx, pay_col = find_header(pay_rows, PAY_COLS, '付款信息')
    acc_rows = first_sheet_rows(paths['累计下单费用'], '累计下单费用')
    acc_idx, acc_col = find_header(acc_rows, ACC_COLS, '累计下单费用')
    log(f'账面毛利 {len(gm_rows) - gm_idx - 1} 行，付款信息 {len(pay_rows) - pay_idx - 1} 行，'
        f'累计下单费用 {len(acc_rows) - acc_idx - 1} 行')
    progress(40)
    pay_index = build_pay_index(pay_rows[pay_idx + 1:], pay_col)
    acc_index = build_acc_index(acc_rows[acc_idx + 1:], acc_col)
    stats = Stats()
    rows = build_output_rows(gm_rows[gm_idx + 1:], gm_col, pay_index, acc_index, rate, stats)
    progress(80)
    write_result(rows, output)
    progress(100)

    c = stats.count
    # 参与核算行数 = 结果行数 - 拆出的新行，避免把拆行算成新增数据
    log(f'账面毛利剔除：外购/价外费用/暂未下单 {c["excluded"]} 行，预提 {c["excluded_yuti"]} 行，'
        f'无合同号 {c["no_contract"]} 行；参与核算 {len(rows) - c["split"]} 行，按付款日期拆行 {c["split"]} 行，'
        f'结果共 {len(rows)} 行')
    log(f'付款信息未匹配 {c["no_pay"]} 行；HYX 金额不一致 {c["hyx_mismatch"]} 行')
    log(f'不计费用 {c["no_fee"]} 行（其中累计表已算过 {c["counted"]} 行）；人工检查 {c["manual"]} 行')
    log(f'结果文件：{os.path.abspath(output)}')


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('下单费用核算工具 v1.0')
        self.root.geometry('880x640')
        self.paths = {label: tk.StringVar() for label in FILE_LABELS}
        self.rate_var = tk.StringVar(value=str(load_rate_config()))
        self.output_var = tk.StringVar()
        self.output_auto = True   # 结果文件路径是否仍为自动默认值
        self._build_ui()
        self._show_help()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding='10')
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text='下单费用核算', font=('微软雅黑', 15, 'bold')).pack(anchor=tk.W)
        ttk.Label(main, text='选择三个文件后点击"核算"，生成结果文件', font=('微软雅黑', 10)).pack(
            anchor=tk.W, pady=(2, 10))

        file_frame = ttk.LabelFrame(main, text='文件选择（只读取第一个 sheet）', padding='10')
        file_frame.pack(fill=tk.X)
        for i, label in enumerate(FILE_LABELS):
            ttk.Label(file_frame, text=f'{label}：', width=12, anchor=tk.E).grid(row=i, column=0, padx=5, pady=3, sticky=tk.E)
            ttk.Entry(file_frame, textvariable=self.paths[label], width=80).grid(row=i, column=1, padx=5, pady=3, sticky=tk.EW)
            ttk.Button(file_frame, text='浏览...', width=10, command=lambda l=label: self.browse(l)).grid(
                row=i, column=2, padx=5, pady=3)
        file_frame.columnconfigure(1, weight=1)

        param_frame = ttk.LabelFrame(main, text='核算参数', padding='10')
        param_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(param_frame, text='超贷利率：', width=12, anchor=tk.E).grid(row=0, column=0, padx=5, pady=3, sticky=tk.E)
        ttk.Entry(param_frame, textvariable=self.rate_var, width=12).grid(row=0, column=1, padx=5, pady=3, sticky=tk.W)
        ttk.Label(param_frame, text='年利率，可写 0.0828 或 8.28%；发货款天数差>8 天时按此利率计财务费用',
                  font=('微软雅黑', 9)).grid(row=0, column=2, columnspan=2, padx=5, sticky=tk.W)
        ttk.Label(param_frame, text='结果文件：', width=12, anchor=tk.E).grid(row=1, column=0, padx=5, pady=3, sticky=tk.E)
        ttk.Entry(param_frame, textvariable=self.output_var, width=80).grid(row=1, column=1, padx=5, pady=3, sticky=tk.EW)
        ttk.Button(param_frame, text='浏览...', width=10, command=self.browse_output).grid(row=1, column=2, padx=5, pady=3)
        param_frame.columnconfigure(1, weight=1)

        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=(10, 5))
        self.calc_btn = ttk.Button(btn_frame, text='核算', width=14, command=self.start)
        self.calc_btn.pack(side=tk.LEFT)
        self.open_btn = ttk.Button(btn_frame, text='打开结果文件', width=16, command=self.open_output, state=tk.DISABLED)
        self.open_btn.pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text='清空日志', width=12, command=self.clear_log).pack(side=tk.RIGHT)

        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(main, variable=self.progress_var, maximum=100).pack(fill=tk.X, pady=(0, 10))

        log_frame = ttk.LabelFrame(main, text='处理日志', padding='5')
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=('Consolas', 9), state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.status_var = tk.StringVar(value='就绪')
        ttk.Label(main, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, pady=(8, 0))

    def _show_help(self):
        self._append_log('使用说明：\n'
                 '  1. 分别选择 账面毛利、付款信息、累计下单费用 三个 xlsx 文件\n'
                 '  2. 按需调整超贷利率与结果文件保存路径（默认与账面毛利同目录），点击"核算"\n'
                 '  3. 规则：付款日期不同按日期拆行（最早日期留原行）；HYX 合同金额不一致、'
                 '|核查-使用激励|>5 标"人工检查"；\n'
                 '     累计表已算过 / 开单成本≤0 / 培训区域囤券类项目 标"不计费用"')

    def browse(self, label):
        path = filedialog.askopenfilename(title=f'选择{label}文件', parent=self.root,
                                          filetypes=[('Excel文件', '*.xlsx'), ('所有文件', '*.*')])
        if path:
            self.paths[label].set(path)
            self.status_var.set(f'已选择：{os.path.basename(path)}')
            # 用户未手动改过结果文件路径时，始终跟随账面毛利所在目录
            if label == FILE_LABELS[0] and self.output_auto:
                self.output_var.set(os.path.join(os.path.dirname(path), OUTPUT_NAME))

    def browse_output(self):
        # 默认从当前结果路径或账面毛利目录开始选
        current = self.output_var.get().strip()
        initialdir = os.path.dirname(current) or os.path.dirname(self.paths[FILE_LABELS[0]].get().strip()) or os.getcwd()
        path = filedialog.asksaveasfilename(
            title='选择结果文件保存位置', parent=self.root, defaultextension='.xlsx', initialdir=initialdir,
            initialfile=os.path.basename(current) or OUTPUT_NAME, filetypes=[('Excel文件', '*.xlsx')])
        if path:
            self.output_var.set(path)
            self.output_auto = False

    def _append_log(self, message):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def clear_log(self):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def start(self):
        # 先校验输入与参数，再启动后台线程，避免跑到一半才报错
        paths = {label: var.get().strip() for label, var in self.paths.items()}
        for label, path in paths.items():
            if not path or not os.path.isfile(path):
                messagebox.showerror('错误', f'请选择有效的{label}文件', parent=self.root)
                return
        try:
            rate = parse_rate(self.rate_var.get())
            save_rate_config(rate)
        except AppError as e:
            messagebox.showerror('错误', str(e), parent=self.root)
            return
        output = self.output_var.get().strip()
        if not output:
            output = os.path.join(os.path.dirname(paths['账面毛利']), OUTPUT_NAME)
        if not os.path.splitext(output)[1]:
            output += '.xlsx'
        self.output_var.set(output)
        self.output_auto = True
        inputs = {os.path.normcase(os.path.abspath(p)) for p in paths.values()}
        if os.path.isdir(output):
            messagebox.showerror('错误', '结果文件路径不能是文件夹，请填写完整的文件名', parent=self.root)
            return
        # 禁止覆盖输入文件，防止把过程表写坏
        if os.path.normcase(os.path.abspath(output)) in inputs:
            messagebox.showerror('错误', '结果文件不能覆盖输入文件，请重新选择保存位置', parent=self.root)
            return
        # 已存在时确认覆盖，避免误覆盖上期结果
        if os.path.exists(output) and not messagebox.askyesno('确认', f'结果文件已存在，是否覆盖？\n{output}', parent=self.root):
            return
        self.clear_log()
        self.calc_btn.configure(state=tk.DISABLED)
        self.open_btn.configure(state=tk.DISABLED)
        self.progress_var.set(0)
        self.status_var.set('正在核算...')
        # 工作线程只往队列写消息，界面更新统一由主线程轮询执行（Tk 控件不能跨线程操作）
        self.queue = queue.Queue()
        threading.Thread(target=self._run, args=(paths, output, rate), daemon=True).start()
        self.root.after(100, self._poll)

    def _poll(self):
        # 逐条消费后台消息；收到 finish 后停止轮询并恢复按钮
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == 'log':
                    self._append_log(payload)
                elif kind == 'progress':
                    self.progress_var.set(payload)
                elif kind == 'error':
                    messagebox.showerror('核算失败', payload, parent=self.root)
                    self.status_var.set('核算失败')
                elif kind == 'done':
                    self.open_btn.configure(state=tk.NORMAL)
                    self.status_var.set(f'核算完成：{payload}')
                elif kind == 'finish':
                    self.calc_btn.configure(state=tk.NORMAL)
                    return
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _run(self, paths, output, rate):
        # 后台只算与写文件，界面更新全部交给 _poll
        q = self.queue
        try:
            compute(paths, output, rate, lambda m: q.put(('log', m)), lambda v: q.put(('progress', v)))
        except Exception as e:
            q.put(('log', f'[错误] {e}'))
            q.put(('error', str(e)))
        else:
            q.put(('done', output))
        finally:
            q.put(('finish', None))

    def open_output(self):
        output = self.output_var.get().strip()
        if output and os.path.exists(output):
            os.startfile(output)

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    App().run()
