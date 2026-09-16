#!/usr/bin/env python3
"""
ETL for the 2025 full-year source files (first batch of prior-year data for
YoY comparison). Kept as a separate script from etl.py (2026 H1) rather than
merged in, so the working 2026 pipeline is never at risk while this batch is
still being validated -- more 2025/2024 files are expected before this is final.

Reuses the exact same helper conventions (add_record shape, brand/model
classification, date parsing) as etl.py so records.json + records_2025.json
can eventually be concatenated with zero schema drift.
"""
import os
import openpyxl, re, json
from collections import defaultdict, Counter
from datetime import datetime, date

SRC = os.path.join(os.environ["SALES_RAW_DIR"], "")  # patched: was the cloud upload dir

records = []
issues = []
def note(msg):
    issues.append(msg)
    print("NOTE:", msg)

# ---------------------------------------------------------------- helpers (mirrors etl.py)
def parse_dmy(s):
    if not s:
        return None
    s = str(s).strip()
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', s)
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return date(y, mo, d)
    except ValueError:
        return None

def to_date(v):
    if isinstance(v, (date, datetime)):
        return v.date() if isinstance(v, datetime) else v
    return parse_dmy(v)

def num(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', '').strip())
    except ValueError:
        return 0.0

def resolve_order_amount(line_total, paid, raw_discount):
    """Resolve an order's actual collected amount, cross-checking 'Payment
    amount' against the recorded 'Discount' to guard against occasional
    broken exports seen across real files: some registers log only part of
    a split/voucher payment in Payment amount while Discount stays an
    explicit '0.0' (Payment amount is then unreliable -- trust the line
    total instead); rarely a stray fractional Discount typo shows up that
    is not really a THB amount (Discount is then unreliable -- trust
    Payment amount instead, as already handled by requiring discount >= 1
    below). Percentage-formatted Discount values ('40.00%') are ignored
    entirely since they can't be compared in THB.
    """
    if not line_total:
        return 0.0
    discount = None
    if raw_discount not in (None, '') and not (isinstance(raw_discount, str) and '%' in raw_discount):
        discount = num(raw_discount)
    discount_implied = (line_total - discount) if (discount is not None and discount >= 1) else None
    if paid is None:
        return discount_implied if discount_implied is not None else line_total
    # Only distrust 'paid' when it looks suspiciously TOO LOW versus what
    # Total/Discount implies -- never when it's higher (a legitimate
    # marketplace-subsidized voucher, tip, or other case we don't model can
    # make actual payment exceed a naive total-minus-discount; only a value
    # far below that floor is a sign of a truncated/partial payment read).
    expected = discount_implied if discount_implied is not None else (line_total if discount == 0 else None)
    if expected is not None:
        tol = max(500.0, 0.15 * line_total)
        if paid < expected - tol:
            return expected
    return paid

BRAND_PREFIX = {
    'VFF': 'VFF', 'BFJ': 'BFJ', 'CP': 'Coolcore', 'CC': 'Coolcore',
    'OLN': 'Oleno', 'MTB': 'TabiRela',
    'KK': 'Others', 'SW': 'Others', 'KA': 'Others', 'LN': 'Others', 'TUB': 'Others',
    # new-in-2025 codes confirmed by user as "Others": KEEN (3rd-party brand),
    # Pet House (ST), TBE (TM); IS spans several small accessory lines (Pet Cool
    # Tank, Blanket, Mesh Cloth, Stole, generic Socks) disambiguated by name below.
    'KE': 'Others', 'ST': 'Others', 'TM': 'Others', 'IS': 'Others',
}
OTHERS_SUBBRAND_PREFIX = {'KK': 'Klean Kanteen', 'SW': 'Swans', 'KA': 'Knockaround', 'LN': 'LUNA', 'TUB': 'Tube',
                           'KE': 'KEEN', 'ST': 'Pet House', 'TM': 'TBE'}
NON_PRODUCT_PREFIX = {'DC', 'CON', 'P'}

def normalize_channel(v):
    if not v:
        return None
    s = str(v).strip()
    if not s:
        return None
    if s.startswith('Shopee'):
        return 'Shopee'
    return s

def classify_by_code(code, name=None):
    if not code:
        return None, None
    m = re.match(r'^[A-Za-z]+', str(code))
    prefix = m.group(0) if m else str(code)
    if prefix in NON_PRODUCT_PREFIX:
        return 'EXCLUDE', None
    brand = BRAND_PREFIX.get(prefix)
    sub = OTHERS_SUBBRAND_PREFIX.get(prefix)
    if brand == 'Others' and sub is None and name and 'PET' in str(name).upper():
        # "IS" spans multiple small-accessory lines -- Pet Cool Tank shares the
        # Pet House sub-brand grouping, disambiguated by product name.
        sub = 'Pet House'
    return brand, sub

def model_from_name(name):
    if not name:
        return 'Other'
    return str(name).split('(')[0].strip()

NAME_VFF_KEYWORDS = ['V-SOUL','V-RUN','V-TREK','V-ALPHA','V-TRAIN','V-AQUA','V-TRAIL','KSO','BREEZANDAL',
                      'SPIDRWALK','SCRAMKEY','TRAILOPE','GROUNDSPLAY','GRASPIFIER','SOCKS MINI CREW',
                      'SOCKS CREW','SOCKS HIGH CREW','HIGH CREW','CVT HEMP','KMD','VFF','FURO','ALITZA',
                      'SPYRIDON']
# name-based non-VFF items confirmed by user (2025 batch review) -- rolled into
# Others with a descriptive sub-brand rather than left as Unknown.
NAME_OTHERS_SUBBRAND = {
    'KEEN': 'KEEN', 'PET HOUSE': 'Pet House', 'PET COOL TANK': 'Pet House',
}
# small residual accessory items confirmed as non-VFF "Others" with no
# meaningful sub-brand grouping (matched by exact base name only).
NAME_OTHERS_EXACT = {'TBE', 'MESH CLOTH', 'BLANKET', 'SOCKS', 'STOLE'}
def classify_by_name(name):
    if not name:
        return None, None, 'Other'
    n = str(name).upper()
    # hyphens vs spaces vary across source files for the same model (e.g.
    # "CVT-HEMP" here vs "CVT HEMP" elsewhere) -- normalize only for keyword
    # MATCHING, not for the display name, so "V-Trail" doesn't fragment into
    # a separate "V Trail" model from the hyphenated spelling used elsewhere.
    n_match = n.replace('-', ' ')
    base = n.split('(')[0].strip().title()
    if 'BFJ' in n:
        return 'BFJ', None, base
    if n.startswith('OLENO') or n.startswith('OLN'):
        return 'Oleno', None, base
    if n.startswith('TABI'):
        return 'TabiRela', None, ('Marugo ' + base if not base.upper().startswith('MARUGO') else base)
    # El-X Knit / Classic are confirmed VFF footwear lines despite non-standard
    # naming (no "V-" prefix) -- matched by exact base name, not substring, to
    # avoid over-matching unrelated future items that happen to contain "Classic".
    # "El-X Knit" also appears misspelled several ways across years/sheets
    # ("El X-Knit", "Elx-Knit", "Exl-Knit") -- canonicalize all of them to one
    # display name so they roll up as a single model, not fragment into three.
    base_match = base.upper().replace('-', ' ')
    if base_match in ('EL X KNIT', 'ELX KNIT', 'EXL KNIT'):
        return 'VFF', None, 'VFF El-X Knit'
    if base_match == 'CLASSIC':
        return 'VFF', None, ('VFF ' + base if not base.upper().startswith('VFF') else base)
    for kw in NAME_VFF_KEYWORDS:
        # keywords themselves may contain hyphens (e.g. "V-SOUL") -- normalize
        # the same way as n_match so hyphenated keywords still match.
        if kw.replace('-', ' ') in n_match:
            return 'VFF', None, ('VFF ' + base if not base.upper().startswith('VFF') else base)
    for kw, sub in NAME_OTHERS_SUBBRAND.items():
        if kw in base_match:
            return 'Others', sub, base
    if base_match in NAME_OTHERS_EXACT:
        return 'Others', None, base
    return 'Unknown', None, base

VFF_SIZE_TOKEN_RE = re.compile(r'^([MWUmwu]?)(\d+)$')
def vff_shoe_gender(text):
    if not text:
        return False, None
    s = str(text)
    m = re.search(r'\(([^()]*)\)', s)
    if m:
        for token in (p.strip() for p in m.group(1).split(',')):
            gm = VFF_SIZE_TOKEN_RE.match(token)
            if gm:
                g = gm.group(1).upper()
                return True, ('Women' if g == 'W' else 'Men' if g == 'M' else 'Unisex')
    fm = re.search(r'([MWUmwu])(\d{2,3})\)?\s*$', s)
    if fm:
        g = fm.group(1).upper()
        return True, ('Women' if g == 'W' else 'Men' if g == 'M' else 'Unisex')
    return False, None

def add_record(store, category, date_, brand, model, sub, qty, amount, order_id,
                channel=None, payment=None, entity=None, vff_source_text=None):
    if date_ is None:
        return
    if brand == 'EXCLUDE':
        return
    is_shoe, gender = (False, None)
    if brand == 'VFF':
        is_shoe, gender = vff_shoe_gender(vff_source_text)
    records.append(dict(
        store=store, category=category, date=date_.isoformat(),
        month=date_.strftime('%Y-%m'), brand=brand or 'Unknown', model=model,
        is_vff_shoe=is_shoe, gender=gender,
        sub=sub, qty=qty, amount=amount, order_id=order_id,
        channel=channel, payment=payment, entity=entity,
    ))

# CODE_TO_MODEL for short-code files (Central Total Department) -- rebuilt from
# the same well-labeled 2026 files used in etl.py (product names don't change
# year to year for the same SKU, so this lookup is shared across both years).
CODE_TO_MODEL = {}
def _scan_codes(fn, sheet, header_row_idx, code_key, name_key):
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[header_row_idx]
    idx = {h: i for i, h in enumerate(header) if h}
    for r in rows[header_row_idx + 1:]:
        if not any(r):
            continue
        code = r[idx.get(code_key)] if code_key in idx else None
        name = r[idx.get(name_key)] if name_key in idx else None
        if not code:
            continue
        m = re.match(r'^([A-Za-z]+)0*(\d+)', str(code).upper())
        if not m:
            continue
        prefix, num_code = m.group(1), int(m.group(2))
        model = model_from_name(name)
        key = (prefix, num_code)
        if model and model not in ('Other',) and key not in CODE_TO_MODEL:
            CODE_TO_MODEL[key] = model

_scan_codes(SRC + 'a0d9bc79-Sales_K_village_JanJun_26.xlsx', 'Orders', 1, 'Product code', 'Product name')
_scan_codes(SRC + '835c1952-Sales_Thaniya_JanJun_26.xlsx', 'Orders', 1, 'Product code', 'Product name')
_scan_codes(SRC + 'cf220ee8-BFT_Shopee_Lazada_Facebook_JanJun_26.xlsx', 'Orders', 1, 'Product code', 'Product name')
_scan_codes(SRC + '2932d908-Sales_Paradies_Park_JanJun_26.xlsx', 'Orders', 1, 'Product code', 'Product name')
print(f"Built brand code->model lookup with {len(CODE_TO_MODEL)} entries")


# ================================================================== 1. BFT consignment 2025 (English headers, prorated Grand Total)
# Different schema from the 2026 file (Thai headers, single 'Amount' field
# present only on an invoice's first line). This file instead carries, on
# EVERY line, list-price 'Pre-VAT Amount' + 'VAT Amount' (pre invoice-level
# discount), and on only the invoice's first line, the invoice-level
# 'Discount %' + the actual post-discount 'Grand Total'. Summing the
# per-line Pre-VAT+VAT directly overstates revenue (verified: sums to
# 1,545,276 vs the correct Grand-Total sum of 1,231,130 across the file) --
# so each line's share of Grand Total is prorated by its own Pre-VAT+VAT
# share of the invoice total, per user confirmation.
def load_bft_consignment_2025():
    fn = SRC + '7250e176-BFT_consignment__2025.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['รายงานใบแจ้งหนี้']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]  # single header row (unlike the 2026 file's 2-row Thai header)

    inv_pretax = defaultdict(float)
    inv_grand = {}
    for r in data:
        if not any(r):
            continue
        doc, pc = r[0], r[3]
        if not doc or not pc:
            continue
        prevat, vat = num(r[8]), num(r[9])
        inv_pretax[doc] += prevat + vat
        gt = r[13]
        if gt not in (None, ''):
            inv_grand[doc] = num(gt)

    n = 0
    for r in data:
        if not any(r):
            continue
        doc, pc = r[0], r[3]
        if not doc or not pc:
            continue
        d = parse_dmy(r[1])
        if d is None:
            continue
        qty = num(r[5])
        prevat, vat = num(r[8]), num(r[9])
        line_pretax = prevat + vat
        total_pretax = inv_pretax.get(doc, 0.0)
        grand = inv_grand.get(doc, total_pretax)
        amt = (line_pretax / total_pretax * grand) if total_pretax else 0.0
        pname = r[4]
        brand, sub = classify_by_code(pc, pname)
        model = model_from_name(pname)
        add_record('BFT Consignment', 'consignment', d, brand, model, sub, qty, amt, doc, vff_source_text=pc)
        n += 1
    print(f"loaded {n} rows -> BFT Consignment 2025")
load_bft_consignment_2025()


# ================================================================== 2/3. BFT & EDV Online/Event/Store/Consignment 2025
# Both files share an identical flat schema (unlike the old multi-row-per-order
# "Orders" export -- every field, including Sales channel/Warehouse/Branch, is
# repeated on every product line here, so no order-level forward-fill is
# needed). Warehouse/Branch is the dimension that used to be a separate file
# per store/channel in the 2026 batch; map it explicitly per confirmed rules.
# 'Total amount' is each line's PRE-discount total; the order's actual
# collected revenue is 'Payment amount', which (like the Central LP 3F Thai
# export) only appears once, on the order's first product line. Confirmed
# 2026-08 against store-level payment-method ledgers: summing 'Total amount'
# directly overstates revenue by the order-level discount, so each line is
# prorated to its share of the order's actual amount paid. Orders with no
# 'Payment amount' anywhere (a handful per file) fall back to their own
# line-total sum unprorated.
def load_flat_branch_file(fn, sheet, branch_map, label):
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[1]  # row 0 is a merged "Product data" title band
    idx = {h: i for i, h in enumerate(header) if h}
    data = rows[2:]
    data = [r for r in data if any(r) and r[idx.get('Product code')]]

    order_line_total = defaultdict(float)  # order_id -> sum('Total amount') across its lines
    order_paid, order_discount = {}, {}    # order_id -> 'Payment amount' / 'Discount' (first line only)
    for r in data:
        order_id = r[idx.get('Sales order No.')]
        order_line_total[order_id] += num(r[idx.get('Total amount')])
        pay = r[idx.get('Payment amount')]
        if pay not in (None, ''):
            order_paid[order_id] = num(pay)
        disc = r[idx.get('Discount')]
        if disc not in (None, '') and order_id not in order_discount:
            order_discount[order_id] = disc

    order_resolved = {}
    n, skipped_unmapped = 0, 0
    for r in data:
        branch = r[idx.get('Warehouse/Branch')]
        mapping = branch_map.get(branch)
        if mapping is None:
            skipped_unmapped += 1
            note(f"{label}: unmapped Warehouse/Branch value {branch!r} (1 row skipped)")
            continue
        store, cat = mapping
        d = to_date(r[idx.get('Date')])
        if d is None:
            continue
        pc = r[idx.get('Product code')]
        order_id = r[idx.get('Sales order No.')]
        qty = num(r[idx.get('Quantity')])
        line_amt = num(r[idx.get('Total amount')])
        order_total = order_line_total[order_id]
        if order_id not in order_resolved:
            order_resolved[order_id] = resolve_order_amount(
                order_total, order_paid.get(order_id), order_discount.get(order_id))
        resolved = order_resolved[order_id]
        amt = (line_amt / order_total * resolved) if order_total else 0.0
        channel = normalize_channel(r[idx.get('Sales channel')])
        pname = r[idx.get('Product name')]
        brand, sub = classify_by_code(pc, pname)
        model = model_from_name(pname)
        add_record(store, cat, d, brand, model, sub, qty, amt, order_id, channel=channel, vff_source_text=pc)
        n += 1
    overridden = sum(1 for oid, p in order_paid.items()
                      if abs(order_resolved.get(oid, p) - p) > 0.01)
    print(f"loaded {n} rows -> {label} ({skipped_unmapped} unmapped, "
          f"{len(order_paid)}/{len(order_line_total)} orders prorated to actual amount paid, "
          f"{overridden} Payment-amount overridden as unreliable)")

BFT_BRANCH_MAP_2025 = {
    'Online': ('Online', 'online'),
    'Event 1': ('Event', 'event'),
    'Event 2': ('Event', 'event'),
    'Event': ('Event', 'event'),
    # Pre-Order lines fulfilled from the online store, per user confirmation
    'Pre-Order Free Bag': ('Online', 'online'),
    'Pre-Order V-Soul Ivory': ('Online', 'online'),
    'Pre-Order': ('Online', 'online'),
    # "Main Warehouse" -- treated as Online per user confirmation
    'คลังสินค้าหลัก': ('Online', 'online'),
}
load_flat_branch_file(SRC + '789168d2-BFT_Sale_Online_Shopee_Lazada_Event_2025.xlsx', 'Orders',
                       BFT_BRANCH_MAP_2025, 'BFT Online/Event 2025')

EDV_BRANCH_MAP_2025 = {
    'Thaniya': ('Thaniya', 'store'),
    'Coollabo Cen LP 3F': ('Central Ladprao 3F (Coollabo)', 'store'),
    'Central LP 3F': ('Central Ladprao 3F (Coollabo)', 'store'),
    'Kvillage': ('K Village', 'store'),
    'K village': ('K Village', 'store'),
    'Event 1': ('Event', 'event'),
    'Event': ('Event', 'event'),
    # individual consignment-partner shop names -- confirmed to roll up into EDV Consignment
    'Banana Run': ('EDV Consignment', 'consignment'),
    'Avarin': ('EDV Consignment', 'consignment'),
    'Mega Bangna': ('EDV Consignment', 'consignment'),
    'Runnercart': ('EDV Consignment', 'consignment'),
    'Caveman': ('EDV Consignment', 'consignment'),
    'Anvil Camp': ('EDV Consignment', 'consignment'),
    'Highlandner': ('EDV Consignment', 'consignment'),
    'Pathwild': ('EDV Consignment', 'consignment'),
    'Art of Golf': ('EDV Consignment', 'consignment'),
    'Consignment': ('EDV Consignment', 'consignment'),
    # "Main Warehouse" -- treated as Online per user confirmation
    'คลังสินค้าหลัก': ('Online', 'online'),
}
load_flat_branch_file(SRC + '7ec8bd46-EDV_Sales_Online_Shopee_Lazada_K_village_Central_LP_Thaniya_Consignment_2025.xlsx',
                       'Orders', EDV_BRANCH_MAP_2025, 'EDV Multi-store 2025')


# ================================================================== 4. Siam Discovery 2025 (monthly sheets, richer schema than the 2026 file)
# Columns (most months): Date, running ID#, Product name, Qty, <cost col>,
# Sell price, Net amount, <blank padding...>. Net amount is consistently 68%
# of sell price across every sampled row (e.g. 5400*0.68=3672.00,
# 500*0.68=340.00) -- the consignment-net figure, same semantic role as the
# single 'Net' column the 2026 file used directly.
#
# December's sheet (12-25) SHIFTS this layout left by one column -- it drops
# the running-ID column entirely, so Product name lands at index 1 instead of
# 2 (verified: December row 0 is (Date, 'KSO EVO...', 1, 4940, 5200, 3536.0)
# vs November's (Date, 9, 'KSO EVO...', 1, 4940, 5200, 3536.0)). Reading every
# sheet at fixed positions silently fed price figures in as quantities for
# all of December (a qty total of 206,435 for one month, vs ~1-2K for every
# other month, was the tell). Detect the layout per sheet instead of
# assuming it's fixed. Each sheet also has a handful of trailing
# subtotal/reference rows with columns shifted again and no date -- requiring
# a parseable date in column 0 filters those out uniformly regardless of
# which layout is in play.
# Per user confirmation, uses the 12 individual monthly sheets (01-25..12-25),
# not the "รวม 25" summary tab (which undercounts by 49 rows against these).
def load_siam_discovery_2025():
    fn = SRC + '1656eb05-Sales_Record_Siam_discovery.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    months = ['01-25','02-25','03-25','04-25','05-25','06-25','07-25','08-25','09-25','10-25','11-25','12-25']
    n = 0
    for sn in months:
        ws = wb[sn]
        rows = list(ws.iter_rows(values_only=True))
        name_idx = None
        for r in rows:
            if not r or to_date(r[0]) is None:
                continue
            if isinstance(r[1], str) and r[1].strip():
                name_idx = 1  # December-style: no ID column
                break
            if len(r) > 2 and isinstance(r[2], str) and r[2].strip():
                name_idx = 2  # normal style: Date, ID#, Name, ...
                break
        if name_idx is None:
            note(f"Siam Discovery 2025 sheet {sn}: could not detect column layout, skipped entirely")
            continue
        # net_idx: column layout is Date/[ID]/Name/Qty/NetSales/Price/GP32%.
        # Per user confirmation 2026-08, Siam Discovery's recognized revenue
        # is net of their 32% consignment fee -- GP32% (always exactly
        # Price*0.68) is the correct "amount" here, 4 columns after Name,
        # not the gross NetSales/Price columns in between.
        qty_idx, net_idx = name_idx + 1, name_idx + 4
        month_n = 0
        for i, r in enumerate(rows):
            if not r:
                continue
            d = to_date(r[0])
            if d is None:
                continue  # also filters the trailing subtotal/reference rows (no date)
            item = r[name_idx] if len(r) > name_idx else None
            if not item or not isinstance(item, str):
                continue
            qty = num(r[qty_idx]) if len(r) > qty_idx else 0.0
            net = num(r[net_idx]) if len(r) > net_idx else 0.0
            brand, sub, model = classify_by_name(item)
            add_record('Siam Discovery', 'store', d, brand, model, sub,
                        qty, net, f'SIAMDIS25-{sn}-{i}', vff_source_text=item)
            n += 1
            month_n += 1
        print(f"  {sn}: name_idx={name_idx}, {month_n} rows")
    print(f"loaded {n} rows -> Siam Discovery 2025")
load_siam_discovery_2025()


# ================================================================== 5. BFT_Central_Total_Department 2025 (Sep-Dec only, 3 of 5 stores)
def load_central_total_department_2025():
    fn = SRC + 'f6f9daad-BFT_Central_Total_Department_2025.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['Export']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    idx = {h: i for i, h in enumerate(header) if h}
    data = rows[1:]
    MONTHS = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
              'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
    STORE_NAME_MAP = {
        'CHIDLOM': 'Central Chidlom',
        'CHIDLOM ONLINE': 'Central Chidlom Online',
        'CENTRAL WORLD-CDS': 'Central World (CDS)',
        'LARDPRAO': 'Central Lardprao (Dept.)',
        'EASTVILLE': 'Central Eastville',
    }
    n = 0
    for r in data:
        if not any(r):
            continue
        store = r[idx['Store Name']]
        cat = r[idx['Catalogue No.']]
        msd = r[idx['Month Sales Date']]
        if not store or not cat or not msd:
            continue
        mo, yr = str(msd).split('-')
        d = date(int(yr), MONTHS[mo], 1)
        qty = num(r[idx['Sales Quantity']])
        amt = num(r[idx['Total Net Sales (Sales Amount)']])
        brand, sub = classify_by_code(cat)
        model = None
        mcode = re.match(r'^([A-Za-z]+)0*(\d+)', str(cat).upper())
        if mcode:
            model = CODE_TO_MODEL.get((mcode.group(1), int(mcode.group(2))), 'Other')
        store_label = STORE_NAME_MAP.get(store, f'Central {store.title()}')
        add_record(store_label, 'central_dept', d, brand, model, sub, qty, amt,
                   order_id=None, vff_source_text=cat)
        n += 1
    print(f"loaded {n} rows -> Central Total Department 2025 (Sep-Dec only, {len(set(r[idx['Store Name']] for r in data if r[idx['Store Name']]))} stores)")
load_central_total_department_2025()


# ================================================================== summary / sanity checks
print()
print("="*80)
print(f"TOTAL 2025 RECORDS: {len(records)}")
total_amt = sum(r['amount'] for r in records)
total_qty = sum(r['qty'] for r in records)
print(f"TOTAL AMOUNT: {total_amt:,.2f}")
print(f"TOTAL QTY: {total_qty:,.0f}")

by_store = defaultdict(lambda: {'amount':0.0,'qty':0.0,'orders':set()})
for r in records:
    by_store[r['store']]['amount'] += r['amount']
    by_store[r['store']]['qty'] += r['qty']
    if r['order_id']:
        by_store[r['store']]['orders'].add(r['order_id'])
print()
print(f"{'Store':40s} {'Amount':>15s} {'Qty':>8s} {'Orders':>8s}")
for store, v in sorted(by_store.items(), key=lambda x: -x[1]['amount']):
    print(f"{store:40s} {v['amount']:>15,.2f} {v['qty']:>8,.0f} {len(v['orders']):>8d}")

by_month = defaultdict(lambda: {'amount':0.0,'qty':0.0})
for r in records:
    by_month[r['month']]['amount'] += r['amount']
    by_month[r['month']]['qty'] += r['qty']
print()
print("Monthly totals:")
for m, v in sorted(by_month.items()):
    print(f"  {m}: amount={v['amount']:>14,.2f}  qty={v['qty']:>8,.0f}")

by_brand = defaultdict(lambda: {'amount':0.0,'qty':0.0})
for r in records:
    by_brand[r['brand']]['amount'] += r['amount']
    by_brand[r['brand']]['qty'] += r['qty']
print()
print("Brand totals:")
for b, v in sorted(by_brand.items(), key=lambda x: -x[1]['amount']):
    print(f"  {b:15s} amount={v['amount']:>14,.2f}  qty={v['qty']:>8,.0f}")

with open(os.path.join(os.environ['SALES_OUT_DIR'], 'records_2025.json'), 'w', encoding='utf-8') as f:
    json.dump(records, f, ensure_ascii=False)
print()
print("Saved", len(records), "records to records_2025.json")
