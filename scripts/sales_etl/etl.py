#!/usr/bin/env python3
"""
ETL for Barefoot Inc sales dashboard (Jan-Jun 2026).
Consolidates 16 source files into canonical transaction records, then aggregates.
"""
import os
import openpyxl, csv, re, json
from collections import defaultdict, Counter
from datetime import datetime, date

SRC = os.path.join(os.environ["SALES_RAW_DIR"], "")  # patched: was the cloud upload dir

records = []  # each: dict(store, category, channel, date(YYYY-MM-DD), month(YYYY-MM),
              #            brand, model, qty, amount, order_id, payment)
issues = []   # data-quality notes to surface

def note(msg):
    issues.append(msg)
    print("NOTE:", msg)

# ---------------------------------------------------------------- helpers
def parse_dmy(s):
    """Parse a D/M/YYYY (day-first, verified across all files so far) string date."""
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
    amount'/'จำนวนเงินที่ชำระ' (paid) against the recorded 'Discount'/'ส่วนลด'
    (raw_discount) to guard against occasional broken exports seen across
    real files: some registers log only part of a split/voucher payment in
    Payment amount while Discount stays an explicit '0.0' (Payment amount is
    then unreliable -- trust the line total instead); rarely a stray
    fractional Discount typo like '0.05' shows up that is not really a THB
    amount (Discount is then unreliable -- trust Payment amount instead, as
    already handled by simply requiring discount >= 1 below). Percentage-
    formatted Discount values ('40.00%') are ignored entirely since they
    can't be compared in THB.
    line_total: this order's summed line-item total (pre-discount).
    paid: 'Payment amount' if any line in the order recorded one, else None.
    raw_discount: the raw 'Discount' cell if any line recorded one, else None.
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
}
OTHERS_SUBBRAND_PREFIX = {'KK': 'Klean Kanteen', 'SW': 'Swans', 'KA': 'Knockaround', 'LN': 'LUNA', 'TUB': 'Tube'}
NON_PRODUCT_PREFIX = {'DC', 'CON', 'P'}  # DC=discount line, CON=consignment doc ref, P=misc asset-sale income (non-product)

def normalize_channel(v):
    if not v:
        return None
    s = str(v).strip()
    if not s:
        return None
    if s.startswith('Shopee'):
        return 'Shopee'
    return s

def normalize_payment(v):
    """Map raw Payment channel text to Cash / Credit Card / QR per confirmed business rule:
       Cash->Cash, credit card (any bank)->Credit Card, KBANK & bank-transfer (เงินโอน)->QR."""
    if not v:
        return None
    s = str(v).strip()
    if not s:
        return None
    if s == 'Cash':
        return 'Cash'
    if 'บัตรเครดิต' in s or 'เครดิต' in s.lower() or 'credit' in s.lower():
        return 'Credit Card'
    if s == 'KBANK' or 'เงินโอน' in s:
        return 'QR Code'
    return 'Other'

def classify_by_code(code):
    if not code:
        return None, None
    m = re.match(r'^[A-Za-z]+', str(code))
    prefix = m.group(0) if m else str(code)
    if prefix in NON_PRODUCT_PREFIX:
        return 'EXCLUDE', None
    brand = BRAND_PREFIX.get(prefix)
    sub = OTHERS_SUBBRAND_PREFIX.get(prefix)
    return brand, sub

def model_from_name(name):
    """Generic model label for ANY brand: full product name up to the first '(',
    e.g. 'VFF V-Soul', 'CP Arm Sleeves', 'Marugo TabiRela', 'BFJ Socks', 'Oleno Ultimate'.
    Kept brand-prefixed (not stripped) so top/bottom-model rankings read the same way
    across every source file."""
    if not name:
        return 'Other'
    return str(name).split('(')[0].strip()

# name-based classification for Siam Discovery (no product codes) -- normalized to the
# same brand-prefixed naming convention used by every other (code-based) source file.
NAME_VFF_KEYWORDS = ['V-SOUL','V-RUN','V-TREK','V-ALPHA','V-TRAIN','V-AQUA','KSO','BREEZANDAL',
                      'SPIDRWALK','SCRAMKEY','TRAILOPE','GROUNDSPLAY','GRASPIFIER','SOCKS MINI CREW',
                      'SOCKS CREW','SOCKS HIGH CREW','HIGH CREW','CVT HEMP','KMD','VFF']
def classify_by_name(name):
    if not name:
        return None, None, 'Other'
    n = str(name).upper()
    base = n.split('(')[0].strip().title()
    if 'BFJ' in n:
        return 'BFJ', None, base
    if n.startswith('OLENO') or n.startswith('OLN'):
        return 'Oleno', None, base
    if n.startswith('TABI'):
        return 'TabiRela', None, ('Marugo ' + base if not base.upper().startswith('MARUGO') else base)
    for kw in NAME_VFF_KEYWORDS:
        if kw in n:
            return 'VFF', None, ('VFF ' + base if not base.upper().startswith('VFF') else base)
    return 'Unknown', None, base

# VFF shoe / non-shoe (socks, furoshiki-wrap accessories with no numeric size, etc.)
# classifier, plus gender inference. VFF footwear codes/names always carry a size token
# shaped like an optional single gender letter + a number in the last parenthesized
# group -- "(BK,W37)" (code) or "(W37, Baby Blue)" (name) -- e.g. W37=Women's 37,
# M43=Men's 43, U/no-letter=Unisex. Sock-type items use letter-only sizing (S/M/L/XL)
# with no digits, which this pattern deliberately does not match.
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
    # fallback for malformed source codes missing the comma/closing paren
    # (seen in Central_Total_Department, e.g. "VFF02(TTBKM42" instead of
    # "VFF02(TTBK,M42)") -- look for a trailing gender-letter + size at the
    # very end of the string.
    fm = re.search(r'([MWUmwu])(\d{2,3})\)?\s*$', s)
    if fm:
        g = fm.group(1).upper()
        return True, ('Women' if g == 'W' else 'Men' if g == 'M' else 'Unisex')
    return False, None

# Color + size extraction for VFF shoes. Two conventions appear across the
# source files, both a single trailing parenthesized group:
#   name-style: 'Model(SizeToken, Color Name)'   e.g. '(W40, Fuchsia)'
#   code-style: 'PREFIX(ColorCode,SizeToken)'    e.g. '(TT/BK,M43)'
# Whichever end of the comma list matches the gender+size token shape
# (letter?+digits) is size; the other end is the color, resolved to a full
# name via COLOR_CODE_TO_NAME (built below) when it's an abbreviated code --
# Siam Discovery's item names and Central_Total_Department's catalogue codes
# are code-style with no separate name field of their own, so their color
# text only ever comes through as a code. Falls back to the raw text when no
# mapping exists (better than dropping a real sale from the breakdown).
COLOR_CODE_TO_NAME = {}
def vff_shoe_size_color(text):
    if not text:
        return None, None
    m = re.search(r'\(([^()]*)\)\s*$', str(text))
    if not m:
        return None, None
    parts = [p.strip() for p in m.group(1).split(',')]
    if len(parts) < 2:
        return None, None
    if VFF_SIZE_TOKEN_RE.match(parts[0]):
        size_tok, color_raw = parts[0], ','.join(parts[1:]).strip()
    elif VFF_SIZE_TOKEN_RE.match(parts[-1]):
        size_tok, color_raw = parts[-1], ','.join(parts[:-1]).strip()
    else:
        return None, None
    size_num = VFF_SIZE_TOKEN_RE.match(size_tok).group(2)
    color = COLOR_CODE_TO_NAME.get(color_raw.upper(), color_raw) if color_raw else None
    if color:
        color = COLOR_NAME_CANON.get(color.upper(), color)
    return size_num, color

def add_record(store, category, date_, brand, model, sub, qty, amount, order_id,
                channel=None, payment=None, entity=None, vff_source_text=None, vff_name_text=None):
    if date_ is None:
        return
    if brand == 'EXCLUDE':
        return
    is_shoe, gender = (False, None)
    size, color = None, None
    if brand == 'VFF':
        is_shoe, gender = vff_shoe_gender(vff_source_text)
        if is_shoe:
            size, color = vff_shoe_size_color(vff_name_text or vff_source_text)
    records.append(dict(
        store=store, category=category, date=date_.isoformat(),
        month=date_.strftime('%Y-%m'), brand=brand or 'Unknown', model=model,
        is_vff_shoe=is_shoe, gender=gender, size=size, color=color,
        sub=sub, qty=qty, amount=amount, order_id=order_id,
        channel=channel, payment=payment, entity=entity,
    ))

# ================================================================== brand-prefix numeric-code -> model lookup
# Central_Total_Department uses short codes like "VFF08"/"BFJ01"/"OLN05" instead of the
# "VFF0008"/"BFJ0001"/"OLN0005"-style codes used elsewhere. Build a (prefix, normalized-numeric)
# -> model-name table from the well-labeled files (product name is present) so short-code files
# can be backfilled, for every brand prefix -- not just VFF.
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
# 2026-09: cf220ee8-BFT_Shopee_Lazada_Facebook_JanJun_26.xlsx replaced by
# 34e59155-BFT_Shopee_Lazada_Facebook_Jan-Jun_26.xlsx (same 'Orders' sheet
# schema; see load_online_receipts() below for why the *loader* itself
# switched to a different, newer file instead of this one).
_scan_codes(SRC + '34e59155-BFT_Shopee_Lazada_Facebook_Jan-Jun_26.xlsx', 'Orders', 1, 'Product code', 'Product name')
_scan_codes(SRC + '2932d908-Sales_Paradies_Park_JanJun_26.xlsx', 'Orders', 1, 'Product code', 'Product name')
print(f"Built brand code->model lookup with {len(CODE_TO_MODEL)} entries")

# ================================================================== color-code -> full color name lookup
# Same source files as CODE_TO_MODEL above, since they're the ones with both
# a product code (abbreviated color, e.g. "TT/BK") and a product name (full
# color name, e.g. "Total Black") on the same row -- feeds
# vff_shoe_size_color()'s code-style branch for sources that only carry the
# abbreviated code (Siam Discovery, Central_Total_Department).
def _scan_colors(fn, sheet, header_row_idx, code_key, name_key):
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
        if not code or not name:
            continue
        cm = re.search(r'\(([^()]*)\)\s*$', str(code))
        nm = re.search(r'\(([^()]*)\)\s*$', str(name))
        if not cm or not nm:
            continue
        c_parts = [p.strip() for p in cm.group(1).split(',')]
        n_parts = [p.strip() for p in nm.group(1).split(',')]
        if len(c_parts) < 2 or len(n_parts) < 2:
            continue
        if not VFF_SIZE_TOKEN_RE.match(c_parts[-1]) or not VFF_SIZE_TOKEN_RE.match(n_parts[0]):
            continue  # not a shoe row (code-style size not last, or name-style size not first)
        color_code = ','.join(c_parts[:-1]).strip().upper()
        color_name = ','.join(n_parts[1:]).strip()
        if not (color_code and color_name):
            continue
        if color_code not in COLOR_CODE_TO_NAME:
            COLOR_CODE_TO_NAME[color_code] = color_name
        # Central_Total_Department uses the same color codes without the '/'
        # separator (e.g. "TTBK" instead of "TT/BK") -- register that variant
        # too so its code-only rows resolve to the same full name.
        stripped = color_code.replace('/', '')
        if stripped and stripped not in COLOR_CODE_TO_NAME:
            COLOR_CODE_TO_NAME[stripped] = color_name

_scan_colors(SRC + 'a0d9bc79-Sales_K_village_JanJun_26.xlsx', 'Orders', 1, 'Product code', 'Product name')
_scan_colors(SRC + '835c1952-Sales_Thaniya_JanJun_26.xlsx', 'Orders', 1, 'Product code', 'Product name')
_scan_colors(SRC + '34e59155-BFT_Shopee_Lazada_Facebook_Jan-Jun_26.xlsx', 'Orders', 1, 'Product code', 'Product name')
_scan_colors(SRC + '2932d908-Sales_Paradies_Park_JanJun_26.xlsx', 'Orders', 1, 'Product code', 'Product name')
print(f"Built color code->name lookup with {len(COLOR_CODE_TO_NAME)} entries")

# Casing canonicalizer: some sources spell color names in ALL CAPS (Central
# Total Department's July 'SKU Name' column, e.g. "BLACK") while others use
# proper case (K Village etc., "Black") for the exact same color -- without
# this, the two would fragment into separate breakdown entries. Built from
# COLOR_CODE_TO_NAME's own values (the properly-cased names harvested above),
# applied to every vff_shoe_size_color() result regardless of which branch
# produced it.
COLOR_NAME_CANON = {}
for _name in COLOR_CODE_TO_NAME.values():
    COLOR_NAME_CANON.setdefault(_name.upper(), _name)

# ================================================================== 1. VFF Cart LP
# Superseded 2026-08: the original 'fdc407d1-Sale_VFF_Cart_LP_01062026.csv'
# (25 rows, VFF V-Soul only) was confirmed against the store's own payment
# ledger to be a massively incomplete extract -- March-June totals were
# understated by 44,000-70,000 THB/month. Replaced with the store's full
# 'Orders'-schema export (same schema as load_orders_style/load_flat_branch_
# file: Payment amount is the order's actual paid total, present once on the
# order's first line; prorate each line to its share of it). Scoped to
# March-June only -- July is already loaded via etl_jul2026.py's EDV
# multi-store file (confirmed identical July total: 85,866.60).
def load_vff_cart_lp():
    fn = SRC + 'f578885d-EDV_Cart_Central_LP_01072026.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['Orders']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    idx = {h: i for i, h in enumerate(header) if h}
    data = rows[1:]
    data = [r for r in data if any(r) and r[idx.get('Product code')]]

    order_line_total = defaultdict(float)
    order_paid, order_discount = {}, {}
    for r in data:
        oid = r[idx.get('Sales order No.')]
        order_line_total[oid] += num(r[idx.get('Total amount')])
        pay = r[idx.get('Payment amount')]
        if pay not in (None, ''):
            order_paid[oid] = num(pay)
        disc = r[idx.get('Discount')]
        if disc not in (None, '') and oid not in order_discount:
            order_discount[oid] = disc

    order_resolved = {}
    n, skipped_july = 0, 0
    for r in data:
        d = to_date(r[idx.get('Date')])
        if d is None:
            continue
        if d >= date(2026, 7, 1):
            skipped_july += 1
            continue  # July already covered by etl_jul2026.py
        pc = r[idx.get('Product code')]
        order_id = r[idx.get('Sales order No.')]
        line_amt = num(r[idx.get('Total amount')])
        order_total = order_line_total[order_id]
        if order_id not in order_resolved:
            order_resolved[order_id] = resolve_order_amount(
                order_total, order_paid.get(order_id), order_discount.get(order_id))
        resolved = order_resolved[order_id]
        amt = (line_amt / order_total * resolved) if order_total else 0.0
        qty = num(r[idx.get('Quantity')])
        brand, sub = classify_by_code(pc)
        pname = r[idx.get('Product name')]
        model = model_from_name(pname)
        add_record('VFF Cart LP', 'store', d, brand, model, sub, qty, amt, order_id, vff_source_text=pc, vff_name_text=pname)
        n += 1
    print(f"loaded {n} rows -> VFF Cart LP (Mar-Jun 2026; {skipped_july} July rows skipped, already covered)")
# Disabled 2026-08 per user request: VFF Cart LP is an Endeavors-operated
# corner, now represented instead by the consolidated 対EDV wholesale-invoice
# loader below (load_edv_invoice_report()) -- kept here, not deleted, in case
# this is reverted ("一旦" -- for now).
# load_vff_cart_lp()

# ================================================================== generic "Orders" schema loader
def load_orders_style(fn, sheet, store_label, category, header_row_idx=1,
                       has_channel=False, has_payment_channel=False,
                       is_online_split=False, entity=None):
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[header_row_idx]
    idx = {h: i for i, h in enumerate(header) if h}
    data = rows[header_row_idx + 1:]
    data = [r for r in data if any(r) and r[idx.get('Product code')]]

    # order-level fields (Sales channel / Payment channel / Payment amount) are
    # only populated on an order's first line in these exports; forward-fill
    # them per order_id first. 'Total amount' is each line's PRE-discount
    # total -- the order's actual collected revenue is 'Payment amount'.
    # Confirmed 2026-08 against store-level payment-method ledgers (Central
    # LP 3F, Thaniya, K Village): summing 'Total amount' directly overstates
    # revenue by the order-level discount, so each line is prorated to its
    # share of the order's actual amount paid. Orders with no 'Payment
    # amount' anywhere fall back to their own line-total sum unprorated.
    order_channel, order_payment_channel = {}, {}
    order_line_total = defaultdict(float)
    order_paid, order_discount = {}, {}
    for r in data:
        oid = r[idx.get('Sales order No.')]
        order_line_total[oid] += num(r[idx.get('Total amount')])
        if 'Payment amount' in idx:
            pay = r[idx['Payment amount']]
            if pay not in (None, ''):
                order_paid[oid] = num(pay)
        if 'Discount' in idx:
            disc = r[idx['Discount']]
            if disc not in (None, '') and oid not in order_discount:
                order_discount[oid] = disc
        if not oid:
            continue
        if has_channel and 'Sales channel' in idx and r[idx['Sales channel']] and oid not in order_channel:
            order_channel[oid] = r[idx['Sales channel']]
        if has_payment_channel and 'Payment channel' in idx and r[idx['Payment channel']] and oid not in order_payment_channel:
            order_payment_channel[oid] = r[idx['Payment channel']]

    order_resolved = {}

    n = 0
    for r in data:
        pc = r[idx.get('Product code')]
        d = to_date(r[idx.get('Date')])
        if d is None:
            continue
        qty = num(r[idx.get('Quantity')])
        order_id = r[idx.get('Sales order No.')]
        line_amt = num(r[idx.get('Total amount')])
        order_total = order_line_total[order_id]
        if order_id not in order_resolved:
            order_resolved[order_id] = resolve_order_amount(
                order_total, order_paid.get(order_id), order_discount.get(order_id))
        resolved = order_resolved[order_id]
        amt = (line_amt / order_total * resolved) if order_total else 0.0
        brand, sub = classify_by_code(pc)
        pname = r[idx.get('Product name')]
        model = model_from_name(pname)
        channel = normalize_channel(order_channel.get(order_id)) if has_channel else None
        payment_raw = order_payment_channel.get(order_id) if has_payment_channel else None
        payment = normalize_payment(payment_raw)
        store = store_label
        cat = category
        if is_online_split:
            store = 'Online'
            cat = 'online'
        add_record(store, cat, d, brand, model, sub, qty, amt, order_id, channel=channel,
                    payment=payment, entity=entity, vff_source_text=pc, vff_name_text=pname)
        n += 1
    overridden = sum(1 for oid, p in order_paid.items()
                      if abs(order_resolved.get(oid, p) - p) > 0.01)
    print(f"loaded {n} rows from {fn.split('/')[-1]} -> {store_label} "
          f"({len(order_paid)}/{len(order_line_total)} orders prorated to actual amount paid, "
          f"{overridden} Payment-amount overridden as unreliable)")

# 2. BFT online -- 2026-09: replaced by load_online_receipts() below, which
# reads a different, newer, more comprehensive source (see that function's
# docstring). The old cf220ee8 loader call and the load_bft_merged_new_only()
# dedup-supplement function it depended on are disabled together.
# load_orders_style(SRC + 'cf220ee8-BFT_Shopee_Lazada_Facebook_JanJun_26.xlsx', 'Orders',
#                    'Online', 'online', has_channel=True, is_online_split=True, entity='BFT')
# 3. BFT_EVENT_3
load_orders_style(SRC + '2514cf15-BFT_EVENT_3.xlsx', 'Orders (2)', 'Event', 'event', has_channel=True)
# 4. Paradise Park
load_orders_style(SRC + '2932d908-Sales_Paradies_Park_JanJun_26.xlsx', 'Orders',
                   'Paradise Park', 'store', has_channel=True)
# 5. Thaniya -- disabled 2026-08, see load_vff_cart_lp() note above.
# load_orders_style(SRC + '835c1952-Sales_Thaniya_JanJun_26.xlsx', 'Orders', 'Thaniya', 'store')
# 6. K Village -- disabled 2026-08, see load_vff_cart_lp() note above.
# load_orders_style(SRC + 'a0d9bc79-Sales_K_village_JanJun_26.xlsx', 'Orders', 'K Village', 'store')
# 7-9. Event files with payment channel (47-col schema); reuse loader (Product code / Date / etc. keys match).
# EDV_EVENT_1 (Endeavors-run event) disabled 2026-08, see load_vff_cart_lp()
# note above; BFT_EVENT_2/1 (Barefoot's own events) are kept.
# load_orders_style(SRC + '9c5663cd-EDV_EVENT_1.xlsx', 'Orders', 'Event', 'event',
#                    has_channel=True, has_payment_channel=True)
# 2026-09: 7abc0f65-BFT_EVENT_2.xlsx + 0165b670-BFT_EVENT_1.xlsx replaced by
# 95bfee20-BFT_EVENT_1.xlsx -- per user confirmation, this is a corrected
# version of the original EVENT_1 file (the original no longer exists to
# compare against), and despite its name it already carries BOTH Event 1 and
# Event 2 warehouse rows (Warehouse/Branch: 419 'Event 2' + 345 'Event 1' of
# 778 total), so one loader call replaces both old ones. Confirmed via a
# separately re-supplied e3662139-BFT_EVENT_2.xlsx: every one of its 300
# order numbers already appears in 95bfee20, so it adds nothing and is
# deliberately not loaded here (would only risk double-counting).
# load_orders_style(SRC + '7abc0f65-BFT_EVENT_2.xlsx', 'Orders', 'Event', 'event',
#                    has_channel=True, has_payment_channel=True)
# load_orders_style(SRC + '0165b670-BFT_EVENT_1.xlsx', 'Orders', 'Event', 'event',
#                    has_channel=True, has_payment_channel=True)
def load_event_2026_09():
    """95bfee20-BFT_EVENT_1.xlsx has an explicit column S, 'Payment amount
    (excl. VAT)', on each order's first line -- per user confirmation
    2026-09, this (not load_orders_style()'s usual resolve_order_amount()
    heuristic over 'Payment amount'/Discount) is each order's authoritative
    amount. Verified: matches the "Jan-Jun_final_without_tax.xlsx" BFT-tab
    Event-column reference exactly, to the cent, for every one of the 6
    months. Column S is blank only for voided/no-payment orders (verified:
    22 of 568 orders, nearly all Status='Voided') -- correctly excluded.
    A handful of orders (5, all in May) have every line's 'Total amount'
    blank too, so there's no per-line weight to prorate the order's S-value
    by; those fall back to a quantity-weighted split instead (excluding
    non-product DC/CON/P lines from that quantity, so e.g. a same-order
    discount line doesn't silently absorb -- and lose -- half the revenue).
    """
    fn = SRC + '95bfee20-BFT_EVENT_1.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['Orders']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[1]
    idx = {h: i for i, h in enumerate(header) if h}
    data = rows[2:]
    data = [r for r in data if any(r) and r[idx.get('Product code')]]

    order_channel = {}
    order_line_total = defaultdict(float)
    order_line_qty = defaultdict(float)  # excludes non-product (DC/CON/P) lines
    order_s = {}
    for r in data:
        oid = r[idx.get('Sales order No.')]
        order_line_total[oid] += num(r[idx.get('Total amount')])
        b_check, _ = classify_by_code(r[idx.get('Product code')])
        if b_check != 'EXCLUDE':
            order_line_qty[oid] += num(r[idx.get('Quantity')])
        s_val = r[idx['Payment amount (excl. VAT)']]
        if s_val not in (None, '') and oid not in order_s:
            order_s[oid] = num(s_val)
        if oid and 'Sales channel' in idx and r[idx['Sales channel']] and oid not in order_channel:
            order_channel[oid] = r[idx['Sales channel']]

    n, skipped = 0, 0
    for r in data:
        pc = r[idx.get('Product code')]
        d = to_date(r[idx.get('Date')])
        if d is None:
            continue
        order_id = r[idx.get('Sales order No.')]
        if order_id not in order_s:
            skipped += 1
            continue  # voided/no-payment order
        qty = num(r[idx.get('Quantity')])
        line_amt = num(r[idx.get('Total amount')])
        order_total = order_line_total[order_id]
        resolved_raw = round(order_s[order_id] * 1.07, 2)  # excl.-VAT -> raw; ser() re-applies /1.07 downstream
        if order_total:
            amt = line_amt / order_total * resolved_raw
        else:
            order_qty = order_line_qty[order_id]
            amt = (qty / order_qty * resolved_raw) if order_qty else 0.0
        brand, sub = classify_by_code(pc)
        if brand == 'EXCLUDE':
            continue
        pname = r[idx.get('Product name')]
        model = model_from_name(pname)
        channel = normalize_channel(order_channel.get(order_id))
        add_record('Event', 'event', d, brand, model, sub, qty, amt, order_id, channel=channel,
                    vff_source_text=pc, vff_name_text=pname)
        n += 1
    print(f"loaded {n} rows -> Event ({skipped} voided/no-payment rows skipped)")
load_event_2026_09()

# Yoshi Run -- a Barefoot-run event with no product sales, just a lump-sum
# participation/booth fee (per user 2026-09): 19,626.28 THB (Jan), 9,813.14
# THB (May), matching the "Jan-Jun_final_without_tax.xlsx" BFT-tab reference
# exactly. Those figures are already VAT-exclusive per that file's own
# convention, so converted back to VAT-inclusive raw amounts here since
# every other add_record() amount in this pipeline is raw/VAT-inclusive
# (ser() applies the /1.07 conversion once, downstream, for everyone).
add_record('Event', 'event', date(2026, 1, 1), None, None, None, 0, round(19626.28 * 1.07, 2), None)
add_record('Event', 'event', date(2026, 5, 1), None, None, None, 0, round(9813.14 * 1.07, 2), None)

# ================================================================== Central LP 3F (Thai headers)
# 'ราคารวม' (line total) is the PRE-discount list price of each line; the
# order's actual collected revenue is 'จำนวนเงินที่ชำระ' (amount paid), which
# only appears once, on the first product line of each order ('รายการ'/order
# ID repeats across an order's lines). Confirmed 2026-08 against the store's
# own payment-method ledger (Cash/Credit Card/QR by day): summing 'ราคารวม'
# directly overstates revenue by the order-level discount ('ส่วนลด'), so each
# line is prorated to its share of the order's actual amount paid -- same
# prorate-by-share pattern as the 2025 BFT consignment loader.
def load_central_lp3f():
    fn = SRC + '4106a54d-Sales_Central_LP_3F_JanJun_26.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['รายการขาย']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[1]
    idx = {h: i for i, h in enumerate(header) if h}
    data = rows[2:]
    data = [r for r in data if any(r) and r[idx['รหัสสินค้า']]]

    order_line_total = defaultdict(float)  # order_id -> sum('ราคารวม') across its lines
    order_paid, order_discount = {}, {}    # order_id -> 'จำนวนเงินที่ชำระ' / 'ส่วนลด' (first line only)
    for r in data:
        order_id = r[idx['รายการ']]
        order_line_total[order_id] += num(r[idx['ราคารวม']])
        paid = r[idx['จำนวนเงินที่ชำระ']]
        if paid not in (None, ''):
            order_paid[order_id] = num(paid)
        disc = r[idx['ส่วนลด']]
        if disc not in (None, '') and order_id not in order_discount:
            order_discount[order_id] = disc

    order_resolved = {}
    n = 0
    for r in data:
        d = to_date(r[idx['วันที่ทำรายการ']])
        if d is None:
            continue
        pc = r[idx['รหัสสินค้า']]
        order_id = r[idx['รายการ']]
        line_amt = num(r[idx['ราคารวม']])
        order_total = order_line_total[order_id]
        if order_id not in order_resolved:
            order_resolved[order_id] = resolve_order_amount(
                order_total, order_paid.get(order_id), order_discount.get(order_id))
        resolved = order_resolved[order_id]
        amt = (line_amt / order_total * resolved) if order_total else 0.0
        qty = num(r[idx['จำนวน']])
        brand, sub = classify_by_code(pc)
        pname = r[idx['ชื่อสินค้า']]
        model = model_from_name(pname)
        add_record('Central Ladprao 3F (Coollabo)', 'store', d, brand, model, sub, qty, amt, order_id, vff_source_text=pc, vff_name_text=pname)
        n += 1
    overridden = sum(1 for oid, p in order_paid.items()
                      if abs(order_resolved.get(oid, p) - p) > 0.01)
    print(f"loaded {n} rows -> Central Ladprao 3F "
          f"(prorated {len(order_paid)}/{len(order_line_total)} orders to actual amount paid, "
          f"{overridden} Payment-amount overridden as unreliable)")
# Disabled 2026-08 -- see load_vff_cart_lp() note above; same reasoning.
# load_central_lp3f()

# ================================================================== BFT consignment (with trap-row guard)
def load_consignment(fn, sheet, store_label):
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[1]
    idx = {h: i for i, h in enumerate(header) if h}
    data = rows[2:]
    n = 0
    for r in data:
        if not any(r):
            continue
        doc = r[idx['เลขที่เอกสาร']]
        pc = r[idx['รหัสสินค้า/บริการ']]
        if not doc or not pc:
            continue  # guards against the trailing "รวม" total row (blank doc & product code)
        d = parse_dmy(r[idx['วันที่ออก']])
        if d is None:
            continue
        qty = num(r[idx['Quantity']])
        # 'Amount' = invoice-level net total, present only on first product row of each invoice
        amt_col = r[idx['Amount']]
        amt = num(amt_col) if amt_col not in (None, '') else 0.0
        brand, sub = classify_by_code(pc)
        pname = r[idx['ชื่อสินค้า/บริการ']]
        model = model_from_name(pname)
        add_record(store_label, 'consignment', d, brand, model, sub, qty, amt, doc, vff_source_text=pc, vff_name_text=pname)
        n += 1
    print(f"loaded {n} rows -> {store_label}")
# Note: 'BFT consignment...' (Thai schema, above) covers Jan-Jun. July-August
# comes from a different, English-schema "Invoice Report" export -- see
# load_bft_consignment_jul2026() in etl_jul2026.py (July) and
# load_bft_consignment_aug2026() below (August).
load_consignment(SRC + '2ded23fc-BFT_consignment__JanJun_26_new.xlsx', 'รายงานใบแจ้งหนี้', 'BFT Consignment')

# 2026-09: August-only follow-up (33610632-BFT_Consignment_Jul-Aug_26.xlsx)
# -- July is already covered by etl_jul2026.py's load_bft_consignment_jul2026()
# (confirmed: this file's July total, 86,420.51 THB net, matches the already-
# loaded figure to a few cents -- same underlying data, just a different
# monthly export). Per user instruction: each order's per-product amount is
# column I ('Pre-VAT Amount'), scaled by the order-level discount/fee
# percentage in column K ('Total Discount', stored as e.g. '20.00%' -- only
# present on an order's first line, forward-filled to its other lines here).
# Column L ('VAT 7%', despite the header -- actually the order-level total
# AFTER that discount, still pre-VAT) is the reference total this reproduces:
# verified line_amount = I * (1 - K%) sums to L exactly for a multi-line
# order. Mathematically equivalent to load_bft_consignment_jul2026()'s
# line/total-ratio proration (same result, different derivation) -- this
# loader follows the user's stated column-based method directly instead.
def load_bft_consignment_aug2026():
    fn = SRC + '33610632-BFT_Consignment_Jul-Aug_26.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['Invoice Report']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]

    n = 0
    cur_doc, cur_pct = None, None
    for r in data:
        if not r or not any(r):
            continue
        doc, pc, pname = r[0], r[3], r[4]
        k_val = r[10]
        if doc:
            cur_doc = doc
            if k_val not in (None, ''):
                cur_pct = float(str(k_val).replace('%', '')) / 100.0
        if not cur_doc or not pc:
            continue
        d = parse_dmy(r[1])
        if d is None or not (d.year == 2026 and d.month == 8):
            continue  # July skipped -- already loaded, see note above
        qty = num(r[5])
        i_val = r[8]
        if i_val is None:
            continue
        amt_net = num(i_val) * (1 - (cur_pct or 0.0))
        amt = amt_net * 1.07
        brand, sub = classify_by_code(pc)
        model = model_from_name(pname)
        add_record('BFT Consignment', 'consignment', d, brand, model, sub, qty, amt, cur_doc,
                    vff_source_text=pc, vff_name_text=pname)
        n += 1
    print(f"loaded {n} rows -> BFT Consignment, August only (July already covered by etl_jul2026.py)")
load_bft_consignment_aug2026()

# ================================================================== EDV consignment (no header row, positional)
def load_edv_consignment():
    fn = SRC + '21cdb27d-EDV_consignment__JanJun_26.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['รายการขาย']
    data = list(ws.iter_rows(values_only=True))
    # positional indices validated against EDV_EVENT_1 schema in the exploration phase
    COL = {'order': 2, 'date': 12, 'amount': 29, 'prodcode': 39, 'prodname': 40, 'qty': 41, 'total': 44}
    n = 0
    for r in data:
        if not any(r):
            continue
        order = r[COL['order']]
        pc = r[COL['prodcode']]
        if not order or not pc:
            continue
        d = parse_dmy(r[COL['date']])
        if d is None:
            continue
        qty = num(r[COL['qty']])
        amt_col = r[COL['amount']]
        amt = num(amt_col) if amt_col not in (None, '') else 0.0
        brand, sub = classify_by_code(pc)
        pname = r[COL['prodname']]
        model = model_from_name(pname)
        add_record('EDV Consignment', 'consignment', d, brand, model, sub, qty, amt, order, vff_source_text=pc, vff_name_text=pname)
        n += 1
    print(f"loaded {n} rows -> EDV Consignment")
# Disabled 2026-08 -- see load_vff_cart_lp() note above; same reasoning.
# load_edv_consignment()

# ================================================================== Siam Discovery (name-based classification)
# Columns: Sale Date(0), Item(1), Sales Quantity(2), Net Sales (Sales
# Amount)(3), Price(4), GP 32%(5). Per user confirmation 2026-08, Siam
# Discovery's recognized revenue is net of their 32% consignment fee -- GP
# 32% (always exactly Price*0.68) is the correct "amount", not the gross
# Net Sales (Sales Amount) column. Matches the convention used for every
# other year's Siam Discovery batch (etl_2025.py/etl_2024.py/
# etl_jul2026.py), which already read the equivalent GP32% column.
# 2026-09: swapped to fa90b27f-Sales_Siam_Dis_Jan-Jun_26.xlsx (replacing
# d7e7cac8-Sales_Siam_Dis_JanJun_26.xlsx) -- per-user reconciliation against
# the "Jan-Jun_final_without_tax.xlsx" BFT-tab reference showed a handful of
# transactions dated a month later/earlier than in the old file (net H1 total
# barely moves, 977,522.25->977,522.24, but Feb/Mar and Apr/May each shift by
# ~3,050 THB); this new file's GP32%/1.07 monthly totals now match the BFT
# tab's Siam Discovery column exactly, to the cent, for every one of the 6
# months.
def load_siam_discovery():
    fn = SRC + 'fa90b27f-Sales_Siam_Dis_Jan-Jun_26.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['Sheet1']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]
    n = 0
    for i, r in enumerate(data):
        if not r or not r[1]:
            continue  # blank Item = zero-sales placeholder day
        d = to_date(r[0])
        item = r[1]
        qty = num(r[2])
        net = num(r[5])
        brand, sub, model = classify_by_name(item)
        add_record('Siam Discovery', 'store', d, brand, model, sub,
                    qty, net, f'SIAMDIS-{i}', vff_source_text=item, vff_name_text=item)
        n += 1
    print(f"loaded {n} rows -> Siam Discovery")
load_siam_discovery()

# 2026-09: August-only follow-up (ed45786c-BFT_Siam_Discovery_Jul-Aug_26.xlsx)
# -- July is already covered by etl_jul2026.py's load_siam_discovery_jul2026()
# (confirmed: this file's July total, 257,986.91 THB, matches the already-
# loaded figure to the cent -- same underlying data, just re-exported with
# an extra 'Gross Sales' column inserted). Same "Item"/SKU-Name-based
# classify_by_name() classification as both other Siam Discovery loaders --
# per user question 2026-09, yes, model/color/gender ARE determinable from
# this data, the same way as always for Siam Discovery (the SKU Name column
# carries the same 'MODEL (ColorCode,SizeToken)' shape, e.g.
# 'V-SOUL (BK,W39)'). Amount is column G ('Before Vat', already VAT-
# exclusive per user instruction) -- multiplied by 1.07 here so add_record's
# downstream VAT division reproduces it exactly, same treatment as every
# other VAT-exclusive source column in this file.
def load_siam_discovery_aug2026():
    fn = SRC + 'ed45786c-BFT_Siam_Discovery_Jul-Aug_26.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['Sheet1']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]
    n = 0
    for i, r in enumerate(data):
        if not r or not r[1]:
            continue
        d = to_date(r[0])
        if d is None or not (d.year == 2026 and d.month == 8):
            continue  # July skipped -- already loaded, see note above
        item = r[1]
        qty = num(r[2])
        before_vat = num(r[6])
        amt = before_vat * 1.07
        brand, sub, model = classify_by_name(item)
        add_record('Siam Discovery', 'store', d, brand, model, sub,
                    qty, amt, f'SIAMDISAUG-{i}', vff_source_text=item, vff_name_text=item)
        n += 1
    print(f"loaded {n} rows -> Siam Discovery, August only (July already covered by etl_jul2026.py)")
load_siam_discovery_aug2026()

# ================================================================== EDV online (Shopee/Lazada/etc.)
def load_simple_online(fn, store_label, category, entity=None):
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['รายการขาย']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    idx = {h: i for i, h in enumerate(header) if h}
    data = rows[1:]
    n = 0
    for r in data:
        if not any(r):
            continue
        sku = r[idx['SKU']]
        if not sku:
            continue
        d = to_date(r[idx['Sale Date']])
        if d is None:
            continue
        qty = num(r[idx['Sales Quantity']])
        amt = num(r[idx['Net']])
        order_id = r[idx['No.']]
        channel = normalize_channel(r[idx['Channel']])
        brand, sub = classify_by_code(sku)
        item = r[idx['Item']]
        model = model_from_name(item)
        add_record(store_label, category, d, brand, model, sub, qty, amt, order_id, channel=channel,
                    entity=entity, vff_source_text=sku, vff_name_text=item)
        n += 1
    print(f"loaded {n} rows -> {store_label}")
    return set(str(r[idx['No.']]).strip() for r in data if any(r) and r[idx.get('No.')])

# Disabled 2026-08 -- see load_vff_cart_lp() note above; same reasoning.
# edv_online_orders = load_simple_online(SRC + 'a024894a-EDV_Shopee_Lazada_Facebook_JanJun_26.xlsx',
#                                         'Online', 'online', entity='EDV')

# ================================================================== BFT Online: receipt-level export (replaces cf220ee8 + 6de07263)
# 2026-09: replaced the old cf220ee8-BFT_Shopee_Lazada_Facebook_JanJun_26.xlsx
# (+ its 6de07263 dedup-supplement) with a receipt-level export with an
# explicit before-VAT/VAT-amount column pair per line -- more reliable than
# reconstructing VAT-exclusive amounts via proration. First
# 4da610ac-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx (customer-name-based
# channel split: name starting with 'Shopee'/'Lazada' -> that marketplace,
# anything else -> 'Online'), then swapped to
# 20e89950-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx (near-identical content,
# per-user-confirmed as the "_new" successor) once user confirmation
# surfaced that the customer-name field is NOT a reliable channel indicator
# for every row: some named-individual/company orders are genuinely Shopee
# (buyer supplied their real name instead of the 'Shopee ... Customer'
# placeholder), which the pure name-prefix rule miscounted as 'Online'.
#
# All six months of 2026 have now been individually verified against the
# "Jan-Jun_final_without_tax.xlsx" BFT-tab reference, each against its own
# re-uploaded "..._new.xlsx" file (a different upload per month, despite the
# identical filename -- see SRC hash prefixes below), each requiring
# targeted, month-specific row inspection since the customer-name field
# alone is NOT a reliable channel indicator: some named-individual/company
# orders are genuinely Shopee/Lazada (buyer supplied their real name instead
# of the 'Shopee ... Customer'/'Lazada Customer' placeholder), which the
# pure name-prefix rule below misclassifies as 'Online'.
#
# January 2026: in sheet '(3)' of the 20e89950 file, January's 386 rows are
# laid out as a clean block -- rows 1-25 (data idx 0-24) = Lazada, rows
# 26-366 (idx 25-365) = Shopee (including those named-individual orders),
# rows 367-386 (idx 366-385) = Online (genuinely anonymous/direct/staff).
# Verified to the cent (Lazada 97,084.11 exact; Shopee/Online within a few
# cents, pure rounding). Handled inline below (the only month still using
# load_online_receipts() directly); February-June all have their own
# load_online_receipts_<month>() function further down.
#
# February 2026 (bc054fa0-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx, loaded
# in load_online_receipts_february() below): sheet '(3)' is a clean 3-way
# block -- Lazada rows 388-431, Shopee 432-712, Online 713-743 -- except
# rows 715-721, a block of named-company orders that the user confirmed are
# genuinely Shopee ("715-721行目をShopeeに振り分けて"). Verified to the
# cent using column 15 (see June note below) -- Online/Shopee/Lazada all
# exact.
#
# March 2026 (26bdd1b3-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx, loaded in
# load_online_receipts_march() below): unlike January/May, March's block
# isn't a clean 3-way split -- it opens with a short anonymous/named-
# individual segment (rows 744-753) before the main Shopee block, then
# Lazada comes LAST (rows 996-1035), not first. Rows 744-747 ('ลูกค้า
# ไม่ประสงค์ออกนาม') and 752-753 ('Website Customer') follow the usual
# name-based Online rule; rows 748-751 needed user disambiguation (748 ->
# Lazada, 749-751 -> Shopee, despite all 4 being named companies/individuals
# with no name-based signal to tell them apart). Verified to the cent
# (Lazada 98,747.66 vs 98,747.67, Shopee 654,849.72 vs 654,849.53, Online
# 18,811.22 vs 18,811.21).
#
# April 2026 (5b81a74c-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx, loaded in
# load_online_receipts_april() below): unlike Jan/Mar/May, this file's sheet
# '(3)' contains ALL SIX months of data (not just April's), and within
# April's own date range the Shopee/Lazada/anonymous rows are already
# interleaved day-by-day rather than laid out in a clean per-channel block
# -- so no row-position override was needed; the plain customer-name rule
# reproduces the reference almost exactly on its own (Lazada 57,233.66 vs
# ref 57,233.65, Shopee 580,629.18 vs 580,628.97, Online 23,805.61 vs
# 23,805.61 -- all within a few cents, pure rounding).
#
# May 2026 (5379d845-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx, loaded in
# load_online_receipts_may() below): verified the same way as January --
# rows 1300-1324 (data idx 1298-1322) = Lazada, rows 1325-1691 (idx
# 1323-1689) = Shopee, rows 1692-1705 (idx 1690-1703) = Online. Verified to
# the cent (Online 42,002.80 exact; Shopee/Lazada within a few cents).
#
# June 2026 (c2321b6a-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx, loaded in
# load_online_receipts_june() below): sheet '(3)' is a clean 3-way block --
# Lazada rows 1706-1753, Shopee 1754-2169, Online 2170-2202 -- except row
# 2170 itself, a named-company order that the user confirmed is genuinely
# Shopee ("2170行目までがShopeeの売上です"). This file also revealed that
# column 15 ('VAT Amount', mislabeled -- it actually holds the line Total,
# before_vat+vat) is a more precise amount source than col8 Before Vat*1.07:
# the latter accumulates tens of THB of per-line rounding drift across a
# few hundred rows (e.g. ~56 THB for June alone), while col15 matches the
# reference to the cent. February's fix reuses column 15 for the same
# reason.
def load_online_receipts():
    fn = SRC + '20e89950-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['รายงานใบเสร็จรับเงิน (3)']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]  # single header row (unlike sheets '' / '(2)', which have a banner row above the header too)

    def channel_for(customer):
        if not customer:
            return 'Online'
        c = str(customer)
        if c.startswith('Shopee'):
            return 'Shopee'
        if c.startswith('Lazada'):
            return 'Lazada'
        return 'Online'

    n = 0
    for i, r in enumerate(data):
        if not r or not r[0]:
            continue
        d = to_date(r[1])
        if d is None:
            continue
        if d.year == 2026 and d.month in (2, 3, 4, 5, 6):
            continue  # February-June are loaded from separate, dedicated files -- see load_online_receipts_february()/_march()/_april()/_may()/_june() below
        if d.year == 2026 and d.month == 1:
            channel = 'Lazada' if i <= 24 else ('Shopee' if i <= 365 else 'Online')
        else:
            channel = channel_for(r[3])
        product_code, product_name, desc = r[4], r[5], r[6]
        qty = num(r[8])
        before_vat = num(r[7])
        amt = before_vat * 1.07  # VAT-exclusive column; convert to raw so ser() applies VAT once, downstream, like every other source
        brand, sub = classify_by_code(product_code)
        if brand == 'EXCLUDE':
            continue
        model = model_from_name(product_name) if product_name else ('Shipping Fee' if desc == 'Shipping fee' else 'Other')
        add_record('Online', 'online', d, brand, model, sub, qty, amt, r[0], channel=channel,
                    vff_source_text=product_code, vff_name_text=product_name)
        n += 1
    print(f"loaded {n} rows -> Online (Shopee/Lazada/Online, receipt-level; January uses the verified row-position split, Feb-Jun excluded -- see load_online_receipts_february()/_march()/_april()/_may()/_june())")
load_online_receipts()

def load_online_receipts_february():
    fn = SRC + 'bc054fa0-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['รายงานใบเสร็จรับเงิน (3)']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]

    def channel_for(customer, excel_row):
        if 715 <= excel_row <= 721:
            return 'Shopee'
        if not customer:
            return 'Online'
        c = str(customer)
        if c.startswith('Shopee'):
            return 'Shopee'
        if c.startswith('Lazada'):
            return 'Lazada'
        return 'Online'

    n = 0
    for i, r in enumerate(data):
        if not r or not r[0]:
            continue
        d = to_date(r[1])
        if d is None or not (d.year == 2026 and d.month == 2):
            continue
        channel = channel_for(r[3], i + 2)
        product_code, product_name, desc = r[4], r[5], r[6]
        qty = num(r[8])
        amt = num(r[14])  # column 15 'VAT Amount' -- actually the raw Total (VAT-inclusive); no *1.07 needed
        brand, sub = classify_by_code(product_code)
        if brand == 'EXCLUDE':
            continue
        model = model_from_name(product_name) if product_name else ('Shipping Fee' if desc == 'Shipping fee' else 'Other')
        add_record('Online', 'online', d, brand, model, sub, qty, amt, r[0], channel=channel,
                    vff_source_text=product_code, vff_name_text=product_name)
        n += 1
    print(f"loaded {n} rows -> Online, February only (verified row-position split, col15 Total)")
load_online_receipts_february()

# March 2026: verified against a third re-uploaded file (26bdd1b3-BFT_Online_
# Shopee_Lazada_Jan-Jun_new.xlsx). Unlike January/May, March's block isn't a
# clean 3-way split -- it opens with a short anonymous/named-individual
# segment (rows 744-753) before the main Shopee block, then Lazada comes
# LAST (rows 996-1035), not first. Rows 744-747 ('ลูกค้า ไม่ประสงค์ออกนาม')
# and 752-753 ('Website Customer') follow the usual name-based Online rule;
# rows 748-751 needed user disambiguation (748 -> Lazada, 749-751 -> Shopee,
# despite all 4 being named companies/individuals with no name-based signal
# to tell them apart) since it's not a name pattern but a one-off manual
# call. Verified: matches the reference to the cent (Lazada 98,747.66 vs
# 98,747.67, Shopee 654,849.72 vs 654,849.53, Online 18,811.22 vs 18,811.21).
def load_online_receipts_march():
    fn = SRC + '26bdd1b3-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['รายงานใบเสร็จรับเงิน (3)']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]

    def channel_for(customer, excel_row):
        if 748 <= excel_row <= 751:
            return 'Lazada' if excel_row == 748 else 'Shopee'
        if not customer:
            return 'Online'
        c = str(customer)
        if c.startswith('Lazada'):
            return 'Lazada'
        if c.startswith('Shopee'):
            return 'Shopee'
        return 'Online'

    n = 0
    for i, r in enumerate(data):
        if not r or not r[0]:
            continue
        d = to_date(r[1])
        if d is None or not (d.year == 2026 and d.month == 3):
            continue
        channel = channel_for(r[3], i + 2)
        product_code, product_name, desc = r[4], r[5], r[6]
        qty = num(r[8])
        before_vat = num(r[7])
        amt = before_vat * 1.07
        brand, sub = classify_by_code(product_code)
        if brand == 'EXCLUDE':
            continue
        model = model_from_name(product_name) if product_name else ('Shipping Fee' if desc == 'Shipping fee' else 'Other')
        add_record('Online', 'online', d, brand, model, sub, qty, amt, r[0], channel=channel,
                    vff_source_text=product_code, vff_name_text=product_name)
        n += 1
    print(f"loaded {n} rows -> Online, March only (verified row-position split)")
load_online_receipts_march()

def load_online_receipts_may():
    fn = SRC + '5379d845-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['รายงานใบเสร็จรับเงิน (3)']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]

    n = 0
    for i, r in enumerate(data):
        if not r or not r[0]:
            continue
        d = to_date(r[1])
        if d is None or not (d.year == 2026 and d.month == 5):
            continue
        channel = 'Lazada' if i <= 1322 else ('Shopee' if i <= 1689 else 'Online')
        product_code, product_name, desc = r[4], r[5], r[6]
        qty = num(r[8])
        before_vat = num(r[7])
        amt = before_vat * 1.07
        brand, sub = classify_by_code(product_code)
        if brand == 'EXCLUDE':
            continue
        model = model_from_name(product_name) if product_name else ('Shipping Fee' if desc == 'Shipping fee' else 'Other')
        add_record('Online', 'online', d, brand, model, sub, qty, amt, r[0], channel=channel,
                    vff_source_text=product_code, vff_name_text=product_name)
        n += 1
    print(f"loaded {n} rows -> Online, May only (verified row-position split)")
load_online_receipts_may()

def load_online_receipts_april():
    fn = SRC + '5b81a74c-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['รายงานใบเสร็จรับเงิน (3)']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]

    def channel_for(customer):
        if not customer:
            return 'Online'
        c = str(customer)
        if c.startswith('Shopee'):
            return 'Shopee'
        if c.startswith('Lazada'):
            return 'Lazada'
        return 'Online'

    n = 0
    for i, r in enumerate(data):
        if not r or not r[0]:
            continue
        d = to_date(r[1])
        if d is None or not (d.year == 2026 and d.month == 4):
            continue
        channel = channel_for(r[3])
        product_code, product_name, desc = r[4], r[5], r[6]
        qty = num(r[8])
        before_vat = num(r[7])
        amt = before_vat * 1.07
        brand, sub = classify_by_code(product_code)
        if brand == 'EXCLUDE':
            continue
        model = model_from_name(product_name) if product_name else ('Shipping Fee' if desc == 'Shipping fee' else 'Other')
        add_record('Online', 'online', d, brand, model, sub, qty, amt, r[0], channel=channel,
                    vff_source_text=product_code, vff_name_text=product_name)
        n += 1
    print(f"loaded {n} rows -> Online, April only (verified customer-name split)")
load_online_receipts_april()

# June 2026 was verified against a fifth re-uploaded file (c2321b6a-BFT_
# Online_Shopee_Lazada_Jan-Jun_new.xlsx). Sheet '(3)' is a clean 3-way block
# for June -- rows 1706-1753 = Lazada, rows 1754-2169 = Shopee, rows
# 2170-2202 = Online -- EXCEPT row 2170 itself, a named-company order
# ('...ริชมอนด์ ลักซ์ชัวรี...') that user confirmed is genuinely Shopee
# ("2170行目までがShopeeの売上です"), same root-cause pattern as January:
# a real customer name recorded instead of the 'Shopee ... Customer'
# placeholder. Also: unlike the other months, June uses column 15 ('VAT
# Amount', mislabeled -- it actually holds the line Total, before_vat+vat)
# as the amount source instead of col8 Before Vat*1.07 -- the latter
# accumulates ~56 THB of per-line rounding drift across June's ~500 rows,
# while col15 matches the BFT-tab reference to the cent (Lazada 182,448.60
# vs 182,448.59, Shopee/Online exact).
def load_online_receipts_june():
    fn = SRC + 'c2321b6a-BFT_Online_Shopee_Lazada_Jan-Jun_new.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['รายงานใบเสร็จรับเงิน (3)']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]

    def channel_for(customer, excel_row):
        if excel_row == 2170:
            return 'Shopee'
        if not customer:
            return 'Online'
        c = str(customer)
        if c.startswith('Shopee'):
            return 'Shopee'
        if c.startswith('Lazada'):
            return 'Lazada'
        return 'Online'

    n = 0
    for i, r in enumerate(data):
        if not r or not r[0]:
            continue
        d = to_date(r[1])
        if d is None or not (d.year == 2026 and d.month == 6):
            continue
        channel = channel_for(r[3], i + 2)
        product_code, product_name, desc = r[4], r[5], r[6]
        qty = num(r[8])
        amt = num(r[14])  # column 15 'VAT Amount' -- actually the raw Total (VAT-inclusive); no *1.07 needed
        brand, sub = classify_by_code(product_code)
        if brand == 'EXCLUDE':
            continue
        model = model_from_name(product_name) if product_name else ('Shipping Fee' if desc == 'Shipping fee' else 'Other')
        add_record('Online', 'online', d, brand, model, sub, qty, amt, r[0], channel=channel,
                    vff_source_text=product_code, vff_name_text=product_name)
        n += 1
    print(f"loaded {n} rows -> Online, June only (verified row-position split, col15 Total)")
load_online_receipts_june()

# ================================================================== BFT_Online_Event_Shopee_Lazada_Paradise_Jul-Aug_26 (first data beyond H1 2026)
# 2026-09: single "Orders"-schema file covering Event/Online(Shopee+Lazada+
# anonymous Online)/Paradise Park for July AND August 2026 -- the first
# H2-2026 data loaded (Paradise Park/Event previously had NO July/August
# data at all; Online's July/August were also missing before this). Row
# position (not the 'Sales channel'/'Warehouse-Branch' columns, which are
# inconsistent -- e.g. some Online rows show 'Paradise Park' or the Thai
# main-warehouse label in Warehouse/Branch) determines both store and,
# within Online, channel, per user-confirmed row ranges:
#   rows 2-196   -> Event
#   rows 198-250 -> Online / Lazada
#   rows 252-661 -> Online / Shopee
#   rows 663-771 -> Paradise Park
#   rows 773-828 -> Online / Online (genuinely anonymous/direct)
# Amount is column S ('Total amount', 19th column) -- a per-LINE total
# (unit price minus any unit discount, times quantity), VAT-inclusive,
# reconciles exactly to the order-level 'Amount' column when summed across
# an order's lines (verified on a sample multi-line order). Used directly as
# the raw amount; no *1.07 needed since it's already VAT-inclusive, unlike
# the VAT-exclusive 'Payment amount' columns used elsewhere.
#
# July was already represented in the dashboard for these 3 categories
# (apparently from an earlier/different version of this same underlying
# data -- Event's July total from this file, 414,045.79 THB, matches the
# already-loaded figure to the cent), so per user confirmation this file's
# July rows REPLACE that existing July data rather than adding to it, same
# as August (fully new). On a from-scratch full pipeline run this
# replace-vs-add distinction doesn't matter -- every row here simply becomes
# a record like any other; it only mattered for the direct-JSON-patch used
# to apply this fix in the running dashboard (see commit history: the
# patch's model/color/gender-level global breakdowns for July could only be
# approximated via revenue-proportional scaling of each store's OLD
# brand-level total, since no per-record archive of the prior July data
# survives to delta against precisely -- store-level and channel-level
# figures are exact; deeper model/color/gender splits for July carry a small
# margin of error as a result. This limitation does not apply to August,
# which is a pure addition with no prior data to reconcile against, nor
# to a future full pipeline run of this loader, which has no such gap.
def load_online_event_paradise_julaug2026():
    fn = SRC + '38a001d6-BFT_Online_Event_Shopee_Lazada_Paradise_Jul-Aug_26.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['Orders']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]

    BLOCKS = [
        (2, 196, 'Event', 'event', None),
        (198, 250, 'Online', 'online', 'Lazada'),
        (252, 661, 'Online', 'online', 'Shopee'),
        (663, 771, 'Paradise Park', 'store', None),
        (773, 828, 'Online', 'online', 'Online'),
    ]
    def block_for(excel_row):
        for start, end, store, cat, channel in BLOCKS:
            if start <= excel_row <= end:
                return store, cat, channel
        return None, None, None

    n = 0
    for i, r in enumerate(data):
        excel_row = i + 2
        store, cat, channel = block_for(excel_row)
        if store is None or not r or not r[0]:
            continue
        d = to_date(r[3])
        if d is None or not (d.year == 2026 and d.month in (7, 8)):
            continue
        product_code, product_name = r[13], r[14]
        qty = num(r[15])
        amt = num(r[18])  # column S ('Total amount') -- per-line, already VAT-inclusive
        brand, sub = classify_by_code(product_code)
        if brand == 'EXCLUDE':
            continue
        model = model_from_name(product_name) if product_name else 'Other'
        add_record(store, cat, d, brand, model, sub, qty, amt, r[0], channel=channel,
                    vff_source_text=product_code, vff_name_text=product_name)
        n += 1
    print(f"loaded {n} rows -> Event/Online/Paradise Park, Jul-Aug 2026 (row-position store+channel split, col S Total)")
load_online_event_paradise_julaug2026()

# ================================================================== BFT_Central_Total_Department (line-item detail, split by Store Name)
# 2026-09: swapped to a76e3d0e-BFT_Central_Total_Department_Jan-Jun_26_new.xlsx
# (replacing 026d19bf-BFT_Central_Total_Department_JanJun_26.xlsx), which is a
# genuinely different (better) export: line-item detail with an explicit
# 'GP 30%' column (= Gross Sales * 0.70, the recognized revenue net of
# Central's 30% concession fee -- confirmed exactly, same pattern as Siam
# Discovery's 32%-fee GP32% column) and a 'Before Vat' column (= GP 30% /
# 1.07). Summed across all 5 stores this ties out to the "Jan-Jun_final_
# without_tax.xlsx" BFT-tab reference's Central column to the cent, for
# every one of the 6 months -- which resolves the previous discrepancy and
# means all 5 stores (not just Lardprao) can be reinstated with confidence.
# The 4 Endeavors-corner stores that were disabled 2026-08 (pending this
# reconciliation) are re-enabled below. Unlike the old aggregate-only file,
# SKU Name here spells out colors in full (e.g. "V-SOUL(W39, SILVER)"), so
# vff_shoe_size_color() resolves them directly -- no COLOR_CODE_TO_NAME
# lookup needed for this source.
def load_central_total_department():
    fn = SRC + 'a76e3d0e-BFT_Central_Total_Department_Jan-Jun_26_new.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['Export']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    idx = {h: i for i, h in enumerate(header) if h}
    data = rows[1:]
    MONTHS = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6}
    STORE_NAME_MAP = {
        'CHIDLOM': 'Central Chidlom',
        'CHIDLOM ONLINE': 'Central Chidlom Online',
        'CENTRAL WORLD-CDS': 'Central World (CDS)',
        'LARDPRAO': 'Central Lardprao (Dept.)',
        'EASTVILLE': 'Central Eastville',
    }
    n, skipped = 0, 0
    for r in data:
        if not any(r):
            continue
        store = r[idx['Store Name']]
        sku_name = r[idx['SKU Name']]
        cat = r[idx['Catalogue No.']]
        msd = r[idx['Month Sales Date']]
        if not store or not cat or not msd:
            skipped += 1  # per-store subtotal marker rows (Store Name/Catalogue No. blank)
            continue
        mo, yr = str(msd).split('-')
        d = date(int(yr), MONTHS[mo], 1)  # first-of-month placeholder (source has no daily granularity)
        qty = num(r[idx['Sales Quantity']])
        amt = num(r[idx['GP 30%']])  # VAT-inclusive recognized revenue (Gross Sales * 0.70); ser() applies VAT later
        brand, sub = classify_by_code(cat)
        model = model_from_name(sku_name)
        store_label = STORE_NAME_MAP.get(store, f'Central {store.title()}')
        add_record(store_label, 'central_dept', d, brand, model, sub, qty, amt,
                   order_id=None, vff_source_text=cat, vff_name_text=sku_name)
        n += 1
    print(f"loaded {n} rows -> Central Total Department (all 5 stores; {skipped} subtotal/marker rows skipped)")
load_central_total_department()

# 2026-09: July-August follow-up (d9a6e5aa-BFT_Central_Total_Department_
# Jul-Aug_26.xlsx), same schema as the H1 file above (Store Name/SKU Name/
# Catalogue No./Month Sales Date/.../GP 30%/Before Vat), amount from column J
# ('Before Vat', already VAT-exclusive per user instruction -- *1.07 here so
# add_record's downstream VAT division reproduces it). Covers 4 of the 5
# Central stores (no EASTVILLE rows in this file) for both July and August.
#
# July REPLACES etl_jul2026.py's load_central_total_department_jul2026(),
# now disabled -- see the note there for why (wrong revenue-basis column,
# and an incorrect assumption that 3 of the 4 stores were already covered by
# 対EDV). Confirmed with the user 2026-09: this file's retail-basis figures
# for CENTRAL WORLD-CDS/CHIDLOM/CHIDLOM ONLINE are a genuinely separate
# revenue stream from 対EDV's wholesale invoice to Endeavors for the same
# locations, not a duplicate -- both should be loaded. 対EDV's own data
# currently only extends through July (a separate August wholesale invoice
# is expected later); this loader's scope is unaffected by that either way.
def load_central_total_department_julaug2026():
    fn = SRC + 'd9a6e5aa-BFT_Central_Total_Department_Jul-Aug_26.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['Export']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    idx = {h: i for i, h in enumerate(header) if h}
    data = rows[1:]
    MONTHS = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
    STORE_NAME_MAP = {
        'CHIDLOM': 'Central Chidlom',
        'CHIDLOM ONLINE': 'Central Chidlom Online',
        'CENTRAL WORLD-CDS': 'Central World (CDS)',
        'LARDPRAO': 'Central Lardprao (Dept.)',
        'EASTVILLE': 'Central Eastville',
    }
    n, skipped = 0, 0
    for r in data:
        if not any(r):
            continue
        store = r[idx['Store Name']]
        sku_name = r[idx['SKU Name']]
        cat = r[idx['Catalogue No.']]
        msd = r[idx['Month Sales Date']]
        if not store or not cat or not msd:
            skipped += 1
            continue
        mo, yr = str(msd).split('-')
        if not (int(yr) == 2026 and MONTHS[mo] in (7, 8)):
            continue
        d = date(int(yr), MONTHS[mo], 1)
        qty = num(r[idx['Sales Quantity']])
        before_vat = num(r[idx['Before Vat']])
        amt = before_vat * 1.07
        brand, sub = classify_by_code(cat)
        model = model_from_name(sku_name)
        store_label = STORE_NAME_MAP.get(store, f'Central {store.title()}')
        add_record(store_label, 'central_dept', d, brand, model, sub, qty, amt,
                   order_id=None, vff_source_text=cat, vff_name_text=sku_name)
        n += 1
    print(f"loaded {n} rows -> Central Total Department, Jul-Aug 2026 (4 of 5 stores, no Eastville in this file; {skipped} subtotal/marker rows skipped)")
load_central_total_department_julaug2026()

# ================================================================== 対EDV: Barefoot -> Endeavors wholesale invoices (Jan-Jul 2026)
# New source added 2026-08, replacing the individual Endeavors-operated
# store/channel feeds above (VFF Cart LP, Central Ladprao 3F, Thaniya,
# K Village, Central Chidlom/Chidlom Online/World (CDS)/Eastville, EDV
# online, EDV_EVENT_1, EDV Consignment -- all disabled above/in
# etl_jul2026.py) with a single consolidated line: what Barefoot actually
# invoiced Endeavors for, at wholesale price. Confirmed with the user
# 2026-08: this is a genuinely different revenue basis from every other
# store in this dashboard (wholesale/transfer price, not retail POS price)
# -- roughly a third of the combined retail-equivalent total for the stores
# it replaces -- not a more granular replacement, and intentional. Every row
# bills the same customer ("เอ็นเดเวอร์ซ Company Limited" / Endeavors) from
# Barefoot's single HeadQuarter branch, so there is no store/branch
# breakdown to preserve -- this can only ever be one combined line ('対EDV'
# store, its own new 'edv' channel-group). Spans the full Jan-Jul 2026 range
# in one file (unlike the etl.py/etl_jul2026.py H1/July split used
# elsewhere), so it's loaded here in full; aggregate.py buckets every record
# by its own date field regardless of which script loaded it, so this
# doesn't need to be duplicated across the two scripts. Has no 2025/2024
# data at all -- per user request, no prior-year comparison is expected for
# this store (the existing null/zero-baseline handling in yoyBadge/
# yoyCellHtml already renders that as "--"/pending automatically).
def load_edv_invoice_report():
    fn = SRC + 'e310e8c5-Barefootinc_to_Endeavors_JanJul_2026.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['Invoice Report']
    rows = list(ws.iter_rows(values_only=True))
    header = rows[11]  # rows 0-10 are a report-info banner (merchant/export date/etc.)
    idx = {h: i for i, h in enumerate(header) if h}
    data = rows[12:]
    data = [r for r in data if any(r) and r[idx.get('Document No.')] and r[idx.get('Product/Service  Code')]]

    # 'Grand Total' (the actual VAT-inclusive amount invoiced, after whatever
    # wholesale discount applies to that invoice -- seen ranging 40%-80% off
    # list price across the 16 invoices here) is only populated once per
    # invoice, on its first product line; prorate every line to its share of
    # the invoice's Pre-VAT Amount, same pattern as every other loader here.
    inv_pretax = defaultdict(float)
    inv_grand = {}
    for r in data:
        doc = r[idx['Document No.']]
        inv_pretax[doc] += num(r[idx['Pre-VAT Amount']])
        gt = r[idx.get('Grand Total')]
        if gt not in (None, ''):
            inv_grand[doc] = num(gt)

    n = 0
    for r in data:
        doc = r[idx['Document No.']]
        d = parse_dmy(r[idx['Issue Date']])
        if d is None:
            continue
        pc = r[idx['Product/Service  Code']]
        pname = r[idx['Product/Service Name']]
        qty = num(r[idx['Quantity']])
        line_pretax = num(r[idx['Pre-VAT Amount']])
        total_pretax = inv_pretax.get(doc, 0.0)
        grand = inv_grand.get(doc, total_pretax)
        amt = (line_pretax / total_pretax * grand) if total_pretax else 0.0
        brand, sub = classify_by_code(pc)
        model = model_from_name(pname)
        # order_id intentionally omitted (None): the 16 wholesale invoices
        # here are batch billing documents to a single distributor, not
        # individual retail customer receipts -- counting them as "orders"
        # would pollute 客数（伝票数）/客単価 with a number that doesn't mean
        # the same thing as everywhere else in this dashboard. Same
        # order_id=None convention already used for Central Total
        # Department, which has the identical no-per-transaction-granularity
        # issue.
        add_record('対EDV', 'edv', d, brand, model, sub, qty, amt, order_id=None, vff_source_text=pc, vff_name_text=pname)
        n += 1
    print(f"loaded {n} rows -> 対EDV ({len(inv_grand)} invoices, sum Grand Total = {sum(inv_grand.values()):,.2f})")
load_edv_invoice_report()

# 2026-09: August-only follow-up (0d28b573-BFT_Inter-company_sales_EDV_Jul-
# Aug_26.xlsx) -- July is already covered by the file above (confirmed:
# this file's July total, 981,712.33 THB net, matches the already-loaded
# figure to a few cents -- same underlying data, just a different monthly
# export). Same "Invoice Report"-style schema as BFT Consignment's Jul-Aug
# follow-up (load_bft_consignment_aug2026() above): per user instruction,
# each line's amount is column I ('Pre-VAT Amount') scaled by the order-
# level discount/fee percentage in column K ('Total Discount', e.g.
# '70.00%', present only on an order's first line -- forward-filled to its
# other lines here). Mathematically equivalent to load_edv_invoice_report()'s
# line/invoice-total ratio proration; this loader follows the user's stated
# column-based method directly. order_id=None, same batch-invoice
# convention as load_edv_invoice_report() (these aren't per-transaction
# retail receipts).
def load_edv_invoice_report_aug2026():
    fn = SRC + '0d28b573-BFT_Inter-company_sales_EDV_Jul-Aug_26.xlsx'
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb['Invoice Report (2)']
    rows = list(ws.iter_rows(values_only=True))
    data = rows[1:]

    n = 0
    cur_doc, cur_pct = None, None
    for r in data:
        if not r or not any(r):
            continue
        doc, pc, pname = r[0], r[3], r[4]
        k_val = r[10]
        if doc:
            cur_doc = doc
            if k_val not in (None, ''):
                cur_pct = float(str(k_val).replace('%', '')) / 100.0
        if not cur_doc or not pc:
            continue
        d = parse_dmy(r[1])
        if d is None or not (d.year == 2026 and d.month == 8):
            continue  # July skipped -- already loaded, see note above
        qty = num(r[5])
        i_val = r[8]
        if i_val is None:
            continue
        amt_net = num(i_val) * (1 - (cur_pct or 0.0))
        amt = amt_net * 1.07
        brand, sub = classify_by_code(pc)
        model = model_from_name(pname)
        add_record('対EDV', 'edv', d, brand, model, sub, qty, amt, order_id=None,
                    vff_source_text=pc, vff_name_text=pname)
        n += 1
    print(f"loaded {n} rows -> 対EDV, August only (July already covered by load_edv_invoice_report())")
load_edv_invoice_report_aug2026()

# ================================================================== Export: VFF shoe overseas wholesale shipment (June 2026)
# Added 2026-09 from a packing-list image (not a spreadsheet -- transcribed
# by hand below), a single wholesale export shipment. USD converted to THB
# at 32.9594 per user instruction; verified the resulting THB total
# (2,722,907.87) matches the "Jan-Jun_final_without_tax.xlsx" BFT-tab
# Export-column figure exactly, resolving what had been an unexplained line
# item in that reference file. That THB figure is VAT-exclusive -- exports
# are zero-rated under Thai VAT law, and it matches every other verified
# BFT-tab figure's convention -- so it's converted back to VAT-inclusive raw
# amounts here, same as every other add_record() call (ser() divides every
# amount by 1.07 exactly once, downstream, uniformly). Gender is inferred
# from the Style No. prefix letter (W/M) per the source's own convention,
# not from the size token the way vff_shoe_gender() does elsewhere (this
# shipment uses plain EU sizes 35-48, already the same numbering the rest of
# the dashboard's "size" field uses -- no conversion needed). order_id=None,
# no per-transaction data -- same convention as 対EDV/Central Lardprao
# (Dept.)/Yoshi Run for this kind of lump-shipment/order-less revenue.
def load_export_2026_09():
    USD_TO_THB = 32.9594
    # (style_no, model, color, unit_price_usd, {eu_size: qty})
    LINES = [
        ('26W7201', 'V-Soul', 'Black', 78.70, {'35':9,'36':15,'37':30,'38':48,'39':42,'40':27,'41':21}),
        ('26W7202', 'V-Soul', 'Brown', 78.70, {'36':9,'37':18,'38':33,'39':24,'40':15,'41':9}),
        ('26W7205', 'V-Soul', 'Baby Blue', 78.70, {'35':6,'36':12,'37':21,'38':36,'39':27,'40':18,'41':12}),
        ('26M7001', 'V-Run', 'Total Black', 93.10, {'39':6,'40':9,'41':21,'42':30,'43':27,'44':21,'45':18,'46':12,'47':9,'48':3}),
        ('26M7001', 'V-Run', 'Total Black', 93.10, {'35':6,'36':12,'37':27,'38':27,'39':24,'40':15,'41':9}),
        ('26M7004', 'V-Run', 'Black-Lime/Black', 93.10, {'40':12,'41':21,'42':24,'43':21,'44':15,'45':12,'46':9,'47':6}),
        ('25M7501', 'Trailope', 'Black', 97.90, {'39':6,'40':12,'41':21,'42':24,'43':21,'44':15,'45':12,'46':6,'47':3}),
    ]
    d = date(2026, 6, 1)  # shipment date not itemized in the source; matches the BFT-tab reference, which carries the whole amount in June
    n = 0
    for style_no, model, color, unit_price, sizes in LINES:
        gender = 'Women' if style_no[2] == 'W' else 'Men'
        for size, qty in sizes.items():
            amt_thb_excl_vat = unit_price * qty * USD_TO_THB
            amt_raw = round(amt_thb_excl_vat * 1.07, 2)
            records.append(dict(
                store='Export', category='export', date=d.isoformat(), month=d.strftime('%Y-%m'),
                brand='VFF', model=model, is_vff_shoe=True, gender=gender, size=size, color=color,
                sub=None, qty=qty, amount=amt_raw, order_id=None,
                channel=None, payment=None, entity=None,
            ))
            n += 1
    print(f"loaded {n} rows -> Export (1 shipment, 948 pairs, USD 82,614.00 -> THB {sum(r['amount'] for r in records if r['store']=='Export')/1.07:,.2f} excl. VAT)")
load_export_2026_09()

# ================================================================== summary / sanity checks
print()
print("="*80)
print(f"TOTAL RECORDS: {len(records)}")
total_amt = sum(r['amount'] for r in records)
total_qty = sum(r['qty'] for r in records)
print(f"TOTAL AMOUNT (all categories): {total_amt:,.2f}")
print(f"TOTAL QTY (all categories): {total_qty:,.0f}")

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

# brand totals
by_brand = defaultdict(lambda: {'amount':0.0,'qty':0.0})
for r in records:
    by_brand[r['brand']]['amount'] += r['amount']
    by_brand[r['brand']]['qty'] += r['qty']
print()
print("Brand totals:")
for b, v in sorted(by_brand.items(), key=lambda x: -x[1]['amount']):
    print(f"  {b:15s} amount={v['amount']:>14,.2f}  qty={v['qty']:>8,.0f}")

# save raw records
with open(os.path.join(os.environ['SALES_OUT_DIR'], 'records.json'), 'w', encoding='utf-8') as f:
    json.dump(records, f, ensure_ascii=False)
print()
print("Saved", len(records), "records to records.json")
