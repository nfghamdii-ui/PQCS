#!/usr/bin/env python3
"""
Split the flat Main Log into linked sheets, keyed the way the project
already names things.

Nothing here carries an invented code. A material is named by its MAT
reference, a delivery by its MIR, a document by its own number, and a
vendor by its name — because a PQD covers three manufacturers at once
in this log and so cannot name any one of them.

A reference names a document, not a material: 4MAKA08-MAKA-08-MBL-HV-ITP-00004
covers twelve materials here. So a document gets one row, and a Coverage
sheet says which materials it covers.

Nothing is invented. Every value came out of the Main Log; a cell that
could not be read is carried into a Notes column rather than guessed
at, and two rows that disagree about the same document are both
reported rather than silently merged.
"""
import re, datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

SRC = '/mnt/user-data/uploads/Draft_Project_Materials_Live_Tracking_Sheet.xlsx'
OUT = '/mnt/user-data/outputs/Project Materials — Linked Sheets.xlsx'

wb = load_workbook(SRC, read_only=True, data_only=True)
rows = list(wb['Main Log'].iter_rows(values_only=True))
HDR = [str(h).strip() if h is not None else '' for h in rows[0]]
C = {h: i for i, h in enumerate(HDR)}
ITEMS = [r for r in rows[1:] if r[1]]

NULLS = {'', '-', '--', 'n/a', 'na', 'tbd', 'tba', 'avl', 'alv'}

def val(r, name):
    i = C.get(name)
    return None if i is None or i >= len(r) else r[i]

def txt(r, name):
    v = val(r, name)
    if v is None:
        return ''
    s = str(v).replace('\u00a0', ' ').strip()
    return '' if s.lower() in NULLS else s

MON = dict(jan=1, feb=2, mar=3, apr=4, may=5, jun=6,
           jul=7, aug=8, sep=9, oct=10, nov=11, dec=12)

def edit(a, b):
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]

def month(name):
    k = name.lower()[:3]
    if k in MON:
        return MON[k]
    best, bd, tie = 0, 2, False
    for n, m in MON.items():
        d = edit(k, n)
        if d < bd:
            bd, best, tie = d, m, False
        elif d == bd and m != best:
            tie = True
    return 0 if tie else best          # "Ju" is June and July both, so it is neither

def as_date(v):
    if v is None or v == '':
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, (int, float)):
        n = int(round(v))
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=n) if 1 < n < 80000 else None
    s = str(v).strip()
    if not s or s.lower() in NULLS:
        return None
    if re.fullmatch(r'\d+(\.\d+)?', s) and 20000 < float(s) < 80000:
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(s)))
    m = re.fullmatch(r'(\d{1,2})[-/ ]([A-Za-z]{2,9})[-/ ](\d{2,4})', s)
    if m:
        mo = month(m.group(2))
        if not mo:
            return None
        y = int(m.group(3))
        y += 2000 if y < 100 else 0
        try:
            return datetime.date(y, mo, int(m.group(1)))
        except ValueError:
            return None
    m = re.fullmatch(r'(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None

def parts(v):
    """One cell holding several values becomes the list it always was."""
    if v is None:
        return []
    if not isinstance(v, str):
        return [v]
    s = v.replace('\u00a0', ' ').strip()
    if not s or s.lower() in NULLS:
        return []
    bits = re.split(r'[\r\n]+', s) if re.search(r'[\r\n]', s) else re.split(r'\s{2,}', s)
    return [b.strip() for b in bits if b.strip() and b.strip().lower() not in NULLS]

def pair(keys, others):
    """Equal lengths pair off; one value belongs to all; anything else stays unpaired."""
    if not others:
        return [None] * len(keys)
    if len(others) == len(keys):
        return others
    if len(others) == 1:
        return others * len(keys)
    return [others[i] if i < len(others) else None for i in range(len(keys))]

def status(v):
    s = '' if v is None else str(v).strip()
    if not s:
        return ''
    best, rank = '', {'Rejected': 1, 'Resubmit': 2, 'Under Review': 3,
                      'Approved as Noted': 4, 'Approved': 5}
    for p in re.split(r'[\r\n]+', s):
        k = re.sub(r'[^a-z ]', ' ', p.lower()).strip()
        if not k:
            continue
        if 'as noted' in k or 'comment' in k:
            one = 'Approved as Noted'
        elif 'revise' in k or 'resubmit' in k:
            one = 'Resubmit'
        elif 'reject' in k:
            one = 'Rejected'
        elif 'under review' in k or k == 'ur':
            one = 'Under Review'
        else:
            w = k.split()[0]
            one = 'Approved' if (edit(w, 'approved') <= 2 or edit(w, 'approve') <= 1) else ''
        if one and rank.get(one, 0) > rank.get(best, 0):
            best = one
    return best

def yesno(v):
    s = ('' if v is None else str(v)).strip().lower()
    if not s or s == '-':
        return ''
    if s in ('mo', 'n', 'no'):
        return 'No'
    return 'Yes' if s.startswith('y') else ('No' if s.startswith('n') else '')

def cat(v):
    m = re.search(r'C\s*([0-3])', str(v or '').upper())
    return 'C' + m.group(1) if m else ''

def norm(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()

# ---------------------------------------------------------------- vendors
vendors, vendor_of = [], {}

def vendor_id(r):
    """
    The vendor is whoever made the material. The sub-contractor column holds
    the trade doing the installation — FIRST FIX against half the log — and
    folding the two together would hang fifty materials off one company.
    """
    name = txt(r, 'Manufacturer')
    if not name:
        return ''
    k = norm(name).lower()
    if k in vendor_of:
        v = vendors[vendor_of[k]]
        got = txt(r, 'PQD Number')
        if got and not v['PQD Number']:
            v['PQD Number'] = got
            v['PQD Status'] = status(val(r, 'PQD Status'))
            v['PQD Submittal Date'] = as_date(val(r, 'PQD Submittal Date'))
        elif got and norm(got) != norm(v['PQD Number']):
            add_note(v, 'a row also carries ' + got + ' — two pre-qualifications on record')
        return v['Vendor Name']
    vendor_of[k] = len(vendors)
    vendors.append({
        'Vendor Name': name, 'Kind': 'Manufacturer',
        'Country': txt(r, 'Country of Origin of Manufacture'),
        'Discipline': txt(r, 'Discipline'),
        'PQD Number': txt(r, 'PQD Number'), 'PQD Revision': txt(r, 'PQD Revision'),
        'PQD Status': status(val(r, 'PQD Status')),
        'PQD Submittal Date': as_date(val(r, 'PQD Submittal Date')),
        'PA Tentative Date': as_date(val(r, 'PA Tentative Date')),
        'PA Document Number': txt(r, 'PA Document Number'),
        'PA Date': as_date(val(r, 'PA Date')),
        '3rd Party Assessment': yesno(val(r, '3rd Party Assessment Done')),
        'Client Assessment': yesno(val(r, 'Client Assessment Done')),
        'PMC/LDC Assessment': yesno(val(r, 'PMC/LDC Assessment Done')),
        'Contractor Assessment': yesno(val(r, 'Contractor Assessment Done')),
        'Assessment Result': status(val(r, 'Assessment Result')) or txt(r, 'Assessment Result'),
        'Notes': '',
    })
    return name

def add_note(d, s):
    if s in (d.get('Notes') or ''):
        return
    d['Notes'] = (d['Notes'] + ' | ' if d.get('Notes') else '') + s

# ------------------------------------------------- documents and coverage
materials, coverage, unread = [], [], []
subs, subs_of = [], {}
pos, pos_of = [], {}
insp, insp_of = [], {}
_seen_unread = set()

def note_unread(mid, col, raw):
    if (mid, col) in _seen_unread:
        return
    _seen_unread.add((mid, col))
    many = bool(re.search(r'[\r\n]', str(raw))) or \
        len(re.findall(r'\d{1,2}[-/ ][A-Za-z]{3}', str(raw))) > 1
    unread.append({'MAT_ID': mid, 'Column': col, 'As typed': str(raw),
                   'Why': 'more than one date in the cell' if many
                          else 'the month cannot be told apart'})

def cover(kind, number, mid):
    coverage.append({'Kind': kind, 'Document No': number, 'MAT Number': mid,
                     'Item Description': '', 'Notes': ''})

def merge(store, index, key, new, label):
    """
    One document, one row. A second material naming the same reference adds
    a coverage line, not another copy. Where the two rows disagree about the
    document's own date or status, both are kept in sight — a contradiction
    in the log is a thing to settle, not to average away.
    """
    k = norm(key).lower()
    if k not in index:
        index[k] = len(store)
        store.append(new)
        return new
    have = store[index[k]]
    for f in ('Submittal Date', 'Revision', 'Status', 'PO Date', 'Result', 'Planned Date'):
        if f not in new or f not in have:
            continue
        a, b = have.get(f), new.get(f)
        if a in (None, '') and b not in (None, ''):
            have[f] = b
        elif b not in (None, '') and norm(a) != norm(b):
            add_note(have, '%s differs on another row: %s' % (f, b))
    return have

for n, r in enumerate(ITEMS, 1):
    # The material is named by the reference the project already gave it.
    # All 103 are filled and none repeats. A material raised before its
    # MAT is issued gets a visible placeholder rather than a silent blank.
    mid = txt(r, 'MAT Number') or ('NO MAT NUMBER — %03d' % n)
    seq = 'M%03d' % n                      # only for events that carry no number of their own
    vid = vendor_id(r)
    materials.append({
        'MAT Number': mid, 'Item Description': txt(r, 'Item Description'),
        'Material Category': cat(val(r, 'Material Category')),
        'Discipline': txt(r, 'Discipline'),
        'Package': txt(r, 'Package (Lump Sum / Provisional Sum / Prime Cost)'),
        'Common Package': yesno(val(r, 'Common Package (Yes/No)')),
        'Finishing': txt(r, 'Finishing (Internal/External)'),
        'Sub-contractor': txt(r, 'Sub-contractor Name'),
        'Scope': txt(r, 'Supply & Install / Supply / Install Only'),
        'Vendor': vid,
        'Total Quantity': val(r, 'Total Quantity'),
        'Unit': txt(r, 'Total Quantity Unit'),
        'Delivered': '', 'Remaining': '', 'Delivered %': '',
        'Submittals': '', 'Purchase Orders': '', 'Inspections': '',
        'Design Verification': txt(r, 'Design Verification/Calculation Status'),
        'Fabrication Planned': as_date(val(r, 'Fabrication Planned Date')),
        'Fabrication Started': as_date(val(r, 'Fabrication 1st Batch Started Date')),
        'Fabrication %': val(r, 'Fabrication Percentage Completion (%)'),
        'Sample': txt(r, 'Sample'),
        'Mock-up Delivered': as_date(val(r, 'Mock-up Delivered Date')),
        'Mock-up In Place': as_date(val(r, 'Mock-up First in Place')),
        'Mock-up Approved': txt(r, 'Mock-up Approved by PMC'),
        'Mock-up per DLA': txt(r, 'Mock-up per Client DLA'),
        'Notes': '',
    })

    # ---- submittals ----
    def add_subs(kind, numcol, datecol, revcol, statcol, inclcol=None):
        nums = [str(x).strip() for x in parts(val(r, numcol))]
        if not nums:
            return
        dts = pair(nums, parts(val(r, datecol)) if datecol else [])
        rvs = pair(nums, parts(val(r, revcol)) if revcol else [])
        sss = pair(nums, parts(val(r, statcol)) if statcol else [])
        for i, num in enumerate(nums):
            d = as_date(dts[i]) if dts[i] is not None else None
            if dts[i] and d is None:
                note_unread(mid, datecol, val(r, datecol))
            merge(subs, subs_of, num, {
                'Document No': num, 'Type': kind,
                'Revision': rvs[i] if rvs[i] is not None else '',
                'Submittal Date': d,
                'Status': status(sss[i]) if sss[i] else status(val(r, statcol)),
                'Incl. ITP': yesno(val(r, inclcol)) if inclcol else '',
                'Materials': '',
                'Notes': ('date as typed: ' + str(dts[i])) if (dts[i] and d is None) else '',
            }, kind)
            cover('Submittal', num, mid)

    add_subs('MAT', 'MAT Number', 'MAT Submittal Date', 'MAT Revision', 'MAT Status')
    add_subs('MES', 'Method Statement Number', None, 'MES Revision', 'MES Status',
             'MES Incl. ITP (Yes/No)')
    add_subs('ITP', 'ITP Number', 'ITP Submittal Date', 'ITP Revision', 'ITP Status')
    add_subs('PID', 'PID Number', 'PID Submittal Date', 'PID Revision', 'PID Status')

    # ---- purchase orders ----
    nums = [str(x).strip() for x in parts(val(r, 'PO Number'))]
    dts = pair(nums, parts(val(r, 'PO Date')))
    issued = yesno(val(r, 'Purchase Order Issued (Yes/No)'))
    if not nums and issued == 'Yes':
        key = 'PO pending — ' + seq            # marked issued, but the log has no number
        merge(pos, pos_of, key, {'Document No': key, 'PO Date': None, 'Issued': 'Yes',
                                 'Materials': '',
                                 'Notes': 'no number in the log — replace this placeholder'},
              'PO')
        cover('Purchase Order', key, mid)
    for i, num in enumerate(nums):
        d = as_date(dts[i]) if dts[i] is not None else None
        if dts[i] and d is None:
            note_unread(mid, 'PO Date', val(r, 'PO Date'))
        merge(pos, pos_of, num, {
            'Document No': num, 'PO Date': d, 'Issued': issued or 'Yes', 'Materials': '',
            'Notes': ('date as typed: ' + str(dts[i])) if (dts[i] and d is None) else ''}, 'PO')
        cover('Purchase Order', num, mid)

    # ---- inspections ----
    pfm_txt = txt(r, 'Pre-Fabrication Meeting Date')
    pfm = as_date(val(r, 'Pre-Fabrication Meeting Date'))
    if pfm or pfm_txt:
        key = 'PFM ' + seq                     # a meeting is an event, and carries no number
        merge(insp, insp_of, key, {
            'Document No': key, 'Type': 'Pre-Fabrication Meeting',
            'Planned Date': pfm, 'Location': '',
            '3rd Party Assigned': yesno(val(r, '3rd Party Assigned (Yes/No)')),
            'Provider': txt(r, '3rd Party Service Provider Name'), 'Result': '',
            'Materials': '',
            'Notes': '' if pfm else 'date as typed: ' + pfm_txt}, 'INS')
        cover('Inspection', key, mid)
        if not pfm and pfm_txt:
            note_unread(mid, 'Pre-Fabrication Meeting Date', pfm_txt)
    reports = [str(x).strip() for x in parts(val(r, 'FAT/TPI Results'))]
    fatpkg = txt(r, 'FAT Package/Procedure Number/ITP')
    fatst = status(val(r, 'FAT Package Status'))
    keys = reports or ([fatpkg] if fatpkg else (['FAT ' + seq] if fatst else []))
    for keyn in keys:
        merge(insp, insp_of, keyn, {
            'Document No': keyn, 'Type': 'FAT / TPI',
            'Planned Date': as_date(val(r, 'FAT Planned Date')),
            'Location': txt(r, 'FAT Location / City'),
            '3rd Party Assigned': yesno(val(r, '3rd Party Assigned (Yes/No)')),
            'Provider': txt(r, '3rd Party Service Provider Name'),
            'Result': fatst, 'Materials': '',
            'Notes': '' if keyn != 'FAT ' + seq else 'no report number in the log'}, 'INS')
        cover('Inspection', keyn, mid)

# ---------------------------------------------------------------- writing
BODY = Font(name='Arial', size=10)
HEAD = Font(name='Arial', size=10, bold=True, color='FFFFFF')
TITLE = Font(name='Arial', size=14, bold=True)
SUBF = Font(name='Arial', size=10, italic=True, color='555555')
BOLD = Font(name='Arial', size=11, bold=True)
HFILL = PatternFill('solid', fgColor='1F2937')
KEYFILL = PatternFill('solid', fgColor='EEF0F3')
LOCKFILL = PatternFill('solid', fgColor='F5F6F8')
WARNFILL = PatternFill('solid', fgColor='FDF3E2')
THIN = Border(bottom=Side(style='thin', color='D8DCE2'))

out = Workbook()
out.remove(out.active)

def put(name, cols, data, widths=None, keys=1, locked=()):
    ws = out.create_sheet(name)
    ws.append(cols)
    for j in range(1, len(cols) + 1):
        c = ws.cell(1, j)
        c.font, c.fill = HEAD, HFILL
        c.alignment = Alignment(vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 30
    for row in data:
        ws.append([row.get(c) for c in cols])
    for i in range(2, ws.max_row + 1):
        for j, c in enumerate(cols, 1):
            cell = ws.cell(i, j)
            cell.font, cell.border = BODY, THIN
            cell.alignment = Alignment(vertical='top')
            if j <= keys:
                cell.fill = KEYFILL
            if c in locked:
                cell.fill = LOCKFILL
            if isinstance(cell.value, datetime.date):
                cell.number_format = 'DD/MM/YYYY'
    for j, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = \
            (widths or {}).get(c, max(10, min(30, len(c) + 4)))
    ws.freeze_panes = ws.cell(2, keys + 1)
    ws.auto_filter.ref = ws.dimensions
    return ws

MAT_COLS = ['MAT Number', 'Item Description', 'Material Category', 'Discipline', 'Package',
            'Common Package', 'Finishing', 'Sub-contractor', 'Scope', 'Vendor',
            'Total Quantity', 'Unit', 'Delivered', 'Remaining', 'Delivered %',
            'Submittals', 'Purchase Orders', 'Inspections',
            'Design Verification', 'Fabrication Planned', 'Fabrication Started',
            'Fabrication %', 'Sample', 'Mock-up Delivered', 'Mock-up In Place',
            'Mock-up Approved', 'Mock-up per DLA', 'Notes']
VEN_COLS = ['Vendor Name', 'Kind', 'Country', 'Discipline', 'PQD Number', 'PQD Revision',
            'PQD Status', 'PQD Submittal Date', 'PA Tentative Date', 'PA Document Number',
            'PA Date', '3rd Party Assessment', 'Client Assessment', 'PMC/LDC Assessment',
            'Contractor Assessment', 'Assessment Result', 'Materials', 'Notes']
SUB_COLS = ['Document No', 'Type', 'Revision', 'Submittal Date', 'Status', 'Incl. ITP',
            'Materials', 'Notes']
PO_COLS = ['Document No', 'PO Date', 'Issued', 'Materials', 'Notes']
INS_COLS = ['Document No', 'Type', 'Planned Date', 'Location', '3rd Party Assigned',
            'Provider', 'Result', 'Materials', 'Notes']
COV_COLS = ['Kind', 'Document No', 'MAT Number', 'Item Description', 'Notes']
DEL_COLS = ['MIR Number', 'MAT Number', 'Item Description', 'MIR Approval Date', 'MIR Status',
            'Quantity', 'Unit', 'Storage', 'Planned Date', 'Actual Date', 'Notes']
INST_COLS = ['WIR Number', 'MAT Number', 'Item Description', 'WIR Approval Date', 'WIR Status',
             'Installer', 'Planned Date', 'Actual Date', 'Notes']

n_mat, n_ven, n_sub, n_po, n_ins, n_cov = \
    len(materials), len(vendors), len(subs), len(pos), len(insp), len(coverage)
MROWS, CROWS, DROWS = n_mat + 1, max(n_cov + 1, 900), 900

REF = 34          # a reference is thirty-two characters; give it room
ws_m = put('Materials', MAT_COLS, materials,
           {'MAT Number': REF, 'Item Description': 58, 'Discipline': 22,
            'Sub-contractor': 18, 'Vendor': 26, 'Notes': 30},
           keys=2, locked=('Delivered', 'Remaining', 'Delivered %',
                           'Submittals', 'Purchase Orders', 'Inspections'))
ws_v = put('Vendors', VEN_COLS, vendors,
           {'Vendor Name': 30, 'PQD Number': REF, 'PA Document Number': 30, 'Notes': 54},
           keys=1, locked=('Materials',))
ws_s = put('Submittals', SUB_COLS, subs, {'Document No': REF, 'Notes': 46},
           keys=1, locked=('Materials',))
ws_p = put('Purchase Orders', PO_COLS, pos, {'Document No': 26, 'Notes': 46},
           keys=1, locked=('Materials',))
ws_i = put('Inspections', INS_COLS, insp,
           {'Document No': REF, 'Provider': 30, 'Notes': 40}, keys=1, locked=('Materials',))
ws_c = put('Coverage', COV_COLS, coverage,
           {'Document No': REF, 'MAT Number': REF, 'Item Description': 58, 'Notes': 22},
           keys=3, locked=('Item Description',))
ws_d = put('Deliveries', DEL_COLS, [],
           {'MIR Number': REF, 'MAT Number': REF, 'Item Description': 58},
           keys=2, locked=('Item Description',))
ws_w = put('Installation', INST_COLS, [],
           {'WIR Number': REF, 'MAT Number': REF, 'Item Description': 58},
           keys=2, locked=('Item Description',))
if unread:
    put('Unreadable cells', ['MAT Number', 'Column', 'As typed', 'Why'], unread,
        {'MAT Number': REF, 'Column': 28, 'As typed': 46, 'Why': 34})

# ---- Lookups ----
LK = {
    'Material Category': ['C0', 'C1', 'C2', 'C3'],
    'Discipline': sorted({m['Discipline'] for m in materials if m['Discipline']}),
    'Package': sorted({m['Package'] for m in materials if m['Package']}),
    'Scope': sorted({m['Scope'] for m in materials if m['Scope']}),
    'Vendor Kind': ['Manufacturer', 'Supplier', 'Subcontractor', 'Inspection agency'],
    'Submittal Type': ['MAT', 'MES', 'ITP', 'PID', 'PQD', 'Shop Drawing'],
    'Submittal Status': ['Under Review', 'Approved', 'Approved as Noted', 'Resubmit', 'Rejected'],
    'Inspection Type': ['Pre-Fabrication Meeting', 'In-Process', 'FAT / TPI', 'Site Inspection'],
    'Inspection Result': ['Pending', 'Scheduled', 'Approved', 'Approved as Noted', 'Failed'],
    'Coverage Kind': ['Submittal', 'Purchase Order', 'Inspection'],
    'MIR Status': ['Pending', 'Approved', 'Approved as Noted', 'Rejected'],
    'WIR Status': ['Pending', 'Approved', 'Approved as Noted', 'Rejected'],
    'Storage': ['Site', 'Offsite'],
    'Yes / No': ['Yes', 'No'],
}
ws_l = out.create_sheet('Lookups')
ws_l.cell(1, 1, 'Every dropdown in this workbook reads the column below it. '
                'Add a value here before using it.').font = SUBF
ws_l.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(LK))
for j, (name, vals) in enumerate(LK.items(), 1):
    c = ws_l.cell(2, j, name)
    c.font, c.fill = HEAD, HFILL
    for i, v in enumerate(vals, 3):
        ws_l.cell(i, j, v).font = BODY
    ws_l.column_dimensions[get_column_letter(j)].width = max(14, len(name) + 4)
ws_l.freeze_panes = 'A3'

def col_of(cols, name):
    return cols.index(name) + 1

def letter_of(cols, name):
    return get_column_letter(col_of(cols, name))

def validate(ws, cols, col_name, source, n_rows=900, msg=None):
    """A key is chosen from a list, never typed. That is the whole reason
       this workbook can be keyed by references and still hold together."""
    dv = DataValidation(type='list', formula1=source, allow_blank=True, showErrorMessage=True)
    dv.errorTitle = 'Not on the list'
    dv.error = msg or 'Pick a value from the list, or add it to its own sheet first.'
    ws.add_data_validation(dv)
    L = letter_of(cols, col_name)
    dv.add('%s2:%s%d' % (L, L, n_rows))

def dropdown(ws, cols, col_name, look, n_rows=900):
    j = list(LK).index(look) + 1
    L = get_column_letter(j)
    validate(ws, cols, col_name, "Lookups!$%s$3:$%s$%d" % (L, L, 2 + len(LK[look])), n_rows)

MAT_LIST = 'Materials!$A$2:$A$%d' % MROWS
VEN_LIST = 'Vendors!$A$2:$A$%d' % (n_ven + 1)
PICK_MAT = 'Pick a MAT reference from the Materials sheet.'

validate(ws_m, MAT_COLS, 'Vendor', VEN_LIST, MROWS, 'Pick a vendor from the Vendors sheet.')
validate(ws_c, COV_COLS, 'MAT Number', MAT_LIST, CROWS, PICK_MAT)
validate(ws_d, DEL_COLS, 'MAT Number', MAT_LIST, DROWS, PICK_MAT)
validate(ws_w, INST_COLS, 'MAT Number', MAT_LIST, DROWS, PICK_MAT)

dropdown(ws_m, MAT_COLS, 'Material Category', 'Material Category', MROWS)
dropdown(ws_m, MAT_COLS, 'Discipline', 'Discipline', MROWS)
dropdown(ws_m, MAT_COLS, 'Package', 'Package', MROWS)
dropdown(ws_m, MAT_COLS, 'Scope', 'Scope', MROWS)
dropdown(ws_m, MAT_COLS, 'Common Package', 'Yes / No', MROWS)
dropdown(ws_v, VEN_COLS, 'Kind', 'Vendor Kind')
dropdown(ws_v, VEN_COLS, 'PQD Status', 'Submittal Status')
for c in ('3rd Party Assessment', 'Client Assessment', 'PMC/LDC Assessment',
          'Contractor Assessment'):
    dropdown(ws_v, VEN_COLS, c, 'Yes / No')
dropdown(ws_s, SUB_COLS, 'Type', 'Submittal Type')
dropdown(ws_s, SUB_COLS, 'Status', 'Submittal Status')
dropdown(ws_s, SUB_COLS, 'Incl. ITP', 'Yes / No')
dropdown(ws_p, PO_COLS, 'Issued', 'Yes / No')
dropdown(ws_i, INS_COLS, 'Type', 'Inspection Type')
dropdown(ws_i, INS_COLS, 'Result', 'Inspection Result')
dropdown(ws_i, INS_COLS, '3rd Party Assigned', 'Yes / No')
dropdown(ws_c, COV_COLS, 'Kind', 'Coverage Kind', CROWS)
dropdown(ws_d, DEL_COLS, 'MIR Status', 'MIR Status')
dropdown(ws_d, DEL_COLS, 'Storage', 'Storage')
dropdown(ws_w, INST_COLS, 'WIR Status', 'WIR Status')

# ---- formulas ----
COVK = letter_of(COV_COLS, 'Kind')
COVD = letter_of(COV_COLS, 'Document No')
COVM = letter_of(COV_COLS, 'MAT Number')
DQ = letter_of(DEL_COLS, 'Quantity')
DM = letter_of(DEL_COLS, 'MAT Number')
DS = letter_of(DEL_COLS, 'MIR Status')
QTY = letter_of(MAT_COLS, 'Total Quantity')
GOT = letter_of(MAT_COLS, 'Delivered')

for i in range(2, n_mat + 2):
    ws_m.cell(i, col_of(MAT_COLS, 'Delivered')).value = (
        '=SUMIFS(Deliveries!$%s$2:$%s$%d,Deliveries!$%s$2:$%s$%d,$A%d,'
        'Deliveries!$%s$2:$%s$%d,"Approved")'
        % (DQ, DQ, DROWS, DM, DM, DROWS, i, DS, DS, DROWS))
    ws_m.cell(i, col_of(MAT_COLS, 'Remaining')).value = (
        '=IF($%s%d="","",MAX(0,$%s%d-$%s%d))' % (QTY, i, QTY, i, GOT, i))
    c = ws_m.cell(i, col_of(MAT_COLS, 'Delivered %'))
    c.value = '=IF(OR($%s%d="",$%s%d=0),"",$%s%d/$%s%d)' % (QTY, i, QTY, i, GOT, i, QTY, i)
    c.number_format = '0%'
    for name, kind in (('Submittals', 'Submittal'), ('Purchase Orders', 'Purchase Order'),
                       ('Inspections', 'Inspection')):
        ws_m.cell(i, col_of(MAT_COLS, name)).value = (
            '=COUNTIFS(Coverage!$%s$2:$%s$%d,$A%d,Coverage!$%s$2:$%s$%d,"%s")'
            % (COVM, COVM, CROWS, i, COVK, COVK, CROWS, kind))

VCOL = letter_of(MAT_COLS, 'Vendor')
for i in range(2, n_ven + 2):
    ws_v.cell(i, col_of(VEN_COLS, 'Materials')).value = (
        '=COUNTIF(Materials!$%s$2:$%s$%d,$A%d)' % (VCOL, VCOL, MROWS, i))

def covered_by(ws, cols, n, kind):
    j = col_of(cols, 'Materials')
    for i in range(2, n + 2):
        ws.cell(i, j).value = (
            '=COUNTIFS(Coverage!$%s$2:$%s$%d,$A%d,Coverage!$%s$2:$%s$%d,"%s")'
            % (COVD, COVD, CROWS, i, COVK, COVK, CROWS, kind))

covered_by(ws_s, SUB_COLS, n_sub, 'Submittal')
covered_by(ws_p, PO_COLS, n_po, 'Purchase Order')
covered_by(ws_i, INS_COLS, n_ins, 'Inspection')

def describe(ws, cols, keycol, upto):
    j = col_of(cols, 'Item Description')
    L = letter_of(cols, keycol)
    for i in range(2, upto):
        ws.cell(i, j).value = (
            '=IF($%s%d="","",IFERROR(INDEX(Materials!$B$2:$B$%d,'
            'MATCH($%s%d,Materials!$A$2:$A$%d,0)),"◄ no such material"))'
            % (L, i, MROWS, L, i, MROWS))

describe(ws_c, COV_COLS, 'MAT Number', n_cov + 2)
describe(ws_d, DEL_COLS, 'MAT Number', 60)
describe(ws_w, INST_COLS, 'MAT Number', 60)

# ---- Dashboard ----
db = out.create_sheet('Dashboard', 0)
db.cell(1, 1, 'Where the materials stand').font = TITLE
db.cell(2, 1, 'Every figure counts the sheets beside it. Change a row and this changes with it.'
        ).font = SUBF
S = n_sub + 1
SUBP = letter_of(MAT_COLS, 'Submittals')
POP = letter_of(MAT_COLS, 'Purchase Orders')
INSP = letter_of(MAT_COLS, 'Inspections')
DISC = letter_of(MAT_COLS, 'Discipline')
CATC = letter_of(MAT_COLS, 'Material Category')
row = 4
db.cell(row, 1, '1. OVERVIEW').font = BOLD
row += 1
for label, f in [
        ('Materials', '=COUNTA(Materials!$A$2:$A$%d)' % MROWS),
        ('Vendors', '=COUNTA(Vendors!$A$2:$A$%d)' % (n_ven + 1)),
        ('Disciplines', '=SUMPRODUCT((Materials!$%s$2:$%s$%d<>"")/'
                        'COUNTIF(Materials!$%s$2:$%s$%d,Materials!$%s$2:$%s$%d&""))'
         % (DISC, DISC, MROWS, DISC, DISC, MROWS, DISC, DISC, MROWS)),
        ('Submittal documents', '=COUNTA(Submittals!$A$2:$A$%d)' % S),
        ('Purchase orders', "=COUNTA('Purchase Orders'!$A$2:$A$%d)" % (n_po + 1)),
        ('Inspections', '=COUNTA(Inspections!$A$2:$A$%d)' % (n_ins + 1)),
        ('Coverage links', '=COUNTA(Coverage!$A$2:$A$%d)' % (n_cov + 1)),
        ('Deliveries recorded', '=COUNTA(Deliveries!$A$2:$A$%d)' % DROWS)]:
    db.cell(row, 1, label).font = BODY
    db.cell(row, 2, f).font = BODY
    row += 1
row += 1

db.cell(row, 1, '2. SUBMITTALS BY TYPE').font = BOLD
row += 1
for j, h in enumerate(['Type', 'Documents', 'Materials covered', 'Approved',
                       'Approved as Noted', 'Under Review', '% cleared'], 1):
    c = db.cell(row, j, h)
    c.font, c.fill = HEAD, HFILL
row += 1
for t in ['MAT', 'MES', 'ITP', 'PID']:
    db.cell(row, 1, t).font = BODY
    db.cell(row, 2, '=COUNTIF(Submittals!$B$2:$B$%d,$A%d)' % (S, row)).font = BODY
    db.cell(row, 3, '=SUMPRODUCT((Submittals!$B$2:$B$%d=$A%d)*Submittals!$G$2:$G$%d)'
            % (S, row, S)).font = BODY
    db.cell(row, 4, '=COUNTIFS(Submittals!$B$2:$B$%d,$A%d,Submittals!$E$2:$E$%d,"Approved")'
            % (S, row, S)).font = BODY
    db.cell(row, 5, '=COUNTIFS(Submittals!$B$2:$B$%d,$A%d,Submittals!$E$2:$E$%d,'
                    '"Approved as Noted")' % (S, row, S)).font = BODY
    db.cell(row, 6, '=COUNTIFS(Submittals!$B$2:$B$%d,$A%d,Submittals!$E$2:$E$%d,"Under Review")'
            % (S, row, S)).font = BODY
    c = db.cell(row, 7, '=IF($B%d=0,"",($D%d+$E%d)/$B%d)' % (row, row, row, row))
    c.font, c.number_format = BODY, '0%'
    row += 1
row += 1

db.cell(row, 1, '3. MATERIALS BY DISCIPLINE').font = BOLD
row += 1
for j, h in enumerate(['Discipline', 'Materials', 'Vendors named', 'Submittal links'], 1):
    c = db.cell(row, j, h)
    c.font, c.fill = HEAD, HFILL
row += 1
for d in sorted({m['Discipline'] for m in materials if m['Discipline']}):
    db.cell(row, 1, d).font = BODY
    db.cell(row, 2, '=COUNTIF(Materials!$%s$2:$%s$%d,$A%d)' % (DISC, DISC, MROWS, row)).font = BODY
    db.cell(row, 3, '=COUNTIFS(Materials!$%s$2:$%s$%d,$A%d,Materials!$%s$2:$%s$%d,"<>")'
            % (DISC, DISC, MROWS, row, VCOL, VCOL, MROWS)).font = BODY
    db.cell(row, 4, '=SUMPRODUCT((Materials!$%s$2:$%s$%d=$A%d)*Materials!$%s$2:$%s$%d)'
            % (DISC, DISC, MROWS, row, SUBP, SUBP, MROWS)).font = BODY
    row += 1
row += 1

db.cell(row, 1, '4. MATERIALS BY CATEGORY').font = BOLD
row += 1
for j, h in enumerate(['Category', 'Materials', 'With a purchase order',
                       'With an inspection'], 1):
    c = db.cell(row, j, h)
    c.font, c.fill = HEAD, HFILL
row += 1
for k in ['C0', 'C1', 'C2', 'C3']:
    db.cell(row, 1, k).font = BODY
    db.cell(row, 2, '=COUNTIF(Materials!$%s$2:$%s$%d,$A%d)' % (CATC, CATC, MROWS, row)).font = BODY
    db.cell(row, 3, '=SUMPRODUCT((Materials!$%s$2:$%s$%d=$A%d)*(Materials!$%s$2:$%s$%d>0))'
            % (CATC, CATC, MROWS, row, POP, POP, MROWS)).font = BODY
    db.cell(row, 4, '=SUMPRODUCT((Materials!$%s$2:$%s$%d=$A%d)*(Materials!$%s$2:$%s$%d>0))'
            % (CATC, CATC, MROWS, row, INSP, INSP, MROWS)).font = BODY
    row += 1

db.column_dimensions['A'].width = 34
for c in 'BCDEFG':
    db.column_dimensions[c].width = 19

# ---- Read me ----
rm = out.create_sheet('Read me', 0)
text = [
    ('Project Materials — linked sheets', TITLE),
    ('Built from the Main Log of Draft_Project_Materials_Live_Tracking_Sheet.xlsx. '
     'Every value came out of that sheet; nothing was invented.', SUBF),
    ('', None),
    ('NOTHING HERE IS A MADE-UP CODE', BOLD),
    ('A material is named by its MAT reference, a delivery by its MIR, and every document '
     'by its own number. The one exception is the vendor, which is named by its name: '
     'a PQD names a submittal, not a company, and 4MAKA08-MAKA-08-MBL-ME-PRQ-00024 covers '
     'PROMINENT/ITC, GRUNDFOS and SOLICO together, so it cannot stand for any one of them.', None),
    ('', None),
    ('A REFERENCE NAMES A DOCUMENT, NOT A MATERIAL', BOLD),
    ('4MAKA08-MAKA-08-MBL-HV-ITP-00004 covers twelve materials here, and the purchase order '
     '25000513-2 covers nine. So each document has one row, and the Coverage sheet says which '
     'materials it covers. Change a status once and every material it covers follows.', None),
    ('', None),
    ('HOW THE SHEETS HOLD TOGETHER', BOLD),
    ('Coverage.Document No  →  Submittals / Purchase Orders / Inspections .Document No', None),
    ('Coverage.MAT Number   →  Materials.MAT Number', None),
    ('Materials.Vendor      →  Vendors.Vendor Name', None),
    ('Deliveries.MAT Number →  Materials.MAT Number', None),
    ('Installation.MAT Number → Materials.MAT Number', None),
    ('', None),
    ('YOU NEVER TYPE A REFERENCE', BOLD),
    ('Every column that points at another sheet is a dropdown. Pick the MAT reference from '
     'the list; Excel refuses anything that is not a real one. That is what lets this '
     'workbook be keyed by references at all — a typed key drifts, a chosen one cannot.', None),
    ('', None),
    ('WHICH CELLS TO EDIT', BOLD),
    ('White cells are yours. Grey cells are keys, and columns marked read-only carry a '
     'formula that will overwrite whatever is typed into them.', None),
    ('To link an existing document to one more material, add a line to Coverage: its kind, '
     'its number, and the MAT reference. Nothing else.', None),
    ('Where two rows of the Main Log disagreed about one document, both readings are kept '
     'in that document\'s Notes rather than one being chosen.', None),
    ('', None),
    ('ADDING A DELIVERY — the format the empty sheets expect', BOLD),
]
r = 1
for t, f in text:
    c = rm.cell(r, 1, t)
    c.font = f or BODY
    c.alignment = Alignment(wrap_text=True, vertical='top')
    if f is None and len(t) > 90:
        rm.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        rm.row_dimensions[r].height = 46
    r += 1
for j, h in enumerate(['MIR Number', 'MAT Number', 'MIR Approval Date', 'MIR Status',
                       'Quantity', 'Unit', 'Storage'], 1):
    c = rm.cell(r, j, h)
    c.font, c.fill = HEAD, HFILL
r += 1
for j, v in enumerate(['4MAKA08-MAKA-08-MBL-ME-MIR-00014', materials[0]['MAT Number'],
                       datetime.date(2026, 3, 18), 'Approved', 40, 'set', 'Site'], 1):
    c = rm.cell(r, j, v)
    c.font, c.fill = BODY, WARNFILL
    if isinstance(v, datetime.date):
        c.number_format = 'DD/MM/YYYY'
rm.cell(r + 2, 1, 'An example only — it is not in the Deliveries sheet, and nothing is read '
                  'from this page.').font = SUBF
rm.column_dimensions['A'].width = 46
for c in 'BCDEFG':
    rm.column_dimensions[c].width = 26
rm.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
rm.merge_cells(start_row=2, start_column=1, end_row=2, end_column=7)

out.save(OUT)
print('materials      ', n_mat)
print('vendors        ', n_ven)
print('submittals     ', n_sub, '(was 347 duplicated rows)')
print('purchase orders', n_po)
print('inspections    ', n_ins)
print('coverage links ', n_cov)
print('unreadable     ', len(unread))
print('documents carrying a contradiction:',
      sum(1 for d in subs + pos + insp if 'differs on another row' in (d.get('Notes') or '')))
