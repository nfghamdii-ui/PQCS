#!/usr/bin/env python3
"""
The morning reconciliation sheet.

The daily log is a log of movements, not of documents: the same
reference comes round again each time something happens to it.
4MAKA08-MAKA-08-MBL-CC-PRQ-00001 appears on the third of October as
CloseOut By MBL and again on the eighth as CloseOut By SAPL. So the
sheet answers two questions, not one — is this reference in my file,
and has its outcome moved since I last wrote it down.

Two sheets are added and nothing existing is touched. Main Log keeps
its seventy-seven columns in their order, its table and its formulas,
because the tracker reads that sheet by position.
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

SRC = '/mnt/user-data/uploads/Draft_Project_Materials_Live_Tracking_Sheet.xlsx'
OUT = '/mnt/user-data/outputs/Project Materials Live Tracking Sheet.xlsx'

LAST = 300          # how far down the Main Log every pasted row looks
ROWS = 10000        # how many log lines can be pasted at once
TOP = 6             # first row of the paste area
END = TOP + ROWS - 1
SWITCH = '$P$3'

# every column of the Main Log that can hold a reference
REF_COLS = ['K', 'O', 'T', 'AG', 'AI', 'AM', 'AQ', 'BB', 'BF', 'BJ', 'BP']

# The thirteen words the log's Outcome Review Status can hold, and what each
# one means for the file. Five are a verdict on the document. Seven are a step
# in the routing — who closed the transmittal, whether it was answered — and
# writing those into a status column would overwrite a real outcome with a
# piece of postal history, so they map to nothing. Terminated is neither: the
# document is dead, and a file still calling it pending is wrong.
#
# This lives in a sheet rather than in a formula because it is a judgement,
# not a fact, and the judgement is yours to change.
STATUS_MAP = [
    ('Approved',            'Approved',           'a verdict'),
    ('Approved as Noted',   'Approved as Noted',  'a verdict'),
    ('Rejected',            'Rejected',           'a verdict'),
    ('Revise & Resubmit',   'Revise & Resubmit',  'a verdict'),
    ('Under Review',        'Under Review',       'a verdict'),
    ('Terminated',          'Terminated',         'the document is dead — do not leave it pending'),
    ('CloseOut By DAR',     '',                   'who closed the transmittal, not a verdict'),
    ('CloseOut By MBL',     '',                   'who closed the transmittal, not a verdict'),
    ('CloseOut By SAPL',    '',                   'who closed the transmittal, not a verdict'),
    ('For Information',     '',                   'why it was issued, not a verdict'),
    ('For Use',             '',                   'why it was issued, not a verdict'),
    ('Responded',           '',                   'the reply was sent, not a verdict'),
    ('Reviewed',            '',                   'read, but nothing decided'),
]

# The file spells approval six ways — Approve, Aprrove, Appoved, Aproved.
# Both sides are read through this before being compared, while each cell
# still shows what it actually says.
FILE_SPELLING = [
    ('Approved', 'Approved'), ('Approve', 'Approved'), ('Aprrove', 'Approved'),
    ('Appoved', 'Approved'), ('Aproved', 'Approved'),
    ('Approved as Noted', 'Approved as Noted'),
    ('Approved with comments', 'Approved as Noted'),
    ('Under Review', 'Under Review'), ('UR', 'Under Review'),
    ('Revise & Resubmit', 'Revise & Resubmit'), ('Resubmit', 'Revise & Resubmit'),
    ('Rejected', 'Rejected'), ('Terminated', 'Terminated'),
]

# where each kind of document keeps its outcome, so the log's status can be
# set beside the one already written down
STATUS_COL = [('-PRQ-', 'M'), ('-MAT-', 'R'), ('-MES-', 'AL'), ('-ITP-', 'AP'),
              ('-TRN-', 'AT'), ('-MIR-', 'BL'), ('-WIR-', 'BR')]

wb = load_workbook(SRC)
log = wb['Main Log']

BODY = Font(name='Arial', size=10)
HEAD = Font(name='Arial', size=10, bold=True, color='FFFFFF')
TITLE = Font(name='Arial', size=14, bold=True)
NOTE = Font(name='Arial', size=10, italic=True, color='555555')
SMALL = Font(name='Arial', size=9, color='6E7783')
BIG = Font(name='Arial', size=18, bold=True)
HFILL = PatternFill('solid', fgColor='1F2937')
PASTEF = PatternFill('solid', fgColor='FFFFFF')
LOCK = PatternFill('solid', fgColor='F5F6F8')
NEWF = PatternFill('solid', fgColor='FDF3E2')
MOVEDF = PatternFill('solid', fgColor='E9F2FA')
DEMO = PatternFill('solid', fgColor='FFF8E6')
THIN = Border(bottom=Side(style='thin', color='D8DCE2'))

# ---------------------------------------------------------------- Index
if 'Index' in wb.sheetnames:
    del wb['Index']
ix = wb.create_sheet('Index')
for col, head in (('A', 'Row'), ('B', 'Every reference on that row of the Main Log')):
    ix[col + '1'] = head
    ix[col + '1'].font, ix[col + '1'].fill = HEAD, HFILL
ix['D1'] = ('Built by formula from the Main Log. Nothing is typed here and only the Inbox '
            'reads it. The sheet is hidden.')
ix['D1'].font = NOTE
joined = '&" "&'.join("'Main Log'!$%s%%d" % c for c in REF_COLS)
for i in range(2, LAST + 1):
    ix.cell(i, 1).value = '=IF(\'Main Log\'!$A%d="","",%d)' % (i, i)
    ix.cell(i, 2).value = (
        '=IF(\'Main Log\'!$A%d="","",SUBSTITUTE(SUBSTITUTE(%s&" ",CHAR(10)," "),CHAR(13)," "))'
        % (i, joined % tuple([i] * len(REF_COLS))))
    ix.cell(i, 1).font = ix.cell(i, 2).font = BODY
ix.column_dimensions['A'].width = 8
ix.column_dimensions['B'].width = 90
ix.freeze_panes = 'A2'
ix.sheet_state = 'hidden'

# ------------------------------------------------------------ Status map
if 'Status map' in wb.sheetnames:
    del wb['Status map']
sm = wb.create_sheet('Status map')
sm['A1'] = 'What the log says'
sm['B1'] = 'What it means in your file'
sm['C1'] = 'Why'
for col in 'ABC':
    sm[col + '1'].font, sm[col + '1'].fill = HEAD, HFILL
for i, (a, b, why) in enumerate(STATUS_MAP, 2):
    sm.cell(i, 1, a).font = BODY
    sm.cell(i, 2, b).font = BODY
    sm.cell(i, 3, why).font = NOTE
    for col in range(1, 4):
        sm.cell(i, col).border = THIN
    if not b:
        sm.cell(i, 2).fill = LOCK
sm['E1'] = 'What your file says'
sm['F1'] = 'The same thing, spelt once'
for col in 'EF':
    sm[col + '1'].font, sm[col + '1'].fill = HEAD, HFILL
for i, (a, b) in enumerate(FILE_SPELLING, 2):
    sm.cell(i, 5, a).font = BODY
    sm.cell(i, 6, b).font = BODY
    sm.cell(i, 5).border = sm.cell(i, 6).border = THIN
sm.cell(len(FILE_SPELLING) + 3, 5,
        'The log and the file do not spell approval the same way. Without this the '
        'comparison would report a change on nearly every row.').font = NOTE
sm.column_dimensions['E'].width = 26
sm.column_dimensions['F'].width = 26
sm.cell(len(STATUS_MAP) + 3, 1,
        'Leave the middle column empty for anything that is not a verdict. A status the log '
        'uses that is missing from this list is reported in the Inbox rather than ignored.'
        ).font = NOTE
sm.column_dimensions['A'].width = 24
sm.column_dimensions['B'].width = 24
sm.column_dimensions['C'].width = 52
sm.freeze_panes = 'A2'
MAPLAST = len(STATUS_MAP) + 40          # room to add more without touching a formula

# ---------------------------------------------------------------- Inbox
if 'Inbox' in wb.sheetnames:
    del wb['Inbox']
inb = wb.create_sheet('Inbox', 0)

inb['A1'] = "This morning's log"
inb['A1'].font = TITLE
inb['A2'] = ('Paste the log straight in, all seven columns, starting at A6. The seven columns '
             'to the right fill themselves. Turn Matching off while you edit the Main Log, '
             'and on again to read the next paste.')
inb['A2'].font = NOTE
inb.merge_cells('A2:J2')
inb.row_dimensions[2].height = 30

counts = [
    ('Pasted', '=COUNTA($B$%d:$B$%d)' % (TOP, END)),
    ('New', '=COUNTIF($H$%d:$H$%d,"NEW*")' % (TOP, END)),
    ('Outcome moved', '=COUNTIF($N$%d:$N$%d,"the outcome moved*")' % (TOP, END)),
    ('Already in the file', '=COUNTIF($H$%d:$H$%d,"already*")' % (TOP, END)),
    ('Superseded', '=COUNTIF($H$%d:$H$%d,"superseded*")' % (TOP, END)),
    ('Unknown status', '=COUNTIF($M$%d:$M$%d,"◄*")' % (TOP, END)),
]
for j, (label, f) in enumerate(counts):
    col = 1 + j * 2
    inb.cell(4, col, label).font = SMALL
    c = inb.cell(3, col, f)
    c.font = BIG
    c.alignment = Alignment(horizontal='left')

sw = inb['P3']
sw.value = 'On'
sw.font = Font(name='Arial', size=14, bold=True)
sw.fill = PatternFill('solid', fgColor='E8F5EE')
sw.alignment = Alignment(horizontal='center')
inb['P4'].value = 'Matching'
inb['P4'].font = SMALL
inb['P4'].alignment = Alignment(horizontal='center')
dv = DataValidation(type='list', formula1='"On,Off"', allow_blank=False, showErrorMessage=True)
dv.errorTitle = 'On or Off'
dv.error = 'Turn matching On to read a paste, Off while you edit the Main Log.'
inb.add_data_validation(dv)
dv.add('P3')

HEADERS = ['Received or Isssued Date', 'Document No.', 'Title / Description',
           'Document Revision', 'Document Type', 'Discipline', 'Outcome Review Status',
           'Verdict', 'Times found', 'Main Log row', 'Which material',
           'Outcome in your file', 'Outcome from the log', 'What to do', 'Note']
for j, h in enumerate(HEADERS, 1):
    c = inb.cell(5, j, h)
    c.font, c.fill = HEAD, HFILL
    c.alignment = Alignment(vertical='center', wrap_text=True)
inb.row_dimensions[5].height = 32

IXB = 'Index!$B$2:$B$%d' % LAST


def chain(tests, r, otherwise='""'):
    """A written-out chain of IFs rather than an array formula: Excel and
       LibreOffice agree on this one without argument."""
    out = ''
    for needle, result in tests:
        out += 'IF(ISNUMBER(SEARCH("%s",$B%d)),%s,' % (needle, r, result)
    return out + otherwise + ')' * len(tests)


for r in range(TOP, END + 1):
    for j in range(1, 8):                       # the seven pasted columns
        c = inb.cell(r, j)
        c.font, c.fill, c.border = BODY, PASTEF, THIN
        c.alignment = Alignment(vertical='top', wrap_text=(j == 3))

    # A trailing space closes the comparison so that ...TRN-1 is not found
    # inside ...TRN-17. The row is looked up first and the count taken only
    # if something was found, so a morning of new references costs one walk
    # of the index rather than two.
    inb.cell(r, 10).value = (
        '=IF(OR($B{r}="",{sw}<>"On"),"",IFERROR(MATCH("*"&TRIM($B{r})&" *",{ix},0)+1,""))'
        .format(r=r, ix=IXB, sw=SWITCH))
    inb.cell(r, 9).value = (
        '=IF(OR($B%d="",$J%d=""),0,COUNTIF(%s,"*"&TRIM($B%d)&" *"))' % (r, r, IXB, r))
    inb.cell(r, 8).value = (
        '=IF($B{r}="","",IF({sw}<>"On","— matching is off",'
        'IF(COUNTIF($B{nxt}:$B${end},$B{r})>0,"superseded further down",'
        'IF($J{r}="","NEW — add it","already in the file"))))'
        .format(r=r, nxt=r + 1, end=END, sw=SWITCH))
    inb.cell(r, 11).value = (
        '=IF($J%d="","",INDEX(\'Main Log\'!$A$1:$A$%d,$J%d))' % (r, LAST, r))

    # the outcome already written down, from whichever column this kind of
    # document keeps it in
    inb.cell(r, 12).value = '=IF($J{r}="","",{c})'.format(
        r=r, c=chain([(k, "INDEX('Main Log'!$%s:$%s,$J%d)" % (col, col, r))
                      for k, col in STATUS_COL], r))

    # read through the Status map, so a word the log uses and this sheet has
    # never seen is reported rather than quietly treated as nothing
    inb.cell(r, 13).value = (
        '=IF($G{r}="","",IFERROR(INDEX(\'Status map\'!$B$2:$B${m},'
        'MATCH(TRIM($G{r}),\'Status map\'!$A$2:$A${m},0))&"","◄ not in the Status map"))'
        .format(r=r, m=MAPLAST))

    # The file writes Approve, Aprrove and Appoved for the same verdict, so
    # both sides are read through the Status map before being compared.
    # Comparing the raw text would report ninety documents as moved.
    same = ("IFERROR(INDEX('Status map'!$F$2:$F${m},MATCH(TRIM($L{r}),"
            "'Status map'!$E$2:$E${m},0)),TRIM($L{r}))").format(r=r, m=MAPLAST)
    inb.cell(r, 14).value = (
        '=IF($B{r}="","",IF($H{r}="NEW — add it",{new},'
        'IF(LEFT($M{r},1)="◄","the log uses a status this sheet does not know — add it to '
        'the Status map",'
        'IF(AND($H{r}="already in the file",$M{r}<>"",$L{r}<>"",TRIM($M{r})<>{same}),'
        '"the outcome moved: "&$L{r}&" → "&$M{r},""))))'
        .format(r=r, same=same, new=chain(
            [('-MAT-', '"a new material — add a row to the Main Log"'),
             ('-PRQ-', '"a vendor document — put it in the PQD columns of the rows it covers"')],
            r, '"add it to the row of the material it belongs to"')))

    for j in range(8, 16):
        c = inb.cell(r, j)
        c.font, c.border = BODY, THIN
        c.alignment = Alignment(vertical='top', wrap_text=(j == 14))
        if j < 15:
            c.fill = LOCK

# four lines from a real morning, so the sheet is answering before it is touched
demo = [
    ('08-Oct-2023', '4MAKA08-MAKA-08-MBL-ME-MAT-00020',
     'MAT for Chemical Dosing System for Chilled Network / Al Maymanah', '00',
     'Material Approval', 'Mechanical', 'Approved as Noted',
     'example — in the file, but the file says Approved'),
    ('08-Oct-2023', '4MAKA08-MAKA-08-MBL-HV-ITP-00004',
     'ITP for HVAC equipment', '00', 'Inspection & Test Plan', 'HVAC', 'Approved',
     'example — one plan shared by twelve materials'),
    ('03-Oct-2023', '4MAKA08-MAKA-08-MBL-CC-PRQ-00001',
     'P4- Makkah-Prequalification-Sodamco-Concrete Admixtures & Mortar Based Solutions '
     'for Construction-Rev.00', '00', 'Pre-Qualification', 'Civil Works (Concrete)',
     'CloseOut By MBL', 'example — not in the file'),
    ('08-Oct-2023', '4MAKA08-MAKA-08-MBL-CC-PRQ-00001',
     'P4- Makkah-Prequalification-Sodamco-Concrete Admixtures & Mortar Based Solutions '
     'for Construction-Rev.00', '00', 'Pre-Qualification', 'Civil Works (Concrete)',
     'CloseOut By SAPL', 'example — the same document again, five days later'),
]
for k, line in enumerate(demo):
    r = TOP + k
    for j, v in enumerate(line[:7], 1):
        inb.cell(r, j).value = v
        inb.cell(r, j).fill = DEMO
    n = inb.cell(r, 15, line[7])
    n.font, n.fill, n.border = NOTE, DEMO, THIN
inb.cell(TOP + len(demo) + 1, 15,
         '↑ delete these four rows before the first real paste').font = NOTE

widths = {'A': 17, 'B': 36, 'C': 52, 'D': 11, 'E': 18, 'F': 20, 'G': 19, 'H': 21, 'I': 12,
          'J': 12, 'K': 46, 'L': 20, 'M': 20, 'N': 44, 'O': 40, 'P': 12}
for col, w in widths.items():
    inb.column_dimensions[col].width = w
inb.freeze_panes = 'C6'

inb.conditional_formatting.add(
    'A%d:O%d' % (TOP, END),
    FormulaRule(formula=['$H%d="NEW — add it"' % TOP], fill=NEWF, stopIfTrue=False))
inb.conditional_formatting.add(
    'A%d:O%d' % (TOP, END),
    FormulaRule(formula=['LEFT($N%d,17)="the outcome moved"' % TOP], fill=MOVEDF,
                stopIfTrue=False))

wb.save(OUT)
print('saved:', OUT)
print('sheets:', wb.sheetnames)
print('Main Log still', log.max_column, 'columns ×', log.max_row, 'rows')
print('Table1:', log.tables['Table1'].ref if 'Table1' in log.tables else 'MISSING')
