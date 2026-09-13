# Material Tracker

Tracking material submittals, vendors and inspections on a construction
project, against the SEVEN procedure, with Excel on one side and Aconex
on the other.

The page itself is a single self-contained HTML file backed by Supabase.
Everything in this repository either extends it or feeds it.

---

## What is here

| File | What it is |
|---|---|
| `index.html` | The tracker page, with `excel-sync.js` already installed. |
| `cats.js` | Attachment 1: the component categories the page reads. **Bring your own.** |
| `excel-sync.js` | Reads and writes the Main Log, and reads the Aconex register. |
| `install.py` | Adds `excel-sync.js` to the page. Safe to run twice. |
| `tools/add-inbox-sheet.py` | Adds a reconciliation sheet to a workbook, for pasting a daily log by hand. |
| `tools/split-main-log.py` | Splits the flat 77-column log into linked sheets. An experiment; see below. |

`cats.js` is not committed here — it is yours and the page warns on its
own if it is missing. `index.html` is committed, and already carries the
one line that loads `excel-sync.js`.

---

## Setting up

```bash
git clone <your repo>
cd <your repo>
cp /wherever/you/keep/cats.js .
```

Then open `index.html` in a browser and sign in. Three buttons appear
under **More**.

`install.py` is only needed when the page is replaced with a newer copy
that has not been patched. It inserts one line before `</body>`, keeps a
dated copy of the original beside it, and does nothing if the line is
already there.

---

## The two directions

### Excel

**Download the workbook** writes every material out with the same
seventy-seven columns in the same order, plus a recomputed summary
sheet. **Upload a workbook** reads that file back.

The round trip is closed: exporting and re-importing reports zero
changes. Each material keeps every cell of its original row, so a column
the page has no opinion about — mock-ups, storage, the installer's name
— survives the trip and comes back out where it went in.

The workbook is read and written without any library. A workbook is a
zip of XML and the browser can already inflate a zip stream, so the
reader is `DecompressionStream` and the writer stores its entries
uncompressed. Nothing is fetched from a CDN, which matters on a site
network that blocks them. Needs Chrome, Edge, or Safari 16.4 and later.

### Aconex

**Upload the register** reads the register export as it comes out of
Aconex — the sheet named `Docs`, with the preamble above the headings.

The register is not the transmittal log. It carries one row per document
and no history, so a reference appears exactly once and the question of
which occurrence is the live one never arises.

What it does:

- A document already named somewhere in the tracker has its outcome,
  revision and date brought up to date.
- A document that is not named anywhere can be **brought in whole**. Every
  record created that way is stamped with the upload it came from and
  marked unreviewed, so it can always be told from what was already
  there, and the whole upload can be taken back in one action.
  Pre-qualifications become vendors, everything else becomes a material.
  Nothing is created without asking: the button says how many.
- A reference that shares a cell with other references is shown and not
  touched, because one cell cannot hold two different outcomes.
- `Terminated` is recorded as what it is. A file still calling a dead
  document "under review" is wrong, and waits forever.

Both sides are reduced to the same words before being compared. The file
writes `Approve`, `Aprrove` and `Approved as Noted`; the register writes
`B - Approved with Comments`, and Aconex itself writes that two ways —
with spaces around the dash and without. Comparing the words rather than
the meanings would report every document as moved, and report it again
after it had been set.

---

## The loop

1. **Upload the register** → **Do it all**. Outcomes that moved are
   written; everything the file has never seen is brought in for review.
2. **Reports** → **The general log** → a workbook in the Main Log shape.
3. Work on it in Excel: rename things, fill in manufacturers and purchase
   orders, add rows.
4. **Upload a workbook** → the edited file goes back in.

A row is matched on its reference, not on its description. The
description is the thing a person edits — merging two entries, fixing a
spelling — and a key that changes when someone tidies the file is not a
key. Rename all you like; the row is still the same row.

A row with a heading and nothing else is a section divider. A row with a
heading and no category is a material that has not been categorised, and
it is kept. Those are different things, and reading the category column
alone confused them until materials without categories existed.

Documents do not appear in a log shaped one row per material, so their
absence from that file means nothing and is not reported as missing.

---

## What is still manual

The register says what each document is and how it stands. It does not
say which material a document belongs to — there is no column for it —
so an inspection plan or a method statement arrives as a record of its
own and has to be moved onto its material by hand. The discipline code
narrows it to that trade; it does not decide it.

**Waiting to be reviewed**, under More, lists everything that came in
this way until you have been through it. Each upload can be taken back
whole from the same screen, as long as its records are still unreviewed.

A pre-qualification names a company in its title, and the title is
written by hand — `PRQ for Dar Al-Rokham - Marble Cladding work`,
`P4- Makkah-Prequalification-Sodamco-Concrete Admixtures…`. The name is
pulled out as well as it can be and the whole title is kept beside it.
Expect to correct some of them.

---

## tools/

Two generators, neither of them part of the running system.

`add-inbox-sheet.py` builds a sheet for pasting a daily transmittal log
into, back when that log had to be copied by hand. It works, and the
register upload has largely replaced it. Keep it if the register is ever
unavailable.

`split-main-log.py` splits the flat log into linked sheets keyed by the
Aconex reference, with a coverage sheet for the documents that serve
several materials at once. An experiment, not the live shape. It is here
because the finding behind it still holds: one inspection plan in this
log covers twelve materials, and one purchase order covers nine.

---

## On the Supabase key in the page

The anon key in `index.html` belongs in a public file. On its own it can
read nothing — every table refuses it until somebody signs in, and
row-level security decides what they see after that. Committing it is
fine. The service key, if you ever generate one, is not: that one never
goes near this repository.

---

## Reading and writing at size

Three thousand records behave differently from a hundred, and each of
these was found by pushing past what had been tried before.

**A table is read a page at a time.** A plain select returns a thousand
rows and says nothing about the rest. Because a save compares memory
against the database, the rows that were never read look deleted — and
the next save deletes them. Reads now page until a short page ends them.

**A save is written in batches of 150.** Two and a half thousand records
in one request is five megabytes of body and, for the conflict check, a
query string of two and a half thousand ids. Both are refused.

**Identifiers come from a counter, not from the clock.** The page's
`uid()` is the millisecond plus a random number, which collides as soon
as a hundred records are made inside one millisecond. The database then
rejects the whole batch with *ON CONFLICT DO UPDATE command cannot
affect row a second time*.

**A failed save stops and says so.** It used to pass as a message that
disappeared in four seconds, after which a reload took the work with it.

The percentage beside the project name says how far a long save has got.

**A save always ends somewhere it can name.** It used to be possible for
one to stop — a conflict raised, a question asked — and leave the word
`saving…` on the screen for ever, because the work was still unsaved and
nothing distinguished stopped from slow. There are three states now:
saving, saved, and *not saved — waiting on you*. The lock is released in
a `finally`, and a watchdog clears one that has been held two minutes.

**An unsigned change is not somebody else's.** The conflict check asked
whether the row had moved and who had moved it. Where the column naming
the author is empty — no trigger fills it — every second write to a row
looked like a stranger's, and the save stopped to ask a question nobody
could answer. A change with no name on it is now nobody's.

**Empty cells do not travel.** A material carries seventy-seven columns
and fills perhaps a dozen; the rest were being written to the database
on every save. An absent key and an empty one mean the same thing here,
so the empty ones are dropped on the way out — about seventy per cent
off the weight of a save.


---

## Documents

A method statement is not a material. It arrives as a record of its own
because no column in the register says which material it belongs to.

Inspection requests have a tab to themselves — twelve hundred of the
fifteen hundred documents in this project are inspection requests, and
leaving them in with the rest buries everything else. The remaining
kinds sit under **Documents**, separated by chips: method statements,
inspection plans, dossiers, pre-qualifications.

What a record is gets read off the middle of its reference, so a record
that came in before any of this existed is sorted out the first time the
workspace opens.

The link is made from the material, which is the thing everything hangs
off: open it, and attach the documents that serve it. One document can
serve many materials; in this project one inspection plan covers twelve.

A reference moved onto its material — an inspection request recorded as
a consignment — is still found by the next upload, so it is not created a
second time.

---

## Reports

Their own tab, at the end of the bar. Excel only — nothing is printed
from here, because every one of these ends up in somebody else's
spreadsheet anyway.

**The general log** is the Main Log shape: one row per material, with
its documents folded back into the columns the project already reads.
Several documents of one kind go into one cell separated by newlines,
which is how the original file already holds them. A document attached
to no material would vanish from a report shaped one row per material,
so those are written to a second sheet rather than dropped.

**Vendors and who brought them** comes back too. Every column on it can
be corrected — the name, the kind, the country, the pre-qualification
number, the ISO certificate and its expiry — and sent back. The first
column is the only thing that must not be touched: it is how a row finds
its way home after everything else about it has changed.

Names pulled out of Aconex titles need correcting. `P4-
Makkah-Prequalification-Sodamco-Concrete Admixtures & Mortar Based
Solutions` yields something recognisable but wrong, and titles are
written by hand so no amount of parsing will fix that. Correct them in
the sheet, or click the name on the record — which is now editable,
having been the one thing on the page that was not.

**A two-week look-ahead** is there as well. More as the shapes you
actually use arrive.

---

## Who brought whom

A pre-qualification qualifies a company, and that company is often the
subcontractor rather than the factory. `4MAKA08-…-ME-PRQ-00024` is
Faisal Abdullah Awad Binladen Contracting; the three manufacturers the
file hangs off it are the makers that subcontractor brought.

So a vendor records who brought it onto the project, and a subcontractor
lists what it brought. Reading the Main Log fills this in on its own:
`FIRST FIX` brought twenty-two of the manufacturers in this file.

It is one field and not a list. No manufacturer here was brought by two
subcontractors, and if one later works through somebody else, the
question this answers is still who brought them first.

---

## Finding things

The search box reads every reference a record holds, not the three
fields it used to. An inspection plan number, a method statement, the
number on a consignment — typing any of them finds the record. It used
to return the whole list, which looks like a search that ran and found
everything.

Vendors have a second row of chips for where the pre-qualification
itself stands: approved, approved as noted, under review, rejected, none
recorded. The row above it reads the whole qualification road — a vendor
approved but with no ISO date on file shows as *in progress* there, and
that is a different question with a different answer.

The board counts materials. A method statement has a status, not a road;
twelve hundred inspection requests each sitting at "step 1 of 2, waiting
on you" turned a real number into four and a half thousand, which is no
number at all.

---

## The table

Its own tab, and five tables kept apart rather than folded into one: a
method statement and a manufacturer share almost no columns, and a table
that holds both is mostly empty.

Every column filters, and the filters add together — category C3 *and*
discipline electrical *and* status approved. A column whose values
repeat offers them as a list with counts beside them, so nobody has to
remember how a status is spelt; a column of references takes typed text.
Click a heading to sort, click again to reverse.

**Columns** chooses what is visible. Materials start with nine of their
eighteen; the rest are a tick away.

A screenful is drawn and the rest follows as it is scrolled to, because
nine hundred rows built at once is a pause nobody asked for.

**Excel** exports what the filters left, with the columns on screen.
**Print** prints the same thing — the filter boxes and the tabs drop
away on paper.

---

## Visits

"Visits to the factory while the material is being made" — the words
were plural and the form held one, with the field labelled *Last visit*.
An in-process inspection happens every fortnight and a final test can be
repeated; keeping only the newest makes a handover file that cannot be
defended.

Those two steps keep a list now, the way deliveries do. The newest visit
is also written onto the step itself, because that is what the road
reads and what one column of the log can carry. A visit already recorded
before this existed becomes the first entry in its list rather than
being left behind.

## Where an approval lives

An inspection agency is not pre-qualified, it is approved — clause
2.2.17 — and its road has no pre-qualification step at all. The
reference was being written to one anyway, so the record read as
untouched while its number sat in a field nothing looked at. Every
reader and writer now asks the same question about which step holds it.

Names pulled out of pre-qualification titles are wrong about a quarter
of the time, down from nearly two fifths. They will never be right:
`JAZEERA PAINTSInterior & Exterior Painting` has no separator between
the company and the work. Correct them in the vendor sheet, where the
whole title sits in the Scope column beside the name — or click the name
on the record.

---

## Not getting in the way

**The number across the middle of the screen** was the paged read
reporting itself — and it reported itself on the quiet reload that
happens when you come back to the tab, which is to say while you were
working, because you had been working somewhere else for a minute. A
background read is nobody's business and says nothing now.

**That reload also threw away where you were reading**, because a redraw
replaces the whole pane. It waits for a real gap now — a quarter of an
hour, and never while a sheet is open or there is unsaved work — and a
redraw of the same record keeps its place. Open a different record and
it starts at the top, as it should.
