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

## The order of operations

Load the Main Log first, then the register. The register has nothing to
compare against until the materials are in.

1. **Upload a workbook** → your Main Log → Apply.
2. **Upload the register** → the Aconex export → read it → Apply.
3. **Export the whole list** from that screen for the documents that need
   linking by hand.

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
