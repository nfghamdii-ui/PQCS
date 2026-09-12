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
| `index.html` | The tracker page. **Not in this repo — bring your own.** |
| `cats.js` | Attachment 1: the component categories the page reads. **Bring your own.** |
| `excel-sync.js` | Reads and writes the Main Log, and reads the Aconex register. |
| `install.py` | Adds `excel-sync.js` to the page. Safe to run twice. |
| `tools/add-inbox-sheet.py` | Adds a reconciliation sheet to a workbook, for pasting a daily log by hand. |
| `tools/split-main-log.py` | Splits the flat 77-column log into linked sheets. An experiment; see below. |

`index.html` and `cats.js` are not committed here because the page
carries a compiled copy of the Supabase client inside it — a quarter of
a megabyte that no one should be retyping. Copy your working page into
the repository root and commit it from there.

---

## Setting up

```bash
git clone <your repo>
cd <your repo>
cp /wherever/you/keep/index.html .
cp /wherever/you/keep/cats.js .
python3 install.py
```

Then open `index.html` in a browser and sign in. Three buttons appear
under **More**.

`install.py` inserts one line before `</body>` and keeps a dated copy of
the original beside it. Running it again does nothing.

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
- A document that is not named anywhere is **shown and left alone**. No
  column in the register says which material it belongs to, and guessing
  that is worse than asking.
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

A new material submittal is a new material, and the register carries its
title, discipline and category. That much could be built.

A new inspection plan or method statement belongs to a material that
already exists — and *which* one is not in any column of the register.
The discipline code narrows it to that trade, but it does not decide it.
That part stays with a person.

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
