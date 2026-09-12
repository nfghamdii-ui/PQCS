#!/usr/bin/env python3
"""
Add excel-sync.js to the tracker page.

The page is a single large file with a library embedded in it, so it is
edited in place rather than rewritten: one script tag is inserted before
the closing body tag and nothing else is touched. Running this twice
changes nothing the second time.

    python3 install.py                 # patches ./index.html
    python3 install.py path/to/page.html
"""
import sys, os, shutil, datetime

TAG = '<script src="excel-sync.js"></script>'
DEFAULT = 'index.html'


def main(path):
    if not os.path.exists(path):
        die('There is no file at %s. Put the tracker page there, or pass its path:\n'
            '    python3 install.py path/to/page.html' % path)

    with open(path, encoding='utf-8') as f:
        html = f.read()

    if 'excel-sync.js' in html:
        print('Already installed — %s loads excel-sync.js. Nothing to do.' % path)
        return

    if not os.path.exists(os.path.join(os.path.dirname(path) or '.', 'excel-sync.js')):
        print('Warning: excel-sync.js is not beside %s. The page will load, but the\n'
              '         two buttons will not appear until the file is there.\n' % path)

    i = html.rfind('</body>')
    if i < 0:
        die('That file has no </body> tag, so it does not look like the tracker page.')

    # a copy of what was there before, because this edits the original
    backup = '%s.before-excel-sync.%s' % (path, datetime.date.today().isoformat())
    if not os.path.exists(backup):
        shutil.copy2(path, backup)
        print('Kept a copy at', backup)

    with open(path, 'w', encoding='utf-8') as f:
        f.write(html[:i] + TAG + '\n' + html[i:])

    print('Installed. Open %s, sign in, and look under "More".' % path)
    print()
    print('  Download the workbook   — the whole log, 77 columns in their order')
    print('  Upload a workbook       — read that file back in')
    print('  Upload the register     — the Aconex export, read as it comes')


def die(msg):
    print(msg)
    sys.exit(1)


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else DEFAULT)
