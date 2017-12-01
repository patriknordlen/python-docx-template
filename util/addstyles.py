#!/usr/bin/env python

import re
import zipfile
import sys
import os
import shutil
from tempfile import mkdtemp

def main(args):
    if len(args) < 3:
        print "Usage: %s <file to patch> <stylefile>" % args[0]
        return

    docxfile = args[1]
    tmpdir = mkdtemp()

    try:
        z = zipfile.ZipFile(docfile)
        z.extractall(tmpdir)
        z.close()
    except:
        raise "ERROR: this doesn't look like a valid docx file!"

    customstylefile = args[2]
    try:
        with open(customstylefile) as stylefile:
            customstyles = stylefile.read().rstrip()
    except:
        raise "ERROR: couldn't read style file %s" % args[2]

    with open("%s/word/styles.xml" % tmpdir,"r+") as docxstylefile:
        docxstyles = docxstylefile.read()
        if re.search('styleId="Heading1"', docxstyles):
            raise "File has already been patched with custom styles!"

        docxstyles = re.sub(r'(</w:styles>)',r'%s\1' % customstyles, docxstyles)
        docxstylefile.write(docxstyles)

    outfile = "%s_patched%s" % os.path.splitext(docxfile)
    z = zipfile.ZipFile(outfile, "w", compression=zipfile.ZIP_DEFLATED)
    os.chdir(tmpdir)
    for dirname, subdirs, files in os.walk('.'):
        z.write(dirname)
        for filename in files:
            z.write(os.path.join(dirname, filename))
    z.close()

    shutil.rmtree(tmpdir)
    print "All patched up and done!"

main(sys.argv)