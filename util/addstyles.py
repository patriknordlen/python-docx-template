#!/usr/bin/env python

import re
import zipfile
import sys
import os
import shutil
from tempfile import mkdtemp
from argparse import ArgumentParser

def main(args):
    argparser = ArgumentParser(description='A tool for patching a docx file with a set of chosen styles ' \
                                           'so that they can be referenced within the document.xml part of ' \
                                           'the docx file.')
    argparser.add_argument('-i', '--infile', required=True, help='docx file to patch')
    argparser.add_argument('-s', '--stylefile', required=True, help='File containing the set of styles that should be added ' \
                                                     '(<w:style> elements)')
    argparser.add_argument('-o', '--outfile', help='Patched docx file to write as output')
    argparser.add_argument('-r', '--replace', help='Write patched results back to original docx file', action='store_true')

    args = argparser.parse_args()

    if args.outfile is None and args.replace is None:
        argparser.print_usage()
        sys.exit(0)

    docxfile = args.infile
    tmpdir = mkdtemp()

    try:
        z = zipfile.ZipFile(docxfile)
        z.extractall(tmpdir)
        z.close()
    except:
        raise Exception("ERROR: this doesn't look like a valid docx file!")

    customstylefile = args.stylefile
    try:
        with open(customstylefile) as stylefile:
            customstyles = stylefile.read().rstrip()
    except:
        raise Exception("ERROR: couldn't read style file %s" % customstylefile)

    with open("%s/word/styles.xml" % tmpdir,"r+") as docxstylefile:
        docxstyles = docxstylefile.read()
        if re.search('styleId="Heading1"', docxstyles):
            raise Exception("File has already been patched with custom styles!")

        docxstyles = re.sub(r'(</w:styles>)',r'%s\1' % customstyles, docxstyles)
        docxstylefile.seek(0)
        docxstylefile.write(docxstyles)

    if args.replace:
        outfile = args.infile
    else:
        outfile = args.outfile

    z = zipfile.ZipFile(outfile, "w", compression=zipfile.ZIP_DEFLATED)
    os.chdir(tmpdir)
    for dirname, subdirs, files in os.walk('.'):
        z.write(dirname)
        for filename in files:
            z.write(os.path.join(dirname, filename))
    z.close()

    shutil.rmtree(tmpdir)
    print "All patched up and done! Resulting file is %s" % outfile

main(sys.argv)