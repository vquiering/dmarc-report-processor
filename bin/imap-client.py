#!/usr/bin/env python3
#
# Copyright (c) 2014, Yahoo! Inc.
# Copyrights licensed under the New BSD License. See the
# accompanying LICENSE.txt file for terms.
#
# Author Binu P. Ramakrishnan
# Created 09/12/2014
# 
# An easy to use python script that
#  1. Dumps emails from a given IMAP folder to a local folder.  
#  2. A option to dump mail attachments only (not contents) 
#  3. Support search criteria. Eg. all mail SINCE '10-Sep-2014'
#

import sys
import os
import email
from imaplib import IMAP4
import argparse
import getpass

# global object
args = ""


def vprint(msg):
    global args
    if args.quiet: return
    if args.verbose: print(msg)


def process_mailbox(mail):
    # dumps emails/attachments in the folder to output directory.
    global args
    count = 0
    vprint(args.search)
    mail.select(args.folder)

    ret, data = mail.search(None, '(' + args.search + ')')
    if ret != 'OK':
        print("ERROR: No messages found", file=sys.stderr)
        return 1

    if not os.path.exists(args.outdir):
        os.makedirs(args.outdir)

    for num in data[0].split():
        ret, data = mail.fetch(num, '(RFC822)')
        if ret != 'OK':
            print("ERROR getting message from IMAP server", num, file=sys.stderr)
            return 1

        if not args.attachmentsonly:
            vprint("Writing message " + num)
            fp = open('%s/%s.eml' % (args.outdir, num), 'wb')
            fp.write(data[0][1])
            fp.close()
            count = count + 1
            print(args.outdir + "/" + num + ".eml")

        else:
            m = email.message_from_bytes(data[0][1])
            if m.get_content_maintype() == 'multipart' or \
                    m.get_content_type() == 'application/zip' or \
                    m.get_content_type() == 'application/gzip':
                for part in m.walk():

                    # find the attachment part
                    if part.get_content_maintype() == 'multipart': continue
                    if part.get('Content-Disposition') is None: continue

                    # save the attachment in the given directory
                    filename = part.get_filename()
                    if not filename: continue
                    filename = args.outdir + "/" + filename
                    with open(filename, 'wb') as fp:
                        fp.write(part.get_payload(decode=True))
                    print(filename)
                    count = count + 1
                    if os.path.exists(filename):
                        mail.store(num, '+FLAGS', '\\Deleted')
                    else:
                        print("ERROR file was not written", filename, file=sys.stderr)

    if args.attachmentsonly:
        print("\nTotal attachments downloaded: ", count)
    else:
        print("\nTotal mails downloaded: ", count)


def main():
    global args
    options = argparse.ArgumentParser(epilog='Example: \
  %(prog)s  -s imap.example.com -u dmarc@example.com -f inbox -o ./mymail -S \"SINCE \\\"8-Sep-2014\\\"\" -P ./paswdfile')
    options.add_argument("-v", "--verbose", help="increase output verbosity", action="store_true")
    options.add_argument("--attachmentsonly", help="download attachments only", action="store_true")
    options.add_argument("--disablereadonly", help="enable state changes on server; Default readonly",
                         action="store_true")
    options.add_argument("--quiet", help="supress all comments (stdout)", action="store_true")
    options.add_argument("-s", "--host", help="imap server; eg. imap.mail.yahoo.com", required=True)
    options.add_argument("-p", "--port", help="imap server port; Default is 143", default=143)
    options.add_argument("-u", "--user", help="user's email id", required=True)
    options.add_argument("-f", "--folder", help="mail folder from which the mail to retrieve; Default is INBOX",
                         default="INBOX")
    options.add_argument("-o", "--outdir", help="directory to output", required=True)
    options.add_argument("-S", "--search",
                         help="search criteria, defined in IMAP RFC 3501; eg. \"SINCE \\\"8-Sep-2014\\\"\"",
                         default="ALL")
    options.add_argument("-P", "--pwdfile",
                         help="A file that stores IMAP user password. If not set, the user is prompted to provide a passwd")
    args = options.parse_args()

    # redirect stdout to /dev/null
    if args.quiet:
        f = open(os.devnull, 'w')
        sys.stdout = f

    if args.pwdfile:
        infile = open(args.pwdfile, 'r')
        firstline = infile.readline().strip()
        args.pwd = firstline
    else:
        args.pwd = getpass.getpass()

    mail = IMAP4(args.host, args.port)
    mail.starttls()
    mail.login(args.user, args.pwd)
    ret, data = mail.select(args.folder, True)
    if ret == 'OK':
        vprint("Processing mailbox: " + args.folder)
        if process_mailbox(mail):
            mail.expunge()
            mail.close()
            mail.logout()
            sys.exit(1)

        mail.close()
    else:
        print("ERROR: Unable to open mailbox ", ret, file=sys.stderr)
        mail.logout()
        sys.exit(1)

    mail.logout()


# entry point
if __name__ == "__main__":
    main()
