"""IMAP Attachment extractor.

Works on IMAP SSL server by replacing an email after extracting the attachments.
Use keyring tool to store password. Initialize with 'keyring set imap_aex:<host> <user>'.

Usage: imap_aex [--help] [--verbose] [--debug] [--extract-dir=<e>] [--extract-only] [--folder=<f>] [--date=<d>] [options] [HOST] [USER]

Arguments:
  HOST                      IMAP host name.
  USER                      IMAP user name.


Options:
     --all                  Fetch all mail from the folder. Not recommended, prefer '--date' use.
  -c --conf=<c>             Optional configuration file, containing any of the command line options values. [Default: config.ini]
  -d --date=<d>             Date defininition, formatted as [<>]date[to][date]. Date can be y, y-m, or y-m-d.
                              - 2012-12-21 : on this day
                              - 2012-12 : on this month
                              - 2012 : on this year
                              - >2012 : since this year
                              - <2012-12-21 : before this day
                              - 2012-04 to 2012-10 : between those months

     --dry-run              Dump running information and leave the server intact.
     --debug                Debug mode: append modified messages to the server, but don't delete the source message.
     --dir-reg=<r>          Replace Regular expression to be applied before creating directories.
                            Optional replacement separated by ">>", otherwise delete the match.
                            Multiple expressions separated by "::".
                              - ^INBOX\/?                   Remove "INBOX" and "INBOX/" from subdirectory.
                              - ^INBOX$>>Inbox::INBOX\/     Replace "INBOX" by "Inbox", and ignore it as subdirectory.
                              - ^\[Gmail\]\/?               Remove "[Gmail]" and "[Gmail]/" from subdirectory.
                              - ^Drafts$>>Brouillons        Replace by translation.

  -e --extract-dir=<d>      Extract attachment to this directory. [Default: ./]
     --extract-only         Don't detach attachments, extract only.
     --flagged=<f>          Flagged/starred mail behaviour: [Default: skip]
                              - detach:     detach as a normal mail.
                              - extract:    extract only, leave message intact.
                              - skip:       don't extract, leave message intact.
  -f --folder=<f>           Mail folder. [Default: INBOX]

     --help                 Display this help message then exit.
     --inline-images        Handle inline images as attachments.
     --list                 List the server folders then exit.
     --max-size=<m>         Extract attachments bigger than this size. [Default: 100K]
     --no-subdir            Don't create folder subdirectories inside extract dir.
     --password             Prompt for password instead of looking in the keyring.
  -p --port=<p>             IMAP host SSL port. [Default: 993]
     --run                  Force running, even if dry-run found in config file.
     --thunderbird          Implements thunderbird detach mode, pointing to the extracted file local URL.
  -v --verbose              Display more information.
"""
import calendar
import imaplib
import sys
from binascii import Error as BinasciiError
import os
import re
from base64 import b64decode, b64encode
import getpass
from configparser import ConfigParser

from datetime import date, datetime
from email import message_from_bytes, policy
from email.message import EmailMessage, Message
from email.header import decode_header
from imaplib import IMAP4_SSL, IMAP4, Time2Internaldate, ParseFlags
from time import time

import keyring
from docopt import docopt, parse_defaults


class ImapAttachmentExtractor:
    def __init__(self, host, login, port=993, folder='INBOX', extract_dir="./", no_subdir=False, dir_reg=None,
                 thunderbird_mode=False, max_size='100K', flagged_action="skip", extract_only=False,
                 inline_images=False, dry_run=False, ask_password=False, debug=False, verbose=False):
        """IMAP Attachment extractor.

        :param str host: IMAP host name.
        :param str login: IMAP host login.
        :param int port: IMAP SSL port. (default: 993)
        :param str folder: Mail folder. (default: INBOX)
        :param str extract_dir: Extract directory. (default: ./)
        :param bool no_subdir: Don't create folder sub-directories inside extract dir. (default: False)
        :param str dir_reg: Optional regexp to apply when creating subdirectories.
                            Optional replacement separated by ">>", otherwise delete the match.
                            Multiple expressions separated by "::".
        :param bool thunderbird_mode: Thunderbird detach mode. (default: False)
        :param int|str max_size: Extract only attachment bigger than this size.  (default: 100K)
        :param str flagged_action: Flagged/starred message behaviour. [detach, extract, skip] (default: skip)
        :param bool extract_only: Extract only attachments, don't detach from messages. (default: False)
        :param bool inline_images: Handle inline images as attachments. (default: False)
        :param bool dry_run: Dump information, and leave the server intact. (default: False)
        :param bool ask_password: Prompt for password instead of looking in keyring. (default: False)
        :param bool debug: Debug mode, don't delete original message. (default: False)
        :param bool verbose: Display more information. (default: False)
        """

        # server init
        self.imap = None  # type: IMAP4
        self.host = host
        self.port = port
        self.login = login

        # check basic configuration
        if not self.host or not self.login:
            raise RuntimeWarning("Missing host and login configuration.")

        # max size
        if type(max_size) == int:
            self.max_size = max_size
        else:
            self.max_size = human_readable_size_to_bytes(max_size)

        # configuration
        self.folder = folder
        self.thunderbird_mode = thunderbird_mode
        self.extract_dir = os.path.abspath(extract_dir)
        self.no_subdir = no_subdir
        self.dir_reg = []
        if dir_reg is not None:
            for reg in dir_reg.split('::'):
                self.dir_reg.append(tuple(reg.split('>>')))

        subdir_path = self.folder
        if not self.no_subdir and subdir_path:
            if self.dir_reg:
                for dr in self.dir_reg:
                    m = dr[0]
                    s = dr[1] if len(dr) > 1 else ''
                    subdir_path = re.sub(m, s, subdir_path)

            self.extract_dir = os.path.join(self.extract_dir, *subdir_path.split('/'))

        self.inline_images = inline_images
        self.dry_run = dry_run
        self.flagged_action = flagged_action
        self.extract_only = extract_only
        self.ask_password = ask_password
        self.debug = debug
        self.verbose = verbose or debug
        self.gmail_mode = "gmail" in self.host
        
        # stats
        self.extracted_nb = 0
        self.extracted_from_nb = 0
        self.extracted_size = 0

        self.password = None

    def list(self):
        """List IMAP folders"""
        status, list_data = self.imap.list()
        if status != "OK":
            raise RuntimeWarning("Could not list folders")

        folders = []
        max_len = 0
        for folder in list_data:
            folder = folder.split(b' "/" ')[1].decode()
            folder = imaputf7decode(folder)
            folder = re.sub('(^"|"$)', '', folder)
            max_len = max(max_len, len(folder))
            folders.append(folder)

        print("Folder list:")
        print('  Folder'.ljust(max_len+4), 'Extract subdirectory')
        print('  '.ljust(max_len+4, '-'), ''.ljust(30, '-'))
        for folder in folders:
            extract_dir = self.extract_dir
            subdir_path = folder
            if not self.no_subdir and subdir_path:
                if self.dir_reg:
                    for dr in self.dir_reg:
                        m = dr[0]
                        s = dr[1] if len(dr) > 1 else ''
                        subdir_path = re.sub(m, s, subdir_path)

                extract_dir = os.path.join(*subdir_path.split('/'))

            print('  '+folder.ljust(max_len+2), extract_dir)

        print()

    @staticmethod
    def parse_date(date_def, date_format="imap"):
        """Parse and format a date definition.

        :param str date_def: The date(s) definition.
        :param str date_format: The date format to output. [imap, ymd] (default: imap)
        :return: A tuple of ON, SINCE, and BEFORE dates formatted according to IMAP requirements.
        :rtype: (str, str, str)
        """
        date_on = None
        date_since = None
        date_before = None

        ret_date_on = None
        ret_date_since = None
        ret_date_before = None

        if date_format == "imap":
            date_format = "%d-%b-%Y"
        elif date_format == "ymd":
            date_format = "%Y-%m-%d"

        match = re.match("^([<>])?\s*(\d{4})-?(\d{2})?-?(\d{2})?(\s*to\s*)?(\d{4})?-?(\d{2})?-?(\d{2})?", date_def)
        if not match:
            raise RuntimeWarning("Invalid date definition: %s." % date_def)

        modifier = match.group(1)
        y1, m1, d1 = match.group(2), match.group(3), match.group(4)
        to = match.group(5)
        y2, m2, d2 = match.group(6), match.group(7), match.group(8)

        if modifier == ">":
            # after
            y1 = int(y1)
            m1 = int(m1) if m1 else 1
            d1 = int(d1) if d1 else 1
            date_since = date(y1, m1, d1)

        elif modifier == "<":
            # before
            y1 = int(y1)
            m1 = int(m1) if m1 else 12
            d1 = int(d1) if d1 else calendar.monthrange(y1, m1)[1]
            date_before = date(y1, m1, d1)

        else:
            # single date
            y1 = int(y1)

            if m1 is None:
                date_since = date(y1, 1, 1)
                date_before = date(y1, 12, 31)
            elif d1 is None:
                m1 = int(m1)
                date_since = date(y1, m1, 1)
                date_before = date(y1, m1, calendar.monthrange(y1, m1)[1])
            else:
                m1 = int(m1)
                d1 = int(d1)
                date_on = date(y1, m1, d1)
            
            if to:
                # range
                y2 = int(y2)

                if not date_since:
                    date_since = date(y1, m1, d1)

                if m2 is None:
                    date_before = date(y2, 12, 31)
                elif d2 is None:
                    m2 = int(m2)
                    date_before = date(y2, m2, calendar.monthrange(y2, m2)[2])
                else:
                    m2 = int(m2)
                    d2 = int(d2)
                    date_before = date(y2, m2, d2)

        if type(date_since) == date:
            ret_date_since = date_since.strftime(date_format)

        if type(date_before) == date:
            ret_date_before = date_before.strftime(date_format)

        if type(date_on) == date:
            ret_date_on = date_on.strftime(date_format)

        return ret_date_on, ret_date_since, ret_date_before

    def extract(self, date_def=None, fetch_all=False):
        """Parse email and extract attachments.

        :param str date_def: Date definition as [<>]date[to][date] where date is year, or year-month, or year-month-day.
        :param bool fetch_all: Fetch all messages. (default: False)
        """
        if not self.dry_run:
            os.makedirs(self.extract_dir, exist_ok=True)

            if self.verbose:
                print("Create extract dir %s." % self.extract_dir)

        else:
            print("[Dry-run] Create extract dir %s." % self.extract_dir)

        folder = self.folder
        if " " in self.folder:
            folder = '"%s"' % folder

        status, select_data = self.imap.select(imaputf7encode(folder))
        if status != "OK":
            raise RuntimeWarning("Could not select %s" % folder)

        total_mail = int(select_data[0])
        print("Selected folder '%s' (%d mails)." % (self.folder, total_mail))

        if not date_def and not fetch_all:
            raise RuntimeWarning("Date criteria not found, use 'all' parameter to fetch all messages.")
        elif date_def:
            date_on, date_since, date_before = self.parse_date(date_def)
            dates = self.parse_date(date_def, "ymd")

            date_crit = []
            if date_on:
                date_crit.append('ON "%s"' % date_on)

            if date_since:
                date_crit.append('SINCE "%s"' % date_since)

            if date_before:
                date_crit.append('BEFORE "%s"' % date_before)
        else:
            date_crit = []

        # status, search_data = self.imap.search("UTF-8", 'UNDELETED', 'ON "27-aug-2018"')
        if not self.gmail_mode:
            status, search_data = self.imap.search("UTF-8", 'UNDELETED', *date_crit)
        else:
            status, search_data = self.imap.search("UTF-8", 'UNDELETED', 'X-GM-RAW', 'has:attachment', *date_crit)
        if status != "OK":
            raise RuntimeWarning("Could not search in %s" % folder)

        if search_data[0]:
            uids = search_data[0].split(b' ')  # type: list
        else:
            uids = []

        print("%d messages corresponding to search." % len(uids))

        if not uids:
            exit(0)

        to_fetch = []

        try:
            status, fetch_data = self.imap.fetch(b','.join(uids), 'BODYSTRUCTURE[PEEK]')
        except IMAP4.error:
            status, fetch_data = self.imap.fetch(b','.join(uids), 'BODYSTRUCTURE')

        if status != "OK":
            print("Could not fetch messages.")

        merge_previous = False
        previous_structure = b''

        for structure in fetch_data:  # type: bytes
            if type(structure) in (list, tuple):
                if len(structure) == 2:
                    structure = structure[0] + b'"'+structure[1]+b'"'
                else:
                    structure = b' '.join(structure)

            if not structure.endswith(b')'):
                previous_structure = b''+structure
                merge_previous = True
                continue

            if merge_previous:
                structure = previous_structure + structure
                merge_previous = False

            reg_attachment = "attachment|application"
            if self.inline_images:
                reg_attachment = reg_attachment+"|image"

            try:
                has_attachments = re.match('^.*"(%s)".*$' % reg_attachment, structure.decode("utf-8"), re.IGNORECASE) is not None
            except AttributeError as e:
                print("Failed to process message. Error: %s" % e)
                has_attachments = False
            if not has_attachments:
                continue
            uid = structure.split(b' ')[0]

            if uid:
                to_fetch.append(uid)
            else:
                # Probably an eml attachment
                pass

        print("%d messages with attachments." % len(to_fetch))
        print()

        if not to_fetch:
            exit(0)

        for uid in to_fetch:
            try:
                status, fetch_data = self.imap.fetch(uid, '(FLAGS RFC822)')
            except imaplib.IMAP4.error as e:
                print("Encountered error when reading mail uid %s: %s" % (uid, repr(e)))
                continue

            if status != "OK":
                print("Could not fetch messages")

            skip_flags = False
            for i in range(len(fetch_data)):
                if skip_flags:
                    # previous mail flags part
                    skip_flags = False
                    continue

                fetch = fetch_data[i]
                if b')' == fetch:
                    continue

                flags = ParseFlags(bytes(fetch[0]))
                if not flags and i + 1 < len(fetch_data) and type(fetch_data[i+1]) == bytes:
                    # check if flags are in the next fetch_data
                    flags = ParseFlags(fetch_data[i+1])
                    if flags:
                        skip_flags = True

                if flags:
                    flags = tuple(map(lambda x: x.decode("utf-8"), flags))

                mail = message_from_bytes(fetch[1])  # type: EmailMessage

                subject, encoding = decode_header(mail.get("Subject"))[0]
                if encoding:
                    if not encoding.startswith('unknown'):
                        subject = subject.decode(encoding)
                    else:
                        subject = str(subject)
                elif type(subject) == bytes:
                    subject = subject.decode()

                is_flagged = "\\Flagged" in flags

                mail_date = mail.get("Date", None)
                date_match = re.match("^(.*\d{4} \d{2}:\d{2}:\d{2}).*$", mail_date)
                if date_match:
                    mail_date = date_match.group(1)

                try:
                    mail_date = datetime.strptime(mail_date, "%a, %d %b %Y %H:%M:%S")
                except ValueError:
                    mail_date = datetime.strptime(mail_date, "%d %b %Y %H:%M:%S")

                if dates is not None:
                    date_ok = False
                    check_date = mail_date.strftime("%Y-%m-%d")

                    if dates[0]:
                        date_ok = check_date == dates[0]
                    else:
                        if dates[1] and not dates[2]:
                            date_ok = check_date >= dates[1]
                        elif dates[2] and not dates[1]:
                            date_ok = check_date <= dates[2]
                        elif dates[1] and dates[2]:
                            date_ok = dates[1] <= check_date <= dates[2]

                    if not date_ok:
                        if self.verbose:
                            print("\nSkip email: '%s' [%s] (possible previous extract)." % (subject, mail_date))
                        continue

                new_mail = EmailMessage()
                new_mail._headers = mail._headers

                has_alternative = False
                nb_alternative = 0
                nb_extraction = 0
                part_nb = 1

                to_print = []  # print buffer

                to_print.append("")
                to_print.append("Parsing mail: '%s' [%s]" % (subject, mail_date))
                if is_flagged and 'skip' == self.flagged_action:
                    if self.verbose:
                        to_print.append("  Skip flagged mail.")
                    continue

                for part in mail.walk():  # type: Message
                    if part.get_content_type().startswith("multipart/"):
                        if part.get_content_type() == "multipart/alternative":
                            new_mail.attach(part)  # add text/plain and text/html alternatives
                            has_alternative = True
                        continue

                    if has_alternative and nb_alternative < 2 and (part.get_content_type() == "text/plain" or part.get_content_type() == "text/html"):
                        nb_alternative = nb_alternative + 1
                        continue  # text/plain and text/html already added in multipart/alternative

                    is_attachment = part.get_content_disposition() is not None and part.get_content_disposition().startswith("attachment")
                    if not is_attachment:
                        new_mail.attach(part)
                        continue

                    part_nb = part_nb + 1

                    if not part.get_filename():
                        attachment_filename = "part.%d" % part_nb
                    else:
                        attachment_filename, encoding = decode_header(part.get_filename())[0]
                        if encoding:
                            attachment_filename = attachment_filename.decode(encoding)
                        elif type(attachment_filename) == bytes:
                            attachment_filename = attachment_filename.decode()

                    if "AttachmentDetached" in part.get("X-Mozilla-Altered", ""):
                        if self.verbose:
                            to_print.append("  Attachment '%s' already detached." % attachment_filename)
                        continue

                    if part.get("Content-Transfer-Encoding", "").lower() == "base64":
                        try:
                            attachment_content = b64decode(part.get_payload())
                        except BinasciiError:
                            to_print.append("  Error when decoding attachment '%s', leave intact." % attachment_filename)
                            new_mail.attach(part)
                            continue
                    else:
                        if isinstance(part.get_payload(), list):
                            attachment_content = part.get_payload(1).encode("utf-8")
                        else:
                            attachment_content = part.get_payload().encode("utf-8")

                    attachment_size = len(attachment_content)

                    if attachment_size < self.max_size:
                        if self.verbose:
                            to_print.append("  Attachment '%s' size (%s) is smaller than defined threshold (%s), leave intact." % (attachment_filename, human_readable_size(attachment_size), human_readable_size(self.max_size)))
                        new_mail.attach(part)
                        continue

                    filename = "%s - %s" % (mail_date.strftime("%Y-%m-%d"), attachment_filename)
                    idx = 0
                    while os.path.exists(os.path.join(self.extract_dir, filename)):
                        idx = idx + 1
                        filename = re.sub("^(\d{4}-\d{2}-\d{2}) (?:\(\d+\) )?- (.*)$", "\g<1> (%02d) - \g<2>" % idx, filename)

                    if not self.dry_run:
                        with open(os.path.join(self.extract_dir, filename), "wb") as file:
                            file.write(attachment_content)

                        to_print.append("  Extracted '%s' (%s) to '%s'" % (attachment_filename, human_readable_size(attachment_size), os.path.join(self.extract_dir, filename)))
                    else:
                        to_print.append("  [Dry-run] Extracted '%s' (%s) to '%s'" % (attachment_filename, human_readable_size(attachment_size), os.path.join(self.extract_dir, filename)))

                    self.extracted_nb = self.extracted_nb + 1
                    self.extracted_size = self.extracted_size + attachment_size

                    nb_extraction = nb_extraction + 1

                    if self.thunderbird_mode and ('detach' == self.flagged_action or not is_flagged):
                        # replace attachement by local file URL
                        headers_str = ""
                        try:
                            headers_str = "\n".join(map(lambda x: x[0]+": "+x[1], part._headers))
                        except Exception as e:
                            to_print.append("  Error when serializing headers: %s" % repr(e))

                        new_part = Message()
                        new_part._headers = part._headers
                        new_part.set_payload("You deleted an attachment from this message. The original MIME headers for the attachment were:\n%s" % headers_str)

                        new_part.replace_header("Content-Transfer-Encoding", "")
                        url_path = "file:///%s/%s" % (self.extract_dir.replace("\\", "/"), filename)
                        new_part.add_header("X-Mozilla-External-Attachment-URL", url_path)
                        new_part.add_header("X-Mozilla-Altered",  'AttachmentDetached; date=%s' % Time2Internaldate(time()))

                        new_mail.attach(new_part)

                if nb_extraction:
                    self.extracted_from_nb = self.extracted_from_nb + 1

                if to_print:
                    if nb_extraction > 0 or self.verbose:
                        print("\n".join(to_print))

                if nb_extraction > 0 and not self.extract_only and ('detach' == self.flagged_action or not is_flagged):
                    if not self.dry_run:
                        print("  Extracted %s attachment%s, replacing email." % (nb_extraction, "s" if nb_extraction > 1 else ""))
                        status, append_data = self.imap.append(imaputf7encode(folder), " ".join(flags), '', new_mail.as_bytes(policy=policy.SMTPUTF8))
                        if status != "OK":
                            print("  Could not append message to IMAP server.")
                            continue

                        if self.verbose:
                            print("  Append message on IMAP server.")
                    else:
                        print("  [Dry-run] Extracted %s attachment%s, replacing email." % (nb_extraction, "s" if nb_extraction > 1 else ""))

                        if self.verbose:
                            print("  [Dry-run] Append message on IMAP server.")

                    if not self.debug and not self.dry_run:
                        status, store_data = self.imap.store(uid, '+FLAGS', '\Deleted')
                        if status != "OK":
                            print("  Could not delete original message from IMAP server.")
                            continue

                        if self.verbose:
                            print("  Delete original message.")

                    elif self.dry_run:
                        if self.verbose:
                            print("  [Dry-run] Delete original message.")
                    else:
                        print("  Debug: would delete original message.")

                elif self.extract_only and nb_extraction > 0:
                    print("  Extracted %s attachment%s." % (nb_extraction, "s" if nb_extraction > 1 else ""))
                elif 'extract' == self.flagged_action:
                    if self.verbose:
                        print("  Flagged message, leave intact.")
                else:
                    if nb_extraction > 0 or self.verbose:
                        print("  Nothing extracted.")

        print()
        print('Extract finished.')

        if self.extracted_size > 0:
            if not self.dry_run:
                print("  Extracted %d files from %s messages, %s gain." % (self.extracted_nb, self.extracted_from_nb, human_readable_size(self.extracted_size)))
            else:
                print("  [Dry-run] Extracted %d files from %s messages, %s gain." % (self.extracted_nb, self.extracted_from_nb, human_readable_size(self.extracted_size)))

            if self.verbose:
                print("  Thunderbird headers used: %s." % ("yes" if self.thunderbird_mode else "no"))

        print()

    def connect(self):
        if self.password is None:
            if self.ask_password:
                # prompt for password
                self.password = getpass.getpass()
            else:
                # fetch password from keyring
                self.password = keyring.get_password("imap_aex:%s" % self.host, self.login)
                if not self.password:
                    raise RuntimeWarning("Password not found for user {user} on {host} . Use 'keyring set imap_aex:{host} {user}' to set.".format(host=self.host, user=self.login))

        # connect to host
        self.imap = IMAP4_SSL(host=self.host, port=self.port)
        login_status, login_details = self.imap.login(self.login, self.password)
        if login_status != 'OK':
            raise RuntimeWarning("Could not login %s on %s:%s : %s " % (self.login, self.host, self.port, login_details))

    def __enter__(self):
        self.connect()

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.imap.state == "SELECTED":
            self.imap.expunge()
            self.imap.close()
            self.imap.logout()


def human_readable_size(num, suffix='B'):
    """Human readable file size.

    :param int num: Input size to format.
    :param str suffix: The suffix to append after size unit. (default: 'B')
    :return: Formatted string representation of the size.
    :rtype: str
    """
    for unit in ['', 'K', 'M', 'G', 'T', 'P', 'E', 'Z']:
        if abs(num) < 1024.0:
            return "%3.1f%s%s" % (num, unit, suffix)
        num /= 1024.0
    return "%.1f%s%s" % (num, 'Yi', suffix)


def human_readable_size_to_bytes(size_label, suffix='B'):
    """Get the numeric size corresponding to a human readable size label.

    :param str size_label: The size label to parse.
    :param str suffix: The suffix to use. (default: 'B')
    :return: The file size.
    :rtype: int
    """
    if size_label is None:
        return None

    size_label = size_label.upper()

    match = re.match("^(\d[\d.]*)(K|M|G|T|P|E|Z)?(%s)?$" % suffix, size_label)

    if not match:
        raise SyntaxWarning("Wrong size %s." % size_label)
    if match.group(3) and match.group(3) != suffix:
        raise SyntaxWarning("Invalid suffix %s." % match.group(3))

    size, unit = float(match.group(1)), match.group(2)

    if unit in ['K', 'M', 'G', 'T', 'P', 'E', 'Z']:
        size *= 1024
    if unit in ['M', 'G', 'T', 'P', 'E', 'Z']:
        size *= 1024
    if unit in ['G', 'T', 'P', 'E', 'Z']:
        size *= 1024
    if unit in ['T', 'P', 'E', 'Z']:
        size *= 1024
    if unit in ['P', 'E', 'Z']:
        size *= 1024
    if unit in ['E', 'Z']:
        size *= 1024
    if unit in ['Z']:
        size *= 1024

    size = int(size)

    return size


def b64padanddecode(b):
    """Decode unpadded base64 data"""
    b += (-len(b) % 4) * '='  # base64 padding (if adds '===', no valid padding anyway)
    return b64decode(b, altchars='+,', validate=True).decode('utf-16-be')


def imaputf7decode(s):
    """Decode a string encoded according to RFC2060 aka IMAP UTF7.

    Minimal validation of input, only works with trusted data"""
    lst = s.split('&')
    out = lst[0]
    for e in lst[1:]:
        u, a = e.split('-', 1)  # u: utf16 between & and 1st -, a: ASCII chars folowing it
        if u == '':
            out += '&'
        else:
            out += b64padanddecode(u)
        out += a
    return out


def imaputf7encode(s):
    """"Encode a string into RFC2060 aka IMAP UTF7"""
    s = s.replace('&', '&-')
    unipart = out = ''
    for c in s:
        if 0x20 <= ord(c) <= 0x7f:
            if unipart != '':
                out += '&' + b64encode(unipart.encode('utf-16-be')).decode('ascii').rstrip('=') + '-'
                unipart = ''
            out += c
        else:
            unipart += c
    if unipart != '':
        out += '&' + b64encode(unipart.encode('utf-16-be')).decode('ascii').rstrip('=') + '-'
    return out


def parse_configuration(conf_file='config.ini'):
    """Parse configuration file.

    :param str conf_file: The configuration file.
    :return: The config parser, or None if not found.
    :rype: ConfigParser|None
    """
    if not os.path.exists(conf_file):
        return ConfigParser()

    config = ConfigParser()
    config.read(conf_file)
    return config


def main(options, defaults):
    # parse options not in configuration
    conf_file = options['--conf']  # type: str
    list_folders =  options['--list']  # type: bool
    run = options['--run']  # type: bool

    # arguments definition
    arguments = {
        ('host',                'HOST',                     '[imap] host',                      str),
        ('login',               'USER',                     '[imap] login',                     str),
        ('port',                '--port',                   '[imap] port',                      int),

        ('date_def',            '--date',                   '[parameters] date',                str),
        ('fetch_all',           '--all',                    '[parameters] all',                 bool),

        ('folder',              '--folder',                 '[parameters] folder',              str),
        ('extract_dir',         '--extract-dir',            '[parameters] extract-dir',         str),
        ('max_size',            '--max-size',               '[parameters] max-size',            str),
        ('flagged_action',      '--flagged',                '[parameters] flagged',             str),
        ('dir_reg',             '--dir-reg',                '[parameters] dir-reg',             str),

        ('no_subdir',           '--no-subdir',              '[options] no-subdir',              bool),
        ('thunderbird_mode',    '--thunderbird',            '[options] thunderbird',            bool),
        ('extract_only',        '--extract-only',           '[options] extract-only',           bool),
        ('inline_images',       '--inline-images',          '[options] inline-images',          bool),
        ('dry_run',             '--dry-run',                '[options] dry-run',                bool),
        ('ask_password',        '--password',               '[options] password',               bool),
        ('debug',               '--debug',                  '[options] debug',                  bool),
        ('verbose',             '--verbose',                '[options] verbose',                bool),
    }

    kwargs = {}

    # parse configuration file
    config = parse_configuration(conf_file)
    for a in arguments:
        # ConfigParser section and option name
        match = re.match('^\[(.*)\] (.*)$', a[2])
        section = match.group(1)
        option = match.group(2)

        if a[3] == int:
            kwargs[a[0]] = config.getint(section, option, fallback=None)
        elif a[3] == bool:
            kwargs[a[0]] = config.getboolean(section, option, fallback=None)
        else:
            kwargs[a[0]] = config.get(section, option, fallback=None)

    # calculate overwrite
    overwrite = {}
    for a in arguments:
        overwrite[a[0]] = (options.get(a[1], None) != defaults.get(a[1], None))

    # parse options
    for a in arguments:
        if kwargs[a[0]] is None or overwrite[a[0]]:
            if a[3] == int:
                kwargs[a[0]] = int(options[a[1]])
            else:
                kwargs[a[0]] = options[a[1]]

    # handle --run option
    if run and kwargs.get('dry_run', False):
        kwargs['dry_run'] = False

    # extract arguments
    extract_kwargs = {
        'date_def':     kwargs.pop('date_def'),
        'fetch_all':    kwargs.pop('fetch_all')
    }

    # Launch imap attachment extractor
    with ImapAttachmentExtractor(**kwargs) as imap:
        # folder list and exit
        if list_folders:
            imap.list()
            return

        # launch extractor
        imap.extract(**extract_kwargs)


def cli():
    options = docopt(__doc__)
    defaults = dict(map(lambda x: (x.long, x.value), parse_defaults(__doc__)))

    try:
        main(options, defaults)
    except RuntimeWarning as w:
        print("  Warning: %s" % w, file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    cli()
