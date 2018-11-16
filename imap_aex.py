"""IMAP Attachment extractor.

Works on IMAP SSL server by replacing an email after extracting the attachments.
Use keyring tool to store password. Initialize with 'keyring set imap_aex:<host> <user> <password>'.

Usage: imap_aex [--help] [--verbose] [--debug] [--extract-dir=<e>] [--extract-only] [--folder=<f>] [--date=<d>] [options] [HOST] [USER]

Arguments:
  HOST                      IMAP host name.
  USER                      IMAP user name.


Options:
     --all                  Fetch all mail. Not recommended, prefer --date usage.
  -c --conf=<c>             Optional configuration file, containing any of the command line options values. [Default: config.ini]
  -d --date=<d>             Date defininition, formatted as [<>]date[to][date]. Date can be y, y-m, or y-m-d. Examples:
                              - 2012-12-21 : on this day
                              - 2012-12 : on this month
                              - 2012 : on this year
                              - >2012 : since this year
                              - <2012-12-21 : before this day
                              - 2012-04 to 2012-10 : between those months

     --dry-run              Dump running information and leave the server intact.
     --debug                Debug mode, don't delete messages.
  -e --extract-dir=<d>      Extract attachment to this directory. [Default: ./]
     --extract-only         Don't detach attachments, only extract.
     --flagged=<f>          Flagged/starred mail behaviour: [Default: skip]
                              - detach:     detach as a normal mail.
                              - extract:    extract only, leave message intact.
                              - skip:       don't extract, leave message intact.

  -f --folder=<f>           Mail folder. [Default: INBOX]
     --ignore-inbox-subdir  When creating subdirectories, ignore INBOX as first directory.
     --inline-images        Handle inline images as attachments.
     --list                 List the server folders and exit.
     --max-size=<m>         Extract attachments bigger than this size. [Default: 100K]
     --no-subdir            Don't create subdirectories inside extract dir, corresponding to mail folder.
     --password             Prompt for password instead of looking in keyring.
  -p --port=<p>             IMAP host SSL port. [Default: 993]
     --thunderbird          Implements thunderbird detach mode, pointing to the extracted file local URL.
  -v --verbose              Display more information.
"""
import calendar
import sys
from binascii import Error as BinasciiError
import os
import re
from base64 import b64decode
import getpass
from configparser import ConfigParser

from datetime import date, datetime
from email import message_from_bytes
from email.message import EmailMessage, Message
from email.header import decode_header
from imaplib import IMAP4_SSL, IMAP4, Time2Internaldate, ParseFlags
from time import time

import keyring
from docopt import docopt


class ImapAttachmentExtractor:
    def __init__(self, host, login, port=993, folder='INBOX', extract_dir="./", use_subdir=True, ignore_inbox_subdir=False,
                 thunderbird_mode=True, max_size='100K', flagged_action="skip", extract_only=False,
                 extract_inline_images=False, dry_run=False, ask_password=False, debug=False, verbose=False):
        """IMAP Attachment extractor.

        :param str host: IMAP host name.
        :param str login: IMAP host login.
        :param int port: IMAP SSL port. (default: 993)
        :param str folder: Mail folder. (default: INBOX)
        :param str extract_dir: Extract directory. (default: ./)
        :param bool use_subdir: Create folder sub-directories inside extract dir. (default: True)
        :param bool ignore_inbox_subdir: When creating subdirectories, ignore INBOX as first directory. (default: False)
        :param bool thunderbird_mode: Thunderbird detach mode. (default: True)
        :param int|str max_size: Extract only attachment bigger than this size.  (default: 100K)
        :param str flagged_action: Flagged/starred message behaviour. [detach, extract, skip] (default: skip)
        :param bool extract_only: Extract only attachments, don't detach from messages. (default: False)
        :param bool extract_inline_images: Handle inline images as attachments. (default: False)
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
        self.subdir_path = list(map(lambda x: x.capitalize(), self.folder.split("/")))
        self.thunderbird_mode = thunderbird_mode
        self.extract_dir = os.path.abspath(extract_dir)
        self.use_subdir = use_subdir
        self.ignore_inbox_subdir = ignore_inbox_subdir
        if self.use_subdir and self.subdir_path:
            if self.ignore_inbox_subdir and len(self.subdir_path) > 1 and self.subdir_path[0] == 'Inbox':
                self.subdir_path.pop(0)

            self.extract_dir = os.path.join(self.extract_dir, *self.subdir_path)
        self.extract_inline_images = extract_inline_images
        self.dry_run = dry_run
        self.flagged_action = flagged_action
        self.extract_only = extract_only
        self.ask_password = ask_password
        self.debug = debug
        self.verbose = verbose
        
        # stats
        self.extracted_nb = 0
        self.extracted_size = 0

        self.password = None

    def list(self):
        """List IMAP folders"""
        status, list_data = self.imap.list()
        if status != "OK":
            raise RuntimeWarning("Could not list folders")

        print("Folder list:")
        for f in list_data:
            print('  '+f.split(b' "/" ')[1].decode('utf-8'))

        print()

    @staticmethod
    def parse_date(date_def):
        """Parse and format a date definition.

        :param str date_def: The date(s) definition.
        :return: A tuple of ON, SINCE, and BEFORE dates formatted according to IMAP requirements.
        :rtype: (str, str, str)
        """
        date_on = None
        date_since = None
        date_before = None

        ret_date_on = None
        ret_date_since = None
        ret_date_before = None

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
            ret_date_since = date_since.strftime("%d-%b-%Y")

        if type(date_before) == date:
            ret_date_before = date_before.strftime("%d-%b-%Y")

        if type(date_on) == date:
            ret_date_on = date_on.strftime("%d-%b-%Y")

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

        status, select_data = self.imap.select(folder)
        if status != "OK":
            raise RuntimeWarning("Could not select %s" % folder)

        total_mail = int(select_data[0])
        print("Selected folder '%s' (%d mails)." % (self.folder, total_mail))

        if not date_def and not fetch_all:
            raise RuntimeWarning("Date criteria not found, use 'all' parameter to fetch all messages.")
        elif date_def:
            date_on, date_since, date_before = self.parse_date(date_def)

            date_crit = []
            if date_on:
                date_crit.append('ON "%s"' % date_on)

            if date_since:
                date_crit.append('SINCE "%s"' % date_since)

            if date_before:
                date_crit.append('BEFORE "%s"' % date_before)
        else:
            date_crit = []

        # status, search_data = self.imap.search("UTF-8", 'UNDELETED', 'SINCE "%s"' % since)
        # status, search_data = self.imap.search("UTF-8", 'UNDELETED', 'ON "15-nov-2018"')
        # status, search_data = self.imap.search("UTF-8", 'UNDELETED', 'ON "31-oct-2018"')
        # status, search_data = self.imap.search("UTF-8", 'UNDELETED', 'ON "26-aug-2018"')
        # status, search_data = self.imap.search("UTF-8", 'UNDELETED', 'ON "27-aug-2018"')

        status, search_data = self.imap.search("UTF-8", 'UNDELETED', *date_crit)
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

        status, fetch_data = self.imap.fetch(b','.join(uids), 'BODYSTRUCTURE[PEEK]')
        if status != "OK":
            print("Could not fetch messages.")

        for structure in fetch_data:  # type: bytes
            multipart = b"BODYSTRUCTURE (((" in structure
            if not multipart:
                continue  # skip simple messages

            reg_attachment = "attachment|application"
            if self.extract_inline_images:
                reg_attachment = reg_attachment+"|image"

            has_attachments = re.match('^.*"(%s)".*$' % reg_attachment, structure.decode("utf-8")) is not None
            if not has_attachments:
                continue

            to_fetch.append(structure.split(b' ')[0])

        print("%d messages with attachments." % len(to_fetch))
        print()

        if not to_fetch:
            exit(0)

        status, fetch_data = self.imap.fetch(b','.join(to_fetch), '(FLAGS RFC822)')
        if status != "OK":
            print("Could not fetch messages")

        for fetch in fetch_data:
            if fetch == b')':
                continue

            uid = fetch[0].split(b' ')[0]

            flags = ParseFlags(bytes(fetch[0]))
            if flags:
                flags = tuple(map(lambda x: x.decode("utf-8"), flags))

            mail = message_from_bytes(fetch[1])  # type: EmailMessage

            subject, encoding = decode_header(mail.get("Subject"))[0]
            if encoding:
                subject = subject.decode(encoding)
            elif type(subject) == bytes:
                subject = subject.decode()

            is_flagged = "\\Flagged" in flags

            mail_date = mail.get("Date", None)
            date_match = re.match("^(.*\d{4} \d{2}:\d{2}:\d{2}).*$", mail_date)
            if date_match:
                mail_date = date_match.group(1)
            mail_date = datetime.strptime(mail_date, "%a, %d %b %Y %H:%M:%S")

            new_mail = EmailMessage()
            new_mail._headers = mail._headers

            has_alternative = False
            nb_alternative = 0
            nb_extraction = 0
            part_nb = 1

            print("\nParsing mail: '%s' [%s]" % (subject, mail_date))
            if is_flagged and 'skip' == self.flagged_action:
                print("  Skip flagged mail.")
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
                    print("  Attachment '%s' already detached." % attachment_filename)
                    continue

                if part.get("Content-Transfer-Encoding", "") == "base64":
                    try:
                        attachment_content = b64decode(part.get_payload())
                    except BinasciiError:
                        print("  Error when decoding attachment '%s', leave intact." % attachment_filename)
                        new_mail.attach(part)
                        continue
                else:
                    attachment_content = part.get_payload().encode("utf-8")

                attachment_size = len(attachment_content)

                if attachment_size < self.max_size:
                    print("  Attachment '%s' size (%s) is smaller than defined threshold (%s), leave intact." % (attachment_filename, human_readable_size(attachment_size), human_readable_size(self.max_size)))
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

                    print("  Extracted '%s' (%s) to '%s'" % (attachment_filename, human_readable_size(attachment_size), os.path.join(self.extract_dir, filename)))
                else:
                    print("  [Dry-run] Extracted '%s' (%s) to '%s'" % (attachment_filename, human_readable_size(attachment_size), os.path.join(self.extract_dir, filename)))

                self.extracted_nb = self.extracted_nb + 1
                self.extracted_size = self.extracted_size + attachment_size

                nb_extraction = nb_extraction + 1

                if self.thunderbird_mode and ('detach' == self.flagged_action or not is_flagged):
                    # replace attachement by local file URL
                    headers_str = ""
                    try:
                        headers_str = "\n".join(map(lambda x: x[0]+": "+x[1], part._headers))
                    except Exception as e:
                        print("  Error when serializing headers: %s" % repr(e))

                    new_part = Message()
                    new_part._headers = part._headers
                    new_part.set_payload("You deleted an attachment from this message. The original MIME headers for the attachment were:\n%s" % headers_str)

                    new_part.replace_header("Content-Transfer-Encoding", "")
                    new_part.add_header("X-Mozilla-External-Attachment-URL", "file:///%s/%s" % (self.extract_dir.replace("\\", "/"), filename))
                    new_part.add_header("X-Mozilla-Altered",  'AttachmentDetached; date=%s' % Time2Internaldate(time()))

                    new_mail.attach(new_part)

            # print(new_mail)

            if nb_extraction > 0 and not self.extract_only and ('detach' == self.flagged_action or not is_flagged):
                if not self.dry_run:
                    print("  Extracted %s attachment%s, replacing email." % (nb_extraction, "s" if nb_extraction > 1 else ""))
                    status, append_data = self.imap.append(self.folder, " ".join(flags), '', str(new_mail).encode("utf-8"))
                    # status, append_data = self.imap.append(self.folder, '', '', str(new_mail).encode("utf-8"))
                    if status != "OK":
                        print("  Could not append message to IMAP server.")
                        continue

                    if self.verbose:
                        print("  Append message on IMAP server.")
                else:
                    print("  [Dry-run] Extracted %s attachment%s, replacing email." % (nb_extraction, "s" if nb_extraction > 1 else ""))
                    print("  [Dry-run] Append message on IMAP server.")

                if not self.debug and not self.dry_run:
                    status, store_data = self.imap.store(uid, '+FLAGS', '\Deleted')
                    if status != "OK":
                        print("  Could not delete original message from IMAP server.")
                        continue

                    if self.verbose:
                        print("  Delete original message.")

                elif self.dry_run:
                    print("  [Dry-run] Delete original message.")
                else:
                    print("  Debug: would delete original message.")

            elif self.extract_only:
                print("  Extracted %s attachment%s." % (nb_extraction, "s" if nb_extraction > 1 else ""))
            elif 'extract' == self.flagged_action:
                print("  Flagged message, leave intact.")
            else:
                print("  Nothing extracted.")

        print()
        print('Extract finished.')

        if self.extracted_size > 0:
            if not self.dry_run:
                print("Extracted %d files, %s gain." % (self.extracted_nb, human_readable_size(self.extracted_size)))
            else:
                print("[Dry-run] Extracted %d files, %s gain." % (self.extracted_nb, human_readable_size(self.extracted_size)))

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


def main(options):
    # parse options not in configuration
    conf_file = options['--conf']  # type: str
    list_folders =  options['--list']  # type: bool
    
    # parse configuration file
    config = parse_configuration(conf_file)

    host =                  config.get('imap', 'host', fallback=None)
    login =                 config.get('imap', 'login', fallback=None)
    port =                  config.getint('imap', 'port', fallback=None)

    date_def =              config.get('parameters', 'date', fallback=None)
    fetch_all =             config.getboolean('parameters', 'all', fallback=None)

    folder =                config.get('parameters', 'folder', fallback=None)
    extract_dir =           config.get('parameters', 'extract-dir', fallback=None)
    max_size =              config.get('parameters', 'max-size', fallback=None)
    flagged_action =        config.get('parameters', 'flagged', fallback=None)

    use_subdir =            not config.getboolean('options', 'no-subdir', fallback=None)
    ignore_inbox_subdir =   config.getboolean('options', 'ignore-inbox-subdir', fallback=None)
    thunderbird_mode =      config.getboolean('options', 'thunderbird', fallback=None)
    extract_only =          config.getboolean('options', 'extract-only', fallback=None)
    inline_images =         config.getboolean('options', 'inline-images', fallback=None)
    dry_run =               config.getboolean('options', 'dry-run', fallback=None)
    ask_password =          config.getboolean('options', 'password', fallback=None)

    debug =                 config.getboolean('options', 'debug', fallback=None)
    verbose =               config.getboolean('options', 'verbose', fallback=None)
        
    # parse options        
    host =                  options['HOST'] if host is None else host  # type: str
    login =                 options['USER'] if login is None else login  # type: str
    port =                  int(options['--port']) if port is None else port  # type: int

    date_def =              options['--date'] if date_def is None else date_def  # type: str
    fetch_all =             options['--all'] if fetch_all is None else fetch_all  # type: bool

    folder =                options['--folder'] if folder is None else folder  # type: str
    extract_dir =           options['--extract-dir'] if extract_dir is None else extract_dir  # type: str
    max_size =              options['--max-size'] if max_size is None else max_size  # type: str
    flagged_action =        options['--flagged'] if flagged_action is None else flagged_action  # type: str

    use_subdir =            not options['--no-subdir'] if use_subdir is None else use_subdir  # type: bool
    ignore_inbox_subdir =   options['--ignore-inbox-subdir'] if ignore_inbox_subdir is None else ignore_inbox_subdir  # type: bool
    thunderbird_mode =      options['--thunderbird'] if thunderbird_mode is None else thunderbird_mode  # type: bool
    extract_only =          options['--extract-only'] if extract_only is None else extract_only  # type: bool
    inline_images =         options['--inline-images'] if inline_images is None else inline_images  # type: bool
    dry_run =               options['--dry-run'] if dry_run is None else dry_run  # type: bool
    ask_password =          options['--password'] if ask_password is None else ask_password  # type: bool

    debug =                 options['--debug'] if debug is None else debug  # type: bool
    verbose =               options['--verbose'] if verbose is None else verbose  # type: bool

    with ImapAttachmentExtractor(
        host=host, login=login, port=port,
        folder=folder, extract_dir=extract_dir, max_size=max_size,
        flagged_action=flagged_action, use_subdir=use_subdir, ignore_inbox_subdir=ignore_inbox_subdir,
        thunderbird_mode=thunderbird_mode, extract_only=extract_only,
        extract_inline_images=inline_images, dry_run=dry_run,
        ask_password=ask_password, debug=debug, verbose=verbose
    ) as imap:

        if list_folders:
            imap.list()
            return

        imap.extract(date_def, fetch_all)


def cli():
    options = docopt(__doc__)
    try:
        main(options)
    except RuntimeWarning as w:
        print("  Warning: %s" % w, file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    cli()