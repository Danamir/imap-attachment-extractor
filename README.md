# imap-attachment-extractor
IMAP attachment extractor, with optional Thunderbird detach mode.


#### Prequisites

- Python 3.x
- Use keyring to store password on the system (_Optional_)


#### Installation

##### Installation from release :

 - Download and extract the latest [release](https://github.com/Danamir/imap-attachment-extractor/releases).
 - Open a terminal to the extracted directory.

##### Installation from sources :

```bash
$ curl --location https://github.com/Danamir/imap-attachment-extractor/archive/master.zip --output imap-attachment-extractor-master.zip
$ unzip imap-attachment-extractor-master.zip
$ mv imap-attachment-extractor-master/ imap-attachment-extractor/
$ cd imap-attachment-extractor
```

##### Setup :

Configure Python virtual environment :
```bash
$ python -m venv .env
$ . .env/bin/activate  # on Linux 
-or-
$ .env/Script/activate.bat  # on Windows
```

Install :
```bash
$ python setup.py develop
$ imap_aex --help
-or-
$ python imap_aex.py --help
```

Use configuration file _(optional but recommended)_ :
```bash
$ cp config.ini-dist config.ini  # on Windows use 'copy' instead of 'cp'
```

#### Running

You can use `imap_aex` either only with the CLI, or with the configuration file, or a combination of the two.
When using both, the command line will override the options found in the configuration file.

The only mandatory arguments are `HOST` and `USER` (or `[imap]` section `host` and `user` in the configuration file).

_First run with configuration file_ :  
By default the configuration template is in `dry-run` mode. You can force the execution in CLI with the `--run`
option, or comment the `dry-run=yes` configuration line.

##### Password handling

You can use `keyring` to store a password in with the system secured library. Supported on Linux / Windows / MacOS :
```bash
$ keyring set imap_aex:<HOST> <USER>
```

Or you can prompt the user for the password on run with the `--password` option.

##### Mozilla Thunderbird 'detach' mode

It is recommended to use the `--thunderbird` option if you use Mozilla Thunderbird. The extractor will then use the
extended messages headers `X-Mozilla-External-Attachment-URL` and `X-Mozilla-Altered` to link to the local extracted
file inside the modified message. This simulates the use of _Detach_ action on an attachment in Mozilla Thunderbird.

##### Running options

Please refer to the CLI documentation. The configuration entries have the same name as the CLI options, but are
placed in specific section.

_CLI only options (those options are not in the configuration file):_
 - `--conf=<c>` : Use a specific configuration file, otherwise use `config.ini`. Can be used to configure multiple
                  hosts, each with its own configuration file.
 - `--help` : Display the CLI help, then exit.
 - `--list` : List the server folders and corresponding extraction paths, then exit.
 - `--run` : Force running, even if `dry-run` found in configuration file.
 
##### CLI documentation
```
IMAP Attachment extractor.

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
                              - ^INBOX\/?                   Remove "INBOX" and "INBOX/" subdirectory completely
                              - ^INBOX$>>Inbox::INBOX\/     Replace "INBOX" by "Inbox", and ignore it as subdirectory.
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
```