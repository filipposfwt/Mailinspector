# MailInspector
A forensics tool to aid Outlook email analysis. Enables headers analysis and search for suspicious content in the email body or attached documents. The tool can be used for live collection of the emails on a running machine with Outlook installed or for offline analysis by providind a PST(Personal Storage Table) file as input. While in live collection, the keyword search is done only on the filename and attachments are saved for further investigation.

# Install 
In order to install all required dependencies please run:
```
pip install -r requirements.txt 
```
# Usage

## Live Collection
In order to use the tool for live collection is should be run in a computer with Outlook application installed. The input (-i) flag should be ommitted. In this mode the user can choose to analyze Inbox or Junk by defining the -f flag (0 corresponds to Inbox and 1 to Junk).

## Imported PST file
To use the tool for offline analysis the user should provide a PST file by using the -i flag. In this mode the analysis can only be done in the Inbox folder.
