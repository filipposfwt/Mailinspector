# Mailinspector
A command-line forensics tool to aid Outlook email analysis. Enables headers analysis and search for suspicious content in the email body or attached documents. The tool can be used for live collection of the emails on a running machine with Outlook installed or for offline analysis by providind a PST(Personal Storage Table) file as input. While in live collection, the keyword search is done only on the filename and attachments are saved for further investigation.

# Install 
In order to install all required dependencies please run:
```
pip install -r requirements.txt 
```
# Usage
To view the available options, one can run:
```
python mail_inspector.py -h
```

```      
        ..--""|
        |     |
        | .---'                                _  _   _                                 _               
  (\-.--| |---------.                         (_)| | (_)                               | |              
 / \) \ | |          \       _ __ ___    __ _  _ | |  _  _ __   ___  _ __    ___   ___ | |_  ___   _ __ 
 |:.  | | |           |     | '_ ` _ \  / _` || || | | || '_ \ / __|| '_ \  / _ \ / __|| __|/ _ \ | '__|
 |:.  | |o|           |     | | | | | || (_| || || | | || | | |\__ \| |_) ||  __/| (__ | |_| (_) || |   
 |:.  | `"`           |     |_| |_| |_| \__,_||_||_| |_||_| |_||___/| .__/  \___| \___| \__|\___/ |_|   
 |:.  |_ __  __ _  __ /                                             | |
         |=`|                                                       |_|
         |=_|
                 
A forensics tool to aid Outlook email analysis. Enables headers analysis and search for suspicious content in the email body or attached documents.
Attachments are saved for further investigation. The tool can be used offline by specifying a pst file or for live collection by omitting it.

Usage: mailinspector.py [options]

OPTIONS

-f, --folder <0 or 1>:
	 Specify which folder to search, 0 corresponds to Inbox and 1 to Junk. Default is set to Inbox.

-b, --bkeyword <body search keyword>:
	 Specify a keyword to search in the body and attachments of emails

-a, --akeyword <attachment search keyword> :
	Specify a keyword to search in the email body. The tool can be used as a live response tool. It is fetching
	emails direcly by using the Outlook API (MAPI).

-l, --links:
	 Use this flag to display URLs found inside an email.
-o, --output:
	 Specify an output file to save the results.
-i, <input>

	 Specify a .pst backup file to analyze. Without this option the tool performs a live collection from outlook api.
```



## Live Collection
In order to use the tool for live collection is should be run in a computer with Outlook application installed. The input (-i) flag should be ommitted. In this mode the user can choose to analyze Inbox or Junk by defining the -f flag (0 corresponds to Inbox and 1 to Junk).

## Imported PST file
To use the tool for offline analysis the user should provide a PST file by using the -i flag. In this mode the analysis can only be done in the Inbox folder.

# Example 
In the following example the tool was used in live collection mode by running:
```
python mail_inspector.py -f 0 -b click,money -a docx,pdf -o email_dump.txt
```
It analyzed the Inbox folder and search the body of each message for the keywords "click" and "money" and the attachments for "docx" or "pdf" in their name. We also specified a text file to save the results for each message. In the following picture, the results for a message are shown:

![alt text](https://github.com/filipposfwt/Mailinspector/blob/main/screenshots/results_analysis.png?raw=true)

All attachments are saved in a folder on the same directory:


![alt text](https://github.com/filipposfwt/Mailinspector/blob/main/screenshots/attachments_folder.png?raw=true)

![alt text](https://github.com/filipposfwt/Mailinspector/blob/main/screenshots/attachments.png?raw=true)

The results are written on the comman-line window as well as in the email_dump.txt which was specified:

![alt text](https://github.com/filipposfwt/Mailinspector/blob/main/screenshots/email_dump.png?raw=true)
