#!/usr/bin/python3

import getopt
import os
from colorama import Fore, Style
from colorama import init as colorama_init
import win32com.client
from tabulate import tabulate
import re
import sys
from datetime import date
import pypff

colorama_init(autoreset=True)
print("""        
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
""")

# Function to save attachments in live collection mode.
def saveAttachments(email: object, foldername, input):

    for attachedFile in email.Attachments:
        try:
            filename = attachedFile.FileName
            today = date.today()
            date_formatted = today.strftime("%d/%m/%Y")
            path = os.getcwd() + "\\attachments_" + date_formatted.replace("/", "_") + "_" + foldername + "\\"
            os.makedirs(path, exist_ok=True)
            attachedFile.SaveAsFile(path + filename)
        except Exception as e:
            print(e)

#Function to print the email analysis
def printAnalytics(messages,foldername,input,output,links):

    #Iterate over every message.
    for message in messages:
        #Headers initialiazation.
        meta = {
            "message-id": "-",
            "spf-record": False,
            "dkim-record": False,
            "dmarc-record": False,
            "spoofed": False,
            "ip-address": "-",
            "ip_name": "-",
            "ip_country": "-",
            "ip_city": "-",
            "ip_registrar": "-",
            "sender-client": "-",
            "spoofed-mail": "-",
            "dt": "-",
            "content-type": "-",
            "mime-type": "-"
        }

        #Getting headers according to method used, MAPI or via PST file.
        if input == "not specified":
            headers = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F")
        else:
            headers = message.transport_headers

        #Catching an exception causes when no headers are found in a PST file.
        try:
            headers_list = headers.splitlines()
        except Exception as e:
            print("No headers found! Skipping message...")
            continue


        # Analyzing headers
        for h in headers_list:
            # Message ID
            if re.search("Message-ID:", h):
                meta["message-id"] = h.strip("Message-ID: ")

            if (re.search("spf=pass", h)):
                meta["spf-record"] = True

            if (re.search("dkim=pass", h)):
                meta["dkim-record"] = True

            if (re.search("dmarc=pass", h)):
                meta["dmarc-record"] = True

            if (re.search("does not designate", h)):
                meta["spoofed"] = True

            if (re.search("(\d{1,3}\.){3}\d{1,3}", h)):
                ip = re.search("(\d{1,3}\.){3}\d{1,3}", h)
                meta["ip-address"] = str(ip.group())

            if (re.search("Reply-to:", h)):
                meta["spoofed-mail"] = h

            if (re.search("Date:", h)):
                meta["dt"] = h.strip("Date: ")

            if (re.search("Content-Type:", h)):
                meta["content-type"] = h.strip("Content-Type")

            if (re.search("MIME-Version:", h)):
                meta["mime-type"] = h.strip("MIME-Version:")

        # Getting email body,subject,sender and address in a string
        if input == "not specified":
            body = message.body
            subject = message.subject
            sender = message.sender.Name
            address = message.sender.Address
        else:
            try:
                body = message.plain_text_body.decode()
                subject = message.subject
                sender = message.sender_name
                address = "-" #imported pst's does not contain an address field
            except Exception as e:
                print("Email body not found! Leaving it blank...")
                body = "-"

        #Regex to find all links within a message's body
        urls = re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', body)

        #Initialize attachments list
        attachments = []
        # In live collection we download each attachment for further analysis.
        if input == "not specified":
            attachedFile: object
            for attachedFile in message.Attachments:
                attachment_name = attachedFile.FileName
                attachments.append(attachment_name)
        else:
        # When importing messages from a file we check the attachment's content directly.
            if message.number_of_attachments > 0:
                for i in range(message.number_of_attachments):
                    attachment = message.get_attachment(i)
                    size = message.get_attachment(i).get_size()
                    attachment_content = (message.get_attachment(i).read_buffer(size)).decode('ascii', errors='ignore')
                    attachments.append(attachment_content)

        #Printing results and outputting simultaneously into a file as specified.
        with open(output, "a") as f:
            print(
                Fore.MAGENTA + "\n=========================Results========================================================\n" + Style.RESET_ALL)
            print(
                Fore.MAGENTA + "\n=========================Results========================================================\n" + Style.RESET_ALL,
                file=f)

            print(Fore.GREEN + "[+] Message ID: " + meta["message-id"])
            print(Fore.GREEN + "[+] Message ID: " + meta["message-id"], file=f)

            if (meta["spf-record"]):
                print(Fore.GREEN + "[+] SPF Records: PASS")
                print(Fore.GREEN + "[+] SPF Records: PASS", file=f)
            else:
                print(Fore.RED + "[+] SPF Records: FAIL")
                print(Fore.RED + "[+] SPF Records: FAIL", file=f)

            if (meta["dkim-record"]):
                print(Fore.GREEN + "[+] DKIM: PASS")
                print(Fore.GREEN + "[+] DKIM: PASS", file=f)
            else:
                print(Fore.RED + "[+] DKIM: FAIL")
                print(Fore.RED + "[+] DKIM: FAIL", file=f)

            if (meta["dmarc-record"]):
                print(Fore.GREEN + "[+] DMARC: PASS")
                print(Fore.GREEN + "[+] DMARC: PASS", file=f)
            else:
                print(Fore.RED + "[+] DMARC: FAIL")
                print(Fore.RED + "[+] DMARC: FAIL", file=f)

            if (meta["spoofed"] and (not meta["spf-record"]) and (not meta["dkim-record"]) and (
            not meta["dmarc-record"])):
                print(Fore.RED + "[+] Spoofed Email Received:")
                print(Fore.RED + "[+] Spoofed Email Received:", file=f)
                print(Fore.RED + "\t[+] Email Address: " + meta["spoofed-mail"])
                print(Fore.RED + "\t[+] Email Address: " + meta["spoofed-mail"], file=f)
                print(Fore.RED + "\t[+] IP-Address:  " + meta["ip-address"])
                print(Fore.RED + "\t[+] IP-Address:  " + meta["ip-address"], file=f)
            else:
                print(Fore.GREEN + "[+] Authentic Email Received:")
                print(Fore.GREEN + "[+] Authentic Email Received:", file=f)
                print(Fore.GREEN + "\t[+] IP-Address:  " + meta["ip-address"])
                print(Fore.GREEN + "\t[+] IP-Address:  " + meta["ip-address"], file=f)

            print(Fore.GREEN + "[+] Provider: " + meta["sender-client"])
            print(Fore.GREEN + "[+] Provider: " + meta["sender-client"], file=f)
            print(Fore.GREEN + "[+] Content-Type: " + meta["content-type"])
            print(Fore.GREEN + "[+] Content-Type: " + meta["content-type"], file=f)
            print(Fore.GREEN + "[+] MIME-Version: " + meta["mime-type"])
            print(Fore.GREEN + "[+] MIME-Version: " + meta["mime-type"], file=f)
            print(Fore.GREEN + "[+] Date and Time: " + meta["dt"])
            print(Fore.GREEN + "[+] Date and Time: " + meta["dt"], file=f)
            print(Fore.GREEN + "[+] Subject: " + subject)
            print(Fore.GREEN + "[+] Subject: " + subject, file=f)
            print(Fore.GREEN + "[+] From: " + sender + " <" + address + ">")
            print(Fore.GREEN + "[+] From: " + sender + " <" + address + ">", file=f)
            print(Fore.GREEN + "[+] Body:")
            print(Fore.GREEN + "[+] Body:", file=f)
            if links:
                print(Fore.GREEN + "[+] Urls:")
                print(Fore.GREEN + "[+] Urls:", file=f)
                for url in urls:
                    print(Fore.GREEN + "\t "+url)
                    print(Fore.GREEN + "\t "+url, file=f)

            for keyword in bkeyword:
                if re.search(keyword, body):

                    print(Fore.RED + "\t'" + Fore.RED + keyword + "' was found in email body")
                    print(Fore.RED + "\t'" + Fore.RED + keyword + "' was found in email body", file=f)
                else:
                    print(Fore.GREEN + "\t'" + keyword + "' was not found in email body")
                    print(Fore.GREEN + "\t'" + keyword + "' was not found in email body", file=f)

            print(Fore.GREEN + "[+] Attachments:")
            print(Fore.GREEN + "[+] Attachments:", file=f)

            if (len(attachments) > 0):
                if input == "not specified":
                    for attachment in attachments:
                        for keyword in akeyword:
                            if re.search(keyword, attachment):
                                print(Fore.RED + "\t" + attachment + Fore.RED + "(Attachment's filename contains " + Fore.RED + keyword + Fore.RED + ")")
                                print(Fore.RED + "\t" + attachment + Fore.RED + "(Attachment's filename contains " + Fore.RED + keyword + Fore.RED + ")",file=f)

                            else:
                                print(Fore.GREEN + "\t" + attachment)
                                print(Fore.GREEN + "\t" + attachment, file=f)
                        saveAttachments(message, foldername, input)
                else:
                    for attachment in attachments:
                        for keyword in akeyword:
                            if re.search(keyword,message.get_attachment(0).read_buffer(size).decode('ascii', errors='ignore')):
                                print(
                                    Fore.RED + "\t" + "(Attachment contains " + Fore.RED + keyword + Fore.RED + ")")
                                print(
                                    Fore.RED + "\t" + "(Attachment contains " + Fore.RED + keyword + Fore.RED + ")",
                                    file=f)
                            else:
                                print(Fore.GREEN + "\t" + "(Attachment does not contain specified keyword")
                                print(Fore.GREEN + "\t" + "(Attachment does not contain specified keyword ", file=f)
            else:
                print(Fore.GREEN + "\tNone")
                print(Fore.GREEN + "\tNone", file=f)

#Main function
def main(argv):
    global akeyword
    global bkeyword
    global searchfolder
    links = False
    pst_file = "not specified"
    output = "not specified"
    try:
        opts, args = getopt.getopt(argv, "f:b:a:o:i:lh")
    except getopt.GetoptError:
        print('Wrong format: mailinspector.py -m <method> -f <folder> -b <keyword> -a <keyword> -i <inputfile> -l ')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print('Usage: mailinspector.py [options]\n')
            print('OPTIONS\n')
            print('-f, --folder <0 or 1>:')
            print('\t Specify which folder to search, 0 corresponds to Inbox and 1 to Junk. Default is set to Inbox.\n')
            print('-b, --bkeyword <body search keyword>:')
            print('\t Specify a keyword to search in the body and attachments of emails\n')
            print('-a, --akeyword <attachment search keyword> :')
            print('\tSpecify a keyword to search in the email body. The tool can be used as a live response tool. It is fetching\n'
                '\temails direcly by using the Outlook API (MAPI).\n')
            print('-l, --links:')
            print('\t Use this flag to display URLs found inside an email.')
            print('-o, --output:')
            print('\t Specify an output file to save the results.')
            print('-pst, <filename>\n')
            print('\t Specify a .pst backup file to analyze. Without this option the tool performs a live collection from outlook api.\n')
            sys.exit()
        elif opt in ("-f," "--folder"):
            if arg == "0":
                searchfolder = 6
            elif arg == "1":
                searchfolder = 23
        elif opt in ("-b", "--bkeyword"):
            bkeyword = arg.split(",")
        elif opt in ("-a", "--akeyword"):
            akeyword = arg.split(",")
        elif opt in ("-o", "--output"):
            output = arg
        elif opt in ("-i","--input"):
            pst_file = arg
        elif (opt == "-l"):
            links = True
    if searchfolder == 6:
        foldername = 'Inbox'
    elif searchfolder == 23:
        foldername = "Junk"
    print(tabulate([[foldername,bkeyword, akeyword, links,pst_file,output]], ["Folder","Body keyword(s)", "Attachments keyword(s)", "Display URLs","Input File","Output File"],
                   tablefmt="grid"))

    #If no file is given as input.
    if pst_file == "not specified":
        Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Getting the desired folder,23 corresponds to Junk, 6 to inbox
        Inbox = Outlook.GetDefaultFolder(searchfolder)
        messages = Inbox.Items

    #If a file is given as input. PST analysis only supports Inbox folder.
    else:
        opst = pypff.open(pst_file)
        root = opst.get_root_folder()
        folder = root.get_sub_folder(1)
        inbox = folder.get_sub_folder(2)
        email_count = inbox.get_number_of_sub_items()
        messages = []
        for i in range(email_count):
            msg = inbox.get_sub_item(i)
            messages.append(msg)

    printAnalytics(messages,foldername,pst_file,output,links)


if __name__ == "__main__":
    main(sys.argv[1:])


