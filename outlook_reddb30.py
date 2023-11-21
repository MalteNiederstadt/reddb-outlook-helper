import win32com.client
import pyperclip
import re

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
folder = outlook.Folders.Item("audiencesubscription-data")
access_anfragen = folder.Folders.Item("Dashboard Request Access Anfragen Digital")
reddb = access_anfragen.Folders.Item("RedDB3.0")
messages = reddb.Items




mail_adressen = []

while True:
    user_input = input("Drücken Sie bitte die 1 für FunkeMedien-Adressen, Drücken Sie bitte die 2 für alle anderen Adressen: ")
    if user_input == "1":
        pattern = r'\b[A-Za-z0-9._%+-]+@funkemedien\.de\b'
        break
    elif user_input == "2":
        pattern = r'\b[A-Za-z0-9._%+-]+@(?!funkemedien\.de)[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
        break
    else:
        print('Ungültige Eingabe')


email_addresses = []
for message in messages:
    email_body = message.Body   
    # Use regular expression to find email addresses matching the pattern
    email_addresses = re.findall(pattern, email_body.lower())

    if email_addresses:
        for email_address in email_addresses:
            mail_adressen.append(email_address)
mail_adressen = list(set(mail_adressen))
print(mail_adressen)
if user_input == "1":
    list_to_copy = '\n'.join(mail_adressen)
if user_input == "2":
    list_to_copy = ';'.join(mail_adressen)

pyperclip.copy(list_to_copy)
print("#"*50)
print('Mail Address Copied to Clipboard')
print("#"*50)

if user_input == "2":
    bool_send_email = False
    while True:
        input_send_email = input(f"Email mit Anleitung an {mail_adressen} senden? [y/n] ")
        if input_send_email == "y":
            bool_send_email = True
            break
        elif input_send_email == "n":
            bool_send_email = False
            break
        else:
            print('Ungültige Eingabe')


    if bool_send_email:
        input_name =  input("Dein Vorname oder Vor und Nachname (wird am Ende an die Mail angehängt) ")
        file_path = "mail_body.txt"
        with open(file_path, "r", encoding="utf-8") as file:
            text_content = file.read()

        #print(text_content)
        #create E-Mail
        outlook2 = win32com.client.Dispatch("Outlook.Application")
        mail = outlook2.CreateItem(0) 

        # Set email properties
        mail.Subject = "Deine Dashboard Anfrage"
        mail.Body = text_content + f"\n {input_name}"

        # List of recipients
        recipients = mail_adressen
        mail.BCC = ";".join(recipients)

        # Add an attachment
        #attachment_path = "funkemedien_email_mit_google_verbinden.pptx"
        #attachment = mail.Attachments.Add(attachment_path)

        # Send the email
        mail.Send()
    else:
        print("Thank you, goodbye")
else:
    print("Thank you, goodbye")