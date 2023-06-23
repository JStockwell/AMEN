# Automated Mass Email Notifications (AMEN)
# Programmed by Jaime Stockwell Mendoza (JStockwell on GitHub)

# Requirements:
# Template for email, with all fields under curly brackets with the fields with this format: %(FIELDNAME)s
# Excel file containing allowed terms

# TEMPLATE.TXT
# Fields = in accepted_terms.json
# First line = Subject
# Second line = Body

import win32com.client, pandas, json

# TODO Take terms from accepted_terms.json

accepted_terms = json.load(open("accepted_terms.json"))

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return template_file_content

def send_emails(usr_dict):
    cc = input("Enter CC email addresses separated by semicolons: ")

    for i in range(0, len(usr_dict["User ID"])):
        vars = {}
        vars["userid"] = usr_dict["User ID"][i]
        vars["firstname"] = usr_dict["First Name"][i]
        vars["password"] = usr_dict["Password"][i]
        vars["email"] = usr_dict["Email"][i]
        send_email(vars, cc)

def send_email(vars, cc):
    email_data = read_template("template.txt") % vars

    ol = win32com.client.Dispatch("Outlook.Application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)

    newmail.Subject= email_data.partition('\n')[0]
    newmail.To=vars["email"]
    newmail.CC=cc
    newmail.Body= email_data.split('\n',1)[1]

    # attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
    # newmail.Attachments.Add(attach)

    newmail.Send()

def print_email(usr_dict):
    vars = {}
    vars["userid"] = usr_dict["User ID"][0]
    vars["firstname"] = usr_dict["First Name"][0]
    vars["password"] = usr_dict["Password"][0]
    vars["email"] = usr_dict["Email"][0]

    email_data = read_template("template.txt") % vars

    subject = email_data.partition('\n')[0]
    body = email_data.split('\n',1)[1]
    email = vars["email"]
    cc = input("Enter CC email addresses separated by semicolons: ")

    print("\n-----------------------------------\n")
    print("Subject: " + subject + "\n\nEmail: " + email + "\n\nCC: " +  cc + "\n\nBody:\n" + body)

workbook_path = input("Enter path to excel file: ")

sheet_name = input("Enter excel sheet name (optional): ")
mode = input("Enter 1 to send emails, 2 to print emails: ")

if sheet_name == "":
    sheet = pandas.read_excel(workbook_path)

else:
    sheet = pandas.read_excel(workbook_path, sheet_name=sheet_name)

if mode == "1":
    if input("Are you sure? Y or N:").lower() == "y":
        send_emails(sheet.to_dict())
elif mode == "2":
    print_email(sheet.to_dict())