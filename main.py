# Automatic User Creation Notification
# Programmed by Jaime Stockwell Mendoza (JStockwell on GitHub)

# Components and requirements:
# Template for email, with all fields under curly brackets with the fields with this format: %(FIELDNAME)s
# Excel file with the following fields: User ID, First Name, Password, Email
# Ask for parameters for emails to add to CC list
# Config file for the connection to the SMTP server

# TEMPLATE.TXT
# Fields = userid, firstname, password, email
# First line = Subject
# Second line = Body

import win32com.client, pandas

sheet = pandas.read_excel("users.xlsx")

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
    newmail.Body= email_data.split("\n",1)[1]

    # attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
    # newmail.Attachments.Add(attach)

    newmail.Send()

send_emails(sheet.to_dict())