# really Automated Mass Email Notifications (rAMEN)
# Programmed by Jaime Stockwell Mendoza (JStockwell on GitHub)
# Support me on Ko-Fi at https://ko-fi.com/jstockwell

# Requirements:
# Template for email, with all fields under curly brackets with the fields with this format: %(FIELDNAME)s
# Excel file containing the terms and all the values for each email

# TEMPLATE.TXT
# Fields = excel sheet column names in lowercase with the spaces replaced with underscores (_)
# First line = Subject
# Second line = Body

# Excel worksheet MUST contain an Email column with the accounts that will recieve each email

import pandas
import win32com.client

# ------------------------------ FUNCTIONS ------------------------------ #

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return template_file_content

def compose_email(template, i, usr_dict):
    vars = {}

    for key in usr_dict.keys():
        vars[key.lower().replace(" ", "_")] = usr_dict[key][i]

    return (template % vars, usr_dict["Email"][i])

def send_email(compose_res, attachments, cc):
    email_data = compose_res[0]

    ol = win32com.client.Dispatch("Outlook.Application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)

    newmail.Subject= email_data.partition('\n')[0]
    newmail.To=compose_res[1]
    newmail.CC = cc
    newmail.Body= email_data.split('\n',1)[1]
    
    for attach in attachments.split(";"):
        if attach != "":
            newmail.Attachments.Add(attach)

    newmail.Send()

def print_email(compose_res, attachments, cc):
    email_data = compose_res[0]

    subject = email_data.partition('\n')[0]
    body = email_data.split('\n',1)[1]
    email = compose_res[1]

    print("\n-----------------------------------\n")
    print("Subject: " + subject + "\n\nEmail: " + email + "\n\nCC: " +  cc + "\n\nBody:\n" + body + "\n\nAttachments: " + attachments)

# ------------------------------ MAIN FUNCTION ------------------------------ #

def ramen():
    template_path = input("Enter path to template text file: ")
    workbook_path = input("Enter path to excel file: ")
    sheet_name = input("Enter excel sheet name (optional): ")

    if sheet_name == "":
        sheet = pandas.read_excel(workbook_path)

    else:
        sheet = pandas.read_excel(workbook_path, sheet_name=sheet_name)

    mode = input("Enter 1 to send emails, 2 to print emails: ")

    if (mode == "1" and input("Are you sure you want to SEND emails? Y or N:").lower() == "y") or mode == "2":
        usr_dict = sheet.to_dict()

        c_flag = False
        if not usr_dict.keys().__contains__("CC"):
            cc = input("Enter CC email addresses separated by semicolons (optional): ")

        else:
            c_flag = True

        a_flag = False
        if not usr_dict.keys().__contains__("Attachments"):
            attachments = input("Enter path to attachments, separated by semicolons (optional): ")

        else:
            a_flag = True

        template = read_template(template_path)

        for i in range(0, sheet.shape[0]):
            if a_flag:
                attachments = usr_dict["Attachments"][i]

            if c_flag:
                cc = usr_dict["CC"][i]
            
            if mode == "1":
                send_email(compose_email(template, i, usr_dict), attachments, cc)
            elif mode == "2":
                print_email(compose_email(template, i, usr_dict), attachments, cc)

# ------------------------------ EXECUTION ------------------------------ #

ramen()

# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣀⣤⡀⠀⠀⠀⢀⣀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣠⣾⠋⠉⣿⣤⣤⡾⠏⠛⢷⡀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢨⣿⣳⣿⠡⣠⣄⣚⣿⡟⣗⡄⠐⣈⣻⣄⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣀⣀⣀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣴⠟⠁⣀⠛⢋⡿⢻⡿⢻⠟⠁⣿⣿⣿⣇⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⡴⠛⠉⠉⠉⠉⠓⠶⣶⠲⠖⠒⠛⠒⠛⠒⠓⠓⠒⠒⠒⠶⠿⠶⣿⣿⣶⡿⢁⣾⣿⣄⠀⠀⠀⠀⠀⣿⡆⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⢀⣠⣴⡟⠀⠀⠀⠀⣀⣤⣄⣀⠈⠳⣄⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣠⠟⢁⡾⠛⢿⠿⣧⣄⡉⠀⠈⠿⣿⡄⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⢀⣠⡴⠚⠉⣽⠃⠀⠀⢀⡴⠋⠉⠉⠙⢦⡀⠹⣦⣀⣀⣀⣀⣀⣀⣀⣀⣀⣀⠀⣀⡴⠋⣠⠞⠁⠀⠈⠀⠉⠙⠳⣦⣀⠠⣿⣧⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⢀⣴⠛⠁⠀⠀⢀⣿⢀⠀⠀⢸⠀⠀⠀⠀⠀⠀⢳⠀⠹⣇⠔⡁⠀⣰⠌⢀⣈⣭⠿⠛⢉⣴⣾⣷⣤⣄⣀⡀⠀⠀⠀⠀⠀⠙⢷⣟⢿⡇⠀⠀⠀⠀⠀⠀⠀⠀
# ⢀⡿⠁⠀⠀⠀⣠⡼⣷⠸⡄⠀⠉⢷⡀⠀⠀⠀⢀⡼⠀⠀⣿⢀⣉⣷⡶⠿⢿⣏⣤⠶⣿⣿⡽⣶⣭⣉⠻⣯⡙⠳⢤⣀⠀⠀⠀⠀⢹⣾⣧⠀⠀⠀⠀⠀⠀⠀⠀
# ⢸⡇⠀⠀⣠⢾⡟⠂⢻⡆⢳⣄⡀⠀⠙⠶⠤⠴⠛⠡⠔⢻⣿⣻⣿⡴⠶⠛⢻⣏⣡⣾⠿⠷⣤⣦⣴⣽⣷⣞⢿⡌⠘⢏⢳⣄⠀⠀⢀⣿⡟⠀⠀⠀⠀⠀⠀⠀⠀
# ⠈⣷⡄⢰⡏⠰⠓⢠⡴⠿⡾⠿⢷⡀⠀⠀⠀⠀⠀⠀⣠⣾⣟⣿⡌⣷⣤⣴⡿⢛⣩⡴⠶⠳⣿⠿⢿⣿⠾⣿⡟⠋⠁⠀⠲⢽⡆⢀⣼⣿⠃⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⣿⣿⣶⣿⣄⠀⠸⣧⣀⣀⣀⣽⣟⣶⠶⡶⢶⣶⣫⣿⣋⣉⣙⣿⠌⣿⢷⡿⢻⡇⠀⠀⠀⠘⣻⠶⣶⠾⠟⢹⠆⠀⠀⢀⣼⣷⣿⣿⠏⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⢻⡌⠻⢿⣿⣷⣶⣬⣍⢻⣏⠻⣿⣟⠻⠿⢛⣻⣧⣌⣉⣉⣭⠟⢠⣿⡐⢿⣆⠻⣦⣀⣀⣴⠟⣠⡟⠀⢀⣀⣤⣴⣾⣿⣿⠟⢁⣿⣦⡀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⢿⣦⣀⠉⠛⠿⣿⣻⣿⣿⠷⢿⣿⣿⣽⣟⣻⣤⣉⣛⣫⣥⡀⢀⣸⣇⣼⣟⣶⣤⣭⣭⣴⣾⣿⠶⢿⢛⣏⣹⡴⠛⠁⢀⣴⣾⠿⣿⣟⣷⣦⠀⠀⠀⠀⠀⠀
# ⠀⠀⠈⣿⣿⣿⣶⣤⣀⡈⠉⠛⠛⠶⠶⣦⣭⣎⣭⣋⣟⣙⣛⣛⣻⣛⣛⣛⣏⣻⣍⣯⣭⣱⡶⠶⠿⠓⠛⠋⠉⣀⡄⠀⣠⣟⣾⠃⠀⠈⠙⠛⠁⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠈⠻⣿⣷⢯⡿⣽⢿⣶⣶⡲⣤⠤⣄⣀⣀⣁⣈⣉⡉⠉⠁⠉⠉⠉⣈⣁⣈⣀⣀⣀⡤⣤⠴⣲⠒⣏⣿⠏⠀⡴⣧⡿⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠙⢿⣿⣽⣻⣟⣮⣟⣷⡌⢳⡘⢬⢲⡡⢎⠭⣙⠏⢯⡹⢭⡙⢥⠚⡔⡣⣜⢡⠞⣠⢛⡤⢯⡴⢬⡷⣞⣷⠏⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠙⢷⣿⣟⣾⡽⣾⢿⡆⡝⣆⠣⡜⣡⢋⢆⡛⢦⠱⢦⡙⢦⠛⣬⠱⡌⠮⡜⢥⣋⡜⢯⣤⣞⣵⠟⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠙⠻⣯⣿⣯⣟⣿⣶⣌⠳⡜⢤⢋⠦⡙⢦⡙⠦⡙⢦⠹⢤⠓⣍⠞⡸⢆⡱⢎⣽⠿⠋⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⡼⠛⠻⣿⣾⣿⣷⣯⡘⢣⠞⡲⢩⠦⣙⠲⣍⠲⡙⢦⡙⢦⣩⣵⣾⠿⡟⢣⡀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣿⡄⠠⠰⠘⢿⣿⣿⡿⣿⣷⣮⣵⣧⣾⣤⣷⣬⣧⣽⣶⣿⣿⣿⣿⠃⢃⠀⠂⣿⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣀⣀⣤⣤⣤⣄⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠈⠻⣍⡀⠂⠌⠛⢿⣿⣿⣿⣽⣯⣿⣽⣯⣿⣽⣿⣽⣷⣿⠿⠋⠡⠈⠁⠈⠼⠃⢀⣀⣠⣤⠤⠶⠶⠚⠛⠋⠉⠉⠀⠀⣐⣿⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠙⠳⠦⣔⣀⠊⠉⠉⠛⠛⠛⠿⠻⠟⠟⠻⠛⠛⠩⠙⣀⣂⣤⣤⣶⠶⠶⠒⠚⠋⠉⠉⠀⠀⢀⣀⣀⣴⣤⣶⠶⠶⢿⡟⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠈⠁⠉⠐⠀⠉⠐⢠⣥⣬⣭⠶⠷⠞⠛⠋⠉⠉⣀⣀⣀⣠⣒⣤⣥⣬⠶⠶⢿⠛⠻⠭⠉⠁⠀⠀⠈⠉⠀⣀⣀⣤⡀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⣀⣀⣤⡤⠤⠶⠒⠛⠋⢉⠉⣁⣀⣠⣤⣤⣴⠶⠾⠾⠟⠛⠛⠙⠉⠉⠀⠀⢀⣀⣀⣠⣤⠤⠤⠶⠖⠚⠛⠛⠉⠉⠉⠁⢀⣿
# ⠀⠀⠀⠀⠀⠀⠀⠀⢀⣶⠛⠛⠉⢉⣀⣀⣠⣤⣤⣴⠶⠾⠛⠛⠛⠉⠉⠀⠀⣀⣀⣀⣤⣤⡤⠶⠶⠖⠚⠛⠋⠉⠉⠁⠀⢀⣀⣠⣤⣠⣤⣤⣤⡶⠶⠿⠟⣛⡏
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠻⠶⠛⠛⠋⠉⠉⣁⣀⣀⣄⡤⡤⠤⠶⠖⠒⠚⠛⠉⠉⢉⡀⣀⣀⣠⣤⣤⣤⣴⠶⠶⠷⠟⠾⠛⠛⠉⠉⠉⠉⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣴⠖⠚⠛⠉⠉⠉⣁⣀⣤⣄⣠⣤⣤⣴⠶⠶⠶⠛⠛⠛⠋⠙⠉⠉⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
# ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠻⠶⠶⠞⠛⠛⠛⠛⠉⠉⠉⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
#
# Ramen ASCII art from https://emojicombos.com/