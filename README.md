# Automated Mass Email Notifications (AMEN) <!-- omit from toc -->

Automatically sends emails en masse using a plain text template and an excel sheet with the variations.

---

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [How to use AMEN](#how-to-use-amen)
  - [Template.txt](#templatetxt)
    - [Example template:](#example-template)
    - [List of accepted terms](#list-of-accepted-terms)
- [Excel Workbook](#excel-workbook)

## Prerequisites
- Python 3.11
- Outlook
- Microsoft Excel

## Installation
Execute the console command ```pip install -r requirements.txt``` while in the root folder of the project.
You must also have an instance of Outlook open on your PC, with the email you wish to send the emails from open.

## How to use AMEN
There are two necessary files that are necessary for AMEN: **template.txt and an excel workbook with the variations**.
Template.txt is the template Email that will be used. Users.xlsx is the list of users that will be notified.

### Template.txt
Template.txt is divided in two sections: subject and body. The first line of the txt file is the subject, the rest is the body.
To insert a term from the excel sheet it must follow the following format: ```%(term_name)s```

#### Example template:
>Example Email Subject
><br>
>Good day %(firstname)s,
><br><br>
>Here is your user and password:
><br><br>
>User: %(userid)s
>Password: %(password)s
><br><br>
>Thank you for your time,

#### List of accepted terms

<em>Format of the list: [Excel column header]: [Template term]</em>

- "First Name": "firstname"
- "Last Name": "lastname"
- "Email": "email"
- "Password": "password"
- "User ID": "userid"

## Excel Workbook
This workbook must contain in one sheet a table with all the used terms and the different values. The only necessary field is "Email", as it is used to *send* the emails.

Example worksheet:
| Email | Username | First Name | Last Name | Password |
| juan.galvan@gmail.com | jgalvan | Juan | Galván | password1234 |