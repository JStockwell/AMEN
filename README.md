# Automated Mass Email Notifications (AMEN) <!-- omit from toc -->

Automatically sends emails en masse using a plain text template and an excel sheet with the variations.

---

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [How to use AMEN](#how-to-use-amen)
  - [Template.txt](#templatetxt)
    - [List of accepted terms](#list-of-accepted-terms)


## Prerequisites
- Python 3.11
- Outlook
- Microsoft Excel

## Installation
Execute the console command ```pip install -r requirements.txt``` while in the root folder of the project
You must also have an instance of Outlook open on your PC, with the email you wish to send the emails from open.

## How to use AMEN
There are two necessary files that are necessary for AMEN: **template.txt and users.xlsx**.

Template.txt is the template Email that will be used. Users.xlsx is the list of users that will be notified.

### Template.txt
Template.txt is divided in two sections: subject and body. The first line of the txt file is the subject, the rest is the body.

To insert a term from the excel sheet it must follow the following format: ```%(term_name)s```

Example template:
>Example Email Subject
Good day %(firstname)s,
>
>Here is your user and password:
>
>User: %(userid)s
Password: %(password)s
>
>Thank you for your time,

#### List of accepted terms

<em>Format of the list: [Excel column header]: [Template term]</em>

- "First Name": "firstname"
- "Last Name": "lastname"
- "Email": "email"
- "Password": "password"
- "User ID": "userid"