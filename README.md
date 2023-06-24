# really Automated Mass Email Notifications (rAMEN) <!-- omit from toc -->
Automatically sends emails en masse using a plain text template and an excel sheet with the variations.

## Table of contents <!-- omit from toc -->
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [How to use rAMEN](#how-to-use-ramen)
  - [Template.txt](#templatetxt)
    - [Example template](#example-template)
  - [Excel Workbook](#excel-workbook)
    - [Example worksheet](#example-worksheet)
  - [Executing rAMEN](#executing-ramen)
  - [Advanced functionalities](#advanced-functionalities)
    - [CCs](#ccs)
    - [Attachments](#attachments)

## Prerequisites
- Python 3.11
- Outlook
- Microsoft Excel

## Installation
Execute the console command ```pip install -r requirements.txt``` while in the root folder of the project.
You must also have an instance of Outlook open on your PC, with the email you wish to send the emails from open.

## How to use rAMEN
There are two necessary files that are necessary for rAMEN: **template.txt and an excel workbook with the variations**.
Template.txt is the template Email that will be used. Users.xlsx is the list of users that will be notified.

### Template.txt
Template.txt is divided in two sections: subject and body. The first line of the txt file is the subject, the rest is the body.

To insert a term from the excel sheet it must follow the following format: ```%(term name)s```. The term name is the Excel worksheet's column name in lowercase, with all the spaces replaced by underscores (_). Some examples are: the column First Name would have the term name _first_name_. The column UserID would have the term name _userid_. The column First Name and Surname would have the term name _first_name_and_surname_.

#### Example template
>Example Email Subject
><br>
>Good day %(first_name)s,
><br><br>
>Here is your user and password:
><br><br>
>User: %(user_id)s
>Password: %(password)s
><br><br>
>Thank you for your time,

### Excel Workbook
This workbook must contain in one sheet a table with all the used terms and the different values. The only necessary field is "Email", as it is used to *send* the emails.

The workbook can contain more sheets than the one used by rAMEN and the used sheet can also have more fields that are not used in rAMEN. The fields do not need to be in a specific order, but there must always be a value for every row for the utilized fields. The sheet must only contain the table of fields.

#### Example worksheet
| Email                       | Username | First Name | Last Name | Password     |
| --------------------------- | -------- | ---------- | --------- | ------------ |
| juan.galvan@gmail.com       | jgalvan  | Juan       | Galván    | password1234 |
| alex.paredes@outlook.es     | aparedes | Alex       | Paredes   | password1234 |
| paolo.sanchez@hotmail.co.uk | psanchez | Paolo      | Sánchez   | password1234 |

### Executing rAMEN
Run the command ```python .\ramen.py``` in the root folder of the project.

### Advanced functionalities

#### CCs

When prompted to add CCs, add the emails, separated by semicolons (;) without any spaces. Example: ```cc_email_1@domain.com;cc_email_2@domain.net```

For different CCs per email, the excel worksheet must contain the column CC (case sensitive), where you introduce the necessary CCs in the same format as before.

#### Attachments

When prompted to add attachments, add the path to the files, with no spaces at the beginning and separated by semicolons (;). Example: ```C:\Users\USER\Downloads\FILE_TO_SEND_1;C:\Users\USER\Downloads\FILE_TO_SEND_2```.

For different files per email, the excel worksheet must contain the column Attachments (case sensitive), where you introduce the necessary attachments in the same format as before.