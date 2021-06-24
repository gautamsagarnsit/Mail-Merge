# Mail Merge Desktop Application

Mail Merge is desktop application which lets the user to send bulk emails. Mail merge application takes an Excel file(.xlsx) as source of data to be inserted inside placeholders placed in a Microsoft word file(.docx). This application is free to use and share. Read more about Mail Merge here:
- https://en.wikipedia.org/wiki/Mail_merge
- https://support.microsoft.com/en-us/office/use-mail-merge-for-bulk-email-letters-labels-and-envelopes-f488ed5b-b849-4c11-9cff-932c49474705


# Table of Contents
  - [Required Files](#1)
  - [How to use](#2)
  - [TODO](#3)
  - [Contact](#4)

<br />

<a id="1"></a>
# Required Files
## Python Files
- mail_merge_project.py
- mail_mail_gui.py
## Modules
- tkinter
- mammoth
- bs4
- email
- smtplib
- ssl
- os
- docx2txt
- docx
- shutil
- win32com
- sys


<a id="2"></a>
## How to use

###  Preparing Source(Excel, xlsx) File
- Open Microsoft Excel file which you want to use as the source file.
- Data in the first row will be treated as place holders in docx file.
- Check all the Data like Name, Email etc. and save the file
- Close the file.
- See "source.xlsx" for reference. 

### Preparing Template(Word, Docx) File
- Add placeholders wherever you want in a docx file.
- Format for placeholder: {{placeholder_name}}
- One placeholder can not be nested into another placeholder.
- "placeholder_name" must be exactly same as the header name in the excel file. Case of placeholders are ignored and all every header and placeholder is treated as upper case strings.
- You can add as many as images and text you want.
- Save and close the file.
- See "template.docx" for reference

### Before moving forward you need to allow "Less Secure app acess". Click on this link to know more about Turning on "Less Secure app access": 
https://support.google.com/accounts/answer/6010255?hl=en#zippy=%2Cif-less-secure-app-access-is-on-for-your-account%2Cif-less-secure-app-access-is-off-for-your-account

Click on <b> If "Less secure app access" is off for your account</b>. Under the section click on the link given to <b>turn it back on</b>.

### Desktop Application
- Open mail_merge_gui.exe
- Enter your Gmail Email Id and password and click on Login to Validate
- Upload Source and Template File
- Optional: Add Attachments
- Optional: Add Subject
- Finally click on Send Mail Button and wait for Mail Merge to send all mails at once.
- Sometimes Windows protection or antivirus blocks the exe file from running. You can either bypass the block or use the method mentioned in next sub-section.
- mail_merge_gui.exe may not be updated. TO convert python script to .exe file use this command:<b>pip install pyinstaller</b> then after installation run: <b>pyinstaller --onefile -w 'mail_merge_gui.py'</b>.

### Running directly from script
- Open mail_merge_gui.py file
- Run the file to open the GUI
- Repeat steps mentioned in above section(Desktop Application)

<a id="3"></a>
## TODO
- Integrating support for multiple domains other than gmail.
- Registration feature to avoid login every time.
- Improving performance metrics like speed and space. Current version is approx of 17MB. It can be made smaller.
- Support for older version of Excel(.xls) and Word(.Doc) file.


<a id="4"></a>
## Contact
- For any suggestions or doubt regarding this project you can mail me @ gautamsagarnsit@gmail.com. You can also add suggestions in TODO section above.
- Linkedin: www.linkedin.com/in/gautamsagarnsit
- Developers interested in contributing to this project are most welcome.


