Welcome to Tae's Bulk Email Sender (Outlook for Windows and Mac)  
================================================

## Example of excel file [Windows]  
Headers/columns must be First Name, Last Name, Email, Attachments1, Attachments2, Attachments3...  
Each email is an excel row, attachment values are paths.  
![preview](https://github.com/teatae/BulkEmailSender/blob/main/excel.png?raw=true)  

## Example of excel file [macOS]  
Headers/columns must be First Name, Last Name, Email, Signature, Attachments1, Attachments2, Attachments3...  
Each email is an excel row, attachment values are paths.  
![preview](https://github.com/teatae/BulkEmailSender/blob/main/excel2.png?raw=true)

## Preview of application  
![preview](https://github.com/teatae/BulkEmailSender/blob/main/preview.png?raw=true) 
![preview](https://github.com/teatae/BulkEmailSender/blob/main/preview2.png?raw=true) 

## Installing modules to run python file  
pip install --upgrade pip  
pip install tkinter pandas smtplib   

## Installing modules to convert python file into standalone executable with no windows console  
pip install pyinstaller  
pyinstaller --noconsole emailer.py  
Inside your dist > emailer folder, you will have your program as emailer.exe

## How to use  
Launch emailer.exe  (emailer.py is for windows, emailermac.py is for mac os)
Enter Subject in field  
Write template such as {first_name} is first name and {last_name} is last name  
Import excel file (Load excel file for preview button)  
Verify preview (email tabs)  
Click "Send Emails" to send mass personalized emails with attachements  
