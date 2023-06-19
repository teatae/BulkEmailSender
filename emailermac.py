import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import re
from tkinterhtml import HtmlFrame
from tkinterdnd2 import DND_FILES, TkinterDnD

excel_data = None

def solve(s):
    pat = "^[a-zA-Z0-9-_]+@[a-zA-Z0-9]+\.[a-z]{2,}$"
    if re.match(pat, s):
        return True
    return False

def on_drop(event):
    global excel_data
    try:
        file_path = event.data
        # Process the Excel file here
        # You can load and handle the Excel file as per your requirements
        print("File dropped:", file_path)
        excel_data = pd.read_excel(file_path)
        print("Excel data:", excel_data)  # Print excel_data for debugging
        if excel_data.empty:
            # DataFrame is empty
            messagebox.showerror("Error", "No file at" + str(file_path))
            return
        preview_function()
    except tk.TclError:
        print("No valid selection or form 'STRING' not defined")


def load_excel_data():
    global excel_data
    # Read the Excel file
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    print("File path:", file_path)  # Print file path for debugging
    excel_data = pd.read_excel(file_path)
    print("Excel data:", excel_data)  # Print excel_data for debugging
    if excel_data.empty:
        # DataFrame is empty
        messagebox.showerror("Error", "No file at" + str(file_path))
        return

def preview_function():
    # Load the email template from the input field
    email_template = email_template_entry.get("1.0", tk.END)

    # Clear existing tabs/pages
    for child in email_preview_notebook.winfo_children():
        child.destroy()

    # Iterate over the rows of the Excel file
    for idx, row in excel_data.iterrows():
        first_name = row[0].title()
        last_name = row[1].title()
        email = row[2]
        signature = row[3]

        # Replace placeholders in the email template with data
        body = email_template.replace("{first_name}", first_name).replace("{last_name}", last_name).replace('\n', '<br>')

        email_subject = subject_entry.get()

        # Create a new tab/page for the email preview
        email_page = ttk.Frame(email_preview_notebook)
        email_preview_notebook.add(email_page, text="Email " + str(idx+1))

        # Create an HtmlFrame widget for the email preview
        preview_frame = HtmlFrame(email_page)
        preview_frame.pack(fill=tk.BOTH)

        attachments = ""

        # Iterate over the attachment columns and add attachments to the email preview
        for i in range(4, len(row)):
            attachment_path = row[i]
            if isinstance(attachment_path, str) and attachment_path.strip() != '':
                attachments += os.path.basename(attachment_path) + "<br>"

        if solve(email):
            preview_frame.set_content(f"To: {email}<br>Subject: {email_subject}<br><br>{body}{signature}<br>Attachments:<br>{attachments}")
        else:
            preview_frame.set_content(f"To invalid email: {email}<br>Subject: {email_subject}<br><br>{body}{signature}<br>Attachments:<br>{attachments}")

    messagebox.showinfo("Preview Complete", "Email preview completed successfully.")

def preview_emails():
    load_excel_data()
    # Read the Excel file
    global excel_data
    # Check if the Excel data is loaded
    if excel_data is None:
        messagebox.showerror("Error", "Please load an Excel file first.")
        return

    preview_function()




def send_emails():
    # Get SMTP variables from input fields
    smtp_username = smtp_username_entry.get()
    smtp_password = smtp_password_entry.get()

    # Get email subject from input field
    email_subject = subject_entry.get()

    # Read the Excel file
    global excel_data
    # Check if the Excel data is loaded
    if excel_data is None:
        messagebox.showerror("Error", "Please load an Excel file first.")
        return

    # Load the email template from the input field
    email_template = email_template_entry.get("1.0", tk.END)

    # Create a connection to the SMTP server
    server = smtplib.SMTP("smtp.office365.com", "587")
    server.starttls()

    try:
        if smtp_password.strip() != '' and solve(smtp_username):
            server.login(smtp_username, smtp_password)
        else:
            messagebox.showerror("Authentication Error", "Invalid SMTP credentials. Please check your username and password.")
            return
    except smtplib.SMTPAuthenticationError as e:
        messagebox.showerror("Authentication Error", "Invalid SMTP credentials. Please check your username and password.")
        return

    # Iterate over the rows of the Excel file
    for _, row in excel_data.iterrows():
        first_name = row[0].title()
        last_name = row[1].title()
        email = row[2]
        signature = row[3]

        # Replace placeholders in the email template with data
        body = email_template.replace('{first_name}', first_name).replace('{last_name}', last_name).replace('\n', '<br>')
        body += '<br>'
        body += signature

        # Create the email message
        msg = MIMEMultipart()
        msg['From'] = smtp_username
        msg['To'] = email
        msg['Subject'] = email_subject
        msg.attach(MIMEText(body, 'html'))

        # Iterate over the attachment columns and add attachments to the email
        for i in range(4, len(row)):
            attachment_path = row[i]
            if isinstance(attachment_path, str) and attachment_path.strip() != '':
                attachment = MIMEApplication(open(attachment_path, 'rb').read())
                attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                msg.attach(attachment)

        # Send the email using the SMTP server
        try:
            server.sendmail(smtp_username, email, msg.as_string())
        except smtplib.SMTPException as e:
            messagebox.showerror("Error", f"Failed to send emails:\n{str(e)}")
            return

    messagebox.showinfo("Success", "Emails sent successfully!")

    # Close the SMTP connection
    server.quit()


def open_preview_emails():
    preview_emails()

def open_send_emails():
    send_emails()

# Create the main window
window = tk.Tk()
window.title("Bulk Email Sender")

icon_path = "C:/Temp/TAE/Projects/22. Bradley emails/icon.ico"

# Set the window icon
window.iconbitmap(icon_path)

# Create a Tkinter window using TkinterDnD
window = TkinterDnD.Tk()

# Bind the drop event to the window
window.drop_target_register(DND_FILES)
window.dnd_bind('<<Drop>>', on_drop)

# SMTP Server Settings
smtp_username_label = tk.Label(window, text="SMTP Username:")
smtp_username_label.pack()
smtp_username_entry = tk.Entry(window)
smtp_username_entry.pack()

smtp_password_label = tk.Label(window, text="SMTP Password:")
smtp_password_label.pack()
smtp_password_entry = tk.Entry(window, show="*")
smtp_password_entry.pack()

# Email Subject
subject_label = tk.Label(window, text="Email Subject:")
subject_label.pack()
subject_entry = tk.Entry(window)
subject_entry.pack()

# Email Template
email_template_label = tk.Label(window, text="Email Template:")
email_template_label.pack()
email_template_entry = tk.Text(window, width=69, height=10)
email_template_entry.pack()

# Buttons
preview_emails_button = tk.Button(window, text="Preview Emails", command=open_preview_emails)
preview_emails_button.pack()

send_emails_button = tk.Button(window, text="Send Emails", command=open_send_emails)
send_emails_button.pack()

# Create a notebook to display email previews
email_preview_notebook = ttk.Notebook(window)
email_preview_notebook.pack(fill=tk.BOTH, expand=True)

window.mainloop()


#/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
#(echo; echo 'eval "$(/usr/local/bin/brew shellenv)"') >> /Users/tae/.zprofile
#eval "$(/usr/local/bin/brew shellenv)"
#brew install pip 
#brew install python 
#python3.11 -m pip install --upgrade pip 
#brew install python-tk
#pip3 install --upgrade pip
#export PATH="/usr/local/opt/python/libexec/bin:$PATH"
