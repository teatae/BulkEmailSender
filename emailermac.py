import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

excel_data = None

def load_excel_data():
    global excel_data
    # Read the Excel file
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    print("File path:", file_path)  # Print file path for debugging
    excel_data = pd.read_excel(file_path)
    print("Excel data:", excel_data)  # Print excel_data for debugging
    if excel_data.empty:
        # DataFrame is empty
        messagebox.showerror("Error", "no file at" + str(file_path))
        return

def preview_emails():
    load_excel_data()
    # Read the Excel file
    global excel_data
    # Check if the Excel data is loaded
    if excel_data is None:
        messagebox.showerror("Error", "Please load an Excel file first.")
        return

    # Load the email template from the input field
    email_template = email_template_entry.get("1.0", tk.END)

    # Clear existing tabs/pages
    for child in email_preview_notebook.winfo_children():
        child.destroy()

    # Keep track of the maximum number of lines in an email preview
    max_lines = 0

    # Iterate over the rows of the Excel file
    for idx, row in excel_data.iterrows():
        first_name = row[0].title()
        last_name = row[1].title()
        email = row[2]
        
        # Replace placeholders in the email template with data
        email_body = email_template.replace("{first_name}", first_name).replace("{last_name}", last_name)

        email_subject = subject_entry.get()

        # Create a new tab/page for the email preview
        email_page = ttk.Frame(email_preview_notebook)
        email_preview_notebook.add(email_page, text="Email " + str(idx+1))

        # Create a text widget for the email preview
        preview_text = tk.Text(email_page)
        preview_text.pack(fill=tk.BOTH, expand=True)
        preview_text.insert(tk.END, f"To: {email}\nSubject: {email_subject}\n\n{email_body}\nAttachments:\n")

        # Iterate over the attachment columns and add attachments to the email preview
        for i in range(3, len(row)):
            attachment_path = row[i]
            if isinstance(attachment_path, str) and attachment_path.strip() != '':
                preview_text.insert(tk.END, os.path.basename(attachment_path) + "\n")

        # Disable editing in the preview text widget
        preview_text.config(state=tk.DISABLED)

        # Calculate the number of lines in the text widget
        num_lines = int(preview_text.index('end-1c').split('.')[0])
        # Update the maximum number of lines
        max_lines = max(max_lines, num_lines)

    # Calculate the desired window height based on the maximum number of lines
    desired_window_height = 400+max_lines * 20  # Adjust the factor as needed

    # Update the window height
    window.geometry(f"{window_width}x{desired_window_height}")

    messagebox.showinfo("Preview Complete", "Email preview completed successfully.")



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
        if smtp_username.strip() != '' and smtp_password.strip() != '' and "@" in smtp_username and "." in smtp_username:
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

        # Replace placeholders in the email template with data
        email_body = email_template.replace('{first_name}', first_name).replace('{last_name}', last_name).replace('\n', '<br>')
        email_body = email_body + '<br>'

        # Create the email message
        msg = MIMEMultipart()
        msg['From'] = smtp_username
        msg['To'] = email
        msg['Subject'] = email_subject
        msg.attach(MIMEText(email_body, 'html'))

        # Iterate over the attachment columns and add attachments to the email
        for i in range(3, len(row)):
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
    # ...

    # Call the preview_emails function
    preview_emails()

def open_send_emails():
    # ...

    # Call the send_emails function
    send_emails()


# Create the main window
window = tk.Tk()
window.title("Bulk Email Sender")

icon_path = "C:/Temp/TAE/Projects/22. Bradley emails/icon.ico"

# Set the window icon
window.iconbitmap(icon_path)

# Set the initial window width and height
window_width = 800
window_height = 600

# Set the window geometry
window.geometry(f"{window_width}x{window_height}")

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
template_label1 = tk.Label(window, text="Email Template:")
template_label1.pack()
email_template_entry = tk.Text(window, height=10, width=50)
email_template_entry.pack()

# Buttons
preview_emails_button = tk.Button(window, text="Preview Emails", command=open_preview_emails)
preview_emails_button.pack()

send_emails_button = tk.Button(window, text="Send Emails", command=open_send_emails)
send_emails_button.pack()

# Create a notebook for the email previews
email_preview_notebook = ttk.Notebook(window)
email_preview_notebook.pack(fill=tk.BOTH, expand=True)

# Run the main event loop
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