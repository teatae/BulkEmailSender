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
        first_name = row[0].strip().title()
        last_name = row[1].strip().title()
        email = row[2].strip()
        signature = row[3]

        # Replace placeholders in the email template with data
        body = email_template.replace("{first_name}", first_name).replace("{last_name}", last_name).replace('\n', '<br>')

        email_subject = subject_entry.get().replace("{first_name}", first_name).replace("{last_name}", last_name).replace('\n', '<br>')

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
                print(attachment_path)

        if solve(email):
            #content = f"<span style='font-size: 16pt;'>To: {email}<br>Subject: {email_subject}<br><br>{body}{signature}<br>Attachments:<br>{attachments}</span>"
            content = f"To: {email}<br>Subject: {email_subject}<br><br>{body}<br>{signature}<br>Attachments:<br>{attachments}"
        else:
            #content = f"<span style='font-size: 16pt;'>To INVALID EMAIL: {email}<br>Subject: {email_subject}<br><br>{body}{signature}<br>Attachments:<br>{attachments}</span>"
            content = f"To invalid email: {email}<br>Subject: {email_subject}<br><br>{body}<br>{signature}<br>Attachments:<br>{attachments}"

        preview_frame.set_content(content)

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

    if smtp_username and smtp_password:
        file = open("temp.txt","w")
        file.write(smtp_username+"\n"+smtp_password)
        file.close()

    # Get email subject from input field
    email_subject = subject_entry.get()

    # Read the Excel file
    global excel_data
    # Check if the Excel data is loaded
    if excel_data is None:
        messagebox.showerror("Error", "Please load an Excel file first.")
        return False

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
            return False
    except smtplib.SMTPAuthenticationError as e:
        messagebox.showerror("Authentication Error", "Invalid SMTP credentials. Please check your username and password.")
        return False

    # Iterate over the rows of the Excel file
    for _, row in excel_data.iterrows():
        first_name = row[0].strip().title()
        last_name = row[1].strip().title()
        email = row[2].strip()
        signature = row[3]

        # Replace placeholders in the email template with data
        body = email_template.replace('{first_name}', first_name).replace('{last_name}', last_name).replace('\n', '<br>')
        body += '<br>'
        body += signature

        # Create the email message
        msg = MIMEMultipart()
        msg['From'] = smtp_username
        msg['To'] = email
        msg['Subject'] = email_subject.replace("{first_name}", first_name).replace("{last_name}", last_name)
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
            return False

    # Close the SMTP connection
    server.quit()

    return True

def open_preview_emails():
    preview_emails()

def open_send_emails():
    # Create a Toplevel window for the loading message
    loading_window = tk.Toplevel(window)
    loading_window.title("Loading")
    loading_window.attributes('-topmost', True)  # Ensure it's on top of the main window

    # Calculate the position of the loading window
    app_width = window.winfo_width()
    app_height = window.winfo_height()
    loading_window_width = 150
    loading_window_height = 75
    x = window.winfo_x() + (app_width - loading_window_width) // 2
    y = window.winfo_y() + (app_height - loading_window_height) // 2
    loading_window.geometry(f"{loading_window_width}x{loading_window_height}+{x}+{y}")

    # Create a label to display the loading message
    loading_label = tk.Label(loading_window, text="Sending emails...")
    loading_label.pack(pady=20)

    # Simulate the email sending process
    # Replace this with your actual email sending code
    # Here, we're using the `after` method to schedule a callback after 3 seconds
    window.after(5, lambda: complete_send_emails(loading_window))

def complete_send_emails(loading_window):
    # Perform the actual email sending process here
    # Replace this with your actual email sending code

    # Simulating a successful email sending
    success = send_emails()

    # Destroy the loading window
    loading_window.destroy()

    # Show success message box
    if success:
        messagebox.showinfo("Success", "Emails sent successfully!")
    else:
        messagebox.showerror("Error", "Failed to send emails.")


# Create a Tkinter window using TkinterDnD
window = TkinterDnD.Tk()
window.title("Bulk Email Sender")
# Set the width of the window
window_width = 800  # Desired width
window.wm_geometry(f"650x575")

icon_path = "./icon.ico"

# Set the window icon
window.iconbitmap(icon_path)

# Bind the drop event to the window
window.drop_target_register(DND_FILES)
window.dnd_bind('<<Drop>>', on_drop)

# SMTP Server Settings
smtp_username_label = tk.Label(window, text="Email:")
smtp_username_label.grid(row=0, column=0, sticky=tk.W, padx=20)
smtp_username_entry = tk.Entry(window, width=50)
smtp_username_entry.grid(row=0, column=1, sticky=tk.W, padx=20)

smtp_password_label = tk.Label(window, text="Password:")
smtp_password_label.grid(row=1, column=0, sticky=tk.W, padx=20)
smtp_password_entry = tk.Entry(window, show="*", width=50)
smtp_password_entry.grid(row=1, column=1, sticky=tk.W, padx=20)

# Email Subject
subject_label = tk.Label(window, text="Email Subject:")
subject_label.grid(row=2, column=0, sticky=tk.W, padx=20)
subject_entry = tk.Entry(window, width=50)
subject_entry.grid(row=2, column=1, sticky=tk.W, padx=20)
subject_entry.insert(tk.END, "Congratulations {first_name}!")

# Email Template
email_template_label = tk.Label(window, text="Email Template:")
email_template_label.grid(row=3, column=0, sticky=tk.W, padx=20)
email_template_entry = tk.Text(window, width=65, height=8)
email_template_entry.grid(row=3, column=1, sticky=tk.W, padx=20)
email_template_entry.insert(tk.END, "Hello {first_name} {last_name},\n\nWelcome to teatae's Bulk Email Sender!\n")

# Buttons
preview_emails_label = tk.Label(window, text="Drag your excel file here to load!")
preview_emails_label.grid(row=4, column=0, columnspan=2, padx=5)

send_emails_button = tk.Button(window, text="Send Emails", command=open_send_emails)
send_emails_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

# Create a notebook to display email previews
email_preview_notebook = ttk.Notebook(window)
email_preview_notebook.grid(row=6, column=0, columnspan=2, sticky=tk.NSEW)

# Configure grid weights to allow resizing of the notebook
window.grid_rowconfigure(6, weight=1)
window.grid_columnconfigure(0, weight=1)
window.grid_columnconfigure(1, weight=1)

# Configure column widths to match the email template
window.grid_columnconfigure(1, minsize=email_template_entry.winfo_reqwidth())

try:
    file = open("temp.txt","r")
    lines = file.read().splitlines()
    smtp_username_entry.insert(tk.END, lines[0])
    smtp_password_entry.insert(tk.END, lines[1])
    file.close()
finally:
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
