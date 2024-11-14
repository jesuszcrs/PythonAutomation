import pandas as pd
from tkinter import filedialog, simpledialog, messagebox
import tkinter as tk
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import quote_plus
from tkinter import scrolledtext
from tkinter import messagebox

subject = ""  # Define subject in the global scope
recipient = ""  # Define recipient in the global scope
email_text = ""  # Define email_text in the global scope


def process_data(file_path, preview_text):
    global subject, recipient, email_text  # Access global variables
    # Read the Excel file into a Pandas DataFrame
    df = pd.read_excel(file_path)

    # Clear existing text in the preview window
    preview_text.delete('1.0', tk.END)

    # Group DataFrame by VNDR_ID
    grouped_df = df.groupby('VNDR_ID')

    # Iterate over groups
    for vndr_id, group in grouped_df:
        # Extracting parameters from the group
        vendor_name = group['Vendor Name'].iloc[0]  # Assuming vendor name is the same for all rows in the group
        year = group['Year'].iloc[0]  # Assuming year is the same for all rows in the group
        due_date = group['Due_Date'].iloc[0]  # Assuming due_date is the same for all rows in the group
        debit_week = group['Debit_Week'].iloc[0]  # Assuming debit_week is the same for all rows in the group
        program_info = generate_program_info(group)
        recipient = group['Rebate_Contact'].iloc[0]  # Assuming Rebate_Contact is the recipient email column


        # Generate email text
        subject, recipient, email_text = generate_email_text(vndr_id, year, vendor_name, due_date, debit_week, program_info, recipient)

        # Add email preview to GUI
        preview_text.insert(tk.END, f"Preview for Vendor {vndr_id}:\n\n")
        preview_text.insert(tk.END, f"Subject: {subject}\n")
        preview_text.insert(tk.END, f"Recipient: {recipient}\n")
        preview_text.insert(tk.END, f"Body: {email_text}\n")
        preview_text.insert(tk.END, "-" * 50 + "\n\n")

def generate_program_info(df_group):
    # Initialize program info table
    program_info = f"<table border='1'>\n"
    program_info += "<tr><th>Program</th><th>Year</th><th>Frequency</th><th>Method</th><th>NBRPAYS</th><th>EXPECTPAY</th></tr>\n"

    # Iterate over rows in the group
    for index, row in df_group.iterrows():
        # Add program info for each row
        program_info += f"<tr><td>{row['PROGRAM']}</td><td>{row['Year']}</td><td>{row['FREQUENCY']}</td><td>{row['METHOD']}</td><td>{row['NBRPAYS']}</td><td>{row['EXPECTPAY']}</td></tr>\n"
    
    # Close the table
    program_info += "</table>"
    return program_info

def generate_email_text(vndr_id, year, vendor_name, due_date, debit_week, program_info, recipient):

    # Generate dynamic subject line
    subject_line = f"{year} True Value Rebate Payment for V# {vndr_id} Due by {due_date}"
    
    # Construct email body with HTML content
    email_text = f"RE: Vendor# {vndr_id} - {year} Rebates\n\n"
    email_text += f"{vendor_name}:\n\n"
    email_text += f"This is a reminder that your rebate payment(s) are due by {due_date}. Our records show we have not received all of your rebate payments:\n\n"
    email_text += program_info + "\n\n"
    email_text += f"Please send outstanding payment(s) to address below by {due_date}, otherwise a debit will be initiated the week of {debit_week}.\n\n"
    email_text += "Checks:\t\tTRUE VALUE COMPANY\n"
    email_text += "\t\t1164 PAYSPHERE CIRCLE\n"
    email_text += "\t\tCHICAGO, IL 60674\n\n"
    email_text += "Credit Memos:\tVendorRebates@TrueValue.com\n\n"
    email_text += "Regards,\n\n"
    email_text += "Vendor Ops Team"
    email_text += "<br>"
    email_text += "<img src=\"cid:image1\"><br>"
    email_text += "TRUE VALUE COMPANY<br>"
    email_text += "1164 PAYSPHERE CIRCLE<br>"
    email_text += "CHICAGO, IL 60674<br>"
    return subject_line, recipient, email_text

def select_file_and_process_data():
    # Prompt the user to select the Excel file
    file_path = filedialog.askopenfilename()

    # Process data using the selected file
    process_data(file_path, preview_text)

# Define the LoginDialog class for custom login dialog
class LoginDialog(simpledialog.Dialog):
    def body(self, master):
        tk.Label(master, text="Email Address:").grid(row=0, sticky="e")
        tk.Label(master, text="Password:").grid(row=1, sticky="e")
        self.username_entry = tk.Entry(master)
        self.password_entry = tk.Entry(master, show="*")
        self.username_entry.grid(row=0, column=1)
        self.password_entry.grid(row=1, column=1)
        return self.username_entry  # Focus on username entry field by default

    def apply(self):
        self.smtp_username = self.username_entry.get()
        self.smtp_password = self.password_entry.get()


def send_email(use_popup=True):
    global subject, recipient, email_text  # Access global variables
    # SMTP server configuration
    smtp_server = 'smtp.office365.com'
    smtp_port = 587

    # Sender and recipient email addresses
    sender_email = 'Jesus.Zacarias@TrueValue.com'
    recipient_email = recipient

    # Create message container
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email

    # Attach HTML body
    msg.attach(MIMEText(email_text, 'html'))

    # Hardcoded login credentials (temporary, replace with secure method)
    #smtp_username = 'vendorrebates@truevalue.com'
    #smtp_password = quote_plus('7$noise$ANGRY$30')

    # Use popup for login credentials (uncomment to enable)
    if use_popup:
        # Prompt user for login credentials using custom dialog
        smtp_username, smtp_password = login_popup()

    # Connect to SMTP server and send email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        print("Email sent successfully")
        messagebox.showinfo("Emails Sent", "Emails have been sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {str(e)}")

# Function to display the custom login dialog
def login_popup():
    dialog = LoginDialog(root)
    return dialog.smtp_username, dialog.smtp_password

def cancel_emails():
    global email_sending_in_progress
    email_sending_in_progress = False
    messagebox.showinfo("Emails Canceled", "Email sending process has been canceled.")
    root.after(100, root.destroy)  # Close the GUI window after a delay

# Create the main window
root = tk.Tk()
root.title("Email Preview")

# Create a scrolled text widget to display the email preview
preview_text = scrolledtext.ScrolledText(root, width=80, height=30)
preview_text.pack()

# Create a frame to contain the buttons
button_frame = tk.Frame(root)
button_frame.pack(side=tk.BOTTOM, fill=tk.X)

# Create a button to cancel emails
cancel_emails_button = tk.Button(button_frame, text="Cancel Emails", command=cancel_emails)
cancel_emails_button.pack(side=tk.LEFT, padx=5, pady=10)


# Create a button to select the Excel file and process data
select_file_button = tk.Button(button_frame, text="Select Excel File", command=select_file_and_process_data)
select_file_button.pack(side=tk.LEFT, padx=(10, 5), pady=10)

# Create a button to send emails
send_emails_button = tk.Button(button_frame, text="Send Emails", command=send_email)
send_emails_button.pack(side=tk.RIGHT, padx=(5, 10), pady=10)


# Adjust button sizes
button_width = 15
button_height = 2
select_file_button.config(width=button_width, height=button_height)
send_emails_button.config(width=button_width, height=button_height)
cancel_emails_button.config(width=button_width, height=button_height)

# Start the GUI main loop
root.mainloop()
