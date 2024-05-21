import tkinter as tk
from tkinter import messagebox
import imaplib
import email
from email.header import decode_header
import threading
 
 
# Email account credentials
EMAIL_ADDRESS = "Enter Your Email Address"  # Update with your Hotmail address
PASSWORD = "Enter-Your-Password"  # Update with your Hotmail password or app password if 2FA is enabled
 
 
# IMAP server configuration
IMAP_SERVER = "outlook.office365.com"  # Hotmail IMAP server
IMAP_PORT = 993  # Hotmail IMAP port
 
 
def decode_subject_header(header):
   try:
       decoded_header = email.header.decode_header(header)
       subject = ""
       for part, charset in decoded_header:
           if isinstance(part, bytes):
               if charset is None:
                   # If charset is None, assume utf-8 encoding
                   subject += part.decode('utf-8', errors='replace')
               else:
                   subject += part.decode(charset, errors='replace')
           else:
               subject += part
       return subject
   except Exception as e:
       print(f"Error decoding subject header: {e}")
       return "<Decoding Error>"
 
 
def fetch_emails(fetch_all=False):
   def fetch():
       try:
           # Connect to the IMAP server
           mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
           mail.login(EMAIL_ADDRESS, PASSWORD)
           mail.select('inbox')
 
 
           # Search for unseen or all emails based on the parameter
           if fetch_all:
               result, data = mail.search(None, 'ALL')
           else:
               result, data = mail.search(None, 'UNSEEN')
 
 
           if result == 'OK':
               email_listbox.delete(0, tk.END)  # Clear previous emails
               for num in data[0].split():
                   # Fetch the email headers
                   result, data = mail.fetch(num, '(BODY[HEADER.FIELDS (FROM SUBJECT)])')
                   if result == 'OK':
                       raw_header = data[0][1]
                       msg = email.message_from_bytes(raw_header)
                       subject = msg.get('subject', None)
                       sender = msg.get('from', None)
                       if subject is not None:
                           if isinstance(subject, str):  # Check if subject is a string
                               subject = decode_subject_header(subject)
                           else:
                               subject = str(subject)  # Convert to string if not already
                       else:
                           subject = ""
                       if sender is not None:
                           if isinstance(sender, str):  # Check if sender is a string
                               sender = sender.split()[-1]  # extracting only the email address part
                           else:
                               sender = str(sender)  # Convert to string if not already
                       else:
                           sender = ""
                       email_listbox.insert(tk.END, f"From: {sender}\nSubject: {subject}\n")
                       email_listbox.insert(tk.END, "\n")  # Add empty line for readability
 
 
           mail.close()
           mail.logout()
       except Exception as e:
           messagebox.showerror("Error", str(e))
 
 
   threading.Thread(target=fetch).start()
 
 
def fetch_and_display_selected():
   try:
       selected_indices = email_listbox.curselection()
       if selected_indices:  # Check if any item is selected
           selected_index = selected_indices[0]
           email_content_text.delete(1.0, tk.END)
           selected_email = email_listbox.get(selected_index)
           email_content_text.insert(tk.END, selected_email)
           # Fetch the full content of the selected email
           threading.Thread(target=fetch_email_content, args=(selected_index,)).start()
   except IndexError:
       pass
 
 
def fetch_email_content(index):
   try:
       mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
       mail.login(EMAIL_ADDRESS, PASSWORD)
       mail.select('inbox')
       result, data = mail.search(None, 'SEEN')
       if result == 'OK':
           num_list = data[0].split()
           print("Number of emails:", len(num_list))  # Debug print
           if index < len(num_list):
               num = num_list[index]
               result, data = mail.fetch(num, '(RFC822)')
               if result == 'OK':
                   raw_email = data[0][1]
                   msg = email.message_from_bytes(raw_email)
                   email_content_text.insert(tk.END, f"\n{'-'*50}\n")
                   for part in msg.walk():
                       content_type = part.get_content_type()
                       content_disposition = str(part.get("Content-Disposition"))
                       if content_type == "text/html":
                           from bs4 import BeautifulSoup
                           soup = BeautifulSoup(part.get_payload(decode=True), "html.parser")
                           for link in soup.find_all('a', href=True):
                               email_content_text.insert(tk.END, f"URL: {link['href']}\n")
                       elif content_type == "text/plain":
                           try:
                               # Attempt to decode the email content using UTF-8
                               decoded_content = part.get_payload(decode=True).decode('utf-8')
                               email_content_text.insert(tk.END, decoded_content + "\n")
                           except UnicodeDecodeError:
                               # If decoding with UTF-8 fails, try Latin-1 encoding
                               decoded_content = part.get_payload(decode=True).decode('latin-1')
                               email_content_text.insert(tk.END, decoded_content + "\n")
                   email_content_text.insert(tk.END, "\n" + "-" * 50 + "\n\n")
       mail.close()
       mail.logout()
   except Exception as e:
       messagebox.showerror("Error", str(e))
 
 
# Create the main window
root = tk.Tk()
root.title("Email Reader - The Pycodes")
 
 
# Create widgets
fetch_button = tk.Button(root, text="Fetch Unseen Emails", command=lambda: fetch_emails(fetch_all=False))
fetch_all_button = tk.Button(root, text="Fetch All Emails", command=lambda: fetch_emails(fetch_all=True))
email_listbox = tk.Listbox(root, width=100, height=20)
email_listbox.bind("<<ListboxSelect>>", lambda event: fetch_and_display_selected())
email_scrollbar = tk.Scrollbar(root, orient=tk.VERTICAL, command=email_listbox.yview)
email_listbox.config(yscrollcommand=email_scrollbar.set)
email_content_text = tk.Text(root, width=100, height=20)
 
 
# Layout widgets
fetch_button.pack(pady=5)
fetch_all_button.pack(pady=5)
email_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
email_scrollbar.pack(side=tk.LEFT, fill=tk.Y)
email_content_text.pack(pady=10)
 
 
# Start the Tkinter event loop
root.mainloop()
