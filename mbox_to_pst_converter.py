import pypff
import mailbox
from datetime import datetime

# Function to create a new PST file
def create_pst_file(pst_file):
    pst = pypff.file()
    pst.open(pst_file, "w")
    return pst

# Function to add a message to the PST folder
def add_message_to_pst(pst_folder, subject, body):
    message = pst_folder.add_sub_message()
    message.set_subject(subject)
    message.set_plain_text_body(body)
    message.set_delivery_time(datetime.now())

# Function to convert MBOX to PST
def convert_mbox_to_pst(mbox_file, pst_file):
    mbox = mailbox.mbox(mbox_file)
    pst = create_pst_file(pst_file)
    
    # Create a root folder in PST (e.g., "Inbox")
    root_folder = pst.get_root_folder()
    inbox = root_folder.add_sub_folder("Inbox")
    
    # Loop through MBOX messages and add them to PST
    for msg in mbox:
        subject = msg['subject'] if msg['subject'] else 'No Subject'
        body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
        add_message_to_pst(inbox, subject, body)
    
    # Close PST after completion
    pst.close()

# Specify the paths to your MBOX and PST files
mbox_file = "path/to/your/file.mbox"
pst_file = "path/to/your/file.pst"

# Perform the conversion
convert_mbox_to_pst(mbox_file, pst_file)

print("Conversion completed!")
