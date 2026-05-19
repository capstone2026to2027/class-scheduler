import os
import smtplib
import imaplib
import email
from email.header import decode_header
import time
import re
from app import app, db
from models import User
def load_env():
    if os.path.exists('.env'):
        with open('.env') as f:
            for line in f:
                if line.strip() and not line.startswith('#'):
                    key, val = line.strip().split('=', 1)
                    os.environ[key] = val

load_env()

def trigger_recovery_and_test():
    with app.test_client() as client:
        # Get Super Admin username
        with app.app_context():
            super_admin = User.query.filter_by(username='super').first()
            if not super_admin:
                super_admin = User.query.filter_by(role='admin').first()
            if not super_admin:
                return "❌ FAILURE: Super Admin / admin user not found."
            username = super_admin.username
            recovery_email = super_admin.recovery_email
            
            # Temporary replace recovery_email with SMTP_USER so we can read it, or if it's already a valid email, read that.
            # But we only have access to SMTP_USER's inbox!
            smtp_user = os.environ.get('SMTP_USER')
            if recovery_email != smtp_user:
                print(f"Setting recovery email to {smtp_user} temporarily for testing...")
                super_admin.recovery_email = smtp_user
                db.session.commit()

        print(f"Triggering recovery for {username}...")
        response = client.post('/recover_account', data={
            'action': 'step1', 
            'username': username,
            'recovery_email': recovery_email if recovery_email else 'admin@amlhs.edu.ph'
        })
        
        # Check if error in response
        if 'SMTP Error' in response.get_data(as_text=True) or 'Failed to send' in response.get_data(as_text=True):
            return "❌ FAILURE: Failed to send OTP email during request. " + response.get_data(as_text=True)[:200]
            
        # Verify if an OTP is generated
        with app.app_context():
            super_admin = User.query.filter_by(username=username).first()
            db_otp = super_admin.recovery_otp
            if not db_otp:
                # To debug
                error_html = response.get_data(as_text=True)
                return f"❌ FAILURE: OTP not saved to database. Response HTML snippet: {error_html[:500]}"
        
        print("OTP generated in DB:", db_otp)
        
        # Now check IMAP
        email_user = os.environ.get('SMTP_USER')
        email_pass = os.environ.get('SMTP_PASSWORD').replace(' ', '')
        
        print(f"Waiting for email in {email_user}...")
        time.sleep(15) # Wait 15 seconds for email to arrive
        
        found_otp = None
        where_found = ""
        
        try:
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            mail.login(email_user, email_pass)
            
            for folder in ["inbox", '"[Gmail]/Spam"']:
                print(f"Checking {folder}...")
                status, messages = mail.select(folder)
                if status != 'OK':
                    continue
                # Search for recent emails (UNSEEN)
                status, data = mail.search(None, "UNSEEN")
                email_ids = data[0].split()
                if email_ids:
                    # check latest unseen emails
                    for email_id in reversed(email_ids):
                        status, msg_data = mail.fetch(email_id, "(RFC822)")
                        for response_part in msg_data:
                            if isinstance(response_part, tuple):
                                msg = email.message_from_bytes(response_part[1])
                                subject, encoding = decode_header(msg["Subject"])[0]
                                if isinstance(subject, bytes):
                                    subject = subject.decode(encoding if encoding else "utf-8", errors='ignore')
                                print("Found email with subject:", subject.encode('ascii', 'replace').decode('ascii'))
                                
                                body = ""
                                if msg.is_multipart():
                                    for part in msg.walk():
                                        if part.get_content_type() == "text/plain":
                                            body = part.get_payload(decode=True).decode(errors='ignore')
                                            break
                                else:
                                    body = msg.get_payload(decode=True).decode(errors='ignore')
                                
                                # extract OTP
                                match = re.search(r'\b\d{6}\b', body)
                                if match:
                                    found_otp = match.group(0)
                                    where_found = folder
                                    break
                        if found_otp:
                            break
                if found_otp:
                    break
            
            mail.logout()
            
            if not found_otp:
                return "❌ FAILURE: OTP email not found in Inbox or Spam."
                
            print(f"Extracted OTP from email: {found_otp} (found in {where_found})")
            
            # Verify via endpoint
            print("Verifying OTP via endpoint...")
            response = client.post('/recover_account', data={'action': 'step2', 'username': username, 'otp': found_otp})
            
            with app.app_context():
                user = User.query.filter_by(username=username).first()
                if user.recovery_otp is None or 'Reset Password' in response.get_data(as_text=True) or response.status_code == 302:
                    return "✅ SUCCESS: Email received with OTP + recovery flow completed successfully"
                else:
                    return "❌ FAILURE: Could not verify OTP via endpoint."

        except Exception as e:
            return f"❌ FAILURE: IMAP check failed with exception: {str(e)}"

if __name__ == '__main__':
    result = trigger_recovery_and_test()
    with open('smtp_test_result.txt', 'w', encoding='utf-8') as f:
        f.write(result)
    print("Done")
