import os
import smtplib
import logging
from email.message import EmailMessage

SMTP_SERVER = os.environ.get('SMTP_SERVER')
SMTP_PORT = int(os.environ.get('SMTP_PORT', 587))
SMTP_EMAIL = os.environ.get('SMTP_EMAIL')
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD')

def send_email_with_attachments(to_email, subject, html_body, attachments=None, inline_images=None, sender_name=None, debug=False):
    """Send a single email with optional attachments and inline images.
    inline_images: list of tuples (path, cid)
    Returns (True, None) on success or (False, error_str)."""
    try:
        msg = EmailMessage()
        display_from = f"{sender_name} <{SMTP_EMAIL}>" if sender_name else SMTP_EMAIL
        msg['From'] = display_from
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.set_content('This email requires an HTML-capable client.')
        msg.add_alternative(html_body, subtype='html')

        # Attach files if provided
        for path in (attachments or []):
            try:
                with open(path, 'rb') as f:
                    data = f.read()
                maintype = 'application'
                filename = os.path.basename(path)
                msg.add_attachment(data, maintype=maintype, subtype='octet-stream', filename=filename)
            except Exception as e:
                logging.exception('Attachment read failed: %s', path)
                return False, f'Attachment {path} error: {e}'

        # Add inline images as related to the HTML part
        if inline_images:
            # find the html part
            html_part = None
            for part in msg.iter_parts():
                if part.get_content_type() == 'text/html':
                    html_part = part
                    break
            if html_part is None:
                logging.error('No HTML part found to attach inline images')
            else:
                import mimetypes
                for img_path, cid in inline_images:
                    try:
                        with open(img_path, 'rb') as f:
                            img_data = f.read()
                        mimetype, _ = mimetypes.guess_type(img_path)
                        if mimetype:
                            maintype, subtype = mimetype.split('/', 1)
                        else:
                            maintype, subtype = 'image', 'jpeg'
                        html_part.add_related(img_data, maintype=maintype, subtype=subtype, cid=f'<{cid}>', filename=os.path.basename(img_path))
                    except Exception as e:
                        logging.exception('Inline image attach failed: %s', img_path)
                        return False, f'Inline image {img_path} error: {e}'

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=60) as smtp:
            smtp.ehlo()
            if debug:
                smtp.set_debuglevel(1)
            smtp.starttls()
            smtp.login(SMTP_EMAIL, SMTP_PASSWORD)
            smtp.send_message(msg)

        return True, None
    except Exception as e:
        logging.exception('Failed to send email to %s', to_email)
        return False, str(e)
