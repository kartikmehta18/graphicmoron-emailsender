#!/usr/bin/env python3
"""
Safari Bulk Email Sender
Upload Excel file → Write message → Send to all contacts at once
"""

from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import smtplib
import os
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from werkzeug.utils import secure_filename

app = Flask(__name__, static_folder='.')

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/upload', methods=['POST'])
def upload_excel():
    """Upload Excel and return list of emails found"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if not file.filename.endswith(('.xlsx', '.xls', '.csv')):
        return jsonify({'error': 'Please upload an Excel (.xlsx / .xls) or CSV file'}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        # Read file
        if filename.endswith('.csv'):
            df = pd.read_csv(filepath)
        else:
            # Try all sheets, merge results
            xl = pd.ExcelFile(filepath)
            frames = []
            for sheet in xl.sheet_names:
                try:
                    frames.append(pd.read_excel(filepath, sheet_name=sheet))
                except:
                    pass
            df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

        # Find email column (flexible detection)
        email_col = None
        for col in df.columns:
            col_lower = str(col).lower()
            if any(k in col_lower for k in ['email', 'mail', 'e-mail', 'contact']):
                email_col = col
                break

        if email_col is None:
            # Try to detect by content
            for col in df.columns:
                sample = df[col].dropna().astype(str)
                if sample.str.contains('@').sum() > 0:
                    email_col = col
                    break

        if email_col is None:
            return jsonify({'error': 'Could not find an email column in your file. Make sure a column contains email addresses.'}), 400

        # Extract valid emails
        emails_raw = df[email_col].dropna().astype(str).tolist()
        emails = [e.strip() for e in emails_raw if '@' in e and '.' in e]
        emails = list(dict.fromkeys(emails))  # deduplicate

        # Also grab names if available
        name_col = None
        for col in df.columns:
            if any(k in str(col).lower() for k in ['name', 'business', 'company', 'client']):
                name_col = col
                break

        contacts = []
        if name_col:
            for _, row in df.iterrows():
                email = str(row[email_col]).strip() if pd.notna(row[email_col]) else ''
                name  = str(row[name_col]).strip()  if pd.notna(row[name_col])  else ''
                if '@' in email:
                    contacts.append({'email': email, 'name': name})
        else:
            contacts = [{'email': e, 'name': ''} for e in emails]

        # Deduplicate by email
        seen = set()
        unique_contacts = []
        for c in contacts:
            if c['email'] not in seen:
                seen.add(c['email'])
                unique_contacts.append(c)

        return jsonify({
            'success': True,
            'contacts': unique_contacts,
            'total': len(unique_contacts),
            'email_column': email_col
        })

    except Exception as e:
        return jsonify({'error': f'Error reading file: {str(e)}'}), 500


@app.route('/send', methods=['POST'])
def send_emails():
    """Send bulk emails via SMTP"""
    data = request.get_json(silent=True) or {}

    if not isinstance(data, dict):
        return jsonify({'error': 'Invalid JSON payload'}), 400

    smtp_host     = data.get('smtp_host', 'smtp.gmail.com')
    try:
        smtp_port = int(data.get('smtp_port', 587))
    except (TypeError, ValueError):
        return jsonify({'error': 'SMTP port must be a valid number'}), 400

    sender_email  = data.get('sender_email', '').strip()
    sender_pass   = ''.join(str(data.get('sender_pass', '')).split())
    subject       = data.get('subject', '').strip()
    message_body  = data.get('message', '').strip()
    contacts      = data.get('contacts', [])

    # Validate
    if not sender_email or not sender_pass:
        return jsonify({'error': 'Please enter your email and password/app password'}), 400
    if not subject:
        return jsonify({'error': 'Subject line is required'}), 400
    if not message_body:
        return jsonify({'error': 'Message body is required'}), 400
    if not contacts:
        return jsonify({'error': 'No contacts to send to'}), 400

    results = {'sent': [], 'failed': []}

    server = None

    try:
        tls_context = ssl.create_default_context()

        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port, context=tls_context)
            server.ehlo()
        else:
            server = smtplib.SMTP(smtp_host, smtp_port)
            server.ehlo()
            server.starttls(context=tls_context)
            server.ehlo()

        server.login(sender_email, sender_pass)

        for contact in contacts:
            to_email = contact.get('email', '').strip()
            name     = contact.get('name', '')
            if not to_email:
                continue

            # Personalise greeting if name available
            personalised = message_body
            if name:
                personalised = f"Hi {name},\n\n" + message_body

            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From']    = sender_email
            msg['To']      = to_email

            # Plain text
            part1 = MIMEText(personalised, 'plain')
            # HTML version
            html_body = personalised.replace('\n', '<br>')
            html = f"""
            <html><body style="font-family:Arial,sans-serif;font-size:15px;color:#222;max-width:600px;margin:auto;padding:24px;">
            {html_body}
            </body></html>"""
            part2 = MIMEText(html, 'html')

            msg.attach(part1)
            msg.attach(part2)

            try:
                server.sendmail(sender_email, to_email, msg.as_string())
                results['sent'].append(to_email)
            except Exception as e:
                results['failed'].append({'email': to_email, 'error': str(e)})

        server.quit()

    except smtplib.SMTPAuthenticationError as exc:
        smtp_message = exc.smtp_error.decode('utf-8', errors='replace') if isinstance(exc.smtp_error, bytes) else str(exc.smtp_error)
        hint = 'For Gmail, use a 16-character App Password with 2-Step Verification enabled.'

        if 'application-specific password' in smtp_message.lower():
            hint = 'Gmail is rejecting the normal password. Create and use a Gmail App Password.'
        elif 'username and password not accepted' in smtp_message.lower():
            hint = 'Double-check the sender email, the app password, and the selected SMTP host/port.'
        elif '535' in smtp_message or '5.7.8' in smtp_message:
            hint = 'This usually means the account blocked the login. Use a Gmail App Password or confirm the SMTP provider settings.'

        return jsonify({
            'error': 'SMTP authentication failed.',
            'smtp_code': exc.smtp_code,
            'smtp_message': smtp_message,
            'hint': hint,
        }), 401
    except smtplib.SMTPNotSupportedError as exc:
        return jsonify({'error': f'SMTP AUTH is not supported by this server: {str(exc)}'}), 400
    except smtplib.SMTPResponseException as exc:
        smtp_message = exc.smtp_error.decode('utf-8', errors='replace') if isinstance(exc.smtp_error, bytes) else str(exc.smtp_error)
        return jsonify({'error': f'SMTP server rejected the request: {exc.smtp_code} {smtp_message}'}), 502
    except Exception as e:
        return jsonify({'error': f'SMTP connection error: {str(e)}'}), 500
    finally:
        if server is not None:
            try:
                server.quit()
            except Exception:
                pass

    return jsonify({
        'success': True,
        'sent_count': len(results['sent']),
        'failed_count': len(results['failed']),
        'sent': results['sent'],
        'failed': results['failed']
    })


if __name__ == '__main__':
    print("\n🚀 Safari Bulk Email Sender running at http://localhost:5000\n")
    app.run(debug=True, port=5000)
