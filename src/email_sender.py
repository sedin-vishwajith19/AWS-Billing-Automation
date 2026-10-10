import os
import smtplib
import base64
import boto3
from botocore.exceptions import ClientError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formatdate, parseaddr

def _decode_credential(val):
    if not val:
        return val
    val_str = val.strip()
    try:
        return base64.b64decode(val_str).decode('utf-8')
    except Exception:
        return val_str

def send_email_report(object_url, month_name, file_path=None):
    """
    Sends an email notification via Amazon SES (using SMTP or boto3 client)
    with the S3 object download link and optional Excel report attachment.

    Environment Variables for SES SMTP:
        SES_SMTP_HOST       — (Optional) SES SMTP Endpoint host (e.g. email-smtp.ap-south-1.amazonaws.com).
        SES_SMTP_PORT       — (Optional) SES SMTP Port (defaults to 587).
        SES_SMTP_USERNAME   — (Optional) SES SMTP Username.
        SES_SMTP_PASSWORD   — (Optional) SES SMTP Password.

    Environment Variables for SES API (Fallback):
        SES_SENDER_EMAIL    — (Required) The email address to send from. Must be verified in SES.
        SES_RECIPIENT_EMAIL — (Required) Comma-separated list of recipient emails.
        AWS_REGION          — (Optional) AWS region for SES. Defaults to 'ap-south-1'.
    """
    sender = os.environ.get("SES_SENDER_EMAIL") or os.environ.get("SMTP_SENDER_EMAIL")
    recipient_str = os.environ.get("SES_RECIPIENT_EMAIL") or os.environ.get("SMTP_RECIPIENT_EMAIL")
    region = os.environ.get("AWS_REGION", "ap-south-1")

    if not sender or not recipient_str:
        print("ℹ SES_SENDER_EMAIL or SES_RECIPIENT_EMAIL not set. Skipping email notification.")
        return False

    recipients = [email.strip() for email in recipient_str.split(",") if email.strip()]
    
    if not recipients:
        print("ℹ No valid recipient emails found. Skipping email notification.")
        return False

    custom_subject = os.environ.get("SES_EMAIL_SUBJECT")
    if custom_subject:
        subject = f"{custom_subject} - {month_name}"
    else:
        subject = f"📊 AWS Monthly Cost Report - {month_name}"
    
    body_text_parts = [
        f"AWS Monthly Cost Report – {month_name}\n\n",
        f"Hello Team,\n\n",
        f"Please find attached the AWS Monthly Cost Report for {month_name}, containing the account-level and service-level billing breakdown for our AWS workload accounts.\n\n"
    ]

    if object_url:
        body_text_parts.append(
            f"Report Download URL:\n{object_url}\n\n"
        )

    body_text_parts.append(
        f"Best regards,\n"
        f"DevOps Team"
    )
    body_text = "".join(body_text_parts)

    download_section_html = ""
    if object_url:
        download_section_html = f"""
        <div style="text-align: center; margin: 24px 0 16px 0;">
            <a href="{object_url}" style="background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%); color: #ffffff !important; font-weight: 600; font-size: 15px; text-decoration: none; padding: 12px 28px; border-radius: 6px; display: inline-block; box-shadow: 0 2px 6px rgba(2, 132, 199, 0.3);">📊 Download Excel Report</a>
        </div>
        """

    body_html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>AWS Monthly Cost Report</title>
</head>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #f4f6f9; margin: 0; padding: 0; -webkit-font-smoothing: antialiased;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f4f6f9; padding: 30px 0;">
    <tr>
      <td align="center">
        <table width="100%" cellpadding="0" cellspacing="0" style="max-width: 600px; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05); border: 1px solid #e2e8f0; text-align: left;">
          <!-- Header Banner -->
          <tr>
            <td style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 28px 30px; color: #ffffff;">
              <span style="display: inline-block; background-color: rgba(255, 255, 255, 0.15); color: #38bdf8; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; padding: 4px 10px; border-radius: 20px; margin-bottom: 12px;">AWS Monthly Cost Report</span>
              <h1 style="margin: 0; font-size: 22px; font-weight: 700; letter-spacing: -0.5px; color: #ffffff;">AWS Monthly Cost Report – {month_name}</h1>
            </td>
          </tr>
          <!-- Body Content -->
          <tr>
            <td style="padding: 32px 30px; color: #334155; font-size: 15px; line-height: 1.6;">
              <div style="font-size: 18px; font-weight: 700; color: #0f172a; margin-bottom: 16px;">AWS Monthly Cost Report – {month_name}</div>
              <div style="font-size: 15px; font-weight: 600; color: #0f172a; margin-bottom: 16px;">Hello Team,</div>
              <p style="margin: 0 0 16px 0;">Please find attached the <strong>AWS Monthly Cost Report for {month_name}</strong>, containing the account-level and service-level billing breakdown for our AWS workload accounts.</p>

              {download_section_html}

              <p style="margin-top: 24px; line-height: 1.5;">Best regards,<br><strong style="color: #0f172a;">DevOps Team</strong></p>
            </td>
          </tr>
          <!-- Footer -->
          <tr>
            <td style="background-color: #f8fafc; padding: 20px 30px; border-top: 1px solid #e2e8f0; text-align: center; font-size: 12px; color: #94a3b8;">
              <p style="margin: 0;">© AWS Monthly Cost & Billing Automation</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"""

    # Construct email message
    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ", ".join(recipients)
    msg['Date'] = formatdate(localtime=True)

    # Attach email body (text and HTML)
    msg_body = MIMEMultipart('alternative')
    msg_body.attach(MIMEText(body_text, 'plain', 'utf-8'))
    msg_body.attach(MIMEText(body_html, 'html', 'utf-8'))
    msg.attach(msg_body)

    # Attach Excel file if available
    if file_path and os.path.exists(file_path):
        try:
            with open(file_path, 'rb') as attachment:
                filename = os.path.basename(file_path)
                part = MIMEApplication(
                    attachment.read(),
                    _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                part.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=filename
                )
                msg.attach(part)
                print(f"  ✔ Attached file: {filename}")
        except Exception as e:
            print(f"⚠ Failed to read or attach report file: {e}")
    else:
        if file_path:
            print(f"⚠ Report file not found at: {file_path}")

    # Check if SES SMTP credentials are provided in .env
    smtp_host = os.environ.get("SES_SMTP_HOST") or os.environ.get("SMTP_HOST") or os.environ.get("SES_ENDPOINT")
    smtp_port_val = os.environ.get("SES_SMTP_PORT") or os.environ.get("SMTP_PORT") or "587"
    smtp_user_raw = os.environ.get("SES_SMTP_USERNAME") or os.environ.get("SMTP_USERNAME")
    smtp_pass_raw = os.environ.get("SES_SMTP_PASSWORD") or os.environ.get("SMTP_PASSWORD")
    smtp_user = _decode_credential(smtp_user_raw)
    smtp_pass = _decode_credential(smtp_pass_raw)

    if smtp_host and smtp_user and smtp_pass:
        try:
            smtp_port = int(smtp_port_val)
            print(f"  Sending email to {recipients} from {sender} via SES SMTP ({smtp_host}:{smtp_port})...")
            server = smtplib.SMTP(smtp_host, smtp_port)
            server.starttls()
            server.login(smtp_user, smtp_pass)
            envelope_sender = parseaddr(sender)[1] or sender
            server.sendmail(envelope_sender, recipients, msg.as_string())
            server.quit()
            print("  ✔ Email sent successfully via SES SMTP!")
            return True
        except Exception as e:
            print(f"⚠ Failed to send email via SES SMTP: {e}")
            print("  Falling back to boto3 SES client...")
    elif smtp_host:
        print("⚠ SES_SMTP_HOST is set, but SES_SMTP_USERNAME or SES_SMTP_PASSWORD is missing in environment variables.")
        print("  Falling back to boto3 SES client...")

    # Fallback to AWS SDK boto3 SES client
    print(f"  Sending email to {len(recipients)} recipient(s) via boto3 SES SDK...")
    s3_access_key = os.environ.get("S3_AWS_ACCESS_KEY_ID")
    s3_secret_key = os.environ.get("S3_AWS_SECRET_ACCESS_KEY")

    if s3_access_key and s3_secret_key:
        ses_session = boto3.Session(
            aws_access_key_id=s3_access_key,
            aws_secret_access_key=s3_secret_key,
            region_name=region
        )
        client = ses_session.client('ses')
    else:
        client = boto3.client('ses', region_name=region)

    try:
        response = client.send_raw_email(
            Source=sender,
            Destinations=recipients,
            RawMessage={
                'Data': msg.as_string(),
            }
        )
    except ClientError as e:
        print(f"❌ Failed to send email: {e.response['Error']['Message']}")
        return False
    else:
        print(f"  ✔ Email sent successfully! Message ID: {response['MessageId']}")
        return True


