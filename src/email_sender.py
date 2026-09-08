import os
import boto3
from botocore.exceptions import ClientError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def send_email_report(object_url, month_name, file_path=None):
    """
    Sends an email notification via Amazon SES with the S3 object download link
    and optional Excel report attachment.

    Environment Variables:
        SES_SENDER_EMAIL    — (Required) The email address to send from. Must be verified in SES.
        SES_RECIPIENT_EMAIL — (Required) Comma-separated list of recipient emails.
        AWS_REGION          — (Optional) AWS region for SES. Defaults to 'ap-south-1'.
    """
    sender = os.environ.get("SES_SENDER_EMAIL")
    recipient_str = os.environ.get("SES_RECIPIENT_EMAIL")
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
        f"Hello,\n\n",
        f"The AWS Monthly Cost Report for {month_name} has been successfully generated.\n\n"
    ]

    if object_url:
        body_text_parts.append(
            f"You can download the Excel report using the following link:\n{object_url}\n\n"
        )
    if file_path:
        body_text_parts.append("The report is also attached to this email.\n\n")

    body_text_parts.append(
        f"This report contains the billing breakdown for all monitored AWS accounts.\n\n"
        f"Best regards,\n"
        f"DevOps Team"
    )
    body_text = "".join(body_text_parts)

    download_section_html = ""
    if object_url:
        download_section_html = f"""
        <p>
            <a href="{object_url}" class="button" style="background-color: #0052cc; color: #ffffff; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block; margin: 15px 0;">Download Excel Report</a>
        </p>
        <p style="font-size: 12px; color: #555;">Direct link: <a href="{object_url}">{object_url}</a></p>
        """

    attachment_text_html = "<p>Please find the generated Excel report attached to this email.</p>" if file_path else ""

    body_html = f"""<html>
    <head>
      <style>
        body {{ font-family: Arial, sans-serif; color: #333; line-height: 1.6; }}
        .container {{ padding: 20px; }}
        .footer {{ margin-top: 30px; font-size: 12px; color: #777; }}
      </style>
    </head>
    <body>
      <div class="container">
        <h2>AWS Monthly Cost Report - {month_name}</h2>
        <p>Hello,</p>
        <p>The AWS Monthly Cost Report for <strong>{month_name}</strong> has been successfully generated.</p>
        
        <p>This report contains the service-level billing breakdown for all monitored AWS accounts.</p>
        {download_section_html}
        {attachment_text_html}
        
        <p>Best regards,<br/><strong>DevOps Team</strong></p>
      </div>
    </body>
    </html>
    """

    print(f"  Sending email to {len(recipients)} recipient(s) via SES...")

    # For local testing vs Lambda execution
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

    # Construct email message
    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ", ".join(recipients)

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
                part = MIMEApplication(attachment.read())
                part.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=filename
                )
                msg.attach(part)
        except Exception as e:
            print(f"⚠ Failed to read or attach report file: {e}")

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


