import os
import boto3
from botocore.exceptions import ClientError

def send_email_report(presigned_url, month_name):
    """
    Sends an email with the presigned S3 URL using Amazon SES.

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

    # You can customize the subject using an environment variable, or it defaults to this:
    custom_subject = os.environ.get("SES_EMAIL_SUBJECT")
    if custom_subject:
        subject = f"{custom_subject} - {month_name}"
    else:
        subject = f"📊 AWS Monthly Cost Report - {month_name}"
    
    body_text = (
        f"Hello,\n\n"
        f"The AWS Monthly Cost Report for {month_name} has been successfully generated.\n\n"
        f"You can securely download the Excel report using the following link (valid for 7 days):\n"
        f"{presigned_url}\n\n"
        f"This report contains the billing breakdown for all monitored AWS accounts.\n\n"
        f"Best regards,\n"
        f"AWS Cloud Operations Team"
    )

    body_html = f"""<html>
    <head>
      <style>
        body {{ font-family: Arial, sans-serif; color: #333; line-height: 1.6; }}
        .container {{ padding: 20px; }}
        .button {{ 
            background-color: #0052cc; color: white; padding: 10px 20px; 
            text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;
            margin: 15px 0;
        }}
        .footer {{ margin-top: 30px; font-size: 12px; color: #777; }}
      </style>
    </head>
    <body>
      <div class="container">
        <h2>AWS Monthly Cost Report - {month_name}</h2>
        <p>Hello,</p>
        <p>The AWS Monthly Cost Report for <strong>{month_name}</strong> has been successfully generated.</p>
        
        <p>This report contains the service-level billing breakdown for all monitored AWS accounts.</p>

        <a href="{presigned_url}" class="button">Download Excel Report</a>
        <p style="font-size: 12px; color: #555;"><em>Note: For security reasons, this download link will expire in 7 days.</em></p>
        
        <p>Best regards,<br/><strong>AWS Cloud Operations Team</strong></p>
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

    try:
        response = client.send_email(
            Destination={
                'ToAddresses': recipients,
            },
            Message={
                'Body': {
                    'Html': {
                        'Charset': "UTF-8",
                        'Data': body_html,
                    },
                    'Text': {
                        'Charset': "UTF-8",
                        'Data': body_text,
                    },
                },
                'Subject': {
                    'Charset': "UTF-8",
                    'Data': subject,
                },
            },
            Source=sender,
        )
    except ClientError as e:
        print(f"❌ Failed to send email: {e.response['Error']['Message']}")
        return False
    else:
        print(f"  ✔ Email sent successfully! Message ID: {response['MessageId']}")
        return True
