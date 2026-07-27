from src.config_loader import load_accounts

from src.aws_cost import (
    get_monthly_cost,
    get_last_5_months_cost,
    BILLING_FILTER
)

from src.service_analysis import (
    get_service_costs,
    find_primary_reason
)

from src.report_generator import create_multi_account_report
from src.s3_uploader import upload_report_to_s3
from src.email_sender import send_email_report

import os
import boto3

from datetime import datetime
from dateutil.relativedelta import relativedelta


def process_accounts():
    """
    Core logic — fetch costs for all accounts, generate the
    Excel report, and upload it to S3.

    Returns:
        dict with 'file_path' and 's3_uri' keys on success.
    """

    # Load accounts
    accounts = load_accounts()

    today = datetime.today()

    current_month_start = today.replace(day=1)

    previous_month_start = (
        current_month_start
        - relativedelta(months=1)
    )

    two_months_ago_start = (
        current_month_start
        - relativedelta(months=2)
    )

    all_results = []

    for account in accounts:

        print("\n----------------------------")
        print("Processing account:", account["name"])
        print("----------------------------")

        try:

            # Initialize session based on provided credentials
            if "profile" in account:
                session = boto3.Session(
                    profile_name=account["profile"]
                )
            elif "access_key" in account and "secret_key" in account:
                session = boto3.Session(
                    aws_access_key_id=account["access_key"],
                    aws_secret_access_key=account["secret_key"]
                )
            else:
                raise ValueError(f"No valid credentials found for account {account['name']}")

            prev_start = two_months_ago_start.strftime('%Y-%m-%d')
            prev_end = previous_month_start.strftime('%Y-%m-%d')
            curr_start = previous_month_start.strftime('%Y-%m-%d')
            curr_end = current_month_start.strftime('%Y-%m-%d')

            # ══════════════════════════════════════
            # WITH TAX — Full Cost Explorer cost
            # ══════════════════════════════════════

            print("  Fetching costs with tax...")

            prev_cost_with_tax = get_monthly_cost(
                session, prev_start, prev_end
            )

            curr_cost_with_tax = get_monthly_cost(
                session, curr_start, curr_end
            )

            prev_services_with_tax = get_service_costs(
                session, prev_start, prev_end
            )

            curr_services_with_tax = get_service_costs(
                session, curr_start, curr_end
            )

            reasons_with_tax = find_primary_reason(
                prev_services_with_tax,
                curr_services_with_tax
            )

            history_with_tax = get_last_5_months_cost(
                session
            )

            # ══════════════════════════════════════
            # WITHOUT TAX — Billing (filtered)
            # ══════════════════════════════════════

            print("  Fetching costs without tax...")

            prev_cost_without_tax = get_monthly_cost(
                session, prev_start, prev_end,
                cost_filter=BILLING_FILTER
            )

            curr_cost_without_tax = get_monthly_cost(
                session, curr_start, curr_end,
                cost_filter=BILLING_FILTER
            )

            prev_services_without_tax = get_service_costs(
                session, prev_start, prev_end,
                cost_filter=BILLING_FILTER
            )

            curr_services_without_tax = get_service_costs(
                session, curr_start, curr_end,
                cost_filter=BILLING_FILTER
            )

            reasons_without_tax = find_primary_reason(
                prev_services_without_tax,
                curr_services_without_tax
            )

            history_without_tax = get_last_5_months_cost(
                session,
                cost_filter=BILLING_FILTER
            )

            # ---------- Store Results ----------

            all_results.append({

                "Account Name": account["name"],
                "Account ID": account["id"],

                # With Tax (Cost Explorer)
                "Previous Month (With Tax)": prev_cost_with_tax,
                "Current Month (With Tax)": curr_cost_with_tax,
                "Reasons (With Tax)": reasons_with_tax,
                "History (With Tax)": history_with_tax,

                # Without Tax (Billing)
                "Previous Month (Without Tax)": prev_cost_without_tax,
                "Current Month (Without Tax)": curr_cost_without_tax,
                "Reasons (Without Tax)": reasons_without_tax,
                "History (Without Tax)": history_without_tax,

                # Service-level breakdown (current month)
                "Services (With Tax)": curr_services_with_tax,
                "Services (Without Tax)": curr_services_without_tax,

            })

            print("✔ Success:", account["name"])

        except Exception as e:

            print("❌ Failed:", account["name"])
            print("Error:", str(e))

            # Continue to next account
            continue


    # ---------- Final Report Creation ----------

    if len(all_results) == 0:

        print("\n⚠ No successful accounts processed.")
        print("Report not created.")

        return {
            "status": "error",
            "message": "No successful accounts processed."
        }

    print("\nGenerating Excel report...")

    file_path = create_multi_account_report(all_results)

    print("✔ Report created:", file_path)

    # ---------- Upload to S3 (Optional) ----------
    s3_uri = None
    presigned_url = None

    if os.environ.get("S3_BUCKET_NAME"):
        print("\nUploading report to S3...")
        try:
            s3_result = upload_report_to_s3(file_path)
            s3_uri = s3_result.get("s3_uri")
            presigned_url = s3_result.get("presigned_url")
            print("✔ S3 upload complete:", s3_uri)
        except Exception as e:
            print(f"⚠ S3 upload failed: {e}")
    else:
        print("\nℹ S3_BUCKET_NAME not set — skipping S3 upload.")

    # ---------- Send Email Notification ----------
    month_name = previous_month_start.strftime('%B %Y')
    print("\nSending email notification...")
    send_email_report(file_path, month_name)

    return {
        "status": "success",
        "file_path": file_path,
        "s3_uri": s3_uri,
        "presigned_url": presigned_url
    }


# ═══════════════════════════════════════════
# Lambda Handler
# ═══════════════════════════════════════════

def lambda_handler(event, context):
    """
    AWS Lambda entry point.

    Required environment variables:
        AWS_ACCOUNTS    — Comma-separated account list
                          (name:id:access_key:secret_key)
        S3_BUCKET_NAME  — Target S3 bucket for the report

    Optional environment variables:
        S3_KEY_PREFIX   — S3 folder prefix (default: 'reports/')
        AWS_REGION      — AWS region (default: 'ap-south-1')
    """

    print("Lambda invoked — starting billing report generation")

    try:
        result = process_accounts()

        return {
            "statusCode": 200,
            "body": result
        }

    except Exception as e:

        print(f"❌ Lambda execution failed: {str(e)}")

        return {
            "statusCode": 500,
            "body": {
                "status": "error",
                "message": str(e)
            }
        }


# ═══════════════════════════════════════════
# Local Execution
# ═══════════════════════════════════════════

if __name__ == "__main__":
    result = process_accounts()
    print("\nResult:", result)