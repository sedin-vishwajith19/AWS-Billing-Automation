import os
import boto3
from botocore.exceptions import ClientError


def upload_report_to_s3(file_path):
    """
    Upload the generated report file to S3.

    The S3 upload always targets the 'RF Sandbox' account bucket.
    In Lambda, the execution role provides permissions automatically.
    For local runs, set S3_AWS_ACCESS_KEY_ID and S3_AWS_SECRET_ACCESS_KEY
    to use RF Sandbox credentials explicitly.

    Environment Variables:
        S3_BUCKET_NAME          — (Required) Target S3 bucket name.
        S3_KEY_PREFIX           — (Optional) Prefix/folder in the bucket.
                                  Defaults to 'reports/'.
        S3_OBJECT_ACL           — (Optional) ACL for public access.
                                  Defaults to 'public-read'. Set to 'none' to omit.
        S3_AWS_ACCESS_KEY_ID    — (Optional) RF Sandbox access key
                                  for local runs. Not needed in Lambda.
        S3_AWS_SECRET_ACCESS_KEY— (Optional) RF Sandbox secret key
                                  for local runs. Not needed in Lambda.
        AWS_REGION              — (Optional) AWS region.
                                  Defaults to 'ap-south-1'.

    Returns:
        dict with 's3_uri' and 'object_url' keys on success.

    Raises:
        ValueError  — If S3_BUCKET_NAME is not set.
        Exception   — If the upload fails.
    """

    bucket_name = os.environ.get("S3_BUCKET_NAME")

    if not bucket_name:
        raise ValueError(
            "S3_BUCKET_NAME environment variable is not set. "
            "Please configure it before uploading."
        )

    key_prefix = os.environ.get(
        "S3_KEY_PREFIX", "reports/"
    ).rstrip("/")

    region = os.environ.get("AWS_REGION", "ap-south-1")

    # Build S3 key from prefix + filename
    file_name = os.path.basename(file_path)
    s3_key = f"{key_prefix}/{file_name}"

    print(f"  Uploading to s3://{bucket_name}/{s3_key} ...")

    # Use explicit RF Sandbox credentials if provided
    # (for local testing). In Lambda, the execution role
    # handles authentication automatically.
    s3_access_key = os.environ.get("S3_AWS_ACCESS_KEY_ID")
    s3_secret_key = os.environ.get("S3_AWS_SECRET_ACCESS_KEY")

    if s3_access_key and s3_secret_key:
        s3_session = boto3.Session(
            aws_access_key_id=s3_access_key,
            aws_secret_access_key=s3_secret_key,
            region_name=region
        )
        s3_client = s3_session.client("s3")
    else:
        # In Lambda — uses execution role (RF Sandbox)
        s3_client = boto3.client("s3", region_name=region)

    extra_args = {
        "ContentType": (
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        )
    }

    s3_acl = os.environ.get("S3_OBJECT_ACL", "public-read")
    if s3_acl and s3_acl.lower() != "none":
        extra_args["ACL"] = s3_acl

    try:
        s3_client.upload_file(
            Filename=file_path,
            Bucket=bucket_name,
            Key=s3_key,
            ExtraArgs=extra_args
        )
    except Exception as e:
        # If bucket policy disables ACLs (AccessControlListNotSupported / BucketOwnerEnforced), retry without ACL
        if "AccessControlListNotSupported" in str(e) or "InvalidRequest" in str(e) or "ACL" in str(e):
            print("  ℹ Bucket disables ACLs (Bucket owner enforced) — retrying upload without ACL...")
            extra_args.pop("ACL", None)
            s3_client.upload_file(
                Filename=file_path,
                Bucket=bucket_name,
                Key=s3_key,
                ExtraArgs=extra_args
            )
        else:
            raise e

    s3_uri = f"s3://{bucket_name}/{s3_key}"

    print(f"  ✔ Upload complete: {s3_uri}")

    # Build direct permanent S3 object URL (without expiration)
    s3_base_url = os.environ.get("S3_BASE_URL")
    if s3_base_url:
        object_url = f"{s3_base_url.rstrip('/')}/{s3_key.lstrip('/')}"
    else:
        object_url = f"https://{bucket_name}.s3.{region}.amazonaws.com/{s3_key.lstrip('/')}"

    print(f"  ✔ Object URL: {object_url}")

    return {
        "s3_uri": s3_uri,
        "object_url": object_url
    }

