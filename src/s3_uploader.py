import os
import boto3


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
        S3_AWS_ACCESS_KEY_ID    — (Optional) RF Sandbox access key
                                  for local runs. Not needed in Lambda.
        S3_AWS_SECRET_ACCESS_KEY— (Optional) RF Sandbox secret key
                                  for local runs. Not needed in Lambda.
        AWS_REGION              — (Optional) AWS region.
                                  Defaults to 'ap-south-1'.

    Returns:
        str — The full S3 URI (s3://bucket/key) of the uploaded file.

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

    s3_client.upload_file(
        Filename=file_path,
        Bucket=bucket_name,
        Key=s3_key,
        ExtraArgs={
            "ContentType": (
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        }
    )

    s3_uri = f"s3://{bucket_name}/{s3_key}"

    print(f"  ✔ Upload complete: {s3_uri}")

    # Generate a presigned URL valid for 7 days (604800 seconds)
    try:
        presigned_url = s3_client.generate_presigned_url(
            'get_object',
            Params={'Bucket': bucket_name, 'Key': s3_key},
            ExpiresIn=604800
        )
    except Exception as e:
        print(f"  ⚠ Failed to generate presigned URL: {e}")
        presigned_url = None

    return {
        "s3_uri": s3_uri,
        "presigned_url": presigned_url
    }
