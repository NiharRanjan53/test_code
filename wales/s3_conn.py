import boto3
import os
import glob
import pandas as pd

from io import BytesIO
from botocore.exceptions import ClientError
from configtest import BUCKET_NAME
from openpyxl import load_workbook


# =========================================================
# CONFIG
# =========================================================

# LOCAL -> Read/write from local folder
# S3    -> Read/write from S3

# Example:
# set STORAGE_MODE=LOCAL
# set STORAGE_MODE=S3

LOCAL_DATA_ROOT = os.getenv("LOCAL_DATA_ROOT", os.getcwd())


def get_storage_mode():
    """
    Dynamically fetch storage mode.
    Prevents issue where env variable changes after import.
    """
    return os.getenv("STORAGE_MODE", "LOCAL").upper()


# =========================================================
# S3 SESSION
# =========================================================

# def get_s3_session():
#     return boto3.Session()   # ECS IAM role


def get_s3_session():

    storage_mode = get_storage_mode()

    print(f"====== STORAGE MODE : {storage_mode}")

    # LOCAL MODE
    # S3 client won't be used but returning harmless session
    if storage_mode == "LOCAL":
        return boto3.Session()

    # RUNNING INSIDE AWS
    if os.getenv("AWS_EXECUTION_ENV"):
        return boto3.Session()

    # LOCAL MACHINE WITH AWS PROFILE
    try:
        return boto3.Session(
            profile_name="390070096525_med-av-daas-prod/aif-ces-smba-developer"
        )

    except Exception:
        return boto3.Session()


# =========================================================
# LOCAL PATH RESOLVER
# =========================================================

def _resolve_local_path(path_or_key: str) -> str:
    """
    Convert S3-like path to local filesystem path.

    Example:
    SMBA/Contract/file.xlsx
    ->
    D:/local_data/SMBA/Contract/file.xlsx
    """

    if os.path.isabs(path_or_key):
        return path_or_key

    return os.path.join(
        LOCAL_DATA_ROOT,
        path_or_key.replace("/", os.sep)
    )


# =========================================================
# READ EXCEL
# =========================================================

def read_excel_from_s3(file_path: str, return_bytes=False, **kwargs):
    """
    Fetch Excel file.

    LOCAL MODE:
        Reads from local filesystem.

    S3 MODE:
        Downloads from S3.

    return_bytes=True:
        returns raw bytes

    else:
        returns DataFrame
    """

    storage_mode = get_storage_mode()

    kwargs.setdefault("engine", "openpyxl")

    # =====================================================
    # LOCAL MODE
    # =====================================================

    if storage_mode == "LOCAL":

        local_path = _resolve_local_path(file_path)

        print(f"📂 Reading LOCAL file: {local_path}")

        if not os.path.exists(local_path):
            raise FileNotFoundError(
                f"Local file not found: {local_path}"
            )

        if return_bytes:
            with open(local_path, "rb") as f:
                return f.read()

        return pd.read_excel(local_path, **kwargs)

    # =====================================================
    # S3 MODE
    # =====================================================

    s3 = get_s3_session().client("s3")

    print(
        f"🔎 Fetching from S3 bucket={BUCKET_NAME}, key={file_path}"
    )

    obj = s3.get_object(
        Bucket=BUCKET_NAME,
        Key=file_path
    )

    file_bytes = obj["Body"].read()

    if return_bytes:
        return file_bytes

    return pd.read_excel(
        BytesIO(file_bytes),
        **kwargs
    )


# =========================================================
# LIST FILES
# =========================================================

def list_files_in_s3(directory: str):
    """
    List all files.

    LOCAL MODE:
        Reads local folder recursively.

    S3 MODE:
        Reads S3 prefix recursively.
    """

    storage_mode = get_storage_mode()

    # =====================================================
    # LOCAL MODE
    # =====================================================

    if storage_mode == "LOCAL":

        local_dir = _resolve_local_path(directory)

        print(f"📂 Listing LOCAL directory: {local_dir}")

        files = []

        if os.path.exists(local_dir):

            for f in glob.glob(
                local_dir + "/**/*",
                recursive=True
            ):

                if os.path.isfile(f):

                    files.append(
                        os.path.relpath(
                            f,
                            LOCAL_DATA_ROOT
                        ).replace("\\", "/")
                    )

        return files

    # =====================================================
    # S3 MODE
    # =====================================================

    s3 = get_s3_session().client("s3")

    paginator = s3.get_paginator("list_objects_v2")

    keys = []

    for page in paginator.paginate(
        Bucket=BUCKET_NAME,
        Prefix=directory
    ):

        for obj in page.get("Contents", []):

            keys.append(obj["Key"])

    return keys


# =========================================================
# SAVE EXCEL
# =========================================================

def save_excel_to_s3(data, file_path: str):
    """
    Save DataFrame or dict of DataFrames.

    LOCAL MODE:
        Saves to local filesystem.

    S3 MODE:
        Uploads to S3.
    """

    storage_mode = get_storage_mode()

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # MULTIPLE SHEETS
        if isinstance(data, dict):

            for sheet, df in data.items():

                if not isinstance(df, pd.DataFrame):
                    raise TypeError(
                        f"{sheet} is not a DataFrame"
                    )

                df.to_excel(
                    writer,
                    sheet_name=str(sheet)[:31],
                    index=False
                )

        # SINGLE SHEET
        else:

            if not isinstance(data, pd.DataFrame):
                raise TypeError(
                    "data must be DataFrame or dict of DataFrames"
                )

            data.to_excel(
                writer,
                sheet_name="Sheet1",
                index=False
            )

    output.seek(0)

    # =====================================================
    # LOCAL MODE
    # =====================================================

    if storage_mode == "LOCAL":

        local_out = _resolve_local_path(file_path)

        os.makedirs(
            os.path.dirname(local_out),
            exist_ok=True
        )

        with open(local_out, "wb") as f:
            f.write(output.getvalue())

        print(
            f"✅ Excel successfully saved locally: {local_out}"
        )

        return local_out

    # =====================================================
    # S3 MODE
    # =====================================================

    s3 = get_s3_session().client("s3")

    s3.put_object(
        Bucket=BUCKET_NAME,
        Key=file_path,
        Body=output.getvalue()
    )

    print(
        f"✅ Excel successfully updated in S3: "
        f"s3://{BUCKET_NAME}/{file_path}"
    )

    return f"s3://{BUCKET_NAME}/{file_path}"


# =========================================================
# FILE EXISTS
# =========================================================

def s3_file_exists(file_path: str) -> bool:
    """
    Check if file exists.

    LOCAL MODE:
        Checks local filesystem.

    S3 MODE:
        Checks S3 object.
    """

    storage_mode = get_storage_mode()

    # =====================================================
    # LOCAL MODE
    # =====================================================

    if storage_mode == "LOCAL":

        local_path = _resolve_local_path(file_path)

        return os.path.exists(local_path)

    # =====================================================
    # S3 MODE
    # =====================================================

    s3 = get_s3_session().client("s3")

    try:

        s3.head_object(
            Bucket=BUCKET_NAME,
            Key=file_path
        )

        return True

    except ClientError as e:

        if e.response["Error"]["Code"] in (
            "404",
            "NoSuchKey"
        ):
            return False

        raise


# =========================================================
# OPTIONAL HELPER
# =========================================================

def upload_file_to_s3(local_file_path, s3_key):
    """
    Upload local file to S3.

    Useful for manual sync/testing.
    """

    s3 = get_s3_session().client("s3")

    s3.upload_file(
        local_file_path,
        BUCKET_NAME,
        s3_key
    )

    print(
        f"✅ Uploaded {local_file_path} "
        f"to s3://{BUCKET_NAME}/{s3_key}"
    )