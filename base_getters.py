import io

import pandas as pd
import s3fs


def get_content_from_s3_path(path: str) -> bytes:
    """reads content from an s3 file path

    Args:
        path(str): Full s3 URL
    Returns:
        bytes: The raw file content
    """
    fs = s3fs.S3FileSystem()
    with fs.open(path, mode="rb") as f:
        content = f.read()
    return content


def get_df_from_excel_s3_path(path: str, **kwargs) -> pd.DataFrame:
    """Loads an S3 Excel file into a pandas DataFrame
    Args:
        path(str): Full S3 URL
    Returns:
        pd.DataFrame: The loaded spreadsheet data
    """
    raw_data = get_content_from_s3_path(path)
    content_buffer = io.BytesIO(raw_data)
    df = pd.read_excel(content_buffer, engine="openpyxl", **kwargs)
    return df
