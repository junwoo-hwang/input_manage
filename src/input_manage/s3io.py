"""S3 drive 에 바이트를 읽고 쓴다. 사내 환경에 닿는 자리는 여기뿐이다.

설정과 클라이언트 만드는 법은 포털의 src/dc_ocap/dc_ocap.py 와 같게 맞췄다.
다른 점은 목적뿐이다 -- 거기는 .html 을 내려받아 보여주기만 하고, 여기는
.xlsx 를 읽고 다시 올린다.

자격증명은 환경변수에서만 읽는다. 코드에 적지 않는다.
"""
from __future__ import annotations

import os
import threading

import boto3
import botocore.exceptions

AWS_ACCESS_KEY = os.getenv("AWS_ACCESS_KEY")
AWS_SECRET_KEY = os.getenv("AWS_SECRET_KEY")

BUCKET_NAME = os.getenv("INPUT_S3_BUCKET", "G-DVC")
FOLDER_PATH = os.getenv("INPUT_S3_PREFIX", "2GAPU/input").strip("/")
S3_ENDPOINT = os.getenv("INPUT_S3_ENDPOINT", "http://s3.dataplatform.samsungds.net:9020")

_lock = threading.Lock()
_client = None


def client():
    """클라이언트 하나를 만들어 재사용한다.

    streamlit 은 사람이 버튼 한 번 누를 때마다 스크립트를 처음부터 다시
    돌린다. 그때마다 새로 만들면 매번 연결을 다시 맺느라 저장이 눈에 띄게
    느려진다.
    """
    global _client
    with _lock:
        if _client is None:
            _client = boto3.client(
                service_name="s3",
                aws_access_key_id=AWS_ACCESS_KEY,
                aws_secret_access_key=AWS_SECRET_KEY,
                endpoint_url=S3_ENDPOINT,
            )
        return _client


def get_object(key: str) -> tuple[bytes, str] | None:
    """(내용, 버전표). 없으면 None.

    버전표는 S3 가 주는 ETag(내용 해시)를 그대로 쓴다. '내가 화면에 띄운 뒤
    다른 사람이 먼저 저장했는가' 를 가리는 데 쓰는데, 그 판단 하나 하자고
    파일을 통째로 다시 받아 해시를 뜰 필요가 없다는 것이 ETag 의 쓸모다.
    """
    try:
        got = client().get_object(Bucket=BUCKET_NAME, Key=key)
    except botocore.exceptions.ClientError as err:
        if err.response["Error"]["Code"] in ("NoSuchKey", "404", "NoSuchBucket"):
            return None
        raise
    return got["Body"].read(), got.get("ETag", "").strip('"')


def head_etag(key: str) -> str:
    """지금 올라가 있는 것의 버전표. 없으면 빈 글자."""
    try:
        got = client().head_object(Bucket=BUCKET_NAME, Key=key)
    except botocore.exceptions.ClientError:
        return ""
    return got.get("ETag", "").strip('"')


def put_object(key: str, data: bytes) -> str:
    client().put_object(Bucket=BUCKET_NAME, Key=key, Body=data)
    return head_etag(key)


def list_keys(prefix: str) -> list[str]:
    out: list[str] = []
    paginator = client().get_paginator("list_objects_v2")
    for page in paginator.paginate(Bucket=BUCKET_NAME, Prefix=prefix):
        for obj in page.get("Contents", []):
            if not obj["Key"].endswith("/"):
                out.append(obj["Key"])
    return sorted(out)


def delete_object(key: str) -> None:
    client().delete_object(Bucket=BUCKET_NAME, Key=key)
