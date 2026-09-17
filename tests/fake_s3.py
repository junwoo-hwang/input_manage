"""S3 없이 도는 가짜 저장소. s3io 와 같은 모양이면 된다."""
import hashlib

STORE: dict[str, bytes] = {}


def reset():
    STORE.clear()


def get_object(key):
    if key not in STORE:
        return None
    return STORE[key], hashlib.md5(STORE[key]).hexdigest()


def head_etag(key):
    return hashlib.md5(STORE[key]).hexdigest() if key in STORE else ""


def put_object(key, data):
    STORE[key] = data
    return head_etag(key)


def list_keys(prefix):
    return sorted(k for k in STORE if k.startswith(prefix))


def delete_object(key):
    STORE.pop(key, None)
