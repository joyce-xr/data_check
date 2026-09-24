"""接口请求签名工具：生成 N（nonce）和 S（sign）请求头。"""

import base64
import hashlib
import json
import random
import time
from urllib.parse import urlparse


def _gen_nonce() -> str:
    """请求头 N：13 位毫秒时间戳 + 7 位随机数字。"""
    return str(int(time.time() * 1000)) + "".join(random.choices("0123456789", k=7))


def _jwt_payload(token: str) -> dict:
    """解码 JWT 载荷，取 blockId / userId 用于签名。"""
    b64 = token.split(".")[1]
    b64 += "=" * (-len(b64) % 4)
    return json.loads(base64.urlsafe_b64decode(b64))


def _gen_sign(nonce: str, token: str, path: str) -> str:
    """请求头 S，复刻前端算法：
    MD5(blockId + userId + path + token[time各位数字作索引] + token[倒序random各位数字作索引])
    path 为去掉域名后的完整路径（含 /api 前缀）。
    """
    token = token.replace("Bearer ", "")
    ts, rnd = nonce[:13], nonce[13:]
    payload = _jwt_payload(token)
    sign_str = str(payload["blockId"]) + str(payload["userId"]) + path
    for c in ts:
        sign_str += token[int(c)]
    for c in rnd:
        sign_str += token[len(token) - 1 - int(c)]
    return hashlib.md5(sign_str.encode("utf-8")).hexdigest()


def _build_headers(url: str, key: str) -> dict:
    """按次构建请求头，N/S 每次请求重新生成并配套。"""
    nonce = _gen_nonce()
    return {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json; charset=UTF-8",
        "N": nonce,
        "S": _gen_sign(nonce, key, urlparse(url).path),
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
    }
