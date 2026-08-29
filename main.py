import os

import requests
from dotenv import load_dotenv

# 加载 .env 文件中的配置（不覆盖已存在的环境变量）
load_dotenv()

# 从 .env / 环境变量读取配置，代码中不保留任何敏感信息
API_URL = os.getenv("API_URL", "")
API_KEY = os.getenv("API_KEY", "")
API_NONCE = os.getenv("API_NONCE", "")  # 请求头 N
API_SIGN = os.getenv("API_SIGN", "")  # 请求头 S

HEADERS = {
    "Authorization": f"Bearer {API_KEY}",  # 方式1: Bearer Token
    # "X-API-Key": API_KEY,                # 方式2: 自定义 Key 头
    "Content-Type": "application/json; charset=UTF-8",
    "N": API_NONCE,
    "S": API_SIGN,
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
}


def _print_request(resp: requests.Response) -> None:
    """打印实际发送的完整请求头和请求体（来自 requests 真实发出的请求）。"""
    req = resp.request
    print("=" * 60)
    print(f"{req.method} {req.url}")
    print("--- 请求头 ---")
    for k, v in req.headers.items():
        print(f"{k}: {v}")
    print("--- 请求体 ---")
    body = req.body
    if body is None:
        print("(无请求体)")
    elif isinstance(body, bytes):
        print(body.decode("utf-8", errors="replace"))
    else:
        print(body)
    print("=" * 60)


def call_api(method: str = "POST", params: dict | None = None, json_body: dict | None = None) -> dict:
    """调用指定接口，返回解析后的 JSON 数据。

    Args:
        method: HTTP 方法，如 GET / POST / PUT / DELETE
        params: URL 查询参数
        json_body: 请求体（JSON）
    """
    try:
        resp = requests.request(
            method,
            API_URL,
            headers=HEADERS,
            params=params,
            json=json_body,
            timeout=10,
        )
        resp.raise_for_status()  # 4xx/5xx 抛出异常
        return resp.json()
    except requests.exceptions.Timeout:
        raise RuntimeError("请求超时，请检查网络或接口地址")
    except requests.exceptions.HTTPError as e:
        raise RuntimeError(f"接口返回错误: {e.response.status_code} - {e.response.text[:200]}")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"请求失败: {e}")

def purch_in_not_finished(method: str = "POST", params: dict | None = None, json_body: dict | None = None,fuc_name: str ="") -> dict:
    
    try:
        resp = requests.request(
            method,
            API_URL + fuc_name,
            headers=HEADERS,
            params=params,
            json=json_body,
            timeout=10,
        )
        _print_request(resp)
        resp.raise_for_status()  # 4xx/5xx 抛出异常
        return resp.json()
    except requests.exceptions.Timeout:
        raise RuntimeError("请求超时，请检查网络或接口地址")
    except requests.exceptions.HTTPError as e:
        raise RuntimeError(f"接口返回错误: {e.response.status_code} - {e.response.text[:200]}")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"请求失败: {e}")


def main():
    if not API_KEY:
        print("提示: 未设置 API_KEY 环境变量，请在 PowerShell 中执行:")
        print('  $env:API_URL = "https://实际接口地址"')
        print('  $env:API_KEY = "你的密钥"')
        return


    result = purch_in_not_finished(
        method="POST",
        fuc_name="/purch/in/page",  # API_URL 已含 /api，此处无需重复
        json_body={
            "current": 1,
            "size": 200,
            "sort": "id",
            "order": "descending",
            "model": {"status": [30, 2]},
        },
    )
    print("接口响应:", result)


if __name__ == "__main__":
    main()
