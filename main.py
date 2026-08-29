import os

import requests

# 从环境变量读取配置，避免硬编码敏感信息
API_URL = os.getenv("API_URL", "https://api.example.com/v1/resource")
API_KEY = os.getenv("API_KEY", "")

# 常见的两种 Token 请求头格式，按接口文档选择其一
HEADERS = {
    "Authorization": f"Bearer {API_KEY}",  # 方式1: Bearer Token
    # "X-API-Key": API_KEY,                # 方式2: 自定义 Key 头
    "Content-Type": "application/json",
}


def call_api(method: str = "GET", params: dict | None = None, json_body: dict | None = None) -> dict:
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


def main():
    if not API_KEY:
        print("提示: 未设置 API_KEY 环境变量，请在 PowerShell 中执行:")
        print('  $env:API_URL = "https://实际接口地址"')
        print('  $env:API_KEY = "你的密钥"')
        return

    result = call_api(method="GET")
    print("接口响应:", result)


if __name__ == "__main__":
    main()
