"""自动刷新 .env 中 API_NONCE 的时间戳部分。

N 头结构：前 13 位为毫秒时间戳，其余部分保留原值。
用法：python refresh_nonce.py
"""
import time

from dotenv import dotenv_values, set_key

ENV_FILE = ".env"


def main():
    values = dotenv_values(ENV_FILE)
    old_nonce = values.get("API_NONCE", "")
    if not old_nonce or len(old_nonce) <= 13:
        raise SystemExit("API_NONCE 不存在或格式不符合预期（应多于 13 位）")

    new_nonce = str(int(time.time() * 1000)) + old_nonce[13:]
    set_key(ENV_FILE, "API_NONCE", new_nonce)
    print(f"API_NONCE 已更新:")
    print(f"  旧值: {old_nonce}")
    print(f"  新值: {new_nonce}")


if __name__ == "__main__":
    main()
