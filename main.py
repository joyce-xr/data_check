import base64
import hashlib
import json
import os
import random
import re
import sys
from datetime import datetime, timedelta
import time
from urllib.parse import urlparse

import requests
from openpyxl import Workbook

from api_accounts import API_ACCOUNTS


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


def search_not_finished(
    base_url: str,
    key: str,
    method: str = "POST",
    params: dict | None = None,
    json_body: dict | None = None,
    fuc_name: str = "",
) -> dict:
    """调用接口并返回 JSON 响应。base_url 与 key 来自 api_accounts.py 中的某个账号。"""

    try:
        url = base_url + fuc_name
        resp = requests.request(
            method,
            url,
            headers=_build_headers(url, key),
            params=params,
            json=json_body,
            timeout=10,
        )
        # _print_request(resp)  # 打印请求详情
        resp.raise_for_status()  # 4xx/5xx 抛出异常
        return resp.json()
    except requests.exceptions.Timeout:
        raise RuntimeError("请求超时，请检查网络或接口地址")
    except requests.exceptions.HTTPError as e:
        raise RuntimeError(
            f"接口返回错误: {e.response.status_code} - {e.response.text[:200]}"
        )
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"请求失败: {e}")


def _extract_records(result: dict) -> list:
    """从响应中取记录列表，兼容两种结构：data 为 {records: [...]} 或 data 直接为列表。"""
    data = result.get("data") or []
    if isinstance(data, dict):
        return data.get("records") or []
    return data


def export_result_to_excel(
    result: dict,
    fields: list[str],
    fuc_name: str = "",
    group_field: str = "",
    save_dir: str = "",
) -> list[str]:
    """将接口响应结果导出为 Excel 文件的通用方法。

    从 result 中提取记录列表（兼容 data 为 {records: [...]} 或 data 直接为列表
    两种响应结构），按 fields 指定的字段写入 .xlsx 文件：首行为表头，其后每行
    一条记录。可通过 group_field 按字段值拆分成多个文件。

    Args:
        result (dict): 接口响应的 JSON 字典。
        fields (list[str]): 必传。导出的列，依次作为表头与取值字段。字段名
            必须与记录中的键完全一致（区分大小写），记录中缺失的字段写入空字符串。
        fuc_name (str, optional): 接口路径，用作文件名前缀，其中的 "/" 会被
            替换为 "_"；为空时前缀为 "export"。
        group_field (str, optional): 分组字段。按该字段的值将记录拆分到多个
            文件，字段值（清理 Windows 非法字符后）作为文件名的一部分；
            传空字符串则不分组，所有记录写入同一个文件。
        save_dir (str, optional): 保存目录，不存在时自动创建，缺省为当前目录。

    Returns:
        list[str]: 生成的 Excel 文件路径列表。

    文件命名规则:
        - 分组时: {接口名}_{分组值}_{当前月份}.xlsx，
          如 api_purch_in_page_销售部WLD_202608.xlsx
        - 不分组时: {接口名}_{当前月份}.xlsx

    示例:
        export_result_to_excel(
            result,
            fields=["purchOrganName", "orderDate", "orderNo"],
            fuc_name="/api/purch/in/page",
            group_field="purchOrganName",
            save_dir=r"D:\\account_check_data",
        )
    """
    records = _extract_records(result)

    os.makedirs(save_dir, exist_ok=True)

    func_tag = fuc_name.strip("/").replace("/", "_") or "export"
    month = time.strftime("%Y%m")

    groups: dict[str, list[dict]] = {}
    if group_field:
        for rec in records:
            groups.setdefault(str(rec.get(group_field, "")), []).append(rec)
    else:
        groups[""] = records

    paths = []
    for org, items in groups.items():
        # 清理文件名非法字符及各类括号（沙箱与 Windows 兼容）
        safe_org = re.sub(r"_+", "_", re.sub(r'[\\/:*?"<>|()（）]', "_", org)).strip(
            "_"
        )
        name = (
            f"{func_tag}_{safe_org}_{month}.xlsx"
            if group_field
            else f"{func_tag}_{month}.xlsx"
        )
        path = os.path.join(save_dir, name)
        wb = Workbook()
        ws = wb.active
        ws.append(fields)
        for it in items:
            ws.append([it.get(f, "") for f in fields])
        wb.save(path)
        paths.append(path)
        print(f"已导出 {len(items)} 条 -> {path}")
    return paths


def run_account_tasks(base_url: str, key: str, result_dir: str) -> None:
    """对一个账号（一套 API_URL + API_KEY）执行全部查询与导出。"""

    fuc_name = "/api/purch/in/page"  # 采购入库单查询
    result = search_not_finished(
        base_url=base_url,
        key=key,
        method="POST",
        fuc_name=fuc_name,
        json_body={  # 只看待财务确认
            "current": 1,
            "size": 200,
            "sort": "id",
            "order": "descending",
            "model": {"status": [2]},
        },
    )
    # print("接口响应:", result)
    records = _extract_records(result)
    if records:
        export_result_to_excel(
            result,
            fields=["purchOrganName", "orderDate", "orderNo"],
            fuc_name=fuc_name,
            group_field="purchOrganName",
            save_dir=result_dir,
        )

    func_name2 = "/api/purch/back/page"  # 采购退货单查询
    result = search_not_finished(
        base_url=base_url,
        key=key,
        method="POST",
        fuc_name=func_name2,
        json_body={  # 只看已完结-未确认
            "current": 1,
            "size": 200,
            "sort": "id",
            "order": "descending",
            "model": {"status": [100], "finPushStatus": 0},
        },
    )
    # print("接口响应:", result)
    records = _extract_records(result)
    if records:
        export_result_to_excel(
            result,
            fields=["purchOrganName", "orderDate", "orderNo"],
            fuc_name=func_name2,
            group_field="purchOrganName",
            save_dir=result_dir,
        )

    func_name3 = "/api/sell/out/page"  # 销售出库单查询
    result = search_not_finished(
        base_url=base_url,
        key=key,
        method="POST",
        fuc_name=func_name3,
        json_body={  # 只看待财务确认
            "current": 1,
            "size": 200,
            "sort": "id",
            "order": "descending",
            "model": {"status": [100], "finPushStatus": 0},
        },
    )
    # print("接口响应:", result)
    records = _extract_records(result)
    if records:
        export_result_to_excel(
            result,
            fields=["sellOrganName", "orderDate", "orderNo"],
            fuc_name=func_name3,
            group_field="sellOrganName",
            save_dir=result_dir,
        )

    func_name4 = "/api/sell/back/page"  # 销售退货单查询
    result = search_not_finished(
        base_url=base_url,
        key=key,
        method="POST",
        fuc_name=func_name4,
        json_body={  # 只看已完结-未确认
            "current": 1,
            "size": 200,
            "sort": "id",
            "order": "descending",
            "model": {"status": [2]},
        },
    )
    # print("接口响应:", result)
    records = _extract_records(result)
    if records:
        export_result_to_excel(
            result,
            fields=["sellOrganName", "orderDate", "orderNo"],
            fuc_name=func_name4,
            group_field="sellOrganName",
            save_dir=result_dir,
        )

    # 未下推的费用
    func_name5 = "/api/ship/bill/settleOrganList"  # 分组列表
    result = search_not_finished(
        base_url=base_url,
        key=key,
        method="POST",
        fuc_name=func_name5,
        json_body={
            "current": 1,
            "size": 200,
            "sort": "id",
            "order": "descending",
            "model": {
                "settleOrganId": "",
                "expUserId": "",
                "companyId": "",
                "createDateRang": None,
            },
            "settleOrganId": "",
            "expUserId": "",
            "companyId": "",
            "createDateRange": None,
        },
    )
    # print("接口响应:", result)
    # 暂不导出：该接口返回按结算机构汇总的数据（可用字段：settleOrganName / settleName /
    # totalAmount / totalItem），没有 orderNo / orderDate；确认需求后取消注释：
    # export_result_to_excel(result, fields=["settleOrganName", "settleName", "totalAmount", "totalItem"],
    #                        fuc_name=func_name5, group_field="settleOrganName", save_dir=result_dir)
    # 提取 companyId / expUserId / settleOrganId，供后续费用明细查询作入参
    fee_groups = [
        {
            "companyId": rec.get("companyId", ""),
            "expUserId": rec.get("expUserId", ""),
            "settleOrganId": rec.get("settleOrganId", ""),
        }
        for rec in _extract_records(result)
    ]
    print(f"未下推费用分组: {len(fee_groups)} 组")

    fee_fields = [
        "settleOrganName",
        "expUserId",
        "feeName",
        "companyId",
        "happenDate",
        "createTime",
        "linkBizOrderNo",
        "splitFlag",
    ]
    fee_details = []
    for g in fee_groups:
        result = search_not_finished(
            base_url=base_url,
            key=key,
            method="POST",
            fuc_name="/api/ship/bill/settleFeeList",
            json_body={
                # companyId / expUserId 为 0 表示不限，需转为 null 传给接口
                "companyId": None if g["companyId"] in ("0", 0) else g["companyId"],
                "expUserId": None if g["expUserId"] in ("0", 0) else g["expUserId"],
                "settleOrganId": g["settleOrganId"],
            },
        )
        fee_details.extend(_extract_records(result))

    # 累积所有分组的明细后统一导出，按结算机构分文件，避免同名文件被各分组覆盖
    if fee_details:
        export_result_to_excel(
            {"data": fee_details},
            fields=fee_fields,
            fuc_name="/api/ship/bill/settleFeeList",
            group_field="settleOrganName",
            save_dir=result_dir,
        )

    # 未完结的费用报销单
    func_name6 = "/api/finance/feeSubmit/page"
    result = search_not_finished(
        base_url=base_url,
        key=key,
        method="POST",
        fuc_name=func_name6,
        json_body={
            "current": 1,
            "size": 200,
            "sort": "id",
            "order": "descending",
            "model": {"statusList": ["000", "100"], "onlySelf": False},
        },
    )
    # print("接口响应:", result)
    records = _extract_records(result)
    if records:
        export_result_to_excel(
            result,
            fields=["settleOrganName", "orderDate", "orderNo"],
            fuc_name=func_name6,
            group_field="settleOrganName",
            save_dir=result_dir,
        )

    # 未完结的费用支出单
    func_name7 = "/api/finance/feePayout/page"
    result = search_not_finished(
        base_url=base_url,
        key=key,
        method="POST",
        fuc_name=func_name7,
        json_body={
            "current": 1,
            "size": 200,
            "sort": "id",
            "order": "descending",
            "model": {"statusList": ["000", "100"]},
        },
    )
    # print("接口响应:", result)
    records = _extract_records(result)
    if records:
        export_result_to_excel(
            result,
            fields=["settleOrganName", "orderDate", "orderNo"],
            fuc_name=func_name7,
            group_field="settleOrganName",
            save_dir=result_dir,
        )

    # 库存成本调整记录，是否存在没有改前成本价的记录
    func_name8 = "/api/stock/change/task/page"
    result = search_not_finished(
        base_url=base_url,
        key=key,
        method="POST",
        fuc_name=func_name8,
        json_body={
            "current": 1,
            "size": 10,
            "sort": "id",
            "order": "descending",
            "model": {
                "createDateRange": {
                    "startDate": "2026-08-01",
                    "endDate": "2026-08-31",
                    "type": 30,
                }
            },
        },
    )
    # print("接口响应:", result)
    records = _extract_records(result)
    # 只导出未返回 beforePriceAmtReal 节点的记录
    missing = [r for r in records if "beforePriceAmtReal" not in r]
    if missing:
        export_result_to_excel(
            {"data": missing},
            fields=["ownerOrganName", "createTime", "linkBizOrderNo", "goodCode"],
            fuc_name=func_name8,
            group_field="ownerOrganName",
            save_dir=result_dir,
        )

def run_download_tasks(base_url: str, key: str, result_dir: str) -> None:
    # 下载系统备份报表
    func_name9 = "/api/common/export/page"

    result = search_not_finished(
        base_url=base_url,
        key=key,
        method="POST",
        fuc_name=func_name9,
        json_body={
            "current": 1,
            "size": 10,
            "sort": "id",
            "order": "descending",
            "model": {},
        },
    )
    records = _extract_records(result)
    # 筛选：status=100 且 orderDate=今天（yyyy-mm-dd）
    today = time.strftime("%Y-%m-%d")
    download_list = [
        r
        for r in records
        if str(r.get("status")) == "100" and str(r.get("orderDate", "")) == today
    ]
    download_dir = os.path.join(
        r"D:\data_backup",
        urlparse(base_url).hostname.split(".")[0],  # 二级域名，如 test
    )
    os.makedirs(download_dir, exist_ok=True)
    saved, failed = 0, 0
    for rec in download_list:
        url = rec.get("ossPathUrl", "")
        if not url:
            continue
        fname = (
            os.path.basename(urlparse(url).path)
            or f"backup_{rec.get('orderNo', 'unknown')}.zip"
        )
        fname = re.sub(r'[\\/:*?"<>|]', "_", fname)  # 清理 Windows 非法字符
        fpath = os.path.join(download_dir, fname)
        try:
            dl_resp = requests.get(url, timeout=60)
            dl_resp.raise_for_status()
            with open(fpath, "wb") as f:
                f.write(dl_resp.content)
            print(f"已下载: {fpath}")
            saved += 1
        except requests.exceptions.RequestException as e:
            print(f"下载失败 {url}: {e}")
            failed += 1
    if download_list:
        print(
            f"备份下载汇总: 成功 {saved}, 失败 {failed}, 今日筛选 {today} 共 {len(download_list)} 条"
        )


class _Tee:
    """把写入的数据同时转发到多个流（终端 + 日志文件）。"""

    def __init__(self, *streams) -> None:
        self._streams = streams

    def write(self, data: str) -> int:
        for stream in self._streams:
            stream.write(data)
        return len(data)

    def flush(self) -> None:
        for stream in self._streams:
            stream.flush()


def _setup_logger() -> str:
    """让所有 print 输出及异常信息同时打印到终端并写入 logs 目录的 txt 文件。

    每次运行生成一个带时间戳的日志文件（逐行实时写入），返回日志文件路径。
    """
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, f"run_{time.strftime('%Y%m%d_%H%M%S')}.txt")
    log_file = open(log_path, "a", encoding="utf-8", buffering=1)
    sys.stdout = _Tee(sys.__stdout__, log_file)
    sys.stderr = _Tee(sys.__stderr__, log_file)
    return log_path


# def is_last_day_of_month():
#     today = datetime.now()
#     tomorrow = today + timedelta(days=1)
#     return tomorrow.day == 1  # 明天是1号，说明今天是月末


def main():
    log_path = _setup_logger()
    print(f"日志文件: {log_path}")

    result_dir = r"D:\account_check_data"
    os.makedirs(result_dir, exist_ok=True)

    for account in API_ACCOUNTS:
        name = account.get("name") or account["url"]
        # 以 API_URL 的二级域名（如 test.aipzm.com -> test）为该账号建立独立输出子目录
        host = urlparse(account["url"]).hostname or "default"
        account_dir = os.path.join(result_dir, host.split(".")[0])
        os.makedirs(account_dir, exist_ok=True)
        print(f"\n{'=' * 20} 开始处理: {name} ({account['url']}) {'=' * 20}")
        print(f"输出目录: {account_dir}")
        try:
            run_account_tasks(account["url"], account["key"], account_dir)
            run_download_tasks(account["url"], account["key"], account_dir)
        except RuntimeError as e:
            # 单个账号失败不影响其他账号继续执行
            print(f"[{name}] 执行失败，已跳过: {e}")


if __name__ == "__main__":
    # if not is_last_day_of_month():
    #     sys.exit(0)  # 不是月末，直接退出
    # # ===== 下面放你真正要执行的代码 =====
    # print("月末24点，开始执行...")
    main()
