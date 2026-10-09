import os
import re
import sys
import calendar
from datetime import datetime, timedelta
import time
from urllib.parse import urlparse

import requests
from openpyxl import Workbook, load_workbook

from api_accounts import API_ACCOUNTS, PERIOD, EXPORT_DIR
from api_sign import _build_headers


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


def api_call(
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
    result = api_call(
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
    result = api_call(
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
    result = api_call(
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
    result = api_call(
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
    result = api_call(
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
        result = api_call(
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
    result = api_call(
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
    result = api_call(
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
    result = api_call(

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
                    "startDate": f"{PERIOD}-01",
                    "endDate": f"{PERIOD}-31",
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


def _safe_sheet_name(name: str, used: set) -> str:
    """生成合法且唯一的 Excel sheet 名（≤31字符，不含 []:?*\\/）。"""
    name = re.sub(r'[\\/?*\[\]:]', "_", str(name)).strip() or "Sheet"
    base = name[:31]
    name, i = base, 2
    while name in used:
        suffix = f"_{i}"
        name = base[: 31 - len(suffix)] + suffix
        i += 1
    return name


def _to_amount(value):
    """接口余额为字符串，转为数字；空值返回 None。"""
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return value


def _query_subject_balance(base_url: str, key: str, bulk_id: str, subject_id: str):
    """查询某科目的余额表明细，返回 (本期合计余额, 本年累计余额)。"""
    result = api_call(
        base_url=base_url,
        key=key,
        method="POST",
        fuc_name="/api/account/book/accountDetail",
        json_body={
            "bulkId": bulk_id,
            "beginPeriod": PERIOD,
            "endPeriod": PERIOD,
            "beginSubjectId": None,
            "endSubjectId": None,
            "beginSubjectLevel": 1,
            "endSubjectLevel": 4,
            "assistShow": 0,
            "balanceZeroHide": 0,
            "balanceZeroAndPeriodZeroHide": 0,
            "periodZeroHidePeriodAndYear": 0,
            "leafSubjectShow": 0,
            "otherSideSubjectShow": 0,
            "subjectId": subject_id,
            "assistSubjectId": None,
        },
    )
    period_total, year_total = None, None
    for r in _extract_records(result):
        if r.get("summary") == "本期合计":
            period_total = r.get("balanceAmount")
        elif r.get("summary") == "本年累计":
            year_total = r.get("balanceAmount")
    return period_total, year_total


def _period_date_range() -> tuple[str, str]:
    """返回 PERIOD 月份的第一天和最后一天，格式 yyyy-mm-dd。"""
    year, month = map(int, PERIOD.split("-"))
    last_day = calendar.monthrange(year, month)[1]
    return f"{PERIOD}-01", f"{PERIOD}-{last_day:02d}"


def export_business_data() -> None:
    """通过业务接口获取数据，写入 {PERIOD}_对账明细.xlsx 各 sheet 的科目数据下方。
    布局：科目数据行 -> 空行 -> 业务数据表头行 -> 业务数据行 -> 差额行(公式=科目-业务)。
    """
    start_date, end_date = _period_date_range()
    fpath = os.path.join(EXPORT_DIR, f"{PERIOD}_对账明细.xlsx")

    if not os.path.exists(fpath):
        print(f"文件不存在: {fpath}，请先运行 export_voucher_detail")
        return

    wb = load_workbook(fpath)
    biz_headers = [
        "库存商品(账面)",       # bookPageSummary.costSubtotalAmtReal
        "主营业务收入(收入)",    # incomePageTotal.incomeAmount
        "主营业务成本(成本)",    # incomePageTotal.costAmount
        "销售费用(费用)",       # feeRecordStat.realFeeAmount
        "到货运费(费用)",       # feeRecordStat(realFeeAmount) with freight id
    ]

    for account in API_ACCOUNTS:
        name = account.get("name") or account["url"]
        base_url, key = account["url"], account["key"]

        # 找到对应 sheet（与 export_voucher_detail 中命名逻辑一致）
        sheet_name = _safe_sheet_name(name, set())
        if sheet_name not in wb.sheetnames:
            print(f"[{name}] sheet '{sheet_name}' 不存在，跳过")
            continue
        ws = wb[sheet_name]

        # ---- 获取账套列表 ----
        try:
            bulks = _extract_records(
                api_call(base_url, key, "POST", fuc_name="/api/organize/bulk/account/list", json_body={})
            )
        except RuntimeError as e:
            print(f"[{name}] 账套列表查询失败: {e}")
            continue

        # ---- Step 1: incomePageTotal（收入/成本，账号级汇总）----
        income_amount = cost_amount = None
        try:
            r = api_call(
                base_url, key, "POST",
                fuc_name="/api/report/finance/incomePageTotal",
                json_body={"createDateRange": {"startDate": start_date, "endDate": end_date}},
            )
            data = r.get("data") or {}
            income_amount = _to_amount(data.get("incomeAmount"))
            cost_amount = _to_amount(data.get("costAmount"))
        except RuntimeError as e:
            print(f"[{name}] 收入查询失败: {e}")

        # ---- Step 2: bookPageSummary（库存账面，账号级汇总）----
        stock_amount = None
        try:
            r = api_call(
                base_url, key, "POST",
                fuc_name="/api/stock/query/bookPageSummary",
                json_body={},
            )
            data = r.get("data") or {}
            stock_amount = _to_amount(data.get("costSubtotalAmtReal"))
        except RuntimeError as e:
            print(f"[{name}] 库存查询失败: {e}")

        # ---- Step 3: feeRecordStat（全部费用，账号级汇总）----
        fee_amount = None
        try:
            r = api_call(
                base_url, key, "POST",
                fuc_name="/api/finance/query/feeRecordStat",
                json_body={
                    "current": 1, "size": 200, "sort": "id", "order": "descending",
                    "model": {
                        "viewType": "2", "settleOrganId": None, "feeId": None,
                        "feeIdList": None, "excludeFeeIdList": None,
                        "feeType": None, "financeType": 1,
                        "feeBillLinkUserId": None, "feeLinkBizUserId": None,
                        "receiptSettleCompanyId": None, "receiptUserId": None,
                        "splitFlag": None,
                        "happenDateRange": {"startDate": start_date, "endDate": end_date},
                    },
                },
            )
            data = r.get("data") or {}
            fee_amount = _to_amount(data.get("realFeeAmount"))
        except RuntimeError as e:
            print(f"[{name}] 费用查询失败: {e}")

        # ---- Step 4: listUsed 取运费 feeId ----
        freight_id = None
        try:
            r = api_call(base_url, key, "POST", fuc_name="/api/cfg/fee/listUsed", json_body={})
            for rec in _extract_records(r):
                if rec.get("feeName") == "采购到货运费":
                    freight_id = rec.get("id")
                    break
            if not freight_id:
                print(f"[{name}] 未找到 feeName='运费' 的科目")
        except RuntimeError as e:
            print(f"[{name}] 运费科目查询失败: {e}")

        # ---- Step 5: feeRecordStat 按运费 feeId 过滤（账号级汇总）----
        freight_amount = None
        if freight_id:
            try:
                r = api_call(
                    base_url, key, "POST",
                    fuc_name="/api/finance/query/feeRecordStat",
                    json_body={
                        "current": 1, "size": 200, "sort": "id", "order": "descending",
                        "model": {
                            "viewType": "2", "settleOrganId": None, "feeId": None,
                            "feeIdList": [freight_id], "excludeFeeIdList": None,
                            "feeType": None, "financeType": 1,
                            "feeBillLinkUserId": None, "feeLinkBizUserId": None,
                            "receiptSettleCompanyId": None, "receiptUserId": None,
                            "splitFlag": None,
                            "happenDateRange": {"startDate": start_date, "endDate": end_date},
                        },
                    },
                )
                data = r.get("data") or {}
                freight_amount = _to_amount(data.get("realFeeAmount"))
            except RuntimeError as e:
                print(f"[{name}] 运费查询失败: {e}")

        # ---- 业务数据写入科目数据下方 ----
        data_start = 2                          # 科目数据从第 2 行开始
        data_end = data_start + len(bulks) - 1  # 科目数据结束行
        biz_row_header = data_end + 1           # 业务表头紧接科目数据行
        biz_row_data = biz_row_header + 1       # 业务数据行
        biz_row_diff = biz_row_data + 1         # 差额行

        # 业务数据表头（A列空，B-F列）
        for i, h in enumerate(biz_headers):
            ws.cell(row=biz_row_header, column=2 + i, value=h)

        # 业务数据（A列空，B-F列）
        ws.cell(row=biz_row_data, column=2, value=stock_amount)
        ws.cell(row=biz_row_data, column=3, value=income_amount)
        ws.cell(row=biz_row_data, column=4, value=cost_amount)
        ws.cell(row=biz_row_data, column=5, value=fee_amount)
        ws.cell(row=biz_row_data, column=6, value=freight_amount)

        # 差额行：A列=差额，B-F列=公式 =科目数据第1行 - 业务数据行
        ws.cell(row=biz_row_diff, column=1, value="差额")
        for col_idx in range(2, 7):
            col_letter = chr(64 + col_idx)  # B->'B', C->'C'...
            formula = f"={col_letter}{data_start}-{col_letter}{biz_row_data}"
            ws.cell(row=biz_row_diff, column=col_idx, value=formula)

        print(f"[{name}] 业务数据已写入 R{biz_row_header}-R{biz_row_diff}")

    # ---- 保存 ----
    try:
        wb.save(fpath)
        print(f"业务数据已补充写入: {fpath}")
    except PermissionError:
        ts = time.strftime("%H%M%S")
        alt_path = os.path.join(EXPORT_DIR, f"{PERIOD}_对账明细_{ts}.xlsx")
        try:
            wb.save(alt_path)
            print(f"原文件被占用，已另存为: {alt_path}")
        except PermissionError:
            print(f"保存失败，请关闭 Excel 后重试: {fpath}")


def export_voucher_detail() -> None:
    """生成 {PERIOD}_对账明细.xlsx：
    - 以 API_ACCOUNTS 每个节点的 name 创建 sheet
    - 每个账套一行，写入指定科目的余额
    """
    # (列标题, 科目名称, 是否末级科目, 是否取本年累计)
    # 销售费用取非末级(childCount>0)，其余取末级(childCount=0)
    columns = [
        ("库存商品(本年累计)", "库存商品", True, True),
        ("主营业务收入(本期合计)", "主营业务收入", True, False),
        ("主营业务成本(本期合计)", "主营业务成本", True, False),
        ("销售费用(本期合计)", "销售费用", False, False),
        ("到货运费(本期合计)", "到货运费", True, False),
    ]

    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()

    for account in API_ACCOUNTS:
        name = account.get("name") or account["url"]
        sheet_name = _safe_sheet_name(name, used_names)
        used_names.add(sheet_name)
        ws = wb.create_sheet(sheet_name)
        ws.append(["账套名称"] + [c[0] for c in columns])

        try:
            bulks = _extract_records(
                api_call(
                    base_url=account["url"],
                    key=account["key"],
                    method="POST",
                    fuc_name="/api/organize/bulk/account/list",
                    json_body={},
                )
            )
        except RuntimeError as e:
            print(f"[{name}] 账套列表查询失败: {e}")
            bulks = []

        print(f"\n===== {name}: {len(bulks)} 个账套 =====")
        for bulk in bulks:
            bulk_id = bulk.get("id", "")
            cal_organ_name = bulk.get("calOrganName", "")
            if not bulk_id:
                continue
            try:
                subjects = _extract_records(
                    api_call(
                        base_url=account["url"],
                        key=account["key"],
                        method="POST",
                        fuc_name="/api/account/subject/list",
                        json_body={"bulkId": bulk_id},
                    )
                )
            except RuntimeError as e:
                print(f"  {cal_organ_name} 科目查询失败: {e}")
                subjects = []

            row = [cal_organ_name]
            for _title, sub_name, is_leaf, use_year in columns:
                # childCount==0 即末级科目，>0 即非末级
                matched = [
                    s
                    for s in subjects
                    if s.get("subjectName") == sub_name
                    and (int(s.get("childCount") or 0) == 0) == is_leaf
                ]
                if not matched:
                    row.append(None)
                    continue
                if len(matched) > 1:
                    print(f"  {cal_organ_name} {sub_name} 匹配到 {len(matched)} 个，取第一个")
                try:
                    period_total, year_total = _query_subject_balance(
                        account["url"], account["key"], bulk_id, matched[0]["id"]
                    )
                    row.append(_to_amount(year_total if use_year else period_total))
                except RuntimeError as e:
                    print(f"  {cal_organ_name} {sub_name} 余额查询失败: {e}")
                    row.append(None)
            ws.append(row)
            print(f"  {cal_organ_name}: {row[1:]}")

    os.makedirs(EXPORT_DIR, exist_ok=True)
    fpath = os.path.join(EXPORT_DIR, f"{PERIOD}_对账明细.xlsx")
    try:
        wb.save(fpath)
    except PermissionError:
        # 文件被 Excel 占用，改用带时间戳的备选文件名，避免本次结果丢失
        ts = time.strftime("%H%M%S")
        fpath = os.path.join(EXPORT_DIR, f"{PERIOD}_对账明细_{ts}.xlsx")
        try:
            wb.save(fpath)
            print(f"原文件被占用，已另存为: {fpath}")
            print("提示：请关闭已打开的 2026-08_对账明细.xlsx 后重新运行，即可写入默认文件名")
        except PermissionError:
            print(f"保存失败，请关闭 Excel 后重试: {fpath}")
            return
    print(f"\n对账明细已生成: {fpath}")


def run_download_tasks(base_url: str, key: str, result_dir: str, account_name: str = "") -> None:
    # 下载系统备份报表
    func_name9 = "/api/common/export/page"

    result = api_call(
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
    safe_name = re.sub(r'[\\/:*?"<>|]', "_", account_name) or "default"
    download_dir = os.path.join(
        EXPORT_DIR,
        safe_name,  # API_ACCOUNTS 中配置的账号名称
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

    result_dir = EXPORT_DIR
    os.makedirs(result_dir, exist_ok=True)

    for account in API_ACCOUNTS:
        name = account.get("name") or account["url"]
        # 以 API_ACCOUNTS 中配置的账号名称为该账号建立独立输出子目录
        account_dir = os.path.join(result_dir, name)
        os.makedirs(account_dir, exist_ok=True)
        print(f"\n{'=' * 20} 开始处理: {name} ({account['url']}) {'=' * 20}")
        print(f"输出目录: {account_dir}")
        try:
            # run_account_tasks(account["url"], account["key"], account_dir)
            # run_download_tasks(account["url"], account["key"], account_dir, name)
            pass
        except RuntimeError as e:
            # 单个账号失败不影响其他账号继续执行
            print(f"[{name}] 执行失败，已跳过: {e}")

    # 汇总所有账号，生成 PERIOD_对账明细.xlsx（每个账号一个 sheet）
    export_voucher_detail()
    # 通过业务接口补充数据到同一 Excel 文件
    export_business_data()


if __name__ == "__main__":
    # if not is_last_day_of_month():
    #     sys.exit(0)  # 不是月末，直接退出
    # # ===== 下面放你真正要执行的代码 =====
    # print("月末24点，开始执行...")
    main()
