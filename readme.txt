api_account.py>
	1. 账期月份
	2. 各环境域名及账号

api_sign.py>
	接口请求签名工具：生成 N（nonce）和 S（sign）请求头

data_check_download.py>
	1. search_not_finished 调用接口的公共方法
	2. export_result_to_excel 查询到的未完结单据信息转存到excel文档中
	3. run_account_tasks 执行查询任务
		——采购入库、采购退货、销售出库、销售退货：待财务确认
		——费用待办池：未下推的费用明细
		——未完结的费用报销、费用支出单
		——库存成本调整记录：是否存在没有改前成本价的记录（如果存在，意味着后续影响单据的成本未调整）
	4. export_voucher_detail 导出“库存商品”在指定账期内的科目明细
	5. run_download_tasks 下载系统备份数据

inventory_analysis.py>
	1. 对比上月存货-本月出入库流水-本月存货差异



	按集团统计