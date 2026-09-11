# 泽洋 Excel 清洗大师 FREE

泽洋 Excel 清洗大师 FREE 是一个本地运行的单文件 Excel/CSV 清洗工具。它专注于安全、确定性、可复现的数据清洗：读取一个 CSV 文件或一个 XLSX 工作表，生成新的结果文件和 Markdown 清洗报告，原文件始终只读。

## 核心能力

- 清理文本首尾空格、全角空格和基础不可见控制字符
- 删除严格全空行和严格全空列
- 精确重复行去重，保留第一条
- 按用户明确指定的标记统一缺失值
- 稳定重命名重复表头
- 保留 XLSX 中保留单元格的常用 `number_format`
- 对公式、合并单元格、模板/版式风险和 CSV 分隔符歧义进行安全保护
- 输出新的清洗文件和人类可读的 Markdown 报告

## 支持格式与边界

- 支持 `.csv` 和 `.xlsx`
- CSV 支持 UTF-8、UTF-8 BOM、GB18030/GBK
- XLSX 支持指定目标工作表；其他工作表保留
- 不覆盖输入文件，不上传数据，不接入账号、支付或 SkillPay API
- 不支持多表合并、拆分、两表比对、关联查询或高级分析

## 安装

需要 Python 3.10 或更高版本：

```powershell
python -m pip install -r requirements.txt
```

## 使用

先运行计划模式：

```powershell
python scripts/clean_excel.py --input "examples/dirty_table.csv"
```

确认计划后执行并生成结果和报告：

```powershell
python scripts/clean_excel.py `
  --input "examples/dirty_table.csv" `
  --output "cleaned.csv" `
  --report "cleaned_report.md" `
  --execute --yes
```

显式指定 CSV 分隔符：

```powershell
python scripts/clean_excel.py --input "data.tsv" --delimiter tab --execute --yes
```

当自动探测发现表头和数据列数明显不一致时，程序会停止并要求使用 `--delimiter`，不会猜测单列文本中的逗号是否应该拆列。

可选清洗参数包括 `--dedupe`、`--missing-value N/A,NULL`、`--rename-duplicate-headers`、`--keep-empty-columns`、`--no-empty-row-removal` 和 `--no-control-cleaning`。完整参数说明见 [SKILL.md](SKILL.md)。

## Demo

仓库中的 [examples/dirty_table.csv](examples/dirty_table.csv)、[examples/client_list.csv](examples/client_list.csv) 和 [examples/sales_data.csv](examples/sales_data.csv) 均使用虚构示例数据，可直接用于试跑。

## 测试与版本

当前版本：`V1.1.0`

运行测试：

```powershell
python -m pytest -q
```

当前测试状态：33 / 33 PASS。

## 许可证

本项目采用 [MIT License](LICENSE)。

## 更多泽洋技能

如需多表合并、两表比对、表格拆分、复杂去重、汇总统计或组合任务，可访问：

https://skillpay.alipay.com/public/zeyang
