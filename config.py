"""
工资单助手 - 配置文件
Wage Form Helper - Configuration

STUB: Update these settings once you have the actual Excel files.
请获取到实际Excel文件后，根据文件格式修改以下配置。

How to find the right values / 如何找到正确的配置值:
  1. Open the Excel file in Excel / 用Excel打开文件
  2. Look at the sheet tab name at the bottom → set "sheet_name"
     查看底部工作表标签名 → 设置 sheet_name
  3. Find which row has the column headers → set "header_row"
     找到列标题所在的行号 → 设置 header_row
  4. Copy the exact text of each column header → set col_* fields
     复制每列的确切标题文字 → 设置各 col_* 字段
"""

# ═══════════════════════════════════════════════════════════════
# WAGE FORM (INPUT)  工资表配置（输入文件）
# ═══════════════════════════════════════════════════════════════
WAGE_FORM = {
    # Sheet tab name shown at the bottom of the Excel window
    # Excel窗口底部显示的工作表标签名
    "sheet_name": "Sheet0",       # STUB ← 改为实际工作表名, e.g. "工资表"

    # Row number that contains the column headers (1 = very first row)
    # 列标题所在的行号 (1 = 第一行)
    "header_row": 1,              # STUB ← 通常是 1 或 2

    # Exact header text for each field (copy from Excel, case-sensitive)
    # 各字段的精确列标题（从Excel复制，区分大小写）
    "col_name":         "姓名",        # STUB ← 员工姓名列标题
    "col_id":           "身份证号",    # STUB ← 身份证号码列标题
    "col_bank_account": "银行账号",    # STUB ← 银行账号列标题
    "col_bank_name":    "开户行",      # STUB ← 开户行名称列标题
    "col_amount":       "实发工资",    # STUB ← 实际发放金额列标题
}

# ═══════════════════════════════════════════════════════════════
# BANK FORM (OUTPUT)  银行表配置（输出文件）
# ═══════════════════════════════════════════════════════════════
BANK_FORM = {
    # Sheet tab name in the bank template file
    # 银行模板文件的工作表标签名
    "sheet_name": "Sheet0",       # STUB ← 改为实际工作表名

    # First row where employee data should be written
    # (rows above this are the bank's own title/header rows)
    # 开始写入员工数据的行号（此行以上是银行表格自己的标题行）
    "data_start_row": 2,          # STUB ← 通常是 3 或 4，查看模板确认

    # Column for each field — use letter ("B") or number (2)
    # 各字段所在列 — 用字母（如 "B"）或数字（如 2）均可
    "col_name":         "B",      # STUB ← 姓名所在列
    "col_id":           "C",      # STUB ← 身份证号所在列
    "col_bank_account": "D",      # STUB ← 银行账号所在列
    "col_bank_name":    "E",      # STUB ← 开户行所在列
    "col_amount":       "F",      # STUB ← 金额所在列
}
