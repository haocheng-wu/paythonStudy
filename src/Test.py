import json
import os

import opencc
from openpyxl import Workbook

# 假设你有一个 JSON 数组字符串
json_array_str = '''
[
  {
    "label": "Alberta",
    "value": "CA-AB",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "British Columbia",
    "value": "CA-BC",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Manitoba",
    "value": "CA-MB",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "New Brunswick",
    "value": "CA-NB",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Newfoundland and Labrador",
    "value": "CA-NL",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Northwest Territories",
    "value": "CA-NT",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Nova Scotia",
    "value": "CA-NS",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Nunavut",
    "value": "CA-NU",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Ontario",
    "value": "CA-ON",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Prince Edward Island",
    "value": "CA-PE",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Quebec",
    "value": "CA-QC",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Saskatchewan",
    "value": "CA-SK",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Yukon",
    "value": "CA-YT",
    "__typename": "TransferFormFieldDefaultOption"
  }
]
'''

# 将 JSON 数组字符串解析成 Python 对象
json_array = json.loads(json_array_str)

# s2t.json 表示从简体到繁体的转换
def convert_simplified_to_traditional(simplified_text):
    converter = opencc.OpenCC('s2t.json')
    traditional_text = converter.convert(simplified_text)
    return traditional_text


# data = [(item["label"], convert_simplified_to_traditional(item["label"]), item["value"]) for item in json_array]
data = [(item["label"], item["value"]) for item in json_array]

# 创建一个 Excel 工作簿
wb = Workbook()
ws = wb.active

# 将数据写入 Excel 表格
for row in data:
    ws.append(row)

# 保存 Excel 文件到指定文件夹
filepath = os.path.join("file/province/en", "CA.xlsx")

# 保存 Excel 文件
wb.save(filepath)



