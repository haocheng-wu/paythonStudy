import json
import pandas as pd

# 假设你有一个 JSON 数组字符串
json_array_str = '''
[
  {
    "label": "Almaty",
    "value": "KZ-ALA",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Almaty oblysy",
    "value": "KZ-ALM",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Aqmola oblysy",
    "value": "KZ-AKM",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Aqtöbe oblysy",
    "value": "KZ-AKT",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Atyraū oblysy",
    "value": "KZ-ATY",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Batys Qazaqstan oblysy",
    "value": "KZ-ZAP",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Mangghystaū oblysy",
    "value": "KZ-MAN",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Nur-Sultan",
    "value": "KZ-AST",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Pavlodar oblysy",
    "value": "KZ-PAV",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Qaraghandy oblysy",
    "value": "KZ-KAR",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Qostanay oblysy",
    "value": "KZ-KUS",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Qyzylorda oblysy",
    "value": "KZ-KZY",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Shyghys Qazaqstan oblysy",
    "value": "KZ-VOS",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Shymkent",
    "value": "KZ-SHY",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Soltüstik Qazaqstan oblysy",
    "value": "KZ-SEV",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Türkistan oblysy",
    "value": "KZ-YUZ",
    "__typename": "TransferFormFieldDefaultOption"
  },
  {
    "label": "Zhambyl oblysy",
    "value": "KZ-ZHA",
    "__typename": "TransferFormFieldDefaultOption"
  }
]
'''

# 将 JSON 数组字符串解析成 Python 对象
json_array = json.loads(json_array_str)

# 创建反向字典
reverse_data = {item['value']: item['label'] for item in json_array}

# 根据 value 获取 label 的函数
def get_label_by_value(target_value):
    return reverse_data.get(target_value, None)


# 读取 Excel 文件
excel_file_path = 'file/province/cn/KZ.xlsx'
df = pd.read_excel(excel_file_path)

# 确保 '英文' 列的类型是 object 类型（字符串）
df['英文'] = df['英文'].astype(object)

# 假设 Excel 文件的列标题分别为 '中文简体'、'中文繁体'、'英文'、'编码'
# 遍历每一行，根据编码匹配并填充英文列
for index, row in df.iterrows():
    code = row['编码']
    if pd.isna(row['英文']) and code in reverse_data:
        label = get_label_by_value(code)
        if label:
            df.at[index, '英文'] = label


# 保存更新后的 Excel 文件
df.to_excel('file/province/KZ.xlsx', index=False)



