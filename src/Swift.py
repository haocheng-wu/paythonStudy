import os

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

def scrape_swift_codes(base_url, max_pages=16):
    # 存储所有 Swift Code 和 Bank Name
    swift_codes_and_banks = []

    # 循环遍历每一页
    for page_num in range(1, max_pages + 1):
        # 构建当前页的完整 URL
        if page_num == 1:
            url = f"{base_url}"
        else:
            url = f"{base_url}page/{page_num}/"

        # 发送 GET 请求获取页面内容
        response = requests.get(url)

        if response.status_code == 200:
            # 使用 BeautifulSoup 解析页面内容
            soup = BeautifulSoup(response.text, 'html.parser')

            # 找到包含 Swift Code 的表格行
            table_rows = soup.find_all('tr')

            # 找到表格中的所有 Swift Code 和 Bank Name
            for row in table_rows:
                # 找到表格行中的所有单元格
                cells = row.find_all('td')
                if len(cells) >= 2:
                    # 第一个单元格是 Swift Code
                    swift_code = cells[4].text.strip()
                    # 第二个单元格是 Bank Name
                    bank_name = cells[1].text.strip()
                    swift_codes_and_banks.append((swift_code, bank_name))
        else:
            print(f"Failed to fetch URL: {url}. Status code: {response.status_code}")

    return swift_codes_and_banks

def export_to_excel(data, folder, filename):
    # 创建一个 Excel 工作簿
    wb = Workbook()
    ws = wb.active

    # 在表格中写入数据
    for row in data:
        ws.append(row)

        # 确定文件夹存在
    if not os.path.exists(folder):
        os.makedirs(folder)

    # 保存 Excel 文件到指定文件夹
    filepath = os.path.join(folder, filename)
    wb.save(filepath)

    # # 保存 Excel 文件
    # wb.save(filename)

if __name__ == "__main__":
    base_url = 'https://bank-codes-hk.com/swift-code/japan/'
    swift_codes_and_banks = scrape_swift_codes(base_url)
    if swift_codes_and_banks:
        # 设置 Excel 文件名
        excel_filename = 'JP_SWIFT_CODE.xlsx'
        # 设置保存的文件夹路径
        save_folder = 'file'
        # 导出数据到 Excel
        export_to_excel(swift_codes_and_banks, save_folder, excel_filename)

        print(f"Data exported to {excel_filename}")
    else:
        print("No Swift Codes found.")
