import os
import time

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


def scrape_ach_codes(url):
    # 存储所有 ach Code 和 Bank Name
    ach_codes_and_banks = []

    urls = []

    # 发送 GET 请求获取页面内容
    response = requests.get(url)

    if response.status_code == 200:
        # 使用 BeautifulSoup 解析页面内容
        soup = BeautifulSoup(response.text, 'html.parser')

        # 找到包含 Swift Code 的表格行
        table_rows = soup.find(class_='post_content')

        # 在单元格中查找链接
        for child in table_rows.children:
            if child.name == 'a':
                href = child.get('href')
                name = child.text
                urls.append((href, name))
        # links = table_rows.find_all('a')
        # for link in links:
        #     if link:
        #         href = link.get('href')
        #         name = link.text.strip()
        #     urls.append((href, name))
    else:
        print(f"Failed to fetch URL: {url}. Status code: {response.status_code}")

    for href, name in urls:

        url = "https://bank-codes-hk.com" + href
        second_get(url, ach_codes_and_banks, name)
    return ach_codes_and_banks


def second_get(url, achCodes, name):
    # 发送 GET 请求获取页面内容
    response = requests.get(url)

    if response.status_code == 200:
        # 使用 BeautifulSoup 解析页面内容
        soup = BeautifulSoup(response.text, 'html.parser')

        # 找到包含  bsb cod 的表格行
        table_rows = soup.find_all('tr')

        # 找到表格中的所有 bsb code 和 Bank Name
        for row in table_rows:
            # 找到表格行中的所有单元格
            cells = row.find_all('td')
            if len(cells) >= 2:
                # 第一个单元格是 ach code
                ach_code = cells[1].text.strip()
                achCodes.append((ach_code, name))
    return achCodes


def export_to_excel(data, filename, folder):
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

    print(f"Data exported to {filepath}")


if __name__ == "__main__":
    base_url = 'https://bank-codes-hk.com/us-routing-number/bank/'
    ach_codes = scrape_ach_codes(base_url)
    if ach_codes:
        # 设置 Excel 文件名
        excel_filename = 'US_ACH.xlsx'
        # 设置保存的文件夹路径
        save_folder = 'file'
        # 导出数据到 Excel
        export_to_excel(ach_codes, excel_filename, save_folder)
    else:
        print("No bsb Codes found.")
