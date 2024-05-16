import os
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


def scrape_bsb_codes(url):
    # 存储所有 bsb Code 和 Bank Name
    bsb_codes_and_banks = []

    urls = []

    # 发送 GET 请求获取页面内容
    response = requests.get(url)

    if response.status_code == 200:
        # 使用 BeautifulSoup 解析页面内容
        soup = BeautifulSoup(response.text, 'html.parser')

        # 找到包含 Swift Code 的表格行
        table_rows = soup.find_all('ol')

        # 找到表格中的所有 Swift Code 和 Bank Name
        for row in table_rows:
            cells = row.find_all('li')
            for cell in cells:
                # 在单元格中查找链接
                link = cell.find('a')
                if link:
                    href = link.get('href')
                    urls.append(href)
    else:
        print(f"Failed to fetch URL: {url}. Status code: {response.status_code}")

    for lastUrl in urls:
        url = "https://bank-codes-hk.com" + lastUrl

        second_get(url, bsb_codes_and_banks)
    return bsb_codes_and_banks


def second_get(url, bsbCodes):
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
                # 第一个单元格是 Bank Name
                bank_name = cells[1].text.strip()
                # 第三个单元格是 bsb code
                bsb_code = cells[4].text.strip()
                bsbCodes.append((bsb_code, bank_name))
    return bsbCodes


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
    base_url = 'https://bank-codes-hk.com/australia-bsb-number/bank/'
    bsb_codes = scrape_bsb_codes(base_url)
    if bsb_codes:
        # 设置 Excel 文件名
        excel_filename = 'AU_BANK_TRANSFER.xlsx'
        # 设置保存的文件夹路径
        save_folder = 'file'
        # 导出数据到 Excel
        export_to_excel(bsb_codes, excel_filename, save_folder)
    else:
        print("No bsb Codes found.")
