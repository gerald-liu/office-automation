import pandas as pd
import numpy as np
import os
import re
import pdfplumber

# get the folder path 每月修改两个路径名称即可
# 请不要在文件夹中放置月结单以外的其他PDF
pdf_folder_path = "G:/_Private/EQUITY_DERIVATIVES/合规材料/CSRC业务摸查/24年5月"
# get the output path
output_path = "G:/_Private/EQUITY_DERIVATIVES/合规材料/CSRC业务摸查/24年5月"

def extract_balance_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            # 按照相应的月结单格式查找对应的总资产数字
            # 月结单格式为  中文/英文****HKD***带“，”数字或者纯数字+小数点+纯数字
            balance_match = re.search(r'Total Portfolio in Equivalent\s*HKD\s*((\d{1,3}(?:,\d{3})*)|\d+)\.(\d+)', text, re.IGNORECASE)
            if balance_match:
                #group(1)为子一个整体的数据集((\d{1,3}(?:,\d{3})*)|\d+)，“|”表示或者，group(2)默认为group(1)中的(\d{1,3}(?:,\d{3})*)
                group1 = balance_match.group(1).replace(",","")
                group3 = balance_match.group(3)
                balance = int(str(group1) + str(group3))/100
            else:
                balance_match = re.search(r'投資組合總值\s*HKD\s*((\d{1,3}(?:,\d{3})*)|\d+)\.(\d+)', text, re.IGNORECASE)
                if balance_match:
                    group1 = balance_match.group(1).replace(",","")
                    group3 = balance_match.group(3)
                    balance = int(str(group1) + str(group3))/100
                else:
                    balance_match = re.search(r'投资组合总值\s*HKD\s*((\d{1,3}(?:,\d{3})*)|\d+)\.(\d+)', text, re.IGNORECASE)
                    if balance_match:
                        group1 = balance_match.group(1).replace(",","")
                        group3 = balance_match.group(3)
                        balance = int(str(group1) + str(group3))/100
    return balance

def sum_balances_in_folder(folder_path):
    total_balance = 0
    balance = []
    ID_list = []
    number = 0
    for filename in os.listdir(folder_path):
        if filename.endswith('.PDF'):
            number = number + 1
            print(number)
            pdf_path = os.path.join(folder_path, filename)
            ID_list.append(filename)
            total_balance += extract_balance_from_pdf(pdf_path)
            balance.append(extract_balance_from_pdf(pdf_path))
            print(extract_balance_from_pdf(pdf_path))

    ID_list = np.array(ID_list).tolist()
    ID_list = [elem[:9] for elem in ID_list]
    out_put_balance = pd.DataFrame(balance, index=ID_list)
    print(ID_list)
    return total_balance, out_put_balance

total_balance, balance_csv = sum_balances_in_folder(pdf_folder_path)

print(f'Total balance: {total_balance}')
output_file = os.path.join(output_path, 'out.csv')
balance_csv.to_csv(output_file)
