"""
-*- coding: utf-8 -*-
time     : 2023/9/6
author   : hillking
file     : pdf_ex.py
function : 提取pdf文件文本内容
"""

import pandas as pd
import json

import requests

# "https://github.com/hiyouga/llama-efficient-tuning"

#
# import pdfplumber
#
# pdf_path = 'd:\project\dataset\\2022注会官方教材会计.pdf'
# with pdfplumber.open(pdf_path) as pdf:
#     first_page = pdf.pages[0]
#     print(first_page.extract_text())  # 只能提取文本格式的内容，不能提取提取扫描型的pdf
#







def http_llm_service(url, prompt):
    request_data = {
        "prompt": prompt
        , "max_new_tokens": 2048,
    }

    json_data = json.dumps(request_data)
    # 获取响应结果，以json形式返回
    json_response = requests.post(url, data=json_data).json()
    return json_response


# df1 = pd.read_excel("111.xlsx",  index_col=0)  # header=0,  ,  index_col=0
#
# df = pd.DataFrame(df1)
# print(df)
# print('\n\n')
# prompt = '已知\n' + str(df) + '\n请根据表格回答问题，不要胡编乱造。问题：李良的兴趣爱好有哪些'


# print(df)
# exit()
# list_data = [['John', 28], ['Bob', 24], ['Alice', 31, ['Tom', 22]]
# df = pd.DataFrame(list_data, columns=['Name', 'Age'])
# prompt = '已知\n' + str(df) + '\n请根据表格回答问题，不要胡编乱造。问题：Bob的年龄是多少'


qs = """
核算主体                        核算时点/业务场景              会计科目     科目编码      弹性域    备件要求
新房主体                     预付佣转入到“额外收入”        借：其他应付款_快佣   224146 往来单位：开发商    银行流水
                                            贷：其他业务收入_其他   605198                 
                                      贷：应交税费_应交增值税_销项税额 22211303                 
     剩余未结算结佣保证金需退甲方时，A+系统款项拆分至“退款给甲方”        借：其他应付款_快佣   224146 往来单位：开发商 解约协议、对账
"""
prompt = '已知：\n' + qs + '\n根据表格回答问题，不要胡编乱造。问题：科目编码有哪些？'


url = 'http://test-bigdata-digging.aicloud.ke.com/bclever/bigdata-digging/bac-qa-bc/v1/infer'

json_response = http_llm_service(url, prompt)
print(prompt)
res_total = json_response["text"]
print(res_total)