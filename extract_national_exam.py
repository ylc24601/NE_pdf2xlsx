#!/usr/bin/env python
# coding: utf-8
from pdfminer.high_level import extract_text
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer
import re
import pandas as pd
from itertools import zip_longest
from io import BytesIO
import streamlit as st
import time

def strQ2B(ustring):
    """把字串全形轉半形"""
    ss = []
    for s in ustring:
        rstring = ""
        for uchar in s:
            inside_code = ord(uchar)
            if inside_code == 12288:  # 全形空格直接轉換
                inside_code = 32
            elif (inside_code >= 65281 and inside_code <= 65374):  # 全形字元（除空格）根據關係轉化
                inside_code -= 65248
            rstring += chr(inside_code)
        ss.append(rstring)
    return ''.join(ss)

@st.cache
# Retrieve question and options from pdf
def getQnOptions(file_name):
    '''list questions and options in an ordered sequence'''
    rawtext =''
    for page_layout in extract_pages(file_name):
        temp_list = []
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                temp_list.append((element.bbox[1],element.get_text()))
        temp_list.sort(reverse=True)
        for item in temp_list:
            rawtext += item[1]
    return rawtext


@st.cache
def getAnswer(file_name):
    ans_text = extract_text(file_name)
    strQ2B(ans_text)
    Ans = re.findall(r' ([ABCD#])', strQ2B(ans_text))
    if len(Ans) != 100:
        print('Answer numbers are not equal to 100, please check regex search pattern!')
    else:
        return Ans



def main(QuestionFileName, answerFileName,tags=False):
    extracted_text = getQnOptions(QuestionFileName)
    year = re.findall(r'(\d+)年', extracted_text)
    if re.search("第一次", extracted_text):
        order = "1"
    else:
        order = "2"
    if re.search("類科名稱：醫師", extracted_text):
        kind = "M"
        kind_name = "醫師"
    elif re.search("類科名稱：牙醫師", extracted_text):
        kind = "D"
        kind_name = "牙醫師"

    sequence = (kind,year[0],order)
    st.write("此份考題為",kind_name,year[0],"年第",order,"次考題")
    test = "-".join(sequence)
    if tags:
        Q = re.findall(r'\n(\d+\.[\s\S]*?(?=\n\s*A\.))', extracted_text)
        A = re.findall(r'\n\s*(A\.[\s\S]*?(?=\nB))', extracted_text)
        B = re.findall(r'\n(B\.[\s\S]*?(?=\nC))', extracted_text)
        C = re.findall(r'\n(C\.[\s\S]*?(?=\nD))', extracted_text)
        D = re.findall(r'\n(D\.[\s\S]*?(?=\n\d|\n$))', extracted_text)
    else:
        Q= re.findall(r'\n\d+\.([\s\S]*?(?=\n\s*A\.))', extracted_text)
        A = re.findall(r'\n\s*A\.([\s\S]*?(?=\nB))', extracted_text)
        B = re.findall(r'\nB\.([\s\S]*?(?=\nC))', extracted_text)
        C = re.findall(r'\nC\.([\s\S]*?(?=\nD))', extracted_text)
        D = re.findall(r'\nD\.([\s\S]*?(?=\n\d|\n$))', extracted_text)

    data = list(zip_longest(Q, A, B, C, D, fillvalue=''))
    cleaned_data = [[string.replace('\n', '') for string in row] for row in data]

    # append answers to Questions and Options nested list
    for n, i in enumerate(cleaned_data):
        i.append(getAnswer(answerFileName)[n])
    df = pd.DataFrame(cleaned_data, columns=['Question', 'A', 'B', 'C', 'D', 'Answer'])
    df.index += 1
    # df["Origin"] = "-".join(sequence) + str(df.index)
    origin_check = st.sidebar.checkbox(label="加入Origin欄位:", value= True, help="如M-110-2_3")
    if origin_check:
        df["Origin"] = test + "_" + df.index.astype(str)
    kind_check = st.sidebar.checkbox(label="加入Kind欄位:", value= True, help="醫師M或牙醫師D")
    if kind_check:
        df["Kind"] = kind
    year_check = st.sidebar.checkbox(label="加入Year欄位:", value= True, help="考試年度")
    if year_check:
        df["Year"] = year[0]
    order_check = st.sidebar.checkbox(label="加入Order欄位:", value= True, help="第幾次考試")
    if order_check:
        df["Order"] = order
    qnum_check = st.sidebar.checkbox(label="加入Qnum欄位:", value= True, help="第幾題")
    if qnum_check:
        df["Qnum"] = df.index
    return df, test


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


st.title("國考試題匯整為excel")
st.sidebar.subheader("1. 上傳試題檔")
uploaded_testsheet = st.sidebar.file_uploader("upload a pdf file", key=1)
st.sidebar.subheader("2. 上傳答案檔")
uploaded_answer_sheet = st.sidebar.file_uploader("upload a pdf file", key=2)

if uploaded_answer_sheet and uploaded_answer_sheet is not None:
    df, test_name = main(uploaded_testsheet,uploaded_answer_sheet)
    st.dataframe(df)
    xlsx = to_excel(df)
    csv_clicked = st.download_button(
        label='Download Excel File',
        data=xlsx,
        file_name=f"{test_name}.xlsx")


# if __name__ == '__main__':
#     main('110101_1301.pdf', '110101_ANS1301.pdf', 'M-110-2_wo_tag.xlsx', tags=False)
