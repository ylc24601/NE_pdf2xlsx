#!/usr/bin/env python
# coding: utf-8
from pdfminer.high_level import extract_text
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTFigure
from PyPDF2 import PdfReader
from operator import itemgetter
import re
import pandas as pd
from itertools import zip_longest
from io import BytesIO
import streamlit as st


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

@st.cache_data
# Retrieve question and options from pdf
def getQnOptions(file_name):
    '''list questions and options in an ordered sequence'''
    rawtext =''
    count_dict = {}
    img_list = []
    image_location_list = []
    # take text and figure element in a every page
    for page_layout in extract_pages(file_name):
        temp_list = []
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                rawtext += element.get_text()
                temp_list.append((element.bbox[1],element.get_text()))
            if isinstance(element, LTFigure):
                temp_list.append((element.bbox[1],element)) 
        # sort the coordinates of the element by bbox
        temp_list.sort(reverse=True, key=itemgetter(0)) 
        # locate the figures
        for item in temp_list:
            try: 
                found_number = re.match(r'\n?(\d+)\.', item[1]) 
            except: 
                pass
            if found_number:
                q_num = found_number.group(1)
            if isinstance(item[1], LTFigure):
                #check whether the image within the LTFigure object has benn used in the same pdf
                if item[1].matrix[0:4] not in img_list:
                    img_list.append(item[1].matrix[0:4])
                    # assign the number to images
                    if q_num in count_dict:
                        count_dict[q_num] += 1
                        image_location_list.append(q_num + "_" + str(count_dict[q_num]))
                    else:
                        count_dict[q_num] = 0
                        image_location_list.append(q_num)
    return rawtext, image_location_list


@st.cache_data
def getAnswer(file_name):
    ans_text = extract_text(file_name)
    strQ2B(ans_text)
    Ans = re.findall(r' ([ABCD#])', strQ2B(ans_text))
    return Ans


def get_images(file_name, location_list, test_name):
    reader = PdfReader(file_name)
    image_num = 0
    for page in reader.pages:
        for image_file_object in page.images:
            st.image(image_file_object.data,
                     caption="圖片位於題目" + location_list[image_num] + "; 副檔名為" + image_file_object.name.split(".")[1])
            st.download_button(
                label="Downlaod image",
                data = image_file_object.data,
                file_name = test_name + "_" + location_list[image_num] + "." + image_file_object.name.split(".")[1]
            )
            image_num += 1
    

def main(QuestionFileName, answerFileName,tags=False):
    extracted_text= getQnOptions(QuestionFileName)[0]
    year = re.findall(r'(\d+)\s*年', extracted_text)
    if re.search("第一次", extracted_text[0:10]):
        order = "1"
    elif re.search("第二次", extracted_text[0:10]):
        order = "2"
    if re.search("類科名稱：醫師", extracted_text):
        kind = "M"
        kind_name = "醫師"
    elif re.search("類科名稱：牙醫師", extracted_text):
        kind = "D"
        kind_name = "牙醫師"
    # st.write(extracted_text) # <- this line is for debug
    sequence = (kind,year[0],order)
    st.write("此份考題為",kind_name,year[0],"年第",order,"次考題")
    st.warning("請確認是否正確!")
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
    origin_check = st.sidebar.checkbox(label="加入Origin欄位:", value= True, help="如M-110-2_3", key = "c1")
    if origin_check:
        df["Origin"] = test + "_" + df.index.astype(str)
    kind_check = st.sidebar.checkbox(label="加入Kind欄位:", value= True, help="醫師M或牙醫師D", key = "c2")
    if kind_check:
        df["Kind"] = kind
    year_check = st.sidebar.checkbox(label="加入Year欄位:", value= True, help="考試年度", key = "c3")
    if year_check:
        df["Year"] = year[0]
    order_check = st.sidebar.checkbox(label="加入Order欄位:", value= True, help="第幾次考試", key = "c4")
    if order_check:
        df["Order"] = order
    qnum_check = st.sidebar.checkbox(label="加入Qnum欄位:", value= True, help="第幾題", key = "c5")
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

if uploaded_testsheet and uploaded_answer_sheet is not None:
    df, test_name = main(uploaded_testsheet,uploaded_answer_sheet)
    st.dataframe(df)   
    xlsx = to_excel(df)
    csv_clicked = st.download_button(
        label='Download Excel File',
        data=xlsx,
        file_name=f"{test_name}.xlsx")
    st.markdown("""---""")
    st.subheader("試題中的圖片")
    st.metric(label="找到的圖案數量", value = f"{len(getQnOptions(uploaded_testsheet)[1])}張")
    # st.write(getQnOptions(uploaded_testsheet)[1])
    get_images(uploaded_testsheet, getQnOptions(uploaded_testsheet)[1], test_name)

else:
    st.write("將試題與答案之pdf檔案分別拖曳至左側上傳區")