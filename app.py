# -*- coding: utf-8 -*-
# streamlit run app.py --server.address 0.0.0.0
from time import strftime

import streamlit as st
import pandas as pd
import json
from datetime import date, timedelta
import os

import xlwings as xw
import openpyxl
import re

from win32con import PRINTRATEUNIT_PPM

from excel_manager import *

# 웹 화면 크롤링(selenium)
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ── 페이지 설정 ──────────────────────────────────────────────
st.set_page_config(page_title="IT 업무 스케쥴", page_icon="📋", layout="wide")

EXCEL_FILE = "basic_data.xlsx"
SHEET_NAME1 = "schedule"
SHEET_NAME2 = "holiday"

TODAY = pd.Timestamp(date.today())  # 오늘 날짜
# ── 데이터 로드 ──────────────────────────────────────────────
def read_excel_with_xlwings_IT(filename, sheet_name1, sheet_name2):
    print("=" * 50)
    print("--- 엑셀 파일 열기:", filename)
    print("=" * 50)

    # ── 1) Excel 앱을 통해 파일 열기 ──────────────────
    # visible=False : 엑셀 창을 화면에 띄우지 않음
    # visible=True : 엑셀 창이 실제로 열림 (디버깅할 때 유용)
    app = xw.App(visible=False)
    book = app.books.open(filename)

    try:
        st.subheader("오늘 우리팀의 할일 : ",divider=True)
        # ── 2) 시트 선택 ──────────────────────────────
        sheet = book.sheets[sheet_name1]  # 이름으로 선택
        # sheet = book.sheets[0] # 순서로 선택 (0 = 첫번째)
        print(f"- 시트 선택 완료: [{sheet.name}]")

        # ── 3) 데이터 전체를 DataFrame으로 읽기 ───────
        # used_range : 데이터가 있는 영역을 자동으로 감지
        df = sheet.used_range.options(pd.DataFrame, index=False, header=True).value

        df["Date"] = pd.to_datetime(df["Date"])  # 날짜 열만
        today = pd.Timestamp(date.today())
        today_df = df[df["Date"] == today]
        # print(f" - 오늘은 {today}")
        # print(f" - 가져온 날짜\n {df['date']}")
        # print(f" - 가져온 오늘의 업무\n {today_df}")

        if today_df.empty:
            st.subheader("데이타 없음")
        else:
            for _, row in today_df.iterrows():
                print("오늘 우리팀의 할일 : ")
                print(f"   날짜:  {row['Date'], strftime('%Y-%m-%d')}")
                print(f"   DSR : {row['DSR_Task1']}")
                print(f"   테라 : {row['테라_Task1']}")

        today_df["Date"] = today_df["Date"].dt.strftime('%Y-%m-%d')
        today_df =today_df.fillna("")  # None -> 공백으로 표시
        st.dataframe(today_df, use_container_width=True)

        # print(f"\n 전체 데이터 shape: {df.shape} (행 수, 열 수)")
        # print(f" 컬럼 목록: {list(df.columns)}")
        # print()
        # # ── 4) 한 줄씩 읽기 ───────────────────────────
        # print("-" * 50)
        # print(" 한 줄씩 읽기 시작")
        # print("-" * 50)
        #
        # for row_index, row in df.iterrows():
        #     # row_index : 0, 1, 2, 3 ... (pandas 기준 번호)
        #     # row : 각 행의 데이터 (Series 형태)
        #
        #     print(f"\n[{row_index + 1}번째 줄]")
        #
        #     # 열 이름과 값을 함께 출력
        #     for col_name, value in row.items():
        #         print(f"  {col_name} : {value}")
        #
        # print("\n" + "=" * 50)
        # print(" 읽기 완료!")
        # print("=" * 50)

        #return df # DataFrame을 반환 (다른 곳에서 활용 가능)

        st.subheader("오늘 팀 현황 : ", divider=True)
        # ── 3) 시트 선택 ──────────────────────────────
        sheet = book.sheets[sheet_name2]  # 이름으로 선택
        print(f"- 시트 선택 완료: [{sheet.name}]")
        # used_range : 데이터가 있는 영역을 자동으로 감지
        df = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        df["Date"] = pd.to_datetime(df["Date"])  # 날짜 열만
        today = pd.Timestamp(date.today())
        today_df = df[df["Date"] == today]
        if today_df.empty:
            st.subheader("데이타 없음")
        else:
            for _, row in today_df.iterrows():
                print("")

        today_df["Date"] = today_df["Date"].dt.strftime('%Y-%m-%d')
        today_df =today_df.fillna("")  # None -> 공백으로 표시
        st.dataframe(today_df, use_container_width=True)

    finally:
        # ── 5) 반드시 정리! ───────────────────────────
        # 파일 닫기 + Excel 앱 종료 (안하면 Excel이 백그라운드에 남아있음)
        book.close()
        app.quit()
        print("--- 엑셀 파일 닫기 완료")


# ──────────────────────────────────────────
# 실행
# ──────────────────────────────────────────
st.set_page_config(page_title="PEG1팀 IT 일정", page_icon="", layout="wide")
st.header("PEG1팀 ")

if __name__ == "__main__":

    df = read_excel_with_xlwings_IT(EXCEL_FILE, SHEET_NAME1, SHEET_NAME2)
    with st.sidebar:
        with st.echo():
            st.write("This code will be printed to the sidebar.")


    # 반환된 DataFrame 추가 활용 예시
    print("\n 특정 열만 보기 예시:")
    # print(df["날짜"])  # 날짜 열만
    # print(df.iloc[0]) # 첫 번째 행만
    # print(df.iloc[2, 1]) # 3번째 행, 2번째 열 값만

