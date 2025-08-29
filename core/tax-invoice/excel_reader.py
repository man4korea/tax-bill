# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys

def analyze_excel_structure(file_path):
    """엑셀 파일 구조 분석"""
    try:
        # 엑셀 파일 열기
        excel_file = pd.ExcelFile(file_path)
        print(f"엑셀 파일: {file_path}")
        print(f"시트 개수: {len(excel_file.sheet_names)}")
        print(f"시트 목록: {excel_file.sheet_names}")
        print("\n" + "="*50)
        
        # 각 시트 구조 확인
        for sheet_name in excel_file.sheet_names:
            print(f"\n[시트명: {sheet_name}]")
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"행 개수: {len(df)}")
            print(f"컬럼명: {df.columns.tolist()}")
            
            # 데이터가 있으면 첫 3행 표시
            if len(df) > 0:
                print("첫 3행 데이터:")
                print(df.head(3).to_string())
            else:
                print("데이터 없음")
            print("-" * 30)
        
        return excel_file.sheet_names
        
    except Exception as e:
        print(f"오류: {e}")
        return None

def read_transaction_details(file_path, sheet_name='거래명세표'):
    """거래명세표 시트 데이터 읽기"""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"\n거래명세표 데이터:")
        print(f"총 {len(df)}개의 거래내역")
        
        # 데이터 정리 및 반환
        return df.to_dict('records')  # 리스트 형태로 반환
        
    except Exception as e:
        print(f"거래명세표 읽기 오류: {e}")
        return []

if __name__ == "__main__":
    excel_path = r"C:\Users\man4k\OneDrive\문서\세금계산서.xlsx"
    
    if os.path.exists(excel_path):
        print("엑셀 파일 구조 분석 시작...")
        sheets = analyze_excel_structure(excel_path)
        
        if sheets and '거래명세표' in sheets:
            transactions = read_transaction_details(excel_path)
            print(f"\n거래명세표에서 {len(transactions)}개 데이터 읽기 완료")
        else:
            print("거래명세표 시트를 찾을 수 없습니다.")
    else:
        print(f"엑셀 파일이 존재하지 않습니다: {excel_path}")