# -*- coding: utf-8 -*-
import pandas as pd
import os
from datetime import datetime

class ExcelDataManager:
    """엑셀 데이터 관리 클래스"""
    
    def __init__(self, excel_path=None):
        self.excel_path = excel_path or r"C:\Users\man4k\OneDrive\문서\세금계산서.xlsx"
        self.transaction_data = []
        self.customer_data = []
        
    def load_all_data(self):
        """모든 시트 데이터 로드"""
        try:
            # 거래명세표 데이터 로드
            self.transaction_data = self.load_transaction_details()
            # 거래처 데이터 로드
            self.customer_data = self.load_customer_data()
            
            print(f"데이터 로드 완료:")
            print(f"   거래명세표: {len(self.transaction_data)}개")
            print(f"   거래처: {len(self.customer_data)}개")
            
            return True
            
        except Exception as e:
            print(f"데이터 로드 오류: {e}")
            return False
    
    def load_transaction_details(self):
        """거래명세표 시트 데이터 로드"""
        try:
            df = pd.read_excel(self.excel_path, sheet_name='거래명세표')
            
            # NaN 값 처리
            df = df.fillna('')
            
            # 컬럼명 대신 인덱스로 접근 (한글 인코딩 문제 해결)
            if len(df) == 0:
                print("거래명세표 시트가 비어있습니다.")
                return []
            
            # 컬럼 순서 확인 (첫 행으로 매핑)
            col_names = df.columns.tolist()
            print(f"컬럼 개수: {len(col_names)}")
            
            transactions = []
            for idx, row in df.iterrows():
                # 첫 번째 컬럼이 작성일자인지 확인
                if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == '':
                    continue
                
                transaction = {
                    '작성일자': str(row.iloc[0]).strip(),      # 첫 번째 컬럼
                    '등록번호': str(row.iloc[1]).strip(),      # 두 번째 컬럼
                    '상호': str(row.iloc[2]).strip(),          # 세 번째 컬럼
                    '품목코드': str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else '',
                    '품명': str(row.iloc[4]).strip(),          # 품명
                    '규격': str(row.iloc[5]).strip(),          # 규격
                    '수량': int(float(row.iloc[6])) if pd.notna(row.iloc[6]) else 0,    # 수량
                    '단가': int(float(row.iloc[7])) if pd.notna(row.iloc[7]) else 0,    # 단가
                    '공급가액': int(float(row.iloc[8])) if pd.notna(row.iloc[8]) else 0, # 공급가액
                    '세액': int(float(row.iloc[9])) if pd.notna(row.iloc[9]) else 0,     # 세액
                }
                
                # 계산된 총액 추가
                transaction['총액'] = transaction['공급가액'] + transaction['세액']
                
                transactions.append(transaction)
            
            print(f"거래명세표에서 {len(transactions)}건의 유효 데이터 로드")
            return transactions
            
        except Exception as e:
            print(f"거래명세표 로드 오류: {e}")
            return []
    
    def load_customer_data(self):
        """거래처 시트 데이터 로드"""
        try:
            df = pd.read_excel(self.excel_path, sheet_name='거래처')
            df = df.fillna('')
            
            if len(df) == 0:
                print("거래처 시트가 비어있습니다.")
                return []
            
            customers = []
            for idx, row in df.iterrows():
                # 거래처명이 있는 행만 처리 (세 번째 컬럼)
                if pd.isna(row.iloc[2]) or str(row.iloc[2]).strip() == '':
                    continue
                
                customer = {
                    '순번': row.iloc[0] if pd.notna(row.iloc[0]) else '',
                    '사업자등록번호': str(row.iloc[1]).strip(),   # 사업자등록번호
                    '거래처명': str(row.iloc[2]).strip(),        # 거래처명 
                    '대표자': str(row.iloc[3]).strip(),          # 대표자
                    '사업장주소': str(row.iloc[4]).strip(),      # 사업장주소
                    '업태': str(row.iloc[5]).strip(),           # 업태
                    '종목': str(row.iloc[6]).strip(),           # 종목
                }
                customers.append(customer)
            
            print(f"거래처에서 {len(customers)}개 데이터 로드")
            return customers
            
        except Exception as e:
            print(f"거래처 데이터 로드 오류: {e}")
            return []
    
    def get_transactions_by_date(self, target_date=None):
        """특정 날짜의 거래 조회"""
        if not target_date:
            target_date = datetime.now().strftime('%Y-%m-%d')
        
        filtered = []
        for tx in self.transaction_data:
            tx_date = tx['작성일자']
            # 날짜 형식 통일
            if isinstance(tx_date, str) and len(tx_date) >= 10:
                tx_date_str = tx_date[:10]  # 'YYYY-MM-DD' 부분만 추출
                if tx_date_str == target_date:
                    filtered.append(tx)
        
        return filtered
    
    def get_customer_by_business_number(self, business_number):
        """사업자등록번호로 거래처 조회"""
        for customer in self.customer_data:
            if customer['사업자등록번호'] == business_number:
                return customer
        return None
    
    def get_transaction_summary(self):
        """거래 요약 정보"""
        if not self.transaction_data:
            return {}
        
        total_supply = sum(tx['공급가액'] for tx in self.transaction_data)
        total_tax = sum(tx['세액'] for tx in self.transaction_data)
        total_amount = total_supply + total_tax
        
        unique_customers = set(tx['등록번호'] for tx in self.transaction_data)
        
        return {
            '총_거래건수': len(self.transaction_data),
            '총_공급가액': total_supply,
            '총_세액': total_tax,
            '총액': total_amount,
            '거래처_수': len(unique_customers)
        }
    
    def print_transaction_summary(self):
        """거래 요약 정보 출력"""
        if not self.transaction_data:
            print("로드된 거래 데이터가 없습니다.")
            return
        
        summary = self.get_transaction_summary()
        
        print(f"\n=== 거래 요약 정보 ===")
        print(f"총 거래건수: {summary['총_거래건수']:,}건")
        print(f"총 공급가액: {summary['총_공급가액']:,}원")
        print(f"총 세액: {summary['총_세액']:,}원") 
        print(f"총액: {summary['총액']:,}원")
        print(f"거래처 수: {summary['거래처_수']}개")
        
        # 최근 거래 3건 표시
        print(f"\n=== 최근 거래 3건 ===")
        recent_transactions = sorted(self.transaction_data, 
                                   key=lambda x: x['작성일자'], reverse=True)[:3]
        
        for i, tx in enumerate(recent_transactions, 1):
            print(f"{i}. [{tx['작성일자']}] {tx['상호']} - {tx['품명']}")
            print(f"   {tx['공급가액']:,}원 + {tx['세액']:,}원 = {tx['총액']:,}원")

if __name__ == "__main__":
    # 테스트 실행
    print("엑셀 데이터 관리자 테스트")
    
    manager = ExcelDataManager()
    
    if manager.load_all_data():
        manager.print_transaction_summary()
        
        # 오늘 날짜 거래 조회 테스트
        today = datetime.now().strftime('%Y-%m-%d')
        today_transactions = manager.get_transactions_by_date(today)
        print(f"\n{today} 거래: {len(today_transactions)}건")
    else:
        print("데이터 로드 실패")