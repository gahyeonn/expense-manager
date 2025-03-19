import pandas as pd
from datetime import date

class ExcelProcessor:
    def __init__(self, reading_path, saving_path):
        self.reading_path = reading_path
        self.saving_path = saving_path
        data = None
        
    def process_excel(self):
        # 엑셀 파일의 앞에 3개의 열 삭제
        self.data = self.parse_file()
        print("삭제 완")
        
        # 존재하는 계정과목 확인
        account_nums = self.data['계정코드'].unique()
        print('계정코드 종류: ', account_nums)
        
        # 파일을 저장할 위치 입력받기
        for account_num in account_nums:
            result = self.filter_by_account(account_num)
            self.save_file(result, f'{self.saving_path}/{account_num}_{date.today()}.xlsx') #계정과목_오늘날짜
        
        
    def parse_file(self):
        origin_data = pd.read_excel(self.reading_path)
        origin_df = pd.DataFrame(origin_data)
        
        # 전송여부, 문서번호, 문서항번 컬럼 삭제
        return origin_df.drop(origin_df.columns[[0, 1, 2]], axis=1)
        
    # 계정 과목 분류 함수
    def filter_by_account(self, account_num):
        return self.data[self.data['계정코드'] == account_num]
    
    # 계정 과목 기준으로 분류해서 파일 저장
    def save_file(self, data, file_path):
        data.to_excel(file_path, index=False)
        
    