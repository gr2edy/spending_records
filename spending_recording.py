# 라이브러리
import telegram
from telegram.ext import Updater, MessageHandler, Filters
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import numbers
from datetime import datetime
from config import spending_recording_bot_token
# 토큰
TOKEN = spending_recording_bot_token
bot = telegram.Bot(token=TOKEN)

# 메시지 처리 함수
def handle_message(update, context):
    # 메시지 파싱
    message = update.message.text
    data = message.split()
    
    # 가계부 불러오기
    df = pd.read_excel('가계부.xlsx', skiprows = 1)
    variable_spending = df.iloc[:,16:].dropna()
    this_month_income = int(df.iloc[1,4])
    this_month_spending = int(df.iloc[2,4])
    this_month_saving = int(df.iloc[3,4])
    this_month_summary = int(df.iloc[5,4])

    file_idx = variable_spending.index.max()
    # 입력할 셀 위치
    cell_positions = [f'Q{file_idx + 4}', f'R{file_idx + 4}', f'S{file_idx + 4}', f'T{file_idx + 4}']

    # 입력할 데이터
    input_data = []
    for item in data:
        try:
            input_data.append(item)
        except ValueError:
            pass
    date = input_data[0]        # 날짜
    amount = input_data[1]      # 금액
    main = input_data[2]        # 내용
    spending_type = input_data[3] # 타입
    
    # 엑셀 파일 열기    
    workbook = load_workbook(filename='가계부.xlsx')

    # 시트 선택
    sheet = workbook.active
    
    # 입력할 셀 위치를 반복문을 사용하여 선택하고 데이터 입력
    for idx, cell_pos, data in zip(range(4),cell_positions, input_data):
        if idx == 0:
            date_obj = datetime.strptime(data, '%Y-%m-%d')
            sheet[cell_pos] = date_obj
            sheet[cell_pos].number_format = 'yyyy.mm.dd'
        elif idx == 1:
            sheet[cell_pos] = int(data)
        else:
            sheet[cell_pos] = data
    # 파일 저장
    workbook.save(filename='가계부.xlsx')

    # 응답 전송
    update.message.reply_text('데이터가 입력되었습니다.')
    update.message.reply_text('이번 달 총 지출은 {:,}원 입니다.'.format(-this_month_spending))


# 봇 생성
# 봇 생성
updater = Updater(token=TOKEN, use_context=True)

# 메시지 처리 핸들러 등록
updater.dispatcher.add_handler(MessageHandler(Filters.text, handle_message))

# 봇 시작
updater.start_polling()
updater.idle()
