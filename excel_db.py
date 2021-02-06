from openpyxl import Workbook
from openpyxl import load_workbook

EXCEL_FILE_NAME = 'database.xlsx'
db = load_workbook(filename = EXCEL_FILE_NAME)

response_db = db['Response']

#엑셀로 카카오톡 기본 UI 적용
def get_response(content, user_row):
    user_state = user_row[1].value
    
    data = []
    
    #엑셀파일 Response 시트에서 시나리오 데이터 가져오기
    # row[0] : 자동응답 버턴 명령어의 글자
    # row[1] : message
    # row[2] : photo_url
    # row[3] : message_button
    # row[4] : keyboard
    
    for row in response_db:
        
        if row[0].value == content:
            message = row[1].value
            photo_url = row[2].value
            message_button = row[3].value
            
            if message_button is not None:
                message_button = message_button.split('@')
                message_button_label = message_button[0]
                message_button_url = message_button[1]
            
            buttons = row[4].value
            break
            
        response = {
            "message" : {
                "text" : message
                },
            "keyboard" : {
                "type" : "buttons",
                "buttons" : buttons
                }
            }
        
        if buttons is not None:
            keyboard = {
                "type" : "buttons",
                "buttons" : buttons.split(',')
                }
            response['keyboard'] = keyboard
        else:
            keyboard = {
                "type" : "text"
                }
            response['keyboard'] = keyboard
            
            if photo_url is not None:
                photo = {
                    "url" : photo_url,
                    "width" : 640,
                    "height" : 480
                    }
                response['message']['photo'] = photo
                
            if message_button is not None:
                message_button = {
                    "label" : message_button_label,
                    "url" : message_button_url
                    }
                response['message']['message_button'] = message_button

            return response