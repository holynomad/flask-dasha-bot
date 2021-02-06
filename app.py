from flask import Flask, jsonify, request, json, render_template
import sys
from openpyxl import load_workbook
import excel_db

app = Flask(__name__)

# 챗봇 사용자정보 관리 파일
EXCEL_FILE_NAME = 'database.xlsx'

db = load_workbook(filename=EXCEL_FILE_NAME)
user_db = db['User']


@app.route("/")
def hello():
    #return "Hello, HISmacgyver on goorm.io!"
    return render_template("hello.html")

@app.route("/listlevel", methods=["get"])
def listLevel():
    userLevel = request.args.get("username")
    print(userLevel)
    
    #return render_template("hello.html")


@app.route("/searchlevel", methods = ["get"])
def searchLevel():
    
    usernm = request.args.get("username")
    
    print('searchLevel started... --> ' +usernm)
    
    # 엑셀로 사용자 정보 관리
    for idx, row in enumerate(user_db.rows):
        
        print('[' + str(idx) + ' / ' + str(user_db.max_row) + '] ' + str(row[0].value) + ' , ' + str(row[1].value) + ' , ' + str(row[2].value))
        
        if idx != 0 and row[2].value == usernm:
            user_row = row
            print('[1] user_row catched')
            break
      
        elif idx ==user_db.max_row - 1:
            NEW_INDEX = user_db.max_row + 1
            user_db[NEW_INDEX][0].value = '테스트_유져키'
            user_db[NEW_INDEX][1].value = 0
            db.save(EXCEL_FILE_NAME)
            user_row = user_db[user_db.max_row]
            
            print('[2] user_row catched and db saved')
            
            response = {
                "message": {
                    "text": "{name}님, 처음 방문 감사합니다. 본인의 컨트리뷰션 영역과 관심사 영역을 입력해주시겠어요?".format(name=usernm)
                },
                "keyboard": {
                    "type": "text"
                }
            }
            return jsonify(response)
    
    if user_row[1].value is 0 :
        user_row[1].value = 1
        user_row[2].value = usernm
        db.save(EXCEL_FILE_NAME)

    response = {
        "message": {
            "text" : "{name}님, 반갑습니다. 컨트리뷰션 영역은 {contr}, 관심사는 {curious} 입니다.".format(name=usernm, contr=user_row[3].value, curious=user_row[4].value)
        },
        "keyboard": {
            "type": "buttons",
            "buttons": ["HIS커뮤니티랩 소개", "콘텐츠표", "홈으로"]
        }
    }
    return jsonify(response)
    
    #엑셀로 카카오톡 기본 UI 구현 @ 2021.02.06.
    try:
        response = excel_db.get_response(content, user_row)
    except:
        response = {
            "message": {
                "text" : "다시 시도해 주세요."
            },
            "keyboard": {
                "type": "buttons",
                "buttons": ["홈으로"]
            }
        }
        
    return jsonify(response)
    

# ======= 여기서부턴 카카오톡 챗봇 연동 메소드 정의부 =========


@app.route("/keyboard")
def keyboard():
    
    response = {
        "type": "button",
            "buttons": ["홈으로"]
    }
    return jsonify(response)



@app.route("/message", methods = ["POST"])
def message():
    
    data = json.loads(request.data)
    
    content = data["content"]
    user_key = data["user_key"]
    
    # 엑셀로 사용자 정보 관리
    for idx, row in enumerate(user_db.rows):
        if idx != 0 and row[0].value == user_key:
            user_row = row
            break
      
        if idx ==user_db.max_row - 1:
            NEW_INDEX = user_db.max_row + 1
            user_db[NEW_INDEX][0].value = user_key
            user_db[NEW_INDEX][1].value = 0
            db.save(EXCEL_FILE_NAME)
            user_row = user_db[user_db.max_row]
            
            response = {
                "message": {
                    "text": "처음 방문하셨네요. 이름이 (직군이, 직종이, 직무가..) 어떻게 되세요?"
                },
                "keyboard": {
                    "type": "text"
                }
            }
            return jsonify(response)
            
        if user_row[1].value is 0 :
            user_row[1].value = 1
            user_row[2].value = content
            db.save(EXCEL_FILE_NAME)
            
            response = {
                "message": {
                    "text" : "{name}님, 반갑습니다.".format(name=content)
                },
                "keyboard": {
                    "type": "buttons",
                    "buttons": ["다샤소개", "콘텐츠표", "홈으로"]
                }
            }
            return jsonify(response)
            
    #카카오톡 기본 UI 코딩 --> 엑셀에 콘텐츠 넣어 연동(excel_db) 전환에 따른 주석 @ 2021.02.06.
    '''
    if content == u"홈으로":
        response = {
            "message" : {
                "text" : "원하시는 정보 버튼 눌러주세요."
            },
            "keyboard": {
                "type": "buttons",
                "buttons": ["다샤소개", "콘텐츠표","홈으로"]
            }
        }
    elif content == u"다샤소개":
        response = {
            "message": {
                "text": "다샤는 비즈니스 마이크로러닝 플랫폼 !"
            },
            "keyboard": {
                "type" : "buttons",
                "buttons": ["홈으로"]
            }
        }
    elif content == u"콘텐츠표":
        response = {
            "message": {
                "text": "다샤 콘텐츠 목록입니다.",
                "message_button": {
                    "label": "웹링크로 콘텐츠목록 보기",
                    "url": "https://www/ai-academy.ai/ai-b2c"
                }
            },
            "keyboard": {
                "type": "buttons",
                "buttons":["홈으로"]
            }
        }
    '''
        
    '''
    response = {
        "message": {
            "text": "Hello, World"
        }
    }
    '''
    
    #엑셀로 카카오톡 기본 UI 구현 @ 2021.02.06.
    try:
        response = excel_db.get_response(content, user_row)
    except:
        response = {
            "message": {
                "text" : "다시 시도해 주세요."
            },
            "keyboard": {
                "type": "buttons",
                "buttons": ["홈으로"]
            }
        }
        
    return jsonify(response)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(sys.argv[1]), debug=True)
