# 언어분석, 어휘관계분석, 질문분석, 음성인식 처리

import urllib3
import json

openApiURL = "http://aiopen.etri.re.kr:8000/WiseNLU"
accessKey = "개발자 식별키를 입력해주세요"

def exobrainNLU(type, sentence):
    analysisCode = type
    requestJson = {
        "access_key" : accessKey,
        "argument" : {
            "text" : sentence,
            "analysis_code" : analysisCode
            }
    }
    http = urllib3.PoolManager()
    response = http.request("POST",openApiURL,headers={"Content-Type": "application/json;charset=UTF-8"},body = json.dumps(requestJson, indent= 2))
    data = json.loads(response.data.decode('utf-8'))
    return list(data["return_object"]['sentence'])

# 엑소브레인 개체명 인식 API를 사용하여, 문장에서 날짜, 시간, 장소정보 추출하기

def get_date_from_sentence(sentence):
    keyword_sets = []
    
    # 엑소브레인 API - 개체명 인식 (ner)
    malist = exobrainNLU("ner",sentence)
    
    #print(malist)
    
    for i in range(len(malist[0]['NE'])):
        keyword = malist[0]['NE'][i]['text']
        keyword_type = malist[0]['NE'][i]['type']
        
        print("".join(keyword_type[:2]))
        
        keyword_type = list(map(lambda x,y : "날짜" if x + y == "DT" else "시간" if x + y == "TI" else "장소" if x + y == "LC" else x + y , keyword_type[0],keyword_type[1]))
        keyword_set = (keyword , keyword_type)
        keyword_sets.append(keyword_set)

        print(keyword_sets)
    return keyword_sets