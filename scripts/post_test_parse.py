import requests, json
url='http://127.0.0.1:8000/parse'
file_path='data/clean_data.xlsx'
with open(file_path,'rb') as f:
    files={'file':('clean_data.xlsx',f,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
    r=requests.post(url,files=files, timeout=60)
    print('Status:', r.status_code)
    try:
        print(json.dumps(r.json(), indent=2))
    except Exception as e:
        print('Response text:', r.text)
