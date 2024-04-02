import json 
from urllib.request import urlopen

url_json= 'https://storage.yandexcloud.net/shar-2024/sample_1.json';
data = json.loads(urlopen(url_json).read().decode("utf-8"))

vowels = set('aeiouy')


name_json = url_json.split('/')[-1]
str_for_json = ['deck_mast_port side', 'windlass_hold_hole', 'deckhouse_stern_crossbar sail', 'foresail']

result = {}
count = 0

for elem in str_for_json:
    elem_spl = elem.split('_')
    for elem_det in elem_spl:
        for key in data.keys():
            if data[key] == '10':
                for letter in set(elem_det):
                    if letter in vowels:
                        count += 1
                if count < 2:  
                    result.setdefault(key, []).append(elem_det)
            elif data[key] == '20' and len(elem_det)%2==0:
                result.setdefault(key, []).append(elem_det)
            elif data[key] == '30':
                result.setdefault(key, []).append(elem_det)

print(result)

              
           
           
