"""
Esse scrpit gera um arquivo excel com o nome de "poe_gems.xlsx" listando todas as gemas de path of exile que:
 - tenham "Awakened", "Empower" ou "Enlighten" em seu nome;
 - nao estejam corrompidas
 - tenham lvl 1 ou 5
 - tenham qualidade diferente de 0
mostrando seus atributos de "gemLevel", "chaosValue", "exaltedValue" e "chaosValueDiff", que da a diferenca de preco
entre a mesma gema de lvl 1 e lvl 5. 

"""
import requests
import json
import xlsxwriter

workbook = xlsxwriter.Workbook('poe_gems.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, "Name")
worksheet.write(0, 1, "gemLevel")
worksheet.write(0, 2, "exaltedValue")
worksheet.write(0, 3, "chaosValue")
worksheet.write(0, 4, "chaosValueDiff")

url = "https://poe.ninja/api/data/ItemOverview?league=Scourge&type=SkillGem&language=en"
page = requests.get(url)
pagetext = page.text

pagejson = json.loads(pagetext)
gems = pagejson['lines']

gems_of_interest = []

for i in range(len(gems)):
    try:
        gems[i]['corrupted']
        continue
    except:
        try:
            gemQuality = gems[i]['gemQuality']
            gemName = gems[i]['name']
            words_of_interest = ["Awakened", "Empower", "Enlighten"]
            if any(word in gemName for word in words_of_interest) and (gems[i]['gemLevel'] == 1 or gems[i]['gemLevel'] == 5):
                nice_gem = {'name': gems[i]['name'],
                            'gemLevel': gems[i]['gemLevel'],
                            'exaltedValue': gems[i]['exaltedValue'],
                            'chaosValue': gems[i]['chaosValue']}
                gems_of_interest.append(nice_gem)
            else: continue 
        except: continue

gems_of_interest.sort(key=lambda d: d['name'])
gems_of_most_interest = []

for i in range(len(gems_of_interest)-1):
    if gems_of_interest[i]['name'] == gems_of_interest[i+1]['name']:
        diff = round(abs(gems_of_interest[i]['chaosValue'] - gems_of_interest[i+1]['chaosValue']), 2)
        gems_of_interest[i]['diff'] = diff
        gems_of_interest[i+1]['diff'] = diff
        gems_of_most_interest.append(gems_of_interest[i])
        gems_of_most_interest.append(gems_of_interest[i+1])

gems_of_most_interest.sort(key=lambda d: d['diff'], reverse=True)

line = 1

for i in range(len(gems_of_most_interest)):
    worksheet.write(line, 0, gems_of_most_interest[i]["name"])
    worksheet.write(line, 1, gems_of_most_interest[i]["gemLevel"])
    worksheet.write(line, 2, gems_of_most_interest[i]["exaltedValue"])
    worksheet.write(line, 3, gems_of_most_interest[i]["chaosValue"])
    worksheet.write(line, 4, gems_of_most_interest[i]["diff"])
    line += 1

workbook.close()
