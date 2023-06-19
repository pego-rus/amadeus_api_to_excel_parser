from datetime import datetime
from amadeus import Client, ResponseError
from datetime import datetime
import pandas as pd
from pandas import json_normalize
import os


amadeus = Client(
    client_id='YOUR_AMADEUS_API',
    client_secret='YOUR_AMADEUS_SECRET_KEY'
)

# READ ROUTES FROM FILE IN PROJECT ROOT FOLDER
def inputroute(fileroute):
    excel = pd.read_excel(fileroute)
    originlist = excel['Трехзначный код аэропорта вылета'].tolist()
    destinationlist = excel['Трехзначный код аэропорта прибытия'].tolist()
    return(originlist, destinationlist)

# REBUILT JSON TO WRITE ALL COLUMNS
def format_json(json, route, inputdate):
    new_array_full = []
    for i in range (len(json)):
        new_json = {}
        new_json.update({'DATE':f'{inputdate}'})
        new_json.update({'Route':f'{route}'})
        for j in json[i]:
            if type(json[i][j]) is list:
                if type(json[i][j][0]) is str:
                    new_json.update({f'{j}:0':f"{json[i][j][0]}"})
                else:   
                    for q in json[i][j][0]:
                        if type(json[i][j][0][q]) is str:
                            new_json.update({f'{j}:0:{q}':f"{json[i][j][0][q]}"})
                            # print(i, ':', j ,':', 0,  ':',q,  ':', json[i][j][0][q])
                        elif type(json[i][j][0][q]) is dict:
                            for dict2 in json[i][j][0][q]:
                                new_json.update({f'{j}:0:{q}:{dict2}':f"{json[i][j][0][q][dict2]}"})
                                # print(i, ':', j ,':', 0,  ':',q,  ':', dict2, ":", json[i][j][0][q][dict2])
                        elif type(json[i][j][0][q]) is list:
                            for list2 in range (len(json[i][j][0][q])):
                                for dict3 in json[i][j][0][q][list2]:
                                    if type(json[i][j][0][q][list2][dict3]) is dict:
                                        for dict4 in json[i][j][0][q][list2][dict3]:
                                            new_json.update({f'{j}:0:{q}:{list2}:{dict3}:{dict4}':f"{json[i][j][0][q][list2][dict3][dict4]}"})
                                            # print(i, ':', j ,':', 0,  ':',q,  ':', list2, ":", dict3, ":", dict4, ":", json[i][j][0][q][list2][dict3][dict4])
                                    else:
                                        new_json.update({f'{j}:0:{q}:{list2}:{dict3}':f"{json[i][j][0][q][list2][dict3]}"})
                                        # print(i, ':', j ,':', 0,  ':',q,  ':', list2, ":", dict3, ":", json[i][j][0][q][list2][dict3])
            elif type(json[i][j]) is dict:
                for dict3 in json[i][j]:
                    if type(json[i][j][dict3]) is list:
                        for dict5 in json[i][j][dict3][0]:
                            if type(json[i][j][dict3][0]) is str:
                                new_json.update({f'{j}:{dict3}:{0}':f"{json[i][j][dict3][0]}"})
                                # print(i, ':', j ,':',dict3,':', '0 :',json[i][j][dict3][0])
                            elif type(json[i][j][dict3][0]) is dict:
                                new_json.update({f'{j}:{dict3}:{0}:{dict5}':f"{json[i][j][dict3][0][dict5]}"})
                                # print(i, ':', j ,':',dict3,':', '0 :', dict5, json[i][j][dict3][0][dict5])
                    else:
                        new_json.update({f'{j}:{dict3}':f"{json[i][j][dict3]}"})
                        # print(i, ':', j ,':',dict3,':',json[i][j][dict3])
            else:
                new_json.update({f'{j}':f"{json[i][j]}"})
                # print(i, ':', j ,':',json[i][j])
        new_array_full.append(new_json)
    return(new_array_full)

# WRITE DATAS TO EXCEL
def writeexcel(json, date):
    json.reset_index(inplace=True, drop=True)
    json.to_excel(writer, sheet_name = f'{date}')
    print ('Данные успешно занесены в файл ! \n')



# GET DATAS FROM AMADEUS API; request parameters can be modified according to amadeus api docs
def offersquery(origin, destination, departuredate):
    try:
        response = amadeus.shopping.flight_offers_search.get(
            originLocationCode = origin,
            destinationLocationCode = destination,
            departureDate = departuredate,
            adults = 1,
            travelClass = 'BUSINESS',
            nonStop = 'true'
            )
        return(response.data)
    except ResponseError as error:
        print(error) 
        return None

# MAIN FUNCTION
if __name__ == '__main__':
    a = True
    while a:
        inputdate = input('Введите дату для поиска авиабилетов в формате *гггг-мм-дд*: ')
        fileroute = os.path.abspath('Отправления_Прибытия.xlsx')
        time = str(datetime.now())
        time = time.replace(".","_")
        time = time.replace(" ","_")
        time = time.replace(":","-")
        orig,dest = inputroute(fileroute)       
        writer = pd.ExcelWriter(f'{time}.xlsx', engine = 'xlsxwriter')
        gen_json = pd.DataFrame()
        for i in range (len(orig)):
            if orig[i] != dest[i] or orig[i] != 'nan' or dest[i] != 'nan': 
                datas = offersquery(orig[i],dest[i],inputdate)
                print('Обработка маршрута: ', orig[i], ' -> ', dest[i], '\n')
                if datas != None:
                    route = orig[i] + '-' + dest[i]
                    jsonExcel = (json_normalize(format_json(datas, route, inputdate)))
                    gen_json = pd.concat([gen_json, pd.DataFrame(jsonExcel)])               
                else:
                    print('Авиабилеты не найдены!')
            else: continue
        writeexcel(gen_json, inputdate)
        writer.close()
        if input('Если хотите продолжить, наберите "Да", для выхода нажмите ENTER: ') != 'Да':
            a = False