# -*- coding: utf-8 -*-

import os
import re
import time
from logging import getLogger, StreamHandler, DEBUG

import openpyxl
import requests
from bs4 import BeautifulSoup

logger = getLogger(__name__)
handler = StreamHandler()
handler.setLevel(DEBUG)
logger.setLevel(DEBUG)
logger.addHandler(handler)
logger.propagate = False

# logger.debug('hello')

path = os.getenv("HOMEDRIVE", "None") + os.getenv("HOMEPATH", "None") + "/Dropbox/POG/"

wbpath = (path + "POG_HorseList.xlsx").replace("\\", "/")
htmlpath = (path + "PO_race_horse_list.html").replace("\\", "/")

wb = openpyxl.load_workbook(wbpath)
wshl = wb["POHorseList"]
wsSettings = wb["Settings"]

age = wsSettings["B1"].value
login_id = wsSettings["B2"].value
password = wsSettings["B3"].value

LOGIN_INFO = {
    'pid': 'login',
    'action': 'auth',
    'return_url2': '',
    'mem_tp': '',
    'login_id': login_id,
    'pswd': password,
    'auto_login': ''
}

mysession = requests.Session()
login_url = "https://regist.netkeiba.com/account/"
time.sleep(1)
mypost = mysession.post(login_url, data=LOGIN_INFO)

trow = 1

while wshl.cell(row=trow, column=1).value is not None:
    # print(trow)

    horseNm = wshl.cell(row=trow, column=2).value
    horseNmOrgn = wshl.cell(row=trow, column=3).value

    isHorseNmDtrmnd = False

    if len(horseNm) < 6:
        isHorseNmDtrmnd = True
    elif horseNm[-5] != "の":
        isHorseNmDtrmnd = True

    if isHorseNmDtrmnd and horseNmOrgn is not None:
        trow += 1
        continue

    horseURLsp = wshl.cell(row=trow, column=5).value
    if horseURLsp is None:
        trow += 1
        continue

    time.sleep(1)
    r = mysession.get(horseURLsp)
    soup = BeautifulSoup(r.content, 'lxml')

    horseNmNew = soup.find("p", class_="Name").string
    horseNmOrgnNew = soup.find("th", string="馬名の意味").find_next().string

    wshl.cell(row=trow, column=2).value = horseNmNew
    if horseNmOrgnNew != "-":
        wshl.cell(row=trow, column=3).value = horseNmOrgnNew

    trow += 1

POHList = [[cell.value for cell in row] for row in wshl["A1:F" + str(trow - 1)]]

RHList = []

target_url = 'http://race.netkeiba.com/?rf=navi'
r = mysession.get(target_url)  # requestsを使って、webから取得
soup = BeautifulSoup(r.text, 'lxml')  # 要素を抽出

DateList = soup.find('div', class_='DateList_Box')

for DateItem in DateList.find_all('a'):

    if DateItem.get('href').split('=')[-1][0] in 'np':
        continue

    race_mmdd = DateItem.get('href').split('=')[-1][1:3] + "/" + DateItem.get('href').split('=')[-1][3:5]

    target_url = 'http://race.netkeiba.com' + DateItem.get('href')
    time.sleep(1)
    r = mysession.get(target_url)  # requestsを使って、webから取得
    soup = BeautifulSoup(r.content, 'lxml')  # 要素を抽出

    prev_target_url = None

    for race_url in soup.find_all('a'):

        # logger.info(str(race_url.get("href").split("=")[1])[0:4])

        if not race_url.get("href").startswith("/?pid"):
            break

        if len(race_url.get("href").split("=")[2].split("&")) > 1:
            raceID = race_url.get("href").split("=")[2].split("&")[0]
        else:
            raceID = race_url.get("href").split("=")[2]

        if raceID[0] != "c":
            continue

        race_date = raceID[1:5] + "/" + race_mmdd
        if raceID[5:7] == "01":
            track = "札幌"
        elif raceID[5:7] == "02":
            track = "函館"
        elif raceID[5:7] == "03":
            track = "福島"
        elif raceID[5:7] == "04":
            track = "新潟"
        elif raceID[5:7] == "05":
            track = "東京"
        elif raceID[5:7] == "06":
            track = "中山"
        elif raceID[5:7] == "07":
            track = "中京"
        elif raceID[5:7] == "08":
            track = "京都"
        elif raceID[5:7] == "09":
            track = "阪神"
        elif raceID[5:7] == "10":
            track = "小倉"
        else:
            track = "根岸"

        if str(race_url.get("href").split("=")[1])[0:4] != "race":
            continue

        if race_url.get("title")[0:3] == "３歳上" or (age == 2 and race_url.get("title")[0:2] == "３歳"):
            continue

        target_url = 'http://race.netkeiba.com/?pid=race_old&id=' + raceID

        if target_url == prev_target_url:
            continue

        prev_target_url = target_url
        time.sleep(1)
        r = mysession.get(target_url)  # requestsを使って、webから取得
        soup = BeautifulSoup(r.content, 'lxml')  # 要素を抽出

        dt_list = soup.find_all('dt', limit=2)

        RaceNo = ("0" + dt_list[1].string.strip())[-3:]

        h1_list = soup.find_all('h1')
        RaceName = h1_list[1].contents[0].strip()
        # print(len(h1_list[1].contents))
        if len(h1_list[1].contents) > 1:
            GradeTemp = str(h1_list[1].contents[1])
            Grade = GradeTemp.split('_')[-2]
        else:
            Grade = ''

        RaceAtrb_list = h1_list[1].find_all_next('p', limit=4)
        course = RaceAtrb_list[0].string.strip()
        RaceTime = RaceAtrb_list[1].string[-5:]
        RaceCond1 = RaceAtrb_list[2].string
        RaceCond2 = RaceAtrb_list[3].string

        if len(RaceCond1.split()) > 1:
            if RaceCond1.split()[1][0:2] == "障害" or RaceCond1.split()[1][0:3] == "３歳上" or (
                    age == 2 and RaceCond1.split()[1][0:2] == "３歳"):
                continue

        # logger.debug("PASS1")

        h_list = soup.find_all(class_='bml1')

        for h in h_list:

            t = h.find('td', class_='umaban')
            if t is None:
                HorseNo = '00'
            else:
                HorseNo = ("0" + t.string)[-2:]

            t = h.find('td', class_=re.compile('^waku'))
            if t is None:
                Frame = '0'
            else:
                Frame = t.string

            HorseName = h.find('td', class_="txt_l horsename").find('div').find('a').string

            isFind = False
            for POHItem in POHList:
                # logger.info(HorseName + " " + POHItem[1])
                if HorseName == POHItem[1]:
                    owner = POHItem[0].strip()
                    isFind = True
                    origin = POHItem[2]
                    if POHItem[5] == "封印":
                        isSeal = True
                    else:
                        isSeal = False

            if not isFind:
                continue

            # logger.debug(h)

            HorseURL = h.find('td', class_="txt_l horsename").find('div').find('a').get('href')
            Weight = h.find('td', class_="txt_l horsename").find_next('td').find_next('td').string
            if not h.find_all('td', class_='txt_l', limit=2)[1].find('a'):
                Jockey = None
            else:
                Jockey = h.find_all('td', class_='txt_l', limit=2)[1].find('a').string
            if not h.find('td', class_='txt_r'):
                Odds = None
                PopRank = None
            else:
                Odds = h.find('td', class_='txt_r').string
                PopRank = h.find('td', class_='txt_r').find_next('td').string
            SortKey = race_date + RaceNo + RaceTime + HorseNo + HorseName

            # logger.debug(course)
            # noinspection PyUnboundLocalVariable,PyUnboundLocalVariable,PyUnboundLocalVariable
            RHList.append(
                [SortKey, race_date, RaceTime, track, RaceNo, RaceName, Grade, course, RaceCond1, RaceCond2, HorseNo, Frame,
                 HorseName, Jockey, Odds, PopRank, Weight, target_url, HorseURL, owner, origin,"00", isSeal])

RHList.sort()

f = open(htmlpath, mode="w")

prevDate = None
prev_race_no = None
prevRaceTime = None

for i in RHList:

    f.write("<!--" + str(i) + "-->\n")

    race_date = i[1]
    RaceTime = i[2]
    track = i[3]
    RaceNo = i[4]
    RaceName = i[5]
    GradeTemp = i[6]
    course = i[7].replace("\xa0", " ")
    RaceCond2 = i[9].replace("\xa0", " ")
    HorseNo = i[10]
    BoxNo = i[11]
    HorseName = i[12]
    Jockey = i[13]
    Odds = i[14]
    PopRank = i[15]
    RaceURL = i[17]
    HorseURL = i[18]
    owner = i[19]
    origin = i[20]
    isSeal = i[22]
    sp = ' '

    if GradeTemp == "g1":
        Grade = "[GI]"
    elif GradeTemp == "g2":
        Grade = "[GII}"
    elif GradeTemp == "g3":
        Grade = "[GIII]"
    elif GradeTemp == "jg1":
        Grade = "[JGI]"
    elif GradeTemp == "jg2":
        Grade = "[JGII]"
    elif GradeTemp == "jg3":
        Grade = "[JGIII]"
    elif GradeTemp == "op":
        Grade = "[OP]"
    else:
        Grade = ""

    if BoxNo == "1":
        Frame = '<span style="border: 1px solid; background-color:#ffffff; color:#000000;">1</span> '
    elif BoxNo == "2":
        Frame = '<span style="border: 1px solid; background-color:#000000; color:#ffffff;">2</span> '
    elif BoxNo == "3":
        Frame = '<span style="border: 1px solid; background-color:#ff0000; color:#ffffff;">3</span> '
    elif BoxNo == "4":
        Frame = '<span style="border: 1px solid; background-color:#0000ff; color:#ffffff;">4</span> '
    elif BoxNo == "5":
        Frame = '<span style="border: 1px solid; background-color:#ffff00; color:#000000;">5</span> '
    elif BoxNo == "6":
        Frame = '<span style="border: 1px solid; background-color:#00ff00; color:#ffffff;">6</span> '
    elif BoxNo == "7":
        Frame = '<span style="border: 1px solid; background-color:#ff8000; color:#000000;">7</span> '
    elif BoxNo == "8":
        Frame = '<span style="border: 1px solid; background-color:#ff8080; color:#000000;">8</span> '
    else:
        Frame = None

    s = '<h4>' + race_date + '</h4>\n'
    if prevDate is None:
        f.write(s)
    elif race_date != prevDate:
        f.write('</ul></li></ul>' + s)

    s = '<li> <a href="' + RaceURL + '">' + track + RaceNo + sp + RaceName + Grade + '</a><br />\n'
    s2 = RaceTime + sp + course + sp + RaceCond2 + '<br />\n<ul>'
    if prevDate is None or (prevDate is not None and race_date != prevDate):
        f.write('<ul>' + s)
        f.write(s2)
    elif race_date + RaceNo + RaceTime != prevDate + prev_race_no + prevRaceTime:
        f.write('</ul></li>' + s)
        f.write(s2)

    if isSeal:
        f.write('<li> <a href="' + HorseURL + '"><s>' + HorseName + '</s>' + owner + '</a> <br />\n')
    else:
        f.write('<li> <a href="' + HorseURL + '">' + HorseName + owner + '</a> <br />\n')
    f.write(origin + '<br />\n')
    if Odds is not None:
        f.write(str(Odds) + '倍' + sp + str(PopRank) + '番人気<br />\n')
    if HorseNo != "00":
        f.write(Frame + str(HorseNo) + '番' + sp + Jockey + '騎手<br />\n')
    elif Jockey is not None:
        f.write(Jockey + '騎手<br />\n')
    f.write('</li>\n')

    prevDate = race_date
    prev_race_no = RaceNo
    prevRaceTime = RaceTime

f.write('</ul></li></ul><p>終末オーナーLOVEPOP</p>\n')
f.write('<p>※オッズはnetkeibaより取得したものです。</p>')

f.close()

wb.save(wbpath)
