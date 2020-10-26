import json
import xlwt
import math


def f_piancha(x, mu, simga):
    return (x-mu)/simga


workbook = xlwt.Workbook()
sheet_dao = workbook.add_sheet('每刀偏差值')
sheet_ren = workbook.add_sheet('每人偏差值')
times = [0, 0, 0, 0, 0]
sum_damage = [0, 0, 0, 0, 0]
variance = [0, 0, 0, 0, 0]
with open('下载.json', 'r') as f:
    data = json.loads(f.read())
    # print((data['challenges'][0]['qqid']))
    for i in data['members']:
        for i1 in range(1, 11):#前五个为对对应boss造成的伤害,后五个为对对应boss的出刀数
            i[i1] = 0

    for i in data['challenges']:
        if i['cycle'] != 1:
            if i['health_ramain'] != 0:
                if i['is_continue'] == False:
                    if i['damage'] !=0:
                        times[i['boss_num']-1] += 1
                        sum_damage[i['boss_num']-1] += i['damage']
    for i in range(5):
        sum_damage[i] = sum_damage[i]/times[i]
    for i in data['challenges']:
        if i['cycle'] != 1:
            if i['health_ramain'] != 0:
                if i['is_continue'] == False:
                    if i['damage'] !=0:
                        variance[i['boss_num'] - 1] += (
                            i['damage']-sum_damage[i['boss_num']-1])**2
    for i in range(5):
        variance[i] = math.sqrt(variance[i]/(times[i]-1))
    hang = 0
    for i in data['challenges']:
        if i['cycle'] != 1:
            if i['health_ramain'] != 0:
                if i['is_continue'] == False:
                    for i1 in data['members']:
                        if i1['qqid'] == i['qqid']:
                            fx = f_piancha(
                                i['damage'], sum_damage[i['boss_num']-1], variance[i['boss_num']-1])
                            sheet_dao.write(hang, 0, i1['nickname'])
                            sheet_dao.write(hang, 1, fx)
                            sheet_dao.write(hang, 2, i['damage'])
                            sheet_dao.write(hang, 3, i['boss_num'])
                            i1[i['boss_num']] += fx
                            i1[i['boss_num']+5]+=1
                            hang += 1
                            break
    hang=0
    for i in data['members']:
        total_dao=0
        total_damage=0
        sheet_ren.write(hang, 0, i['nickname'])
        for i1 in range(1,6):
            if i[i1+5]!=0:
                sheet_ren.write(hang, i1, i[i1]/i[i1+5])
                total_dao+=i[i1+5]
                total_damage+=i[i1]
        if total_dao!=0:
            sheet_ren.write(hang, 6, total_damage/total_dao)
        hang+=1
workbook.save('会战分析.xls')
