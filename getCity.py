# -*- coding: utf-8 -*-
import requests, json

def getCities():
    url = 'http://cmshow.qq.com/qqshow/admindata/comdata/vipApollo_clueCity_cityID/xydata.json'
    s = requests.session()
    c = s.get(url).content
    c = json.loads(c)
    cities = c['data']['cityId']
    # for k in cities: k['city'] is name and k['id'] is id
    return cities

def getPointsByCityID(cityID):
    url = 'http://cmshow.qq.com/qqshow/admindata/comdata/vipApollo_clueCity_'+str(cityID)+'/xydata.json'
    s = requests.session()
    c = s.get(url).content
    c = json.loads(c)
    points = c['data']['cityPlayground']
    '''
    An example of a point in points
    "id": 1001,
    "name": "上海市",
    "location": "31.32784,121.35119",
    "address": "宝山区顾村镇海蓝学校",
    "styleName": "Eva扭蛋机",
    '''
    return points

def saveXls(allpoints):
    import xlwt
    wb= xlwt.Workbook(encoding = 'utf-8')
    ws= wb.add_sheet('cmshow')

    ws.write(0, 0, '城市')
    ws.write(0, 1, '位置')
    ws.write(0, 2, '坐标')
    ws.write(0, 3, '类别')

    linecnt = 1
    for city in allpoints:
        points = allpoints[city]
        for point in points:
            ws.write(linecnt, 0, city)
            ws.write(linecnt, 1, point['address'])
            ws.write(linecnt, 2, point['location'])
            ws.write(linecnt, 3, point['styleName'])
            linecnt += 1

    wb.save('cmshow.xls')
    print '文件已保存为cmshow.xls'


def main():
    cities = getCities()
    allpoints = {}
    print '获取到的城市列表'
    for i in cities:
        print 'No.' + str(i['id']) + ' ' + i['city']
        allpoints[i['city']] = getPointsByCityID(i['id'])
    saveXls(allpoints)

if __name__ == '__main__':
    main()
