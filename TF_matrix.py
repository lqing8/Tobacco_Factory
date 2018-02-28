#coding=utf-8
from win32com.client import Dispatch
import win32com.client
import os
import http.client
import json
from urllib.parse import quote_plus

class easyExcel:
    """A utility to make it easier to get at Excel.    Remembering
    to save the data is your problem, as is    error handling.
    Operates on one workbook at a time."""

    def __init__(self, filename=None):  # 打开文件或者新建文件（如果不存在的话）
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        self.xlApp.Visible = True
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def save(self, newfilename=None):  # 保存文件
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):  # 关闭文件
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def getCell(self, sheet, row, col):  # 获取单元格的数据
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def setCell(self, sheet, row, col, value):  # 设置单元格的数据
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def setCellformat(self, sheet, row, col):  # 设置单元格的数据
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Font.Size = 15  # 字体大小
        sht.Cells(row, col).Font.Bold = True  # 是否黑体
        sht.Cells(row, col).Name = "Arial"  # 字体类型
        sht.Cells(row, col).Interior.ColorIndex = 3  # 表格背景
        # sht.Range("A1").Borders.LineStyle = xlDouble
        sht.Cells(row, col).BorderAround(1, 4)  # 表格边框
        sht.Rows(3).RowHeight = 30  # 行高
        sht.Cells(row, col).HorizontalAlignment = -4131  # 水平居中xlCenter
        sht.Cells(row, col).VerticalAlignment = -4160  #

    def deleteRow(self, sheet, row):
        sht = self.xlBook.Worksheets(sheet)
        sht.Rows(row).Delete()  # 删除行
        sht.Columns(row).Delete()  # 删除列

    def getRange(self, sheet, row1, col1, row2, col2):  # 获得一块区域的数据，返回为一个二维元组
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):  # 插入图片
        "Insert a picture in sheet"
        sht = self.xlBook.Worksheets(sheet)
        sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

    def cpSheet(self, before):  # 复制工作表
        "copy sheet"
        shts = self.xlBook.Worksheets
        shts(1).Copy(None, shts(1))

def geocode(address, key):
    #key = '0b00174f6f8ab4ca8d350ac0da105bb9'
    key = '389880a06e3f893ea46036f030c94700'
    #key = 'ee0c2ec9cd719c1c0adaef80f89b5aa8'
    #key = '22d3816e107f199992666d6412fa0691'
    #key = '837a9bdb426d81b6862135983d1d715c'
    #key = '608d75903d29ad471362f8c58c550daf'
    try:
        base = '/v3/geocode/geo'
        path = '{}?address={}&key={}'.format(base, quote_plus(address), key)
        #print(path)
        connection = http.client.HTTPConnection('restapi.amap.com',80)
        connection.request('GET', path)
        rawreply = connection.getresponse().read()
        #print(rawreply)
        reply = json.loads(rawreply.decode('utf-8'))
        print(address + '的经纬度：',reply['geocodes'][0]['location'])
        return reply['geocodes'][0]['location']
    except:
        print('geocode error')

def getDistance(startLonLat, endLonLat, endString, key):
    provincial_capital = ['北京', '天津', '重庆', '上海', '石家庄', '沈阳', '哈尔滨', '杭州', '福州', '济南', '广州', '武汉', '成都', '昆明', '兰州', '台北', '南宁', '银川', '太原',
     '长春', '南京', '合肥', '南昌', '郑州', '长沙', '海口', '贵阳', '西安', '西宁', '呼和浩特', '拉萨', '乌鲁木齐', '澳门', '香港']
    x = None
    if(startLonLat==endLonLat):
        x = 30
        return x
    else:
        try:
            duration = 0;  # 起始地与目的地之间的距离
            # path = '{}?key={}&origins={}&destination={}'.format('http://restapi.amap.com/v3/distance',key,startLonLat,endLonLat)
            path='http://restapi.amap.com/v3/distance?key={}&origins={}&destination={}'.format(key,startLonLat,endLonLat)
            #path = 'http://restapi.amap.com/v3/direction/driving?key={}&origin={}&destination={}'.format(key, startLonLat,endLonLat)
            connection = http.client.HTTPConnection('restapi.amap.com', 80)
            connection.request('GET', path)
            rawreply = connection.getresponse().read()
            # print(rawreply)
            reply = json.loads(rawreply.decode('utf-8'))
            # print(reply['results'][0]['distance'])
            x = int(reply['results'][0]['distance'])/1000
        except:
            print('getDistance error')

    if(endString in provincial_capital): duration = 30
    else: duration = 10
    x = x + duration
    return x

def setDistance(rowS, colD, D, xls, key):
    rowCount = xls.xlBook.Worksheets('base').UsedRange.Rows.Count
    colCount = xls.xlBook.Worksheets('base').UsedRange.Columns.Count

    print(rowS + ',' + colD)
    _col = 5
    while(_col <= colCount):
        src = xls.getCell('base', 2, _col)
        if (rowS == src):
            row = 3
            while(row <= rowCount):
                des = xls.getCell('base', row, 4)
                if(colD == des):
                    return xls.getCell('base',row, _col)
                row = row + 1
        _col = _col + 1
    #return ''
    x = getDistance(geocode(rowS, key), geocode(colD,key), D, key)
    return x

#下面是一些测试代码。
if __name__ == "__main__":
    #key = 'cb649a25c1f81c1451adbeca73623251'     #this key good
    key = '36280aad084f5aa954f04ffe8adc4a20' #my key
    #key = '0b00174f6f8ab4ca8d350ac0da105bb9'
    #key = '389880a06e3f893ea46036f030c94700'
    #key = 'ee0c2ec9cd719c1c0adaef80f89b5aa8'
    #key = '22d3816e107f199992666d6412fa0691'
    #key = '837a9bdb426d81b6862135983d1d715c'
    #key = '608d75903d29ad471362f8c58c550daf'
    #key = '6119e85defa6a97be090a0af41f0613c7'
    # x = getDistance(geocode('合肥', key), geocode('绵阳', key), '绵阳', key)
    #x = geocode('重庆重庆', key)

    continueNum = 1
    ########### get base Excel and 转化 ###########
    xls = easyExcel(os.getcwd() + '\\TF_base.xlsx')
    rowCount = xls.xlBook.Worksheets('base').UsedRange.Rows.Count
    colCount = xls.xlBook.Worksheets('base').UsedRange.Columns.Count

    ########### save result ###########
    filename = os.getcwd() + '\\TF_test.xlsx'
    xls_1 = easyExcel(filename)
    try:
        row = 46
        CN = (row - 3)*344
        continueNum = CN + 1
        while(row<= rowCount):
            p = xls.getCell('base', row, 2)
            s = xls.getCell('base', row, 4)
            xls_1.setCell('sheet1', row -1, 1, p)
            xls_1.setCell('sheet1', row -1, 2, xls.getCell('base', row, 3))
            xls_1.setCell('sheet1', row -1, 3, s)

            xls_1.setCell('sheet1', 1, row + 1, s)

            col = 3
            while (col<= rowCount):
                p1 =  xls.getCell('base', col, 2)
                d = xls.getCell('base', col, 4)
                #xls_1.setCell('sheet1', row -1 , col + 1, s + ',' + d)
                if(continueNum > CN):
                    if(xls_1.getCell('sheet1', row - 1, col + 1) == None):
                        xls_1.setCell('sheet1', row - 1, col + 1, setDistance(p + s,p1 + d, d, xls, key))
                print(continueNum)
                continueNum = continueNum + 1
                col = col + 1

            row = row + 1
    except BaseException as e:
        print(e)
    finally:
        print(continueNum)
        xls_1.save(filename)
        xls.close()
        xls_1.close()
    '''
        while(col<= len(area)):
            y = area[row -1]
            x = area[col -1]
            v = y + ',' + x
            #rs = searchCell(y,x)
            cv = xls.getCell('sheet1', row + 1, col + 1)
            col += 1
            print(v, cv)
    '''
'''
    ########## 高德地图 #########
    key = 'cb649a25c1f81c1451adbeca73623251'

    filename = os.getcwd() + '\\test.xls'
    #os.remove(filename)
    xls = easyExcel()
    #xls.addPicture('Sheet1', PNFILE, 20,20,1000,1000)
    #xls.cpSheet('Sheet1')
    xls.setCell('sheet1',1,1,'')
    row=1
    print("*******beginsetCellformat********")
    x,y = '',''
    while(row<= len(area)):
        col = 1
        while(col<= len(area)):
            x = areacode[col -1]
            y = areacode[row -1]
            v = getDistance(geocode(x,key),geocode(y, key), key)
            if (col == 1):
                xls.setCell('sheet1', row + 1, col, area[row -1])
            if(row == 1):
                xls.setCell('sheet1',row,col + 1,area[col -1])

            xls.setCell('sheet1', row + 1, col + 1, v)
            col += 1
            print("row=%s,col=%s:" %(row,col))
        row += 1
    print("*******endsetCellformat********")
    xls.save(filename)
    xls.close()
'''