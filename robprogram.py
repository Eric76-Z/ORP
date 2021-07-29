# 用于提取某压缩包中某文件内容，并解析
import os
import shutil
import time
import zipfile

import xlrd
import xlwt
from xlutils.copy import copy

from pathmap import pathmap
# 预设变量
BASE_PATH = 'D:\\rob'
PATHRB = 'D:\\rob\\old'  #原备份所在文件夹
PATH_EXPORT = 'D:\\rob\\new'  #重整后文件夹位置

buStandard = {
    'filepath': BASE_PATH,
    'filename': 'BackUpStandard.xls'
}
# 全局变量
targetData = {}

# RobotDatas = {}
# mainData = []
# ExcleDatas = []
# rootlists = []
# dirlists = []


# areakeyword = {
# }
#
# targetFile = 'RobotInfo.xml'
# zipfileTarget = 'C/KRC/Roboter/Rdc/RobotInfo.xml'
# extractTo = 'unzip\\'
# extractedFileList = []  # 列表，储存解压后文件路径
#
# ExclePath = '机器人Rdc数据表V4.xls'


# 获取zip格式文件路径路径
def getZipInfo():
    for root, dirs, files in os.walk(PATHRB):
        for name in files:
            if name.endswith('.zip'):
                originpath = os.path.join(root, name)
                controllername = name.split('.zip')[0]
                # print(controllername[0][-7:])
                workstationname = controllername[-7:]  # 截取的 eg.k2a3a131s460r04 后7位
                targetData[controllername] = {}
                targetData[controllername]['originpath'] = originpath
                targetData[controllername]['controller'] = controllername
                targetData[controllername]['workstation'] = workstationname
                if controllername[0:2] == 'k1':
                    targetData[controllername]['Lv1'] = 'CPH2.1'
                elif controllername[0:2] == 'k2':
                    targetData[controllername]['Lv1'] = 'CPH2.2'
                else:
                    targetData[controllername]['errmsg'] = '没有对应一级地点'
                    raise ('没有对应一级地点')
                if pathmap[controllername[2:6]]:
                    targetData[controllername]['Lv2'] = pathmap[controllername[2:6]]['Lv2']
                    targetData[controllername]['Lv3'] = pathmap[controllername[2:6]]['Lv3']
                else:
                    targetData[controllername]['errmsg'] = '找不到对应区域'
                    raise (controllername + '找不到对应区域')
                newpath = PATH_EXPORT + '\\' + targetData[controllername]['Lv1'] + '\\' + targetData[controllername][
                    'Lv2'] + '\\' + targetData[controllername]['Lv3']
                targetData[controllername]['newpath'] = newpath

                # 移动重整文件
                folder = os.path.exists(targetData[controllername]['newpath'])
                if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
                    os.makedirs(targetData[controllername]['newpath'])  # makedirs 创建文件时如果路径不存在会创建这个路径
                else:
                    pass
                if not os.path.exists(targetData[controllername]['newpath'] + '\\' + name):
                    print(targetData[controllername]['newpath'] + '\\' + controllername)
                    # shutil.move(targetData[controllername]['originpath'], targetData[controllername]['newpath'])
                    shutil.copy2(targetData[controllername]['originpath'], targetData[controllername]['newpath'])
                else:
                    # shutil.move(targetData[controllername]['originpath'],
                    #             targetData[controllername]['newpath'] + '\\' + name + '副本')
                    shutil.copy2(targetData[controllername]['originpath'],
                                 targetData[controllername]['newpath'] + '\\' + name + '副本')


# def mapControllername(controllername):
#
#
#         # 写到这


def backupState():

    standard_filepath = buStandard['filepath'] + '\\' + buStandard['filename']
    book_rd = xlrd.open_workbook(standard_filepath, formatting_info=True)
    sheet_rd = book_rd.sheet_by_index(0)

    book_wt = copy(book_rd)
    worksheet_name = 'rob'
    sheet_wt = book_wt.get_sheet(worksheet_name)
    nrows = sheet_rd.nrows
    for root, dirs, files in os.walk(PATH_EXPORT):
        for name in files:
            controllername = name.split('.zip')[0]

            flag = True
            for i in range(nrows):
                lng = sheet_rd.cell(i - 1, 1).value
                if lng == controllername:
                    sheet_wt.write(i, 3, label='已备份')
                    sheet_wt.write(i, 4, label=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                    flag = False
                    break
            if flag == True:
                nrows_new = sheet_rd.nrows  # 行
                print(nrows_new)
                sheet_wt.write(nrows_new, 1, controllername)
                sheet_wt.write(nrows_new, 3, label='新工位')
                sheet_wt.write(nrows_new, 4, label=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    book_wt.save(BASE_PATH + '\\' + time.strftime("%Y%m%d", time.localtime()) + buStandard['filename'])


# # 解压某文件到指定文件夹
# def extractFile():
#     i = 0
#     for filepath in filepaths:
#         try:
#             src = 'unzip/C/KRC/Roboter/Rdc/RobotInfo.xml'
#             dst = 'unzip/RobotInfo/' + workstationlists[i] + '-' + targetFile
#             extractedFileList.append(dst)
#             f = zipfile.ZipFile(filepath, 'r')
#             f.extract(member=zipfileTarget, path=extractTo, )
#             os.rename(src, dst)
#         except:
#             print(filepath)
#         i = i + 1
#     # print(extractedFile)
#
#     # zippathlist = f.namelist()  #['KRC/R1/Folgen/cell.src', 'KRC/R1/Folgen/folge123.dat', 'KRC/R1/Folgen/folge123.src', 'KRC/R1/Folgen/folge124.dat', .....
#     # print(zippathlist)
#
#     # for zippath in zippathlist:
#     #     if zippath == zipfileTarget:
#     #         f.extract(member=zippath, path=extractTo)
#     #
#
#
# def getfileData():
#     RobotData = []
#     RobotType = []
#     SeriaNum = []
#     print(exportfileData())
#
#     for extractedFile in extractedFileList:
#         try:
#             dom = minidom.parse(extractedFile)
#             # 解析xml
#             # 得到元素对象
#             root = dom.documentElement
#
#             msg = root.getElementsByTagName('RobotData')
#             attribute = msg[0].getAttribute('Timestamp')
#             RobotData.append(attribute)
#
#             msg = root.getElementsByTagName('SerialNumber')
#             SeriaNum.append(msg[0].firstChild.data)
#
#             msg = root.getElementsByTagName('RobotType')
#             RobotType.append(msg[0].firstChild.data)
#         except:
#             RobotData.append('blank')
#             SeriaNum.append('blank')
#             RobotType.append('blank')
#
#     print(RobotData)
#
#     for i in range(0, len(filepaths)):
#         RobotDatas = {
#             'filepath': filepaths[i],
#             'controller': controllerlists[i],
#             'workstation': workstationlists[i],
#             'extractedFile': extractedFileList[i],
#             'RobotDate': RobotData[i],
#             'SeriaNum': SeriaNum[i],
#             'RobotType': RobotType[i],
#         }
#         DATAS = {
#             '序号': i + 1,
#             '机器人名字': controllerlists[i],
#             '工位号': workstationlists[i],
#             '投入运行时间': RobotData[i],
#             '机器人序列号': SeriaNum[i],
#             '机器人类型': RobotType[i],
#         }
#         mainData.append(RobotDatas)
#         ExcleDatas.append(DATAS)
#     print(ExcleDatas)
#
#
# def exportfileData():
#     book = xlwt.Workbook(encoding="utf-8")  # 创建workbook对象
#     sheet = book.add_sheet('机器人Rdc数据表', cell_overwrite_ok=True)
#     col = ("序号", "机器人名字", "工位号", "投入运行时间", "机器人序列号", "机器人类型")
#     for i in range(0, len(col)):
#         sheet.write(0, i, col[i])
#         for j in range(0, len(mainData)):
#             # print()
#             sheet.write(j + 1, i, ExcleDatas[j][col[i]])
#     book.save(ExclePath)  # 保存数据


def main():
    # age = input('输入参数：')
    # print(age)
    # 1.获取路径
    getZipInfo()
    #
    backupState()

    # # 2.解压某文件到指定文件夹
    # extractFile()
    #
    # # 3.解析文件，提取数据
    # getfileData()
    #
    # # 4.导出数据
    # exportfileData()


if __name__ == "__main__":
    main()
