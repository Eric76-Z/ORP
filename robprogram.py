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
BASE_PATH = 'F:\\New folder'
PATHRB = BASE_PATH + '\\' + 'old'  # 原备份所在文件夹
PATH_EXPORT = BASE_PATH + '\\' + 'new'  # 重整后文件夹位置
LOG_FILE_NAME = 'log.txt'

buStandard = {
    'filepath': BASE_PATH,
    'filename': 'BackUpStandard.xls'
}
# 全局变量
targetData = {}
isAll = False  # True:全名对比， False后七位对比，eg:S240R01

# 表格
SHEET_NAME = 'robot'
COMPARE_COL = 5
COMMIT_COL = 6
TIME_COL = 7
Lv3 = 4
Lv2 = 3
Lv1 = 2

# 总结
TOTAL_FILES = 0
DEAL_FILES = 0
ERR_FILES = 0


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


def logWrite(controllername, msg):
    log = open(BASE_PATH + '\\' + LOG_FILE_NAME, 'a')
    log.write(controllername + '[' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ']' + ' :' + msg + '\r\n')
    log.close()


def logWriteTitle(msg):
    log = open(BASE_PATH + '\\' + LOG_FILE_NAME, 'a')
    log.write('==============================' + msg + '==============================' + '\r\n')
    log.close()


# 获取zip格式文件路径路径
def getZipInfo():
    global TOTAL_FILES

    global ERR_FILES
    for root, dirs, files in os.walk(PATHRB):
        for name in files:
            TOTAL_FILES = TOTAL_FILES + 1
            if name.endswith('.zip'):
                originpath = os.path.join(root, name)
                controllername = name.split('.zip')[0]
                workstationname = controllername[-7:].upper()  # 截取的 eg.k2a3a131s460r04 后7位
                # print(workstationname)
                if isAll == True:
                    compare_name = controllername
                else:
                    compare_name = workstationname
                targetData[compare_name] = {}
                targetData[compare_name]['originpath'] = originpath
                targetData[compare_name]['controller'] = compare_name
                targetData[compare_name]['workstation'] = workstationname
                if controllername[0:2] == 'k1':
                    targetData[compare_name]['Lv1'] = 'CPH2.1'
                elif controllername[0:2] == 'k2':
                    targetData[compare_name]['Lv1'] = 'CPH2.2'
                elif controllername[0:2] == 'k3':
                    targetData[compare_name]['Lv1'] = 'CPH2.1'
                    logWrite(controllername, '一级地点为k3')
                else:
                    # print(compare_name)
                    # targetData[compare_name]['errmsg'] = '没有对应一级地点'
                    ERR_FILES = ERR_FILES + 1
                    logWrite(controllername, '没有对应一级地点')
                    continue
                if controllername[2:6] in pathmap:
                    targetData[compare_name]['Lv2'] = pathmap[controllername[2:6]]['Lv2']
                    targetData[compare_name]['Lv3'] = pathmap[controllername[2:6]]['Lv3']
                else:
                    logWrite(controllername, '找不到对应区域')
                    ERR_FILES = ERR_FILES + 1
                    continue
                    # targetData[compare_name]['errmsg'] = '找不到对应区域'
                    # raise (compare_name + '找不到对应区域')
                newpath = PATH_EXPORT + '\\' + targetData[compare_name]['Lv1'] + '\\' + targetData[compare_name][
                    'Lv2'] + '\\' + targetData[compare_name]['Lv3']
                targetData[compare_name]['newpath'] = newpath

                # 移动重整文件
                folder = os.path.exists(targetData[compare_name]['newpath'])
                if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
                    os.makedirs(targetData[compare_name]['newpath'])  # makedirs 创建文件时如果路径不存在会创建这个路径
                else:
                    pass
                if not os.path.exists(targetData[compare_name]['newpath'] + '\\' + name):
                    # print(targetData[compare_name]['newpath'] + '\\' + compare_name)
                    # shutil.move(targetData[compare_name]['originpath'], targetData[compare_name]['newpath'])
                    shutil.copy2(targetData[compare_name]['originpath'], targetData[compare_name]['newpath'])
                else:
                    # shutil.move(targetData[compare_name]['originpath'],
                    #             targetData[compare_name]['newpath'] + '\\' + name + '副本')
                    # shutil.copy2(targetData[compare_name]['originpath'],
                    #              targetData[compare_name]['newpath'] + '\\' + name + '副本')
                    continue


def backupState():
    standard_filepath = buStandard['filepath'] + '\\' + buStandard['filename']
    book_rd = xlrd.open_workbook(standard_filepath, formatting_info=True)
    sheet_rd = book_rd.sheet_by_index(0)
    book_wt = copy(book_rd)
    sheet_wt = book_wt.get_sheet(SHEET_NAME)
    nrows = sheet_rd.nrows
    nrows_compare = nrows
    global DEAL_FILES
    for root, dirs, files in os.walk(PATH_EXPORT):
        for name in files:
            DEAL_FILES = DEAL_FILES + 1
            controllername = name.split('.zip')[0]
            workstationname = controllername[-7:].upper()  # 截取的 eg.k2a3a131s460r04 后7位
            if isAll == True:
                compare_name = controllername
            else:
                compare_name = workstationname
            flag = True

            for i in range(nrows):
                lng = sheet_rd.cell(i - 1, COMPARE_COL).value
                if lng == compare_name:
                    sheet_wt.write(i - 1, COMMIT_COL, label='已备份')
                    sheet_wt.write(i - 1, TIME_COL, label=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                    flag = False
                    break
            if flag == True:
                sheet_wt.write(nrows_compare, COMPARE_COL - 3, targetData[compare_name]['Lv1'])
                sheet_wt.write(nrows_compare, COMPARE_COL - 2, pathmap[controllername[2:6]]['Lv2'])
                sheet_wt.write(nrows_compare, COMPARE_COL - 1, pathmap[controllername[2:6]]['Lv3'])
                sheet_wt.write(nrows_compare, COMPARE_COL, compare_name)
                sheet_wt.write(nrows_compare, COMMIT_COL, label='新工位')
                sheet_wt.write(nrows_compare, TIME_COL, label=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                nrows_compare = nrows_compare + 1
    logWriteTitle('总结')
    log = open(BASE_PATH + '\\' + LOG_FILE_NAME, 'a')
    log.write('备份总数: ' + str(TOTAL_FILES) + '        已处理: ' + str(DEAL_FILES) + '         异常: ' + str(ERR_FILES) + '\r\n')
    log.close()
    logWriteTitle('end')
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
    logWriteTitle('start')
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
