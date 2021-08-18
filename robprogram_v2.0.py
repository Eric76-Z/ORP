import json
import os
import shutil
import time
import xlrd
from xlutils.copy import copy
from pathmap import pathmap

# 配置变量
PATH_BASE = 'F:\\New folder'  # 根目录
PATH_ORIGIN_BACKUP = PATH_BASE + '\\' + 'old'  # 原备份所在文件夹
PATH_EXPORT_TO = PATH_BASE + '\\' + 'new'  # 重整后文件夹位置
LOG_FILE_NAME = 'log.txt'

buStandard = {
    'filepath': PATH_BASE,  # 存储在根目录
    'filename': '机器人备份状态.xls'  # 定义名字
}


# 全局变量
class RobProgramData:
    def __init__(self):
        self.sort = '机器人程序'
        self.path = {
            'path_origin': '',
            'path_new': ''
        }
        self.meta = {
            'title': '',
            'mtime': '',
            'format': '',
        }
        self.data = {
            'controllername': '',
            'workstationname': '',
        }

    def localLv1(self):
        if self.data['controllername'][0:2].lower() == 'k1':
            return 'CPH2.1'
        elif self.data['controllername'][0:2].lower() == 'k2':
            return 'CPH2.2'
        elif self.data['controllername'][0:2].lower() == 'k3':
            return 'CPH2.1'
        else:
            return ''

    def localLv2(self):
        if self.data['controllername'][2:6].lower() in pathmap:
            return pathmap[self.data['controllername'][2:6].lower()]['Lv2']
        else:
            return ''

    def localLv3(self):
        if self.data['controllername'][2:6].lower() in pathmap:
            return pathmap[self.data['controllername'][2:6].lower()]['Lv3']
        else:
            return ''


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
TOTAL_FILES = 0  # 更目录下文件总数
DEAL_FILES = 0
ERR_FILES = 0


def logWrite(controllername, msg):
    log = open(PATH_BASE + '\\' + LOG_FILE_NAME, 'a')
    log.write(controllername + '[' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ']' + ' :' + msg + '\r\n')
    log.close()


def logWriteTitle(msg):
    log = open(PATH_BASE + '\\' + LOG_FILE_NAME, 'a')
    log.write('==============================' + msg + '==============================' + '\r\n')
    log.close()


# 　　'''把时间戳转化为时间: 1479264792 to 2016-11-16 10:53:12'''
def TimeStampToTime(timestamp):
    timeStruct = time.localtime(timestamp)
    return time.strftime('%Y-%m-%d %H:%M:%S', timeStruct)


# 获取zip格式文件路径路径,并移动重整文件
def getZipInfo():
    rob_program_data = RobProgramData()
    global TOTAL_FILES
    global ERR_FILES
    for root, dirs, files in os.walk(PATH_ORIGIN_BACKUP):
        for name in files:
            TOTAL_FILES = TOTAL_FILES + 1
            if name.endswith('.zip'):

                originpath = os.path.join(root, name)
                # ================path==================
                rob_program_data.path['path_origin'] = originpath  # 原始路径
                # ================meta==================
                rob_program_data.meta['title'] = name  # 文件名
                rob_program_data.meta['format'] = name.split('.zip')[1]  # 文件后缀
                rob_program_data.meta['mtime'] = TimeStampToTime(os.path.getmtime(originpath))  # 原数据中修改时间，可视为创建时间
                # ================data==================
                rob_program_data.data['controllername'] = name.split('.zip')[0]  # eg.k2a3a131s460r04
                rob_program_data.data['workstationname'] = rob_program_data.data['controllername'][
                                                           -7:].upper()  # 截取的 eg.s460r04
                rob_program_data.data['localLv1'] = rob_program_data.localLv1()

                def localLv2(self):
                    if self.data['controllername'][2:6].lower() in pathmap:
                        return pathmap[self.data['controllername'][2:6].lower()]['Lv2']
                    else:
                        return ''

                json_str = json.dumps(rob_program_data, default=lambda obj: obj.__dict__)
                print(json_str)
                print(rob_program_data.localLv1())
                ERR_FILES = ERR_FILES + 1
                if rob_program_data.data['localLv1'] == '':
                    logWrite('【异常】' + rob_program_data.data['controllername'], '没有对应一级地点')
                    ERR_FILES = ERR_FILES + 1
                    continue

                elif rob_program_data.data['localLv2'] == '':
                    logWrite('【异常】' + rob_program_data.data['controllername'], '没有对应二级地点')
                    ERR_FILES = ERR_FILES + 1
                    continue
                elif rob_program_data.data['localLv3'] == '':
                    logWrite('【异常】' + rob_program_data.data['controllername'], '没有对应三级地点')
                    ERR_FILES = ERR_FILES + 1
                    continue
                rob_program_data.path['path_new'] = PATH_EXPORT_TO + '\\' + rob_program_data.data['localLv1'] + '\\' + \
                                                    rob_program_data.data['localLv2'] + '\\' + rob_program_data.data[
                                                        'localLv3']

                # # 移动重整文件
                # folder = os.path.exists(targetData[compare_name]['newpath'])
                # if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
                #     os.makedirs(targetData[compare_name]['newpath'])  # makedirs 创建文件时如果路径不存在会创建这个路径
                # else:
                #     pass
                # if not os.path.exists(targetData[compare_name]['newpath'] + '\\' + name):
                #     # print(targetData[compare_name]['newpath'] + '\\' + compare_name)
                #     # shutil.move(targetData[compare_name]['originpath'], targetData[compare_name]['newpath'])
                #     shutil.copy2(targetData[compare_name]['originpath'], targetData[compare_name]['newpath'])
                # else:
                #     # shutil.move(targetData[compare_name]['originpath'],
                #     #             targetData[compare_name]['newpath'] + '\\' + name + '副本')
                #     # shutil.copy2(targetData[compare_name]['originpath'],
                #     #              targetData[compare_name]['newpath'] + '\\' + name + '副本')
                #     continue


def backupState():
    standard_filepath = buStandard['filepath'] + '\\' + buStandard['filename']
    book_rd = xlrd.open_workbook(standard_filepath, formatting_info=True)
    sheet_rd = book_rd.sheet_by_index(0)
    book_wt = copy(book_rd)
    sheet_wt = book_wt.get_sheet(SHEET_NAME)
    nrows = sheet_rd.nrows
    nrows_compare = nrows
    global DEAL_FILES
    for root, dirs, files in os.walk(PATH_EXPORT_TO):
        for name in files:
            DEAL_FILES = DEAL_FILES + 1
            rob_program_data.controllername = name.split('.zip')[0]
            workstationname = rob_program_data.controllername[-7:].upper()  # 截取的 eg.k2a3a131s460r04 后7位
            if isAll == True:
                compare_name = rob_program_data.controllername
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
                sheet_wt.write(nrows_compare, COMPARE_COL - 2,
                               pathmap[rob_program_data.controllername[2:6].lower()]['Lv2'])
                sheet_wt.write(nrows_compare, COMPARE_COL - 1,
                               pathmap[rob_program_data.controllername[2:6].lower()]['Lv3'])
                sheet_wt.write(nrows_compare, COMPARE_COL, compare_name)
                sheet_wt.write(nrows_compare, COMMIT_COL, label='新工位')
                sheet_wt.write(nrows_compare, TIME_COL, label=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                nrows_compare = nrows_compare + 1
    logWriteTitle('总结')
    log = open(PATH_BASE + '\\' + LOG_FILE_NAME, 'a')
    log.write(
        '备份总数: ' + str(TOTAL_FILES) + '        已处理: ' + str(DEAL_FILES) + '         异常: ' + str(ERR_FILES) + '\r\n')
    log.close()
    logWriteTitle('end')
    book_wt.save(PATH_BASE + '\\' + time.strftime("%Y%m%d", time.localtime()) + '机器人备份情况.xls')


def main():
    logWriteTitle('start')
    # 1.获取路径
    getZipInfo()
    #
    # backupState()

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
