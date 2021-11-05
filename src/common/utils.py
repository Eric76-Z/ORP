import datetime
import os
import time

from src.common.rob_program_backup_setting import PATH_TRASH, LOG_TARSH_NAME


def logWrite(wb, controllername, sort, msg):
    sh = wb['log']
    # row_now = sh.max_row
    content = [controllername, datetime.datetime.now(), sort, msg]
    sh.append(content)


def createSheet(wb, sh_name, sh_index, sh_title):
    '''
    wb: workbook实例
    sh_name:要创建的表格名字
    sh_index:sheet的index
    sh_title: sheet第一行标题，为列表
    '''
    wb.create_sheet(title=sh_name, index=sh_index)
    sh_log = wb[sh_name]
    sh_log.append(sh_title)


def logTrash(controllername, msg):
    log = open(PATH_TRASH + '\\' + LOG_TARSH_NAME, 'a')


def logWriteTitle(file, msg):
    log = open(file, 'a')
    log.write('==============================' + msg + '==============================' + '\r\n')
    log.close()


# 　　'''把时间戳转化为时间: 1479264792 to 2016-11-16 10:53:12'''
def TimeStampToTime(timestamp):
    timeStruct = time.localtime(timestamp)
    return time.strftime('%Y-%m-%d %H:%M:%S', timeStruct)


def getFileSize(filePath):
    size = os.path.getsize(filePath)  # 返回的是字节大小
    '''
    为了更好地显示，应该时刻保持显示一定整数形式，即单位自适应
    '''
    if size < 1000:
        return '%i' % size + ' size'
    elif 1000 <= size < 1000000:
        return '%.1f' % float(size / 1000) + ' KB'
    elif 1000000 <= size < 1000000000:
        return '%.1f' % float(size / 1000000) + ' MB'
    elif 1000000000 <= size < 1000000000000:
        return '%.1f' % float(size / 1000000000) + ' GB'
    elif 1000000000000 <= size:
        return '%.1f' % float(size / 1000000000000) + ' TB'


def createFolder(path):
    folder = os.path.exists(path)
    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径
    else:
        pass
