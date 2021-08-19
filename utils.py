import os
import time

from ini import PATH_BASE, LOG_FILE_NAME, PATH_TRASH, LOG_TARSH_NAME


def logWrite(file, msg, controllername):
    log = open(file, 'a')
    log.write(controllername + '[' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ']' + ' :' + msg + '\r\n')
    log.close()


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
