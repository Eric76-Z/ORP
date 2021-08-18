import time

from ini import PATH_BASE, LOG_FILE_NAME


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
