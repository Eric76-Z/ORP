# 配置变量
import pathmap

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