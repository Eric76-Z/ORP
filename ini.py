# 配置变量
import datetime

from pathmap import pathmap

time_now = datetime.datetime.now()

PATH_BASE = 'G:\\机器人备份'  # 根目录
PATH_ORIGIN_BACKUP = PATH_BASE + '\\' + 'old'  # 原备份所在文件夹
PATH_EXPORT_TO = PATH_BASE + '\\' + 'new'  # 重整后文件夹位置
PATH_REPORT = PATH_BASE + '\\' + '备份报告' + '\\' + str(time_now.year) + str(time_now.month) + str(time_now.day)
PATH_STANDARD = PATH_BASE
FILE_STANDARD = 'BackUpStandard.xls'
FILE_REPORT = '机器人备份报告' + str(time_now.year) + str(time_now.month) + str(time_now.day) + str(time_now.hour) + str(
    time_now.minute) + str(time_now.second) + '.xlsx'

LOG_FILE_NAME = 'log.txt'
PATH_TRASH = PATH_BASE + '\\' + '垃圾箱'
LOG_TARSH_NAME = 'trash_log.txt'

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
        self.zipData = {
            'total_files': 0,
            'state': 'null',
            'file_folge_num': 0,
            'file_makro_num': 0,
            'file_up_num': 0,
            'serial_number': 'null',
            'robot_type': 'null',
            'mames_offsets': 'null',
            'version': 'null',
            'tech_packs': 'null',
            'is_axis_7': 'null',
            'seven_axis': 'null',
            'other_E7': 'null',
            'E1': 'null',
            'E2': 'null'
        }

    def localLv1(self):
        if self.data['controllername'][0:2].lower() == 'k1':
            return 'CPH2.1'
        elif self.data['controllername'][0:2].lower() == 'k2':
            return 'CPH2.2'
        elif self.data['controllername'][0:2].lower() == 'k3':
            return 'CPH2.1'
        elif self.data['controllername'][0:2].lower() == 'k4':
            return 'CPH2.2'
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
MOVE_FILES = 0
REPEAT_FILES = 0

# 备份报告
# log
SH_ROB_BACKUP_OV = ['id', '部门', '车间', '区域', '线体', '工位', '创建时间', '大小', '是否损坏', '文件总数', 'folge文件数', 'makro文件数', 'up文件数',
                    '序列号', '机器人型号', '投入运行时间', '系统版本', '安装包', '七轴', 'E1', 'E2', '七轴导轨', '其他E7', '新工位']
SH_ROB_BACKUP_ANA = ['id', '部门', '车间', '区域', '线体', '工位', '总文件数']
SH_LOG_TITLE = ['工位', '时间', '警告等级', '警告内容']
SH_ZIP_LOG_TITLE = ['工位', '时间', '警告等级', '警告内容']
