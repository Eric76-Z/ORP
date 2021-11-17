from src.common.pathmap import pathmap
from src.common.setting import PATH_BASE, time_now

# =============================================机器人程序注释===============================================
PATH_PROGRAM = PATH_BASE + '\\' + '机器人备份'  # 目标机器人被封所在文件夹
SH_ROB_COMMENT_OVERVIEW = ['id', '部门', '车间', '区域', '线体', '工位', '创建时间', '链接']
SH_ROB_COMMENT_ANALYSE = ['id', '区域', '线体', '工位', '信号类', '信号', '现注释', '标准注释', '相似度(标准)', '推荐注释', '相似度(推荐)', '是否修改注释',
                          '是否需要注释', '终判']
SH_BID_DATA_TEMPLE = ['id', '信号', '样本数量', '推荐注释', '数量', '概率']

ROB_COMMENT_OVERVIEW_REPORT_BASE = PATH_BASE + '\\' + '机器人注释解析总览报告' + '\\' + str(time_now.year) + str(
    time_now.month) + str(
    time_now.day)

ROB_COMMENT_OVERVIEW_REPORT = '机器人注释解析报告' + str(time_now.year) + str(time_now.month) + str(time_now.day) + str(
    time_now.hour) + str(
    time_now.minute) + str(time_now.second) + '.xlsx'

STANDARD_COMMENT = PATH_BASE + '\\' + '机器人初始注释.xlsx'
BID_DATA_TEMPLE = PATH_BASE + '\\' + '机器人注释大数据统计.xlsx'


# ROB_COMMENT_ANALYSE_REPORT_BASE = PATH_BASE + '\\' + '机器人注释解析详情' + '\\' + str(time_now.year) + str(
#     time_now.month) + str(
#     time_now.day)
# ROB_COMMENT_ANALYSE_REPORT = '机器人注释解析详情' + str(time_now.year) + str(time_now.month) + str(time_now.day) + str(
#     time_now.hour) + str(
#     time_now.minute) + str(time_now.second) + '.xlsx'


class RobProgramComment:
    def __init__(self):
        self.sort = '机器人注释'
        self.path = {
            'path_origin': '',
        }
        self.meta = {
            'title': '',
            'mtime': '',
            'format': '',
            'size': '',

        }
        self.data = {
            'controllername': '',
            'workstationname': '',
        }
        self.zipData = {
            'create_time': '',
            'total_files': 0,
            'state': 'null',
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
