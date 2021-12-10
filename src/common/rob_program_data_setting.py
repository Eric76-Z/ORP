from src.common.constant import pathmap
from src.common.setting import PATH_BASE, time_now

# =============================================机器人程序注释===============================================
# SUM = 38
PATH_PROGRAM = PATH_BASE + '\\' + 'rob backup reformed'  # 目标机器人备份所在文件夹
SH_ROB_COMMENT_OVERVIEW = ['id', '部门', '车间', '区域', '线体', '工位', '创建时间', '链接']
# SH_ROB_COMMENT_ANALYSE = ['id', '区域', '线体', '工位', '信号类', '信号', '现注释', '标准注释', '相似度(标准)', '推荐注释', '相似度(推荐)', '是否修改注释',
#                           '是否需要注释', '终判']
# SH_BID_DATA_TEMPLE = ['id', '信号', '样本数量', '推荐注释', '数量', '概率']

ROB_PROGARM_DATA_EXPORT_FOLDER_BASE = PATH_BASE + '\\' + 'report' + '\\' + '机器人备份数据提取'
ROB_PROGARM_DATA_EXPORT_OVERVIEW = '机器人备份数据池导出总览.xlsx'


class RobProgramData:
    def __init__(self):
        self.sort = '机器人备份数据'
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
        self.filedata = {
            'am.ini': {
                'version': 'null',
                'tech_packs': 'null'
            },
            'RobotInfo.xml': {
                'serial_number': 'null',
                'rob_type': 'null',
                'time_stamp': 'null'
            },
            'NGAxis': {
                'A1.xml': {
                    'ksp2': 'null'
                },
                'A4.xml': {
                    'ksp1': 'null'
                }
            },
            'SimuAxis': {
                'E1.xml': {
                    'machine_name': 'null'
                },
                'E2.xml': {
                    'machine_name': 'null'
                }
            },
            'ecatms_config.xml': {
                'kpp': {
                    'model': 'null',
                    'revision_no': 'null'
                }

            }

        }
        self.analysedata = {
            'rob': {
                'rob_type': 'null',
                'load_type': 'null',
            },
            'cabinet': {
                'ServoModuleFct': {
                    'ksp2': 'null',
                    'ksp1': 'null',
                    'kpp': 'null'
                },
                'extra_axis': {
                    'is_extra_axis': 'null',
                    'linear_units': 'null',
                    'other_extra_axis': 'null'
                }
            }
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
