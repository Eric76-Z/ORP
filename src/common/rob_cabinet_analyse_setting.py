from src.common.constant import pathmap
from src.common.setting import PATH_BASE, time_now

# =============================================机器人程序注释===============================================
SUM = 38
PATH_PROGRAM = PATH_BASE + '\\' + '机器人备份'  # 目标机器人被封所在文件夹
SH_ROB_CABINET_ANALYSE = ['id', '车间', '区域', '线体', '工位', '机器人型号', 'KSP2_SET', 'KSP1_SET', 'KPP_SET', 'KSP2_ACT',
                          'KSP1_ACT', 'KPP_ACT', '柜箱型号是否正常']

ROB_CABINET_ANALYSE_FOLDER = PATH_BASE + '\\' + '机器人柜箱'
ROB_CABINET_ANALYSE_FILE = '机器人柜箱数据.xlsx'
