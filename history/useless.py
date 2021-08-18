# 用于提取某压缩包中某文件内容，并解析
import os
import zipfile

# from RobotInfo.Constant import PATHRB
from xml.dom import minidom

import xlwt

PATHRB = 'E:\\old'
RobotDatas = {}
mainData = []
ExcleDatas = []
rootlists = []
dirlists = []
filelists = []
# 各文件路径
filepaths = {
    'zipfile': [],  #zip格式文件
}
controllerlists = []  # 控制器名字 eg.k2a3a111s410r05
workstationlists = []  # 工位名字 eg. s410r05

targetFile = 'RobotInfo.xml'
zipfileTarget = 'C/KRC/Roboter/Rdc/RobotInfo.xml'
extractTo = 'unzip\\'
extractedFileList = []  # 列表，储存解压后文件路径

ExclePath = '机器人Rdc数据表V4.xls'


# 获取路径
def getPath():
    # print('11')
    for root, dirs, files in os.walk(PATHRB):
        # print(root)  # 输出各路径
        # print(dirs)  # 输出文件夹名的列表
        # print(files)  # 输出文件名列表

        # for i in root:
        #     rootlists.append(i)
        #
        #
        # for j in dirs:
        #     dirlists.append(j)
        #     # print(j)

        # for k in files:
        #     filelists.append(k)
            # print(k)

        for name in files:
            if name.endswith('.zip'):
                filepaths['zipfile'].append(os.path.join(root, name))
        # print(filepaths)
    print(filepaths)
    print(filepaths['zipfile'].__len__())

    for filename in filelists:
        controllername = filename.split('.zip')
        workstationname = controllername[0][-7:]  # 截取的 eg.k2a3a131s460r04 后7位
        controllerlists.append(controllername[0])
        workstationlists.append(workstationname)
    # print(controllerlists)
    # print(workstationlists)