import json
import os
import re
import shutil

import xml.dom.minidom
import zipfile

import xlrd
from openpyxl import Workbook

from src.common.rob_program_backup_setting import RobProgramData, PATH_ORIGIN_BACKUP, PATH_EXPORT_TO, PATH_TRASH, \
    PATH_REPORT, \
    FILE_STANDARD, \
    PATH_STANDARD, FILE_REPORT, SH_LOG_TITLE, SH_ROB_BACKUP_OV

# 获取RobotProgram对象信息，并写入
from src.common.utils import TimeStampToTime, logWrite, getFileSize, createFolder, createSheet


# 获取机器人数据
def RobotInfo(SUM, wb):
    rob_program_data_json = {}
    for root, dirs, files in os.walk(PATH_ORIGIN_BACKUP):
        for name in files:
            originpath = os.path.join(root, name)
            if zipfile.is_zipfile(originpath):
                rob_program_data = RobProgramData()
                SUM['total_files'] = SUM['total_files'] + 1
                # ================path==================
                rob_program_data.path['path_origin'] = originpath  # 原始路径
                # ================meta==================
                rob_program_data.meta['title'] = name  # 文件名
                rob_program_data.meta['format'] = 'zip'  # 文件后缀
                rob_program_data.meta['mtime'] = TimeStampToTime(os.path.getmtime(originpath))  # 原数据中修改时间，可视为创建时间
                rob_program_data.meta['size'] = getFileSize(originpath)
                # ================data==================
                rob_program_data.data['filename'] = name
                rob_program_data.data['controllername'] = name.split('.zip')[0]  # eg.k2a3a131s460r04
                rob_program_data.data['workstationname'] = rob_program_data.data['controllername'][
                                                           -7:].upper()  # 截取的 eg.s460r04
                rob_program_data.data['localLv1'] = rob_program_data.localLv1()
                rob_program_data.data['localLv2'] = rob_program_data.localLv2()
                rob_program_data.data['localLv3'] = rob_program_data.localLv3()

                # ================error==================
                if rob_program_data.data['localLv1'] == '':
                    logWrite(wb=wb, controllername=rob_program_data.data['controllername'], sort='报错', msg='没有对应的一级地点')
                    SUM['err_files'] = SUM['err_files'] + 1
                    continue
                elif rob_program_data.data['localLv2'] == '':
                    logWrite(wb=wb, controllername=rob_program_data.data['controllername'], sort='报错', msg='没有对应二级地点')
                    SUM['err_files'] = SUM['err_files'] + 1
                    continue
                elif rob_program_data.data['localLv3'] == '':
                    logWrite(wb=wb, controllername=rob_program_data.data['controllername'], sort='报错', msg='没有对应三级地点')
                    SUM['err_files'] = SUM['err_files'] + 1
                    continue

                rob_program_data.path['path_new'] = PATH_EXPORT_TO + '\\' + rob_program_data.data['localLv1'] + '\\' + \
                                                    rob_program_data.data['localLv2'] + '\\' + rob_program_data.data[
                                                        'localLv3'] + '\\' + name
                rob_program_data.path['path_new_path'] = PATH_EXPORT_TO + '\\' + rob_program_data.data[
                    'localLv1'] + '\\' + \
                                                         rob_program_data.data['localLv2'] + '\\' + \
                                                         rob_program_data.data[
                                                             'localLv3']

                # 判断rob_program_data中是否已有这个工位
                if rob_program_data.data['workstationname'] in rob_program_data_json.keys():
                    SUM['repeat_files'] = SUM['repeat_files'] + 1
                    # 进一步判断两者创建时间
                    if rob_program_data.meta['mtime'] == \
                            rob_program_data_json[rob_program_data.data['workstationname']].meta['mtime']:
                        if rob_program_data.meta['size'] == \
                                rob_program_data_json[rob_program_data.data['workstationname']].meta['size']:
                            # 如果文件创建时间一致，大小一致，则将待处理文件移动到删除区
                            shutil.move(rob_program_data.path['path_origin'], PATH_TRASH)
                            logWrite(wb=wb, controllername=rob_program_data.data['controllername'], sort='警告',
                                     msg='重复文件，创建时间一致，大小一致，已将其中一个备份移动至垃圾箱！')
                            SUM['trash_files'] = SUM['trash_files'] + 1
                            continue
                        else:
                            # 如果文件创建时间一致，大小不一致，谁小删谁
                            if int(rob_program_data.meta['size'].split('.')[0]) >= \
                                    int(rob_program_data_json[rob_program_data.data['workstationname']].meta[
                                            'size'].split('.')[0]):
                                shutil.move(rob_program_data_json[rob_program_data.data['workstationname']].path[
                                                'path_origin'], PATH_TRASH)
                                SUM['trash_files'] = SUM['trash_files'] + 1
                            else:
                                shutil.move(rob_program_data.path['path_origin'], PATH_TRASH)
                                SUM['delete_files'] = SUM['delete_files'] + 1
                            logWrite(wb=wb, controllername=rob_program_data.data['controllername'], sort='警告',
                                     msg='重复文件，创建时间一致，大小不一致，且size小的备份被移动至垃圾箱！')
                            continue
                    else:
                        # 备份时间判断
                        if rob_program_data.meta['mtime'] > \
                                rob_program_data_json[rob_program_data.data['workstationname']].meta['mtime']:
                            try:
                                shutil.move(rob_program_data_json[rob_program_data.data['workstationname']].path[
                                                'path_origin'], PATH_TRASH)
                                SUM['trash_files'] = SUM['trash_files'] + 1
                            except:
                                os.remove(rob_program_data_json[rob_program_data.data['workstationname']].path[
                                              'path_origin'])
                        else:
                            try:
                                shutil.move(rob_program_data.path['path_origin'], PATH_TRASH)
                            except:
                                os.remove(rob_program_data.path['path_origin'])
                        logWrite(wb=wb, controllername=rob_program_data.data['controllername'], sort='警告',
                                 msg='备份的创建时间过早，移动至垃圾箱！')
                        continue
                # 分析备份大小
                SUM['avg_file_size'] = (SUM['avg_file_size'] * len(rob_program_data_json) + float(
                    rob_program_data.meta['size'].split(' ')[0])) / (len(
                    rob_program_data_json) + 1)
                SUM['avg_file_size'] = float(format(SUM['avg_file_size'], '.2f'))
                if float(rob_program_data.meta['size'].split(' ')[0]) > float(SUM['max_file_size'].split(' ')[0]):
                    SUM['max_file_size'] = rob_program_data.meta['size']
                elif float(rob_program_data.meta['size'].split(' ')[0]) < float(SUM['min_file_size'].split(' ')[0]):
                    SUM['min_file_size'] = rob_program_data.meta['size']
                if float(rob_program_data.meta['size'].split(' ')[0]) > 40 or float(
                        rob_program_data.meta['size'].split(' ')[0]) < 10:
                    msg = '备份的size大于40M或者小于10M！为：' + rob_program_data.meta['size']
                    logWrite(wb=wb, controllername=rob_program_data.data['controllername'], sort='提示', msg=msg)
                # print('========================解析机器人备份--start========================')
                analysisZip(rob_program_data, wb=wb)
                # print('========================解析机器人备份--end========================')
                # 添加入rob_program_data_json
                rob_program_data_json[rob_program_data.data['workstationname']] = rob_program_data
            else:
                msg = '备份可能损坏!!!源路径为：' + originpath + ';'
                logWrite(wb=wb, controllername=name, sort='警告', msg=msg)
    with open('database/robot_data.json', "w", encoding='utf-8') as f:
        f.write(json.dumps(rob_program_data_json, default=lambda obj: obj.__dict__, indent=4, ensure_ascii=False))
        f.close()


def analysisZip(rob_program_data, wb):
    try:
        filezip = zipfile.ZipFile(rob_program_data.path['path_origin'], "r")
        rob_program_data.zipData['total_files'] = len(filezip.namelist())
        try:
            for file in filezip.namelist():
                if (file.split('/')[-1].endswith('.src')):
                    if (file.split('/')[-1].lower().startswith('folge')):
                        rob_program_data.zipData['file_folge_num'] = rob_program_data.zipData['file_folge_num'] + 1
                    elif (file.split('/')[-1].lower().startswith('makro')):
                        rob_program_data.zipData['file_makro_num'] = rob_program_data.zipData['file_makro_num'] + 1
                    elif (file.split('/')[-1].lower().startswith('up')):
                        rob_program_data.zipData['file_up_num'] = rob_program_data.zipData['file_up_num'] + 1
                elif file.split('/')[-1] == 'RobotInfo.xml':
                    RobotInfoXml = filezip.open(file)
                    dom = xml.dom.minidom.parse(RobotInfoXml)
                    root = dom.documentElement
                    rob_program_data.zipData['serial_number'] = \
                        root.getElementsByTagName('SerialNumber')[0].childNodes[
                            0].data
                    rob_program_data.zipData['robot_type'] = root.getElementsByTagName('RobotType')[0].childNodes[
                        0].data
                    rob_program_data.zipData['mames_offsets'] = root.getElementsByTagName('MamesOffsets')[
                        0].getAttribute('Timestamp')
                elif file.split('/')[-1] == 'am.ini':
                    am_ini = filezip.open(file)
                    mystr = str(am_ini.read())
                    rob_program_data.zipData['version'] = \
                        re.findall(r'Version=(.+)\[', mystr)[0].split('\\r\\n')[0]
                    rob_program_data.zipData['tech_packs'] = re.findall(r'TechPacks=(.+)\|', mystr)[0]
                elif (file.split('/')[-1] == 'E1.xml' or file.split('/')[-1] == 'E2.xml') and file.split('/')[
                    -2] == 'SimuAxis':
                    RobotInfoXml = filezip.open(file)
                    dom = xml.dom.minidom.parse(RobotInfoXml)
                    root = dom.documentElement
                    if file.split('/')[-1] == 'E1.xml':
                        rob_program_data.zipData['E1'] = \
                            root.getElementsByTagName('Machine')[0].getAttribute('Name')
                        rob_program_data.zipData['is_axis_7'] = '7轴'
                        if rob_program_data.zipData['E1'].startswith('#KR'):
                            rob_program_data.zipData['other_E7'] = rob_program_data.zipData['E1']
                        else:
                            rob_program_data.zipData['seven_axis'] = rob_program_data.zipData['E1']
                    elif file.split('/')[-1] == 'E2.xml':
                        rob_program_data.zipData['E2'] = \
                            root.getElementsByTagName('Machine')[0].getAttribute('Name')
                        rob_program_data.zipData['is_axis_7'] = '7轴'
                        if rob_program_data.zipData['E2'].startswith('#KR'):
                            rob_program_data.zipData['other_E7'] = rob_program_data.zipData['E2']
                        else:
                            rob_program_data.zipData['seven_axis'] = rob_program_data.zipData['E2']
                if rob_program_data.zipData['E1'] != 'null' or rob_program_data.zipData['E2'] != 'null':
                    rob_program_data.zipData['is_axis_7'] = '7轴'
                else:
                    rob_program_data.zipData['is_axis_7'] = '非7轴'

            # if rob_program_data.data['workstationname'] == '5330R07':
            #     print(rob_program_data.zipData)
        except Exception as e:
            print(rob_program_data.path['path_origin'] + str(e))
        rob_program_data.zipData['state'] = '备份完好'  # zip文件完好
        filezip.close()
    except Exception as e:
        msg = '备份可能损坏!!!源路径为：' + rob_program_data.path['path_origin'] + ';' + '新路径为:' + rob_program_data.path[
            'path_new']
        logWrite(wb=wb, controllername=rob_program_data.data['controllername'], sort='警告', msg=msg)
        rob_program_data.zipData['state'] = '备份损坏'
        print(rob_program_data.zipData['state'])


def Reforming(SUM):
    ## 移动重整文件
    ## 读取json文件
    with open('database/robot_data.json', 'r', encoding='utf-8') as f:
        info_dict = json.load(f)
        for dict in info_dict:
            folder = os.path.exists(info_dict[dict]['path']['path_new_path'])
            if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
                os.makedirs(info_dict[dict]['path']['path_new_path'])  # makedirs 创建文件时如果路径不存在会创建这个路径
            else:
                pass
            if not os.path.exists(info_dict[dict]['path']['path_new']):
                SUM['move_files'] = SUM['move_files'] + 1
                shutil.copy2(info_dict[dict]['path']['path_origin'], info_dict[dict]['path']['path_new'])
            else:
                SUM['exists_files'] = SUM['exists_files'] + 1
                continue


def backupOverview(wb, SUM):
    # 读取json文件
    book_standard = xlrd.open_workbook(os.path.join(PATH_STANDARD, FILE_STANDARD), formatting_info=True)
    sh_standard = book_standard.sheet_by_name('RobStandard')
    nrows = sh_standard.nrows
    sh = wb['机器人备份总览']
    with open('database/robot_data.json', 'r', encoding='utf-8') as f:
        info_dict = json.load(f)
        workstations = []
        for i in range(1, nrows):
            workstations.append(sh_standard.cell_value(i, 5))
        old_len = len(workstations)
        for dict in info_dict:
            if dict not in workstations:
                workstations.append(dict)
        for i in range(0, len(workstations)):
            try:
                depart = sh_standard.cell_value(i + 1, 1)
                localLv1 = sh_standard.cell_value(i + 1, 2)
                localLv2 = sh_standard.cell_value(i + 1, 3)
                localLv3 = sh_standard.cell_value(i + 1, 4)
            except:
                depart = 'null'
                localLv1 = 'null'
                localLv2 = 'null'
                localLv3 = 'null'
            create_time = 'null'
            size = 'null'
            try:
                state = info_dict[workstations[i]]['zipData']['state']
            except:
                state = '未备份'
            totalFiles = 0
            makro_num = 0
            folge_num = 0
            up_num = 0
            serial_number = 'null'
            robot_type = 'null'
            mames_offsets = 'null'
            version = 'null'
            tech_packs = 'null'
            is_axis_7 = 'null'
            E1 = 'null'
            E2 = 'null'
            is_news = 'null'
            if workstations[i] in info_dict or i >= old_len:
                localLv1 = info_dict[workstations[i]]['data']['localLv1']
                localLv2 = info_dict[workstations[i]]['data']['localLv2']
                localLv3 = info_dict[workstations[i]]['data']['localLv3']
                create_time = info_dict[workstations[i]]['meta']['mtime']
                size = info_dict[workstations[i]]['meta']['size']
                totalFiles = info_dict[workstations[i]]['zipData']['total_files']
                folge_num = info_dict[workstations[i]]['zipData']['file_folge_num']
                makro_num = info_dict[workstations[i]]['zipData']['file_makro_num']
                up_num = info_dict[workstations[i]]['zipData']['file_up_num']
                if info_dict[workstations[i]]['zipData']['state'] == '备份完好':
                    serial_number = info_dict[workstations[i]]['zipData']['serial_number']
                    robot_type = info_dict[workstations[i]]['zipData']['robot_type']
                    mames_offsets = info_dict[workstations[i]]['zipData']['mames_offsets']
                    version = info_dict[workstations[i]]['zipData']['version']
                    tech_packs = info_dict[workstations[i]]['zipData']['tech_packs']
                    is_axis_7 = info_dict[workstations[i]]['zipData']['is_axis_7']
                    E1 = info_dict[workstations[i]]['zipData']['E1']
                    E2 = info_dict[workstations[i]]['zipData']['E2']
                    seven_axis = info_dict[workstations[i]]['zipData']['seven_axis']
                    other_E7 = info_dict[workstations[i]]['zipData']['other_E7']
                if i >= old_len:
                    is_news = '新工位'
            content_1 = [i + 1, depart, localLv1, localLv2, localLv3,
                         workstations[i],
                         create_time, size, state, totalFiles, folge_num, makro_num, up_num, serial_number, robot_type,
                         mames_offsets, version, tech_packs, is_axis_7, E1, E2, seven_axis, other_E7, is_news]
            sh.append(content_1)
    # print(SUM)


def main():
    # 总结
    global SUM
    SUM = {
        'total_files': 0,  # 更目录下文件总数
        'err_files': 0,
        'move_files': 0,
        'exists_files': 0,
        'repeat_files': 0,
        'trash_files': 0,
        'avg_file_size': 0,
        'min_file_size': '1000.0 MB',
        'max_file_size': '00.0 MB'
    }
    createFolder(PATH_TRASH)
    createFolder(PATH_REPORT)
    wb = Workbook(write_only=True)
    createSheet(wb=wb, sh_name='机器人备份总览', sh_index=1, sh_title=SH_ROB_BACKUP_OV)
    createSheet(wb=wb, sh_name='log', sh_index=3, sh_title=SH_LOG_TITLE)

    print('========================获取机器人数据--start========================')
    RobotInfo(SUM, wb)
    print('========================获取机器人数据--end========================')
    print('========================移动重整备份--start========================')
    Reforming(SUM)
    print('========================移动重整备份--end========================')
    print('========================机器人overview--start========================')
    backupOverview(wb, SUM)
    print('========================机器人overview--end========================')
    wb.save(os.path.join(PATH_REPORT, FILE_REPORT))


if __name__ == "__main__":
    main()
