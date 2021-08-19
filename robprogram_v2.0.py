import json
import os
import shutil
import time
import zipfile

from openpyxl import Workbook, load_workbook
from xlutils.copy import copy

from ini import RobProgramData, PATH_ORIGIN_BACKUP, PATH_EXPORT_TO, PATH_TRASH, LOG_TARSH_NAME, PATH_BASE, \
    LOG_FILE_NAME, PATH_REPORT, FILE_STANDARD, PATH_STANDARD
from pathmap import pathmap

# 获取RobotProgram对象信息，并写入
from utils import TimeStampToTime, logWrite, logWriteTitle, getFileSize, createFolder


# 获取机器人数据
def RobotInfo(SUM):
    rob_program_data_json = {}
    for root, dirs, files in os.walk(PATH_ORIGIN_BACKUP):
        for name in files:
            if name.endswith('.zip'):
                rob_program_data = RobProgramData()
                SUM['total_files'] = SUM['total_files'] + 1
                originpath = os.path.join(root, name)
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
                    logWrite(PATH_BASE + '\\' + LOG_FILE_NAME, '【异常】' + rob_program_data.data['controllername'],
                             '没有对应一级地点')
                    SUM['err_files'] = SUM['err_files'] + 1
                    continue
                elif rob_program_data.data['localLv2'] == '':
                    logWrite(PATH_BASE + '\\' + LOG_FILE_NAME, '【异常】' + rob_program_data.data['controllername'],
                             '没有对应二级地点')
                    SUM['err_files'] = SUM['err_files'] + 1
                    continue
                elif rob_program_data.data['localLv3'] == '':
                    logWrite(PATH_BASE + '\\' + LOG_FILE_NAME, '【异常】' + rob_program_data.data['controllername'],
                             '没有对应三级地点')
                    SUM['err_files'] = SUM['err_files'] + 1
                    continue
                rob_program_data.path['path_new'] = PATH_EXPORT_TO + '\\' + rob_program_data.data['localLv1'] + '\\' + \
                                                    rob_program_data.data['localLv2'] + '\\' + rob_program_data.data[
                                                        'localLv3'] + '\\' + name

                # print(int(rob_program_data.meta['size'].split('.')[0]))
                # 判断rob_program_data中是否已有这个工位
                if rob_program_data.data['workstationname'] in rob_program_data_json.keys():
                    SUM['repeat_files'] = SUM['repeat_files'] + 1
                    # 进一步判断两者创建时间
                    if rob_program_data.meta['mtime'] == \
                            rob_program_data_json[rob_program_data.data['workstationname']].meta['mtime']:
                        if rob_program_data.meta['size'] == \
                                rob_program_data_json[rob_program_data.data['workstationname']].meta['size']:
                            # 如果文件创建时间一致，大小一致，则将待处理文件移动到删除区
                            # 重命名文件（目录）
                            msg = '重复文件'
                            logWrite(PATH_TRASH + '\\' + LOG_TARSH_NAME, msg, rob_program_data.data['controllername'])
                            shutil.move(rob_program_data.path['path_origin'], PATH_TRASH)
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
                            msg = '文件size小'
                            logWrite(PATH_TRASH + '\\' + LOG_TARSH_NAME, msg, rob_program_data.data['controllername'])
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
                        msg = '备份时间过早'
                        logWrite(PATH_TRASH + '\\' + LOG_TARSH_NAME, msg, rob_program_data.data['controllername'])
                        continue
                else:
                    pass
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
                    logWrite(PATH_BASE + '\\' + LOG_FILE_NAME, msg, rob_program_data.data['controllername'])
                    continue
                # print('========================解析机器人备份--start========================')
                analysisZip(rob_program_data)
                # print('========================解析机器人备份--end========================')
                # 添加入rob_program_data_json
                rob_program_data_json[rob_program_data.data['workstationname']] = rob_program_data
    with open('robot_data_json.json', "w", encoding='utf-8') as f:
        f.write(json.dumps(rob_program_data_json, default=lambda obj: obj.__dict__, indent=4, ensure_ascii=False))
        f.close()


def Reforming(rob_program_data):
    # # 移动重整文件
    ## 读取json文件
    with open('robot_data_json.json', 'r', encoding='utf-8') as f:
        info_dict = json.load(f)
        for dict in info_dict:
            folder = os.path.exists(info_dict[dict]['path']['path_new'])
            if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
                os.makedirs(info_dict[dict]['path']['path_new'])  # makedirs 创建文件时如果路径不存在会创建这个路径
            else:
                pass
            if not os.path.exists(info_dict[dict]['path']['path_new']):
                SUM['move_files'] = SUM['move_files'] + 1
                shutil.copy2(info_dict[dict]['path']['path_origin'], info_dict[dict]['path']['path_new'])
            else:
                SUM['exists_files'] = SUM['exists_files'] + 1
                continue
    print(SUM)


def analysisZip(rob_program_data):
    try:
        filezip = zipfile.ZipFile(rob_program_data.path['path_origin'], "r")
        rob_program_data.zipData['total_files'] = len(filezip.namelist())
        filezip.close()
    except Exception as e:
        msg = '【警告】备份可能损坏!!!源路径为：' + rob_program_data.path['path_origin'] + ';' + '新路径为:' + rob_program_data.path[
            'path_new']
        logWrite(PATH_TRASH + '\\' + LOG_TARSH_NAME, msg, rob_program_data.data['controllername'])
        print(rob_program_data.path['path_origin'] + ':' + str(e))
    # print(rob_program_data.path['path_origin'])
    # for i in filezip.namelist():
    #     print(i)


def backupState():
    standard_filepath = buStandard['filepath'] + '\\' + buStandard['filename']
    book_rd = xlrd.open_workbook(standard_filepath, formatting_info=True)
    sheet_rd = book_rd.sheet_by_index(0)
    book_wt = copy(book_rd)
    sheet_wt = book_wt.get_sheet(SHEET_NAME)
    nrows = sheet_rd.nrows
    nrows_compare = nrows
    global DEAL_FILES
    for root, dirs, files in os.walk(PATH_EXPORT_TO):
        for name in files:
            DEAL_FILES = DEAL_FILES + 1
            rob_program_data.controllername = name.split('.zip')[0]
            workstationname = rob_program_data.controllername[-7:].upper()  # 截取的 eg.k2a3a131s460r04 后7位
            if isAll == True:
                compare_name = rob_program_data.controllername
            else:
                compare_name = workstationname
            flag = True

            for i in range(nrows):
                lng = sheet_rd.cell(i - 1, COMPARE_COL).value
                if lng == compare_name:
                    sheet_wt.write(i - 1, COMMIT_COL, label='已备份')
                    sheet_wt.write(i - 1, TIME_COL, label=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                    flag = False
                    break
            if flag == True:
                sheet_wt.write(nrows_compare, COMPARE_COL - 3, targetData[compare_name]['Lv1'])
                sheet_wt.write(nrows_compare, COMPARE_COL - 2,
                               pathmap[rob_program_data.controllername[2:6].lower()]['Lv2'])
                sheet_wt.write(nrows_compare, COMPARE_COL - 1,
                               pathmap[rob_program_data.controllername[2:6].lower()]['Lv3'])
                sheet_wt.write(nrows_compare, COMPARE_COL, compare_name)
                sheet_wt.write(nrows_compare, COMMIT_COL, label='新工位')
                sheet_wt.write(nrows_compare, TIME_COL, label=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
                nrows_compare = nrows_compare + 1
    logWriteTitle('总结')
    log = open(PATH_BASE + '\\' + LOG_FILE_NAME, 'a')
    log.write(
        '备份总数: ' + str(SUM['total_files']) + '        已处理: ' + str(DEAL_FILES) + '         异常: ' + str(
            SUM['err_files']) + '\r\n')
    log.close()
    logWriteTitle('end')
    book_wt.save(PATH_BASE + '\\' + time.strftime("%Y%m%d", time.localtime()) + '机器人备份情况.xls')


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
    # createFolder(PATH_TRASH)
    # createFolder(PATH_REPORT)
    # logWriteTitle(PATH_BASE + '\\' + LOG_FILE_NAME, 'start')
    # logWriteTitle(PATH_TRASH + '\\' + LOG_TARSH_NAME, 'start')
    wb = load_workbook(os.path.join(PATH_STANDARD, FILE_STANDARD))
    print(wb.worksheets.)
    ws = wb.active
    print(ws.cell(1, 1))
    # print('========================获取机器人数据--start========================')
    # RobotInfo(SUM)
    # print('========================获取机器人数据--end========================')
    # print('========================移动重整备份--start========================')
    # Reforming(SUM)
    # print('========================移动重整备份--end========================')
    #
    # logWriteTitle(PATH_BASE + '\\' + LOG_FILE_NAME, 'end')
    # logWriteTitle(PATH_TRASH + '\\' + LOG_TARSH_NAME, 'end')
    #
    # backupState()

    # # 2.解压某文件到指定文件夹
    # extractFile()
    #
    # # 3.解析文件，提取数据
    # getfileData()
    #
    # # 4.导出数据
    # exportfileData()


if __name__ == "__main__":
    main()
