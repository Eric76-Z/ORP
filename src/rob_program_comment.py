# 获取机器人数据
import json
import os
import re
import time
import zipfile

from openpyxl import Workbook

from src.common.rob_program_comment_setting import RobProgramComment, SH_ROB_COMMENT_ANALYSE, PATH_PROGRAM, \
    ROB_COMMENT_OVERVIEW_REPORT_BASE, ROB_COMMENT_OVERVIEW_REPORT, \
    SH_ROB_COMMENT_OVERVIEW
from src.common.setting import PATH_BASE, time_now
from src.common.utils import TimeStampToTime, getFileSize, logWrite, createFolder, createSheet


# 1、读取机器人zip相关信息
def RobotInfo(SUM, wb):
    rob_program_comment_json = {}
    for root, dirs, files in os.walk(PATH_PROGRAM):
        for name in files:
            originpath = os.path.join(root, name)
            if zipfile.is_zipfile(originpath):
                rob_program_comment = RobProgramComment()
                SUM['total_files'] = SUM['total_files'] + 1
                # ================path==================
                rob_program_comment.path['path_origin'] = originpath  # 原始路径
                # ================meta==================
                rob_program_comment.meta['title'] = name  # 文件名
                rob_program_comment.meta['format'] = 'zip'  # 文件后缀
                rob_program_comment.meta['mtime'] = TimeStampToTime(os.path.getmtime(originpath))  # 原数据中修改时间，可视为创建时间
                rob_program_comment.meta['size'] = getFileSize(originpath)
                # ================data==================
                rob_program_comment.data['filename'] = name
                rob_program_comment.data['controllername'] = name.split('.zip')[0]  # eg.k2a3a131s460r04
                rob_program_comment.data['workstationname'] = rob_program_comment.data['controllername'][
                                                              -7:].upper()  # 截取的 eg.s460r04
                rob_program_comment.data['localLv1'] = rob_program_comment.localLv1()
                rob_program_comment.data['localLv2'] = rob_program_comment.localLv2()
                rob_program_comment.data['localLv3'] = rob_program_comment.localLv3()
                # print('========================解析机器人备份--start========================')
                analysisZip(rob_program_comment, wb=wb)
                # print('========================解析机器人备份--end========================')
                # 添加入rob_program_comment_json
                rob_program_comment_json[rob_program_comment.data['workstationname']] = rob_program_comment
            else:
                msg = '备份可能损坏!!!源路径为：' + originpath + ';'
                logWrite(wb=wb, controllername=name, sort='警告', msg=msg)
    with open('database/robot_comment.json', "w", encoding='utf-8') as f:
        f.write(json.dumps(rob_program_comment_json, default=lambda obj: obj.__dict__, indent=4, ensure_ascii=False))
        f.close()


# 分析机器人zip相关信息
def analysisZip(rob_program_comment, wb):
    try:
        filezip = zipfile.ZipFile(rob_program_comment.path['path_origin'], "r")
        rob_program_comment.zipData['total_files'] = len(filezip.namelist())
        try:

            for file in filezip.namelist():

                if file.split('/')[-1] == 'RefListe.txt':
                    create_time = filezip.getinfo(file).date_time
                    rob_program_comment.zipData['create_time'] = str(create_time[0]) + '-' + str(create_time[1]) + '-' + \
                                                                 str(create_time[2]) + ' ' + str(create_time[3]) + ':' + \
                                                                 str(create_time[4]) + ':' + str(create_time[5])
                    RefListe = filezip.open(file)

                    # 创建工作表，写入内容
                    wb2 = Workbook(write_only=True)
                    # 返回的数据
                    ret = analyseRefListe(RefListe, wb)
                    # print(SH_ROB_COMMENT_ANALYSE + ret['used_all'])
                    createSheet(wb=wb, sh_name='机器人注释解析', sh_index=1,
                                sh_title=(SH_ROB_COMMENT_ANALYSE + ret['used_all']))

                    # # 写入数据
                    i = 0
                    for ref in ret['ref_liste']:
                        i = i + 1
                        print(ref)
                        localLv1 = rob_program_comment.data['localLv1']
                        localLv2 = rob_program_comment.data['localLv2']
                        localLv3 = rob_program_comment.data['localLv3']
                        signal_sort = backS
                        wb2['机器人注释解析'].append([i, localLv1, localLv2, localLv3])
                    # content_1 = [i + 1, depart, localLv1, localLv2, localLv3,
                    #              workstations[i],
                    #              create_time, size, state, totalFiles, folge_num, makro_num, up_num, serial_number,
                    #              robot_type,
                    #              mames_offsets, version, tech_packs, is_axis_7, E1, E2, seven_axis, other_E7, is_news]
                    # 创建文件夹
                    ROB_COMMENT_ANALYSE_REPORT_BASE = PATH_BASE + '\\' + '机器人注释解析详情报告' + '\\' + str(
                        time_now.year) + str(
                        time_now.month) + str(
                        time_now.day) + '\\' + str(
                        time_now.hour) + str(
                        time_now.minute) + str(time_now.second) + '\\' + rob_program_comment.data[
                                                          'localLv1'] + '\\' + \
                                                      rob_program_comment.data['localLv2'] + '\\' + \
                                                      rob_program_comment.data[
                                                          'localLv3']
                    createFolder(ROB_COMMENT_ANALYSE_REPORT_BASE)

                    ROB_COMMENT_ANALYSE_REPORT = rob_program_comment.data['controllername'] + '.xlsx'
                    wb2.save(os.path.join(ROB_COMMENT_ANALYSE_REPORT_BASE, ROB_COMMENT_ANALYSE_REPORT))
                    wb2.close()
        except Exception as e:
            print(rob_program_comment.path['path_origin'] + str(e))
        rob_program_comment.zipData['state'] = '备份完好'  # zip文件完好
        filezip.close()
    except Exception as e:
        msg = '备份可能损坏!!!源路径为：' + rob_program_comment.path['path_origin'] + ';' + '新路径为:' + rob_program_comment.path[
            'path_new']
        logWrite(wb=wb, controllername=rob_program_comment.data['controllername'], sort='警告', msg=msg)
        rob_program_comment.zipData['state'] = '备份损坏'
        print(rob_program_comment.zipData['state'])


# 分析机器人zip中refliste注释文件
def analyseRefListe(RefListe, wb):
    data = {}
    ref_liste = []
    used_all = []
    for l in RefListe.readlines():
        l = str(l)
        l = l.replace('b\'', '')
        l = l.replace('\\r\\n\'', '')
        if re.search(r"^(E|A|M|I|bin|t|F|T|Makro|UP)\s\d{1,4}\s", l) != None:

            # ========================signal========================
            # print(re.search(r"(E|A|M|I|bin|t|F|T|Makro|UP|Folge)\s\d{1,4}", l))
            signal = re.search(r"^(E|A|M|I|bin|t|F|T|Makro|UP)\s\d{1,4}", l).group()
            signal = signal.replace(" ", "")  # 去除空格

            # ========================comments========================
            if re.search(r'\[.*]', l) != None:
                comments = re.search(r'\[.*]', l).group()[1:-1]
                comments = comments.replace(',', '')  # 去除逗号
            else:
                comments = ''

            # ========================used========================
            if re.search(r']\s*(Folge|Floge|Makro|UP).*$', l) != None:
                used_str = re.search(r']\s*(Folge|Floge|Makro|UP).*$', l).group()
            else:
                used_str = re.search(r'\s\s(Folge|Floge|Makro|UP).*$', l).group()
            # print(used_str)
            used_str = used_str.replace('Floge', 'Folge')
            used_str = used_str.replace('*', '')
            used_str = used_str.replace(']', '')
            used_str = used_str.split('  ')[-1]
            used_str = used_str.replace(' ', '')
            # print(used_str)
            used_list = used_str.split(',')
            # print(used_list)
            used = dealused(used_list)
            ret = {
                'signal': signal,
                'comments': comments,
                'used': used
            }
            ref_liste.append(ret)
        elif re.search(r"^(Folge|Floge)\s\d{1,4}", l) != None:
            signal = 'cell'
            comments = ''
            l = l.replace(' ', '')
            l = l.replace('Floge', 'Folge')
            used_list = l.split(',')

            used = dealused(used_list)
            ret = {
                'signal': signal,
                'comments': comments,
                'used': used
            }
            ref_liste.append(ret)
        else:
            used = []
        used_all = used_all + used
    used_all = list(set(used_all))
    used_all.sort()
    return {
        'ref_liste': ref_liste,
        'used_all': used_all
    }


def dealused(list):
    used = []
    flag = 'N'
    for i in list:
        if 'Folge' in i:
            flag = 'F'
        elif 'Makro' in i:
            flag = 'M'
        elif 'UP' in i:
            flag = 'U'
        if flag == 'F':
            if 'Folge' in i:
                used.append(i)
            else:
                used.append('Folge' + i)
        elif flag == 'M':
            if 'Makro' in i:
                used.append(i)
            else:
                used.append('Makro' + i)
        elif flag == 'U':
            if 'UP' in i:
                used.append(i)
            else:
                used.append('UP' + i)
    # print(used)
    return used


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
    createFolder(ROB_COMMENT_OVERVIEW_REPORT_BASE)
    wb1 = Workbook(write_only=True)
    createSheet(wb=wb1, sh_name='机器人注释解析', sh_index=1, sh_title=SH_ROB_COMMENT_OVERVIEW)

    # createSheet(wb=wb2, sh_name='机器人注释解析', sh_index=1, sh_title=SH_ROB_COMMENT_ANALYSE)

    print('========================获取机器人数据--start========================')
    RobotInfo(SUM, wb1)
    print('========================获取机器人数据--end========================')

    wb1.save(os.path.join(ROB_COMMENT_OVERVIEW_REPORT_BASE, ROB_COMMENT_OVERVIEW_REPORT))
    wb1.close()


if __name__ == "__main__":
    main()
