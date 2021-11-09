# 获取机器人数据
import json
import os
import re
import zipfile

from openpyxl import Workbook

from src.common.rob_program_comment_setting import RobProgramComment, SH_ROB_COMMENT_ANALYSE, PATH_PROGRAM, \
    ROB_COMMENT_OVERVIEW_REPORT_BASE, ROB_COMMENT_ANALYSE_REPORT_BASE, ROB_COMMENT_OVERVIEW_REPORT
from src.common.utils import TimeStampToTime, getFileSize, logWrite, createFolder, createSheet


# 1、读取机器人zip相关性息
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
                    analyseRefListe(RefListe)
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


def analyseRefListe(RefListe):
    for l in RefListe.readlines():
        # if str(l).startswith('b\'E'):  # E开头的型号
        #     print(str(l))
        # print(l)
        l = str(l)
        if len(re.findall(r"E\s\d{1,4}", l)) != 0:
            E1 = re.search(r"E\s\d{1,4}", l).group()
            E1 = E1.replace(" ", "")  # 去除空格
            E2 = re.search(r'\[.*]', l).group()[1:-1]
            E2 = E2.replace(',', '')  # 去除逗号
            E3 = re.search(r']\s*(Folge|Makro|UP).*$', l).group()
            E3 = E3.split('  ')[-1]
            E3 = E3.strip()
            E3 = E3.split('\\r\\n')[0]
            # print(type(E2))
            # print(E1 + '-' + E2)
            print(E3)


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
    createSheet(wb=wb1, sh_name='机器人注释解析', sh_index=1, sh_title=SH_ROB_COMMENT_ANALYSE)
    # createFolder(ROB_COMMENT_ANALYSE_REPORT_BASE)
    # wb2 = Workbook(write_only=True)
    # createSheet(wb=wb2, sh_name='机器人注释解析', sh_index=1, sh_title=SH_ROB_COMMENT_ANALYSE)

    print('========================获取机器人数据--start========================')
    RobotInfo(SUM, wb1)
    print('========================获取机器人数据--end========================')
    # print('========================移动重整备份--start========================')
    # Reforming(SUM)
    # print('========================移动重整备份--end========================')
    # print('========================机器人overview--start========================')
    # backupOverview(wb, SUM)
    # print('========================机器人overview--end========================')
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
    wb1.save(os.path.join(ROB_COMMENT_OVERVIEW_REPORT_BASE, ROB_COMMENT_OVERVIEW_REPORT))


if __name__ == "__main__":
    main()
