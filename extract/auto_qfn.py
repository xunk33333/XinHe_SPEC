import os
import re
import shutil

from extract.cellsearch import exceltool

def extractTable(pdfPath, pageNumber, tableselectRec, outputPath): 
    exe = exceltool()
    exe.get_table_from_pdf(pdfPath, pageNumber, tableselectRec, 'tmp.xlsx')
    data = {}
    def matchKeyNameValue(exe,data,keyword = 'e',name = 'Pads_pitch'):
        cell = exe.cellsearch(keyword,True)
        if cell is not None:
            row = cell.row
            col = cell.col_idx
            tmp = [exe.get_cell(row, col+x).value for x in range(1,4)]
            if  tmp[0] is not None and tmp[1] is not None and tmp[2] is None: #min #max
                data[name] = (float(re.sub(r"[A-Za-z]","",tmp[0]))+float(re.sub(r"[A-Za-z]","",tmp[1])))/2    
            elif tmp[0] is not None and tmp[1] is None and tmp[2] is None:#BSC
                data[name] = tmp[0]
            elif tmp[0] is not None and tmp[1] is  None and tmp[2] is not None:#min #max
                data[name] = (float(re.sub(r"[A-Za-z]","",tmp[0]))+float(re.sub(r"[A-Za-z]","",tmp[2])))/2 
            else:
                data[name] = tmp[1]
        else:
            data[name] = None
    #
    matchKeyNameValue(exe,data,keyword = 'e',name = 'Pads_pitch')
    matchKeyNameValue(exe,data,keyword = 'b',name = 'Pads_width')
    matchKeyNameValue(exe,data,keyword = 'L',name = 'Pads_length')
    matchKeyNameValue(exe,data,keyword = 'D2',name = 'EPad_width')
    matchKeyNameValue(exe,data,keyword = 'E2',name = 'EPad_length')

    #针对D2/E2特殊情况
    if data['EPad_width'] is  None and data['EPad_length'] is  None:
        matchKeyNameValue(exe,data,keyword = 'D2/E2',name = 'EPad_width')
        matchKeyNameValue(exe,data,keyword = 'D2/E2',name = 'EPad_length')

    if data['EPad_width'] is not None and data['EPad_length'] is not None:
        data['EPad_epad'] = "True"
    else:
        data['EPad_epad'] = "False"
    matchKeyNameValue(exe,data,keyword = 'D',name = 'Package_width')
    matchKeyNameValue(exe,data,keyword = 'E',name = 'Package_length')

    #针对D/E特殊情况
    if data['Package_width'] is  None and data['Package_length'] is  None:
        matchKeyNameValue(exe,data,keyword = 'D/E',name = 'Package_width')
        matchKeyNameValue(exe,data,keyword = 'D/E',name = 'Package_length')

    #结果导出
    name = pdfPath[pdfPath.rindex('/') + 1:-4]
    import pandas as pd
    df = pd.DataFrame(data,index=[0 ])
    if not os.path.exists(outputPath):
        os.mkdir(outputPath)

    if 1 == 0:  # 方便测试
        save_path = outputPath + '/' + name + \
            '_page_' + str(pageNumber - 1 + 1) + "_0.csv"
        while os.path.exists(save_path):
            save_path = save_path[:-5] + str(int(save_path[-5]) + 1) + '.csv'
    else:
        save_path = outputPath + '/' + name + \
            '_page_' + str(pageNumber - 1 + 1) + ".csv"
        os.remove('tmp.xlsx')
    df.to_csv(path_or_buf=save_path, sep=',', header=True, index=False)


def del_dir(path):
    filelist = []
    rootdir = path  # 选取删除文件夹的路径,最终结果删除img文件夹
    filelist = os.listdir(rootdir)  # 列出该目录下的所有文件名
    for f in filelist:
        filepath = os.path.join(rootdir, f)  # 将文件名映射成绝对路劲
        if os.path.isfile(filepath):  # 判断该文件是否为文件或者文件夹
            os.remove(filepath)  # 若为文件，则直接删除
            # print(str(filepath)+" removed!")
        elif os.path.isdir(filepath):
            shutil.rmtree(filepath, True)  # 若为文件夹，则删除该文件夹及文件夹内所有文件
            # print("dir "+str(filepath)+" removed!")
    shutil.rmtree(rootdir, True)  # 最后删除img总文件夹