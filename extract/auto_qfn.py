import os
import shutil

from extract.cellsearch import exceltool

def extractTable(pdfPath, pageNumber, tableselectRec, outputPath): 
    exe = exceltool()
    exe.get_table_from_pdf(pdfPath, pageNumber, tableselectRec, 'tmp.xlsx')
    data = {}

    #
    keyword = 'e'
    a = exe.cellsearch(keyword,True)
    i = a.row
    j = a.col_idx + 1 
    data['Pads_pitch'] = exe.get_cell(i, j).value

    #
    keyword = 'b'
    a = exe.cellsearch(keyword,True)
    i = a.row
    j = a.col_idx + 2
    data['Pads_width'] = exe.get_cell(i, j).value


    #
    keyword = 'L'
    a = exe.cellsearch(keyword,True)
    i = a.row
    j = a.col_idx + 2
    data['Pads_length'] = exe.get_cell(i, j).value
 
   #
    keyword = 'D2'
    a = exe.cellsearch(keyword,True)
    i = a.row
    j = a.col_idx + 2
    data['EPad_width'] = exe.get_cell(i, j).value

    #
    keyword = 'E2'
    a = exe.cellsearch(keyword,True)
    i = a.row
    j = a.col_idx + 2
    data['EPad_length'] = exe.get_cell(i, j).value
    
    #
    keyword = 'D'
    a = exe.cellsearch(keyword,True)
    i = a.row
    j = a.col_idx + 2
    data['Package_width'] = exe.get_cell(i, j).value
     

    #
    keyword = 'E'
    a = exe.cellsearch(keyword,True)
    i = a.row
    j = a.col_idx + 2
    data['Package_width'] = exe.get_cell(i, j).value
     

   
     
    
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