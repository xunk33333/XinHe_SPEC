import os
import re
from extract.cellsearch import exceltool


def extractTableBGA(pdfPath, pageNumber, tableselectRec, outputPath):
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
    matchKeyNameValue(exe,data,keyword = 'b',name = 'Pads_size')
    matchKeyNameValue(exe,data,keyword = 'm',name = 'Package_length')
    matchKeyNameValue(exe,data,keyword = 'D',name = 'Package_width')


    #针对D/E特殊情况
    if data['Package_width'] is  None and data['Package_length'] is  None:
        matchKeyNameValue(exe,data,keyword = 'D/E',name = 'Package_width')

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