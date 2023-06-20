import pandas as pd


class Table(object):
    def __init__(self,table) -> None:
        self.tableData = table
    
    def filterTableData(self):
        def filterTableData__(x):
            #无内容的置为None
            if x is None or x.__contains__('–') or x.__contains__('—') or x.__contains__('-'):
                return None
            
            #处理1.35/0.0532类型带 / 的数值数据 只保留前面
            if x.__contains__("/") and x.__contains__("."):
                x = x[:x.index("/")]

            return x.replace(" ","").replace("\n","").replace("BSC","").replace("TYP","")
        self.tableData = [[filterTableData__(x) for x in y] for y in self.tableData]

    def tryReverse(self):
        for row in self.tableData:
            row = [x for x in row if x is not None]
            rowstr ="".join(row).upper()
            if rowstr.__contains__("MIN") and rowstr.__contains__("MAX"):
                return
        self.tableData = list(map(list,zip(*self.tableData))) ##列表转置操作
        print("####该表进行了行列转换#####")
    
    def toExcel(self,outputPath):
            df_detail = pd.DataFrame(self.tableData[1:], columns=self.tableData[0])
            df_detail.to_excel(excel_writer=outputPath, index=False, encoding='utf-8')