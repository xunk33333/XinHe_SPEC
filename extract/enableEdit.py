import os
def pipline(file_path,page_num):
    scale_factor = get_img_from_pdf(file_path,page_num)
    bboxs = get_bbox(file_path)
    table = extract_table_by_pdfplumber(file_path,page_num,bboxs,scale_factor)
    # table_toexcel(table,"tmp.xlsx")
    return table


def get_img_from_pdf(file_path,page_num):
    import fitz
    doc = fitz.open(file_path)
    page = doc[page_num]
    # name = file_path[file_path.rindex('/') + 1:-4]
    scale_factor = 3
    pix = page.get_pixmap(matrix=fitz.Matrix(scale_factor, scale_factor).prerotate(0))
    img_path = r"{}/{}.png".format("tableDet", page_num)
    if not os.path.exists("tableDet"):  # 判断存放图片的文件夹是否存在
        os.makedirs("tableDet")  # 若图片文件夹不存在就创建
    pix.save(img_path)

    page = doc[page_num+1]
    pix = page.get_pixmap(matrix=fitz.Matrix(scale_factor, scale_factor).prerotate(0))
    img_path = r"{}/{}.png".format("tableDet", page_num+1)
    pix.save(img_path)

    return scale_factor

def get_bbox(file_path):
    # name = file_path[file_path.rindex('/') + 1:-4]
    from PaddleTabDet.predict_layout import main
    bboxs = main("tableDet")
    return bboxs

def extract_table_by_pdfplumber(pdfPath,pageNumber,bboxs,scale_factor):
    import pdfplumber
    pdf = pdfplumber.open(pdfPath)                            # readpath:PDF所在位置
    page0 = pdf.pages[pageNumber]
    page0 = page0.crop([x/scale_factor for x in bboxs[0][0]])
    page1 = pdf.pages[pageNumber+1]
    page1 = page1.crop([x/scale_factor for x in bboxs[1][0]])
    table_setting = {
    "vertical_strategy": "lines", 
    "horizontal_strategy": "lines",
    "explicit_vertical_lines": [],
    "explicit_horizontal_lines": [],
    "snap_tolerance": 5,
    "join_tolerance": 3,
    "edge_min_length": 3,
    "min_words_vertical": 1,
    "min_words_horizontal": 1,
    "keep_blank_chars": False,
    "text_tolerance": 0,
    "text_x_tolerance": 0,
    "text_y_tolerance": 0,
    "intersection_tolerance": 3,
    "intersection_x_tolerance": 3,
    "intersection_y_tolerance": 3,
    }
    table0 = page0.extract_table(table_setting)
    # print(table0)
    table1 = page1.extract_table(table_setting)
    # print(table1)
    return table0+table1

def table_toexcel(table,outputPath):
    import pandas as pd
    df_detail = pd.DataFrame(table[1:], columns=table[0])
    df_detail.to_excel(excel_writer=outputPath, index=False, encoding='utf-8')