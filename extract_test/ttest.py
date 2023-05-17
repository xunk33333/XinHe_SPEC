import pickle



if __name__ == '__main__':
    # 存储变量的文件的名字
    filename = 'linedata.data'
    if 10 == 1:
        alldata = extract_with_paddle(r'SGST-S-A0005272207-1/26.png')

        # 以二进制写模式打开目标文件
        f = open(filename, 'wb')
        # 将变量存储到目标文件中区
        pickle.dump(alldata, f)
        # 关闭文件
        f.close()
    else:
        # 以二进制读模式打开目标文件
        f = open(filename, 'rb')
        # 将文件中的变量加载到当前工作区
        lines = pickle.load(f)
    lines_copy = lines.copy()
    for line in lines_copy:
        line['text'] = concat_text(line['spans'])
        if line['spans'].__len__() == 0:
            lines.remove(line)

    concated_lines = []
    # 1 比较字体大小，找到下标字体.
    standard_font_size = lines[0]['spans'][0]['size']
    lines_copy = lines.copy()
    for line in lines_copy:
        if standard_font_size - line['spans'][0]['size'] > 1:  # 2 根据角度dir，区分拼接顺序与筛选条件：5与50
            concated_line = concat_subscript_text(lines, line)
            concated_lines.append(concated_line)
    lines +=concated_lines

    print(lines)
