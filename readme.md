# （0626）代码版本改动
1. 新增了BGA的接口 extractTableBGA
2. 修改了之前QFN接口名称 extractTable -> extractTableQFN
3. 测试界面的按钮第一个是BGA，第二个QFN
# （0621）代码版本改动
## 修复问题
1. 修复取NOM的bug（ATTINY24A.pdf ->P14）
2. 去掉了结果中的BSC和TYP
# （0613）代码版本改动
## 修复问题
1. 纵向表头类型识别错误（efm32g-datasheet_5tables_page_190）
2. 当表格中有“/”间隔时，无法识别（ATTINY24A.pdf ->P14，00002304A_p6_page_45）
3. 部分PDF中NOM没有值的情况下，未取平均值（HC32L130F8UA-QFN32_page_73）

# 0607新增对代码（0524）版本反馈
1. 无法提取或异常 （不可编辑需要OCR）
2. 部分Symbol值没有提取（跨页）
3. 提取结果中出现未预期的字符（背景水印）
4. 部分PDF中NOM没有值的情况下，未取平均值
   
# 0530对代码（0524）版本反馈
1. 纵向表头类型，识别错误
2. 跨页表格无法一次框选进行识别匹配（efm32g-datasheet_5tables_P185）
3. 非通用Symbol标注匹配错误（efm32g-datasheet_5tables_P180）
4. 当表格中有“/”间隔时，无法识别（ATTINY24A.pdf ->P14）
5. 个别表格框选若包含表格边框将无法提取（ATTINY24A_P17）
6. C510874_8CE89A1F3CAD019F269ADB0389E1C8E4.pdf -> P430
