# 主程序
import sys
import traceback


import cv2
from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QVBoxLayout, QFileDialog, QInputDialog, QLabel, \
    QApplication, QMessageBox
import os
import fitz
import pdfplumber
from extract.auto_bga import extractTableBGA

from extract.auto_qfn import extractTableQFN,del_dir
from extract_test.mydatabase import MyDb
from extract_test.manual_boxes import DrawRects, WIN_NAME, onmouse_draw_rect


class PDFReader(QMainWindow):

    def __init__(self):
        super(PDFReader, self).__init__()
        self.db = MyDb()  # 数据库
        self.menubar = self.menuBar()
        self.recentfile = None
        self.generateMenuBar()
        self.generateRecentMenu()
        self.recentfile.triggered.connect(self.onRecentFileClicked)  # 此句话修复了bug
        self.toolbar = self.addToolBar("工具栏")
        self.generateToolBar()
        layout = QVBoxLayout(self)
        self.toc = QTreeWidget()
        self.toc.setFont(QFont("", 13))  # 目录文字的字体及其大小控制，修改字体将字体名放入双引号中（为空表示使用默认字体），字体大小修改数字即可，数字越大字体越大
        self.file_path = ""
        self.page_num = 0
        self.doc = None
        self.book_open = False
        self.note_keeped = ""
        # self.note_loadFromFile = False
        self.note_path = ""
        self.dock = QDockWidget()
        self.generateDockWidget()
        self.addDockWidget(Qt.LeftDockWidgetArea, self.dock)
        self.dock.setVisible(False)
        self.trans_a = 200
        self.trans_b = 200
        self.trans = fitz.Matrix(self.trans_a / 100, self.trans_b / 100).prerotate(0)
        self.scrollarea = QScrollArea(self)
        self.pdfview = QLabel()
        self.tocDict = {}
        self.scrollarea.setWidget(self.pdfview)
        self.generatePDFView()
        self.pagedisplay = QLabel()

        layout.addWidget(self.scrollarea)
        layout.addWidget(self.pagedisplay)

        self.widget = QWidget()
        self.widget.setLayout(layout)
        self.setCentralWidget(self.widget)
        self.setWindowTitle('PDF Reader')
        desktop = QApplication.desktop()
        rect = desktop.availableGeometry()
        self.setGeometry(rect)
        self.setWindowIcon(QIcon('extract_test/icon/reader.png'))
        # self.setGeometry(100, 100, 1000, 600)
        self.show()

    def generatePDFView(self):
        if not self.file_path or not self.doc:
            return
        pix = self.doc[self.page_num].get_pixmap(matrix=self.trans)
        fmt = QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888
        pageImage = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
        pixmap = QPixmap()
        pixmap.convertFromImage(pageImage)
        self.pdfview.setPixmap(QPixmap(pixmap))
        self.pdfview.resize(pixmap.size())

        self.pagedisplay.setFont(QFont('', 22))
        self.pagedisplay.setText('{}/{}'.format(self.page_num + 1, self.doc.page_count))

    def generateFile(self):
        file = self.menubar.addMenu('文件')
        file.setFont(QFont("", 13))
        openFile = QAction(QIcon('extract_test/icon/file.png'), '打开文件', file)
        closeFile = QAction("关闭文件", file)

        openFile.triggered.connect(self.onOpen)
        closeFile.triggered.connect(self.onClose)

        file.addAction(openFile)
        self.recentfile = file.addMenu('最近文件')
        file.addAction(closeFile)

    def generateRecentMenu(self):
        self.recentfile.clear()
        fileList = self.db.getAllRencentFile()
        sortlist = sorted(fileList, key=lambda d: d.opentime, reverse=True)
        for file in sortlist:
            action = QAction(file.path, self.recentfile)
            self.recentfile.addAction(action)
        # self.recentfile.triggered.connect(self.onRecentFileClicked)

    def onRecentFileClicked(self, action):
        self.open_file(action.text())

    def generateMenuBar(self):
        self.menubar.setFont(QFont("", 13))  # 设置菜单栏字体大小
        self.generateFile()

    def generateToolBar(self):
        self.toolbar.setMinimumSize(QSize(200, 200))
        self.toolbar.setIconSize(QSize(100, 100))  # 设置工具栏图标大小
        # self.toolbar.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)  # 文字在图标旁边
        # self.toolbar.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)  # 文字在图标下方
        # 不设置以上两句话默认只显示图标
        ToC = QAction(QIcon('extract_test/icon/目录 (5).png'), '目录', self.toolbar)
        openFile = QAction(QIcon('extract_test/icon/file.png'), '打开文件', self.toolbar)
        prePage = QAction(QIcon('extract_test/icon/分页 下一页 (1).png'), '上一页', self.toolbar)
        nextPage = QAction(QIcon('extract_test/icon/分页 下一页.png'), '下一页', self.toolbar)
        turnPage = QAction(QIcon('extract_test/icon/跳转.png'), '跳转', self.toolbar)
        enlargePage = QAction(QIcon('extract_test/icon/放大 (1).png'), '放大', self.toolbar)
        shrinkPage = QAction(QIcon('extract_test/icon/缩小.png'), '缩小', self.toolbar)
        Pdf2txt = QAction(QIcon('extract_test/icon/阅读.png'), 'BGA提取', self.toolbar)
        extractTable = QAction(QIcon('extract_test/icon/阅读.png'), 'QFN提取', self.toolbar)

        nextPage.setShortcut(Qt.Key_Right)
        prePage.setShortcut(Qt.Key_Left)

        openFile.triggered.connect(self.onOpen)
        ToC.triggered.connect(self.onDock)
        prePage.triggered.connect(self.onPrepage)
        nextPage.triggered.connect(self.nextpage)
        turnPage.triggered.connect(self.turnpage)
        enlargePage.triggered.connect(self.enlargepage)
        shrinkPage.triggered.connect(self.shrinkpage)
        Pdf2txt.triggered.connect(self.pdf2txt)
        extractTable.triggered.connect(self.extracttable)

        self.toolbar.addSeparator()
        self.toolbar.addActions([ToC])
        self.toolbar.addSeparator()
        self.toolbar.addActions([openFile])
        self.toolbar.addSeparator()
        self.toolbar.addActions([prePage, nextPage, turnPage])
        self.toolbar.addSeparator()
        self.toolbar.addActions([enlargePage, shrinkPage])
        self.toolbar.addSeparator()
        self.toolbar.addActions([Pdf2txt])
        self.toolbar.addSeparator()
        self.toolbar.addActions([extractTable])
        self.toolbar.addSeparator()

    def generateDockWidget(self):
        if not self.file_path:
            return
        self.dock.setWidget(self.toc)
        self.generateTreeWidget()

    def generateTreeWidget(self):
        if not self.doc:
            return
        self.toc.setColumnCount(1)
        self.toc.setHeaderLabels(['目录'])
        # tree.setMinimumSize(500, 500)
        self.toc.setWindowTitle('目录')
        toc = self.doc.get_toc()
        nodelist = [self.toc]
        floorlist = [0]
        tempdict = {}
        if not toc:
            return tempdict
        first = True
        for line in toc:
            floor, title, page = line
            if first:
                node = QTreeWidgetItem(self.toc)
                node.setText(0, title)
                nodelist.append(node)
                floorlist.append(floor)
                tempdict[title] = page
                first = False
            else:
                while floorlist[-1] >= floor:
                    nodelist.pop()
                    floorlist.pop()
                node = QTreeWidgetItem(nodelist[-1])
                node.setText(0, title)
                nodelist.append(node)
                floorlist.append(floor)
                tempdict[title] = page
        self.tocDict = tempdict
        self.toc.clicked.connect(self.bookmark_jump)

    def bookmark_jump(self, index):
        item = self.toc.currentItem()
        self.page_num = self.tocDict[item.text(0)] - 1
        self.updatePdfView()

    def onDock(self):
        try:
            if self.dock.isVisible():
                self.dock.setVisible(False)
            else:
                self.dock.setVisible(True)
        except AttributeError:
            pass

    def onDoc2(self):
        try:
            if self.doc2.isVisible():
                self.doc2.setVisible(False)
            else:
                self.doc2.setVisible(True)
        except AttributeError:
            pass

    def onOpen(self):
        fDialog = QFileDialog()
        filename, _ = fDialog.getOpenFileName(self, "打开文件", ".", 'PDF file (*.pdf)')
        self.open_file(filename)

    def open_file(self, filename):
        if not filename:
            return
        if not self.db.fileInDB(filename):
            self.db.addRecentFile(filename)
        else:
            self.db.updateRecentFile(filename)
        if os.path.exists(filename):
            self.file_path = filename
        else:
            QMessageBox.about(self, "提醒", "文件不存在")
            self.db.deleteRecentFile(filename)
        self.toc.clear()
        self.page_num = 0
        self.book_open = True
        self.generateRecentMenu()
        self.getDoc()
        self.generateDockWidget()
        self.generatePDFView()

    def getDoc(self):
        if self.file_path:
            self.doc = fitz.open(self.file_path)

    def onClose(self):
        self.file_path = ""
        self.book_open = False
        self.toc.clear()
        self.pdfview.clear()
        self.getDoc()
        self.generatePDFView()
        self.generateDockWidget()

    def onPrepage(self):
        self.page_num -= 1
        if self.page_num < 0:
            self.page_num += self.doc.page_count
        self.updatePdfView()

    def updatePdfView(self):
        self.scrollarea.verticalScrollBar().setValue(0)
        self.generatePDFView()

    def nextpage(self):
        self.page_num += 1
        if self.page_num >= self.doc.page_count:
            self.page_num -= self.doc.page_count
        self.updatePdfView()

    def turnpage(self):
        if not self.book_open:
            return
        allpages = self.doc.page_count
        print(allpages)
        page, ok = QInputDialog.getInt(self, "选择页面", "输入目标页面({}-{})".format(1, allpages), min=1, max=allpages)
        if ok:
            self.page_num = page - 1
            self.updatePdfView()

    def enlargepage(self):
        self.trans_a += 5  # 每次放大增加5%,修改此参数时注意和下方的tran_b保持一致，否则图片会变形
        self.trans_b += 5
        self.trans = fitz.Matrix(self.trans_a / 100, self.trans_b / 100).prerotate(0)
        self.generatePDFView()

    def shrinkpage(self):
        self.trans_a -= 5
        self.trans_b -= 5
        self.trans = fitz.Matrix(self.trans_a / 100, self.trans_b / 100).prerotate(0)
        self.generatePDFView()

    def pdf2txt(self):
        page = self.doc[self.page_num]
        name = self.file_path[self.file_path.rindex('/') + 1:-4]
        scale_factor = 4
        pix = page.get_pixmap(matrix=fitz.Matrix(scale_factor, scale_factor).prerotate(0))
        img_path = r"{}/{}.png".format(name, self.page_num)
        if not os.path.exists(name):  # 判断存放图片的文件夹是否存在
            os.makedirs(name)  # 若图片文件夹不存在就创建
        pix.save(img_path)

        image = cv2.imread(img_path)
        image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        draw_rects = DrawRects(image, (0, 255, 0), 2)
        cv2.namedWindow(WIN_NAME, 0)
        cv2.setMouseCallback(WIN_NAME, onmouse_draw_rect, draw_rects)

        while True:
            cv2.imshow(WIN_NAME, draw_rects.image_for_show)
            key = cv2.waitKey(30)
            if key == 13:  # Enter    
                break
        cv2.destroyAllWindows()

        
        
        bugBox = draw_rects.rects[0].tl + draw_rects.rects[0].br
        print("圈得到得区域范围:" + str(bugBox))

        box = [x / scale_factor for x in bugBox]

        pdfPath = self.file_path
        outputPath = f"excel"
        pageNumber = self.page_num + 1
        selectRec = box
        try:
            extractTableBGA(pdfPath, pageNumber, selectRec, outputPath)
        except:
            traceback.print_exc()
        finally:
            del_dir('{}'.format(name))
    def extracttable(self):
        # page = self.doc[self.page_num]
        # name = self.file_path[self.file_path.rindex('/') + 1:-4]
        # scale_factor = 4
        # pix = page.get_pixmap(matrix=fitz.Matrix(scale_factor, scale_factor).prerotate(0))
        # img_path = r"{}/{}.png".format(name, self.page_num)
        # if not os.path.exists(name):  # 判断存放图片的文件夹是否存在
        #     os.makedirs(name)  # 若图片文件夹不存在就创建
        # pix.save(img_path)

        # image = cv2.imread(img_path)
        # image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        # draw_rects = DrawRects(image, (0, 255, 0), 2)
        # cv2.namedWindow(WIN_NAME, 0)
        # cv2.setMouseCallback(WIN_NAME, onmouse_draw_rect, draw_rects)

        # while True:
        #     cv2.imshow(WIN_NAME, draw_rects.image_for_show)
        #     key = cv2.waitKey(30)
        #     if key == 13:  # Enter    
        #         break
        # cv2.destroyAllWindows()

        
        
        # bugBox = draw_rects.rects[0].tl + draw_rects.rects[0].br
        # print("圈得到得区域范围:" + str(bugBox))

        # box = [x / scale_factor for x in bugBox]

        pdfPath = self.file_path
        outputPath = f"result"
        pageNumber = self.page_num + 1
        # tableselectRec = box
        try:
            # extractTableQFN(pdfPath, pageNumber, tableselectRec, outputPath)
            extractTableQFN(pdfPath, pageNumber, [], outputPath)
        except:
            traceback.print_exc()
        finally:
            # del_dir('{}'.format(name))
            pass

        

if __name__ == '__main__':
    app = QApplication(sys.argv)
    reader = PDFReader()
    sys.exit(app.exec_())
