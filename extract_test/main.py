import sys
import os
_BASEDIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(_BASEDIR))
from PyQt5.QtWidgets import QApplication

from extract_test.pdfreader import PDFReader

app = QApplication(sys.argv)
reader = PDFReader()
sys.exit(app.exec_())