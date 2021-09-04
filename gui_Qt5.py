import sys, os
from PyQt5.QtWidgets import QApplication, QMainWindow, QListWidget, QListWidgetItem, QPushButton
from PyQt5.QtCore import Qt, QUrl
from PyQt5 import QtGui


class ListboxWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.resize(600, 600)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()

            links = []
            print(f'{event.mimeData().urls()}\n')  # prints in console

            for url in event.mimeData().urls():
                if url.isLocalFile():
                    links.append(url.toString())
                    # links.append(url.toLocalFile())
                else:
                    # links.append(url.toString())
                    pass

            self.addItems(links)

        else:
            event.ignore()


class AppDemo(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('NDT V1.0')
        # self.setWindowIcon('C:\\Users\\hE\\Pictures\\pgcb.png')
        self.setWindowIcon(QtGui.QIcon('C:\\Users\\hE\\Pictures\\pgcb.png'))
        self.resize(1200, 600)

        self.listbox_view = ListboxWidget(self)  # uses the previous class

        self.btn = QPushButton('Get Value', self)
        self.btn.setGeometry(1000 - 5, 200, 200, 50)
        self.btn.clicked.connect(lambda: print(self.getSelectedItem()))  # prints in GUI

    def getSelectedItem(self) -> object:
        item = QListWidgetItem(self.listbox_view.currentItem())
        return item.text()



app = QApplication(sys.argv)
demo = AppDemo()
demo.show()

sys.exit(app.exec_())
