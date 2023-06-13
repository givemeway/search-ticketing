import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QAction, QMenu, QTableWidget, 
                             QMainWindow, QTableWidgetItem, QVBoxLayout, )
from PyQt5 import QtCore


class TestRightClickTableWidget(QWidget): #(QMainWindow):   # (QWidget): #

    def __init__(self):
        super().__init__()

        self.tableWidget = QTableWidget()
        self.tableWidget.setRowCount(4)
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setItem(0, 0, QTableWidgetItem("Cell 1"))
        self.tableWidget.setItem(0, 1, QTableWidgetItem("Cell 2"))
        self.tableWidget.setItem(1, 0, QTableWidgetItem("Cell 3"))
        self.tableWidget.setItem(1, 1, QTableWidgetItem("Cell 4"))
        self.tableWidget.setItem(2, 0, QTableWidgetItem("Cell 5"))
        self.tableWidget.setItem(2, 1, QTableWidgetItem("Cell 6"))
        self.tableWidget.setItem(3, 0, QTableWidgetItem("Cell 7"))
        self.tableWidget.setItem(3, 1, QTableWidgetItem("Cell 8"))
        
        ### This property holds how the widget shows a context menu
        self.tableWidget.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)     # +++
        ### This signal is emitted when the widget's contextMenuPolicy is Qt::CustomContextMenu, 
        ### and the user has requested a context menu on the widget. 
        self.tableWidget.customContextMenuRequested.connect(self.generateMenu) # +++

        self.tableWidget.viewport().installEventFilter(self)

        self.layout = QVBoxLayout() 
        self.layout.addWidget(self.tableWidget)
        self.setLayout(self.layout)

    def eventFilter(self, source, event):
        if(event.type() == QtCore.QEvent.MouseButtonPress and
           event.buttons() == QtCore.Qt.RightButton and
           source is self.tableWidget.viewport()):
            item = self.tableWidget.itemAt(event.pos())
            print('Global Pos:', event.globalPos())
            if item is not None:
                print('Table Item:', item.row(), item.column())
                self.menu = QMenu(self)
                self.menu.addAction(item.text())         #(QAction('test'))
                #menu.exec_(event.globalPos())
        return super(TestRightClickTableWidget, self).eventFilter(source, event)

    ### +++    
    def generateMenu(self, pos):
        print("pos======",pos)
        self.menu.exec_(self.tableWidget.mapToGlobal(pos))   # +++

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = TestRightClickTableWidget()
    ex.show()
    sys.exit(app.exec_())