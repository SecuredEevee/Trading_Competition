import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic, QtGui, QtWidgets, QtCore
from PyQt5.QtCore import Qt, QDate
import pandas as pd
import json
import numpy as np
import time
from bs4 import BeautifulSoup
import requests


def get_bs_obj(stock_code):
    url = "https://finance.naver.com/item/main.nhn?code=" + stock_code
    result = requests.get(url)
    bs_obj = BeautifulSoup(result.content, "html.parser")  # html.parser 로 파이썬에서 쓸 수 있는 형태로 변환
    return bs_obj

def get_KRX_price(stock_code):
    bs_obj = get_bs_obj(stock_code)
    no_today = bs_obj.find("p", {"class": "no_today"})
    blind_now = no_today.find("span", {"class": "blind"})
    return int(blind_now.text.replace(',',''))

json_data =''
isInit = False
df = ''
last_date =''
try:
    with open('result.json', encoding='utf-8') as json_file:
        json_data = json.load(json_file)

    df = pd.DataFrame(json_data)
    with open('info.json') as json_file:
        json_data = json.load(json_file)
        last_date = json_data['last_date']

except Exception as e:
    isInit = True


class MyWindow(QMainWindow):
    def __init__(self, dataframe):
        super(MyWindow, self).__init__()
        self.df = dataframe
        uic.loadUi('mainwindow.ui', self)
        self.setWindowTitle('Trading Competition')
        self.save_button = self.findChild(QtWidgets.QPushButton, 'save_button')  # Find the button
        self.refresh_button = self.findChild(QtWidgets.QPushButton, 'refresh_button')
        self.refresh_rank_button = self.findChild(QtWidgets.QPushButton, 'refresh_rank_button')
        self.load_from_excel_button = self.findChild(QtWidgets.QPushButton, 'load_from_excel_button')
        self.screenshot_button = self.findChild(QtWidgets.QPushButton, 'screenshot_button')

        self.table_view = self.findChild(QtWidgets.QTableView, 'tableView')
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.start_date = self.findChild(QtWidgets.QDateEdit, 'start_date')
        self.last_date = self.findChild(QtWidgets.QDateEdit, 'last_date')
        self.avg_ratio = self.findChild(QtWidgets.QTextBrowser, 'ratiotextBrowser')

        if isInit:
            self.last_date.setDate(QDate.currentDate())
            self.start_date.setDate(QDate.currentDate())
            reply = QMessageBox.question(self, 'Init Open', 'You should load excel file first',
                                         QMessageBox.Yes, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                self.load_from_excel()

                self.start_date.setDate(QDate.currentDate())
                self.last_date.setDate(QDate.currentDate())
                self.refreshView(isInit)
        else:
            self.get_table()
            self.start_date.setDate(QDate.fromString(json_data['start_date'], 'yyyy-MM-dd'))
            self.last_date.setDate(QDate.fromString(json_data['last_date'], 'yyyy-MM-dd'))


        self.save_button.clicked.connect(self.saveButtonPressed)  # Remember to pass the definition/method, not the return value!

        self.refresh_button.clicked.connect(self.refreshResult)
        self.refresh_rank_button.clicked.connect(self.refreshRank)
        self.load_from_excel_button.clicked.connect(self.load_from_excel)
        self.screenshot_button.clicked.connect(self.take_a_screenshot)

        # selection_model = self.table_view.selectionModel()
        # selection_model.currentChanged.connect(self.on_currentChanged)
        self.table_view.model().dataChanged.connect(self.on_dataChanged)
        # selection_model.selectionChanged.connect(self.on_selectionChanged)

    # @QtCore.pyqtSlot('QItemSelection', 'QItemSelection')
    # def on_selectionChanged(self, selected, deselected):
    #     print("selected: ")
    #     for ix in selected.indexes():
    #         print(ix.data())
    #
    #     print("deselected: ")
    #     for ix in deselected.indexes():
    #         print(ix.data())

    # @QtCore.pyqtSlot('QModelIndex', 'QModelIndex')
    # def on_currentChanged(self, current, previous):
    #     print('data changed', current, previous)

    @QtCore.pyqtSlot('QModelIndex', 'QModelIndex')
    def on_dataChanged(self, topleft, bottomright):
        # print(topleft.row(), topleft.column(), topleft.data())
        # print(bottomright.row(), bottomright.column(), bottomright.data())
        if topleft.column() == 11: # hold_date column
            self.df.at[str(topleft.row()), 'hold_date'] = topleft.data()

    def load_from_excel(self):
        fname = QFileDialog.getOpenFileName(self, 'Open base xlsx file', './')
        if fname[0] == '':
            self.load_from_excel()
            return
        dataframe = pd.read_excel(fname[0],dtype={'stock_code':int})
        for idx, row in dataframe.iterrows():
            if pd.isna(row['std_price']):
                dataframe.at[idx, 'std_price'] = int(get_KRX_price('{0:06d}'.format(row['stock_code'])))
                time.sleep(0.2)
            else:
                dataframe.at[idx, 'std_price'] = int(row['std_price'])



        dataframe = dataframe.replace(np.nan, '', regex=True)
        dataframe['last_price'] = dataframe['std_price']
        dataframe['ratio (%)'] = (dataframe['last_price'] - dataframe['std_price']) / dataframe['std_price'] * 100
        dataframe['rank'] = 1
        dataframe['rank_list']= [[1] for _ in range(len(dataframe))]
        # print(df)
        self.df = dataframe.copy()
        self.refreshView(True)


    def take_a_screenshot(self):
        filename = self.last_date.date().toString('yyyy-MM-dd_result.png')
        screen = QtGui.QScreen.grabWindow(app.primaryScreen(), window.winId())  # (메인화면, 현재위젯)
        screen.save(filename, 'png')

    def insert_row(self):
        pass

    def get_table(self):

        self.df['rank'] = self.df['prev_rank'] = 0

        for idx, row in self.df.iterrows():
            self.df.at[idx, 'rank'] = row['rank_list'][-1]
            if len(self.df['rank_list'] ) > 1:
                self.df.at[idx, 'prev_rank'] = row['rank_list'][-2]

        self.df['ratio (%)'] = (self.df['last_price'] - self.df['std_price']) / self.df['std_price'] * 100

        self.refreshView()

    def refreshView(self, _isInit=False):
        if not _isInit:
            self.table_view.setModel(PandasTableModel(self.df.reindex(columns=['rank', 'prev_rank','stock',
                                                                          'name', 'type_of_business',
                                                                          'theme', 'std_price', 'last_price',
                                                                          'ratio (%)', 'market', 'pick_reason',
                                                                          'hold_date']).copy()))
        else:
            self.table_view.setModel(PandasTableModel(self.df.reindex(columns=['rank','stock', 'name',
                                                                          'type_of_business', 'theme',
                                                                          'std_price', 'last_price',
                                                                          'ratio (%)', 'market', 'pick_reason',
                                                                          'hold_date']).copy()))

    def refreshResult(self, donotRefreshView=False):
        # print('refresh')
        for index, row in self.df.iterrows():
            # print(row['last_price'])
            if row['hold_date'] != '':
                continue
            self.df.at[index, 'last_price'] = get_KRX_price('{0:06d}'.format(row['stock_code']))
            self.df.at[index, 'ratio (%)'] = (self.df.at[index, 'last_price'] - row['std_price']) / row['std_price'] * 100

            # print(row['last_price'])
            time.sleep(0.2)


        self.df['rank'] = self.df['ratio (%)'].rank(ascending=False, method='min')
        ratio = self.df['ratio (%)'].mean()

        cursor = self.avg_ratio.textCursor()
        if ratio > 0:
            cursor.insertHtml('''<p><span style="color: red;">{0:.2f}%</span>'''.format((ratio)))
        elif ratio < 0:
            cursor.insertHtml('''<p><span style="color: blue;">{0:.2f}%</span>'''.format((ratio)))
        else:
            cursor.insertHtml('''<p><span style="color: black;">{0:.2f}%</span>'''.format((ratio)))
        self.last_date.setDate(QDate.currentDate())
        if not donotRefreshView:
            self.refreshView()

    def refreshRank(self):
        global last_date
        today = QDate.currentDate().toString('yyyy-MM-dd')
        if last_date == today :
            if self.start_date.date() != self.last_date.date():
                QMessageBox.critical(self, '순위 갱신 에러', '순위 갱신은 내일 해주세요')
                print('순위 갱신은 내일 해주세요')
                return
            else: # 첫날 인 경우 중에 두 번 이상 갱신 하려 하는 경우 리턴
                if len(self.df.at[0, 'rank_list']) > 1:
                    return

        self.df['prev_rank'] = self.df['rank']
        self.refreshResult(donotRefreshView=True)
        for index, row in self.df.iterrows():
            self.df.at[index, 'rank_list'].append(int(row['rank']))
        self.refreshView()
        last_date = today


    def saveButtonPressed(self):
        # This is executed when the button is pressed
        date_info= {'start_date': self.start_date.date().toString('yyyy-MM-dd'),
                    'last_date': self.last_date.date().toString('yyyy-MM-dd')}
        with open('info.json', 'w', encoding='utf-8') as json_file:
            json.dump(date_info, json_file, indent=4)

        df_copy = self.df.copy().drop(columns=['ratio (%)', 'rank', 'prev_rank'])
        df_copy.to_json('result.json',force_ascii=False, indent=4)

        msgBox = QtWidgets.QMessageBox.information(self, ':)', 'Save complete')


class StandardItem(QtGui.QStandardItem):
    def __lt__(self, other):
        return float(self.text()) < float(other.text())

class PandasTableModel(QtGui.QStandardItemModel):
    def __init__(self, data, parent=None):
        QtGui.QStandardItemModel.__init__(self, parent)
        self._data = data
        # self._data = self.df
        # for row in self.df.values.tolist():
        data_list = data.values.tolist()
        total_num_of_person = len(data_list)
        for row in data_list:
            data_row = []
            i = 0
            if len(row) >= 12:
                i+=1
            data_row.append(StandardItem(str(int(row[0]))))
            if i == 1:
                if not pd.isna(row[1]): # If Prev rank existed
                    data_row.append(StandardItem(str(int(row[1]))))
                else:
                    data_row.append(StandardItem(''))
            data_row.extend(QtGui.QStandardItem("{}".format(x)) for x in row[i+1:i+5])
            data_row.append(StandardItem(format(int(row[i+5]), ','))) # std_price
            last_price_item = StandardItem(format(int(row[i+6]), ','))
            if row[i+5] < row[i+6]:
                last_price_item.setForeground(QtGui.QColor('red'))
            elif row[i+5] > row[i+6]:
                last_price_item.setForeground(QtGui.QColor('blue'))
            data_row.append(last_price_item)
            item = StandardItem( '{:.3f}'.format(row[i+7])) # ratio
            if row[i+7] > 0:
                item.setForeground(QtGui.QColor('red'))
            elif row[i+7] < 0:
                item.setForeground(QtGui.QColor('blue'))
            data_row.append(item)
            data_row.extend(QtGui.QStandardItem("{}".format(x)) for x in row[i+8:])
            for x in data_row:
                if int(row[0]) == 1 :
                    x.setBackground(QtGui.QColor(255, 211, 0))
                    # x.setFont(QtGui.QFont('Gothic', 12, QtGui.QFont.Bold))
                elif int(row[0]) == 2 :
                    x.setBackground(QtGui.QColor(192, 192, 192))
                elif int(row[0]) == 3 :
                    x.setBackground(QtGui.QColor(205, 127, 50))
                elif int(row[0]) <= 10:
                    x.setBackground(QtGui.QColor(204,255,255))
                else:
                    if total_num_of_person - 10 < int(row[0]) <= total_num_of_person:
                        x.setBackground(QtGui.QColor(241,156,187))

                x.setTextAlignment(Qt.AlignCenter)
            self.appendRow(data_row)
            self.sort(0, Qt.AscendingOrder)
        return

    def rowCount(self, parent=None):
        return len(self._data.values)

    def columnCount(self, parent=None):
        return self._data.columns.size

    def headerData(self, x, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[x]
        if orientation == Qt.Vertical and role == Qt.DisplayRole:
            return self._data.index[x]
        return None

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow(df)
    window.show()
    app.exec_()