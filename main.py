import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic, QtCore, QtGui, QtWidgets
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
try:
    with open('result.json', encoding='utf-8') as json_file:
        json_data = json.load(json_file)

    df = pd.DataFrame(json_data)
    with open('info.json') as json_file:
        json_data = json.load(json_file)

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
        self.save_to_excel_button = self.findChild(QtWidgets.QPushButton, 'save_to_excel_button')

        self.table_view = self.findChild(QtWidgets.QTableView, 'tableView')
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.start_date = self.findChild(QtWidgets.QDateEdit, 'start_date')
        self.last_date = self.findChild(QtWidgets.QDateEdit, 'last_date')

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


        self.save_button.clicked.connect(
            self.saveButtonPressed)  # Remember to pass the definition/method, not the return value!

        self.refresh_button.clicked.connect(self.refreshResult)
        self.refresh_rank_button.clicked.connect(self.refreshRank)
        self.load_from_excel_button.clicked.connect(self.load_from_excel)


    def load_from_excel(self):
        fname = QFileDialog.getOpenFileName(self, 'Open base xlsx file', './')
        if fname[0] == '':
            self.load_from_excel()
            return
        dataframe = pd.read_excel(fname[0],dtype={'stock_code':int})
        for idx, row in dataframe.iterrows():
            if pd.isna(row['std_price']):
                dataframe.at[idx, 'std_price'] = get_KRX_price('{0:06d}'.format(row['stock_code']))
                time.sleep(0.3)
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


    def save_to_excel(self):
        pass

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

        self.df['rank'] = self.df['ratio (%)'].rank(ascending=False)
        self.last_date.setDate(QDate.currentDate())
        if not donotRefreshView:
            self.refreshView()

    def refreshRank(self):

        today = QDate.currentDate()
        if self.last_date.date() == today :
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
        for row in data.values.tolist():
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
            data_row.append(StandardItem(str(row[i+5]))) # std_price
            last_price_item = StandardItem(str(row[i+6]))
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