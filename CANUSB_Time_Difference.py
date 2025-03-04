import sys  # 系統相關
import pandas as pd  # 資料處理
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QVBoxLayout,
                             QFileDialog, QLabel, QTableWidget, QTableWidgetItem,
                             QHBoxLayout, QLineEdit, QGridLayout, QMessageBox, QComboBox)  # 視窗元件
from PyQt5.QtCore import Qt  # 匯入 Qt 模組以便使用對齊方式

import os  # 原xls檔案轉csv處理


class TimeDifferenceApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('時間差計算工具')
        self.setGeometry(100, 100, 800, 600)  # 加大視窗尺寸

        # --- 輸入欄位 ---
        main_input_layout = QHBoxLayout()  # 使用水平佈局來分割左右兩部分

        left_input_layout = QGridLayout()  # 左邊的佈局 (ID + 選單、最小時間差)
        right_input_layout = QGridLayout()  # 右邊的佈局 (報文 + 選單)

        # 設定左右input_layout的伸縮比例為1
        main_input_layout.addLayout(left_input_layout,1)
        main_input_layout.addLayout(right_input_layout,1)

        # --- 左邊輸入欄位 ---

        # ID1
        self.id1_label = QLabel("ID1:", self)
        self.id1_input = QComboBox(self)
        left_input_layout.addWidget(self.id1_label, 0, 0)  # 在左邊佈局的第一列、第一行
        left_input_layout.addWidget(self.id1_input, 0, 1)  # 在左邊佈局的第一列、第二行
        left_input_layout.setAlignment(self.id1_label, Qt.AlignRight | Qt.AlignVCenter)  # 標籤靠右對齊

        # ID2
        self.id2_label = QLabel("ID2:", self)
        self.id2_input = QComboBox(self)
        left_input_layout.addWidget(self.id2_label, 1, 0)  # 在左邊佈局的第二列、第一行
        left_input_layout.addWidget(self.id2_input, 1, 1)  # 在左邊佈局的第二列、第二行
        left_input_layout.setAlignment(self.id2_label, Qt.AlignRight | Qt.AlignVCenter)  # 標籤靠右對齊

        # 最小時間差輸入欄位
        self.min_time_diff_label = QLabel("最小時間差 (秒):", self)
        self.min_time_diff_input = QLineEdit(self)
        self.min_time_diff_input.setText("0.000")  # 預設值為 0
        left_input_layout.addWidget(self.min_time_diff_label, 2, 0)  # 在左邊佈局的第三列、第一行
        left_input_layout.addWidget(self.min_time_diff_input, 2, 1)  # 在左邊佈局的第三列、第二行
        left_input_layout.setAlignment(self.min_time_diff_label, Qt.AlignRight | Qt.AlignVCenter)  # 標籤靠右對齊

        # --- 右邊輸入欄位 ---

        # 報文1
        self.data1_label = QLabel("報文1:", self)
        self.data1_input = QComboBox(self)
        right_input_layout.addWidget(self.data1_label, 0, 0)  # 在右邊佈局的第一列、第一行
        right_input_layout.addWidget(self.data1_input, 0, 1)  # 在右邊佈局的第一列、第二行
        right_input_layout.setAlignment(self.data1_label, Qt.AlignRight | Qt.AlignVCenter)  # 標籤靠右對齊

        # 報文2
        self.data2_label = QLabel("報文2:", self)
        self.data2_input = QComboBox(self)
        right_input_layout.addWidget(self.data2_label, 1, 0)  # 在右邊佈局的第二列、第一行
        right_input_layout.addWidget(self.data2_input, 1, 1)  # 在右邊佈局的第二列、第二行
        right_input_layout.setAlignment(self.data2_label, Qt.AlignRight | Qt.AlignVCenter)  # 標籤靠右對齊

        # --- 按鈕 ---
        self.loadButton = QPushButton('選擇 CSV 檔案', self)
        self.saveButton = QPushButton('選擇存檔路徑', self)
        self.processButton = QPushButton('計算時間差', self)

        self.loadButton.clicked.connect(self.loadFile)
        self.saveButton.clicked.connect(self.saveFile)
        self.processButton.clicked.connect(self.processData)

        # --- 標籤 ---
        self.fileLabel = QLabel('未選擇檔案', self)
        self.saveLabel = QLabel('未選擇存檔位置', self)

        # --- 表格 ---
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["觸發時間", "結束時間", "時間差 (秒)", "報文1發送次數"])

        # --- 主佈局 ---
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.loadButton)
        main_layout.addWidget(self.fileLabel)
        main_layout.addWidget(self.saveButton)
        main_layout.addWidget(self.saveLabel)
        main_layout.addLayout(main_input_layout)  # 新增：加入主水平佈局
        main_layout.addWidget(self.processButton)
        main_layout.addWidget(self.table)
        self.setLayout(main_layout)

        # --- 儲存檔案路徑 ---
        self.file_path = ""
        self.save_path = ""

        # --- 儲存要擷取的 ID 和報文 (現在從輸入欄位取得) ---
        self.CANID1 = ""
        self.CANID2 = ""
        self.CANData1 = ""
        self.CANData2 = ""
        # 新增最小時間
        self.min_time_diff = 0.0

    def loadFile(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "選擇檔案", "",
                                                   "Excel Files (*.xls *.xlsx);;CSV Files (*.csv);;All Files (*)",
                                                   options=options)
        if file_name:
            self.file_path = file_name
            self.fileLabel.setText(f'選擇的檔案: {file_name}')

            # 判斷是否為 Excel 檔案
            if file_name.endswith(('.xls', '.xlsx')):
                try:
                    # 讀取 Excel 檔案
                    df = pd.read_excel(file_name)

                    # 產生暫時的 CSV 檔名
                    temp_csv_file = os.path.splitext(file_name)[0] + ".csv"

                    # 儲存為 CSV 檔案
                    df.to_csv(temp_csv_file, index=False, encoding="utf-8-sig")

                    # 將檔案路徑更改為暫時的 CSV 檔案
                    self.file_path = temp_csv_file
                    self.fileLabel.setText(f'選擇的檔案: {file_name} (已轉換為: {temp_csv_file})')

                except Exception as e:
                    QMessageBox.critical(self, "錯誤", f"轉換 Excel 檔案失敗: {e}")
                    return
            else:
                if file_name:
                    self.file_path = file_name
                    self.fileLabel.setText(f'選擇的檔案: {file_name}')
            # 之後不管是不是xls都直接去抓CSV檔
        self.populateComboBoxes()

    def saveFile(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "選擇存檔位置", "", "CSV Files (*.csv);;All Files (*)",
                                                   options=options)
        if file_name:
            self.save_path = file_name
            if not self.save_path.endswith(".csv"):
                self.save_path += ".csv"
            self.saveLabel.setText(f'存檔位置: {self.save_path}')

    def populateComboBoxes(self):
        try:
            df = pd.read_csv(self.file_path, encoding="utf-8", low_memory=False)
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"讀取 CSV 檔案失敗: {e}")
            return

        id_column = "帧ID"  # 請確保名稱符合你的 CSV
        data_column = "数据"

        ids = df[id_column].unique().astype(str)
        self.id1_input.clear()
        self.id2_input.clear()
        self.id1_input.addItems(ids)
        self.id2_input.addItems(ids)

        # 立即更新報文下拉式選單
        self.updateDataComboBox1()
        self.updateDataComboBox2()

        # 設置信號
        self.id1_input.currentIndexChanged.connect(self.updateDataComboBox1)
        self.id2_input.currentIndexChanged.connect(self.updateDataComboBox2)

    def updateDataComboBox1(self):
        selected_id = self.id1_input.currentText()
        self.updateDataComboBox(selected_id, self.data1_input)

    def updateDataComboBox2(self):
        selected_id = self.id2_input.currentText()
        self.updateDataComboBox(selected_id, self.data2_input)

    def updateDataComboBox(self, selected_id, data_combo_box):
        try:
            df = pd.read_csv(self.file_path, encoding="utf-8", low_memory=False)
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"讀取 CSV 檔案失敗: {e}")
            return

        data_column = "数据"
        filtered_data = df[df["帧ID"] == selected_id][data_column].dropna().unique().astype(str)
        data_combo_box.clear()
        data_combo_box.addItems(filtered_data)

    def Capture_ID(self):
        # 從輸入欄位取得值
        self.CANID1 = self.id1_input.currentText()
        self.CANID2 = self.id2_input.currentText()
        self.CANData1 = self.data1_input.currentText()
        self.CANData2 = self.data2_input.currentText()
        #修改:讀取最小時間
        try:
            self.min_time_diff = float(self.min_time_diff_input.text())
        except ValueError:
             QMessageBox.critical(self, "錯誤", "請輸入有效的小數點時間")
             self.min_time_diff_input.setText("0.000") #輸入錯誤時復原為0
             return

    def processData(self):
        if not self.file_path:
            self.fileLabel.setText("請先選擇 CSV 檔案!")
            return

        self.Capture_ID()

        if not self.CANData1 or not self.CANData2:
            self.fileLabel.setText("請輸入要截取的報文!")
            return

        # 讀取 CSV
        try:
            df = pd.read_csv(self.file_path, encoding="utf-8", low_memory=False)
        except FileNotFoundError:
            QMessageBox.critical(self, "錯誤", "找不到 CSV 檔案!")
            return
        except pd.errors.EmptyDataError:
            QMessageBox.critical(self, "錯誤", "CSV 檔案為空!")
            return
        except pd.errors.ParserError:
            QMessageBox.critical(self, "錯誤", "CSV 檔案格式錯誤!")
            return

        # 時間和數據列
        time_column = "时间标识"  # 請確保名稱符合你的 CSV
        data_column = "数据"

        # 轉換時間格式
        df[time_column] = pd.to_datetime(df[time_column], format="%H:%M:%S:%f", errors='coerce')

        # 依時間排序
        df = df.sort_values(by=time_column).reset_index(drop=True)

        # 記錄結果
        time_differences = []
        start_time = None  # 記錄第一個報文的時間
        count_data1 = 1  # 新增：記錄報文1發送次數

        if self.CANData1 == self.CANData2:
            QMessageBox.critical(self, "錯誤", "輸入的報文不能相同!")
            return
        else:
            for _, row in df.iterrows():
                if row[data_column] == self.CANData1:
                    if start_time is None:
                        start_time = row[time_column]
                    else:
                        count_data1 += 1
                elif row[data_column] == self.CANData2:
                    if start_time is not None:
                        time_diff = (row[time_column] - start_time).total_seconds()
                         # 修改：只有時間差大於等於最小時間差才記錄
                        if time_diff >= self.min_time_diff:
                            time_differences.append((start_time, row[time_column], time_diff, count_data1))
                        start_time = None  # 重置起點
                        count_data1 = 1
        # 轉換為 DataFrame
        result_df = pd.DataFrame(time_differences, columns=[f"{self.CANData1}時間", f"{self.CANData2}時間", "時間差 (秒)", "報文1發送次數"])

        # 格式化時間
        result_df[f"{self.CANData1}時間"] = pd.to_datetime(result_df[f"{self.CANData1}時間"])
        result_df[f"{self.CANData1}時間"] = result_df[f"{self.CANData1}時間"].dt.strftime("%H:%M:%S.%f").str[:-3]
        result_df[f"{self.CANData2}時間"] = result_df[f"{self.CANData2}時間"].dt.strftime("%H:%M:%S.%f").str[:-3]
        result_df["時間差 (秒)"] = result_df["時間差 (秒)"].apply(lambda x: f"{x:.3f}")

        # 顯示結果到表格
        self.table.setRowCount(len(result_df))
        for i, row in result_df.iterrows():
            self.table.setItem(i, 0, QTableWidgetItem(row[f"{self.CANData1}時間"]))
            self.table.setItem(i, 1, QTableWidgetItem(row[f"{self.CANData2}時間"]))
            self.table.setItem(i, 2, QTableWidgetItem(row["時間差 (秒)"]))
            self.table.setItem(i, 3, QTableWidgetItem(str(row["報文1發送次數"])))

        # 儲存結果
        if self.save_path:
            result_df.to_csv(self.save_path, index=False, encoding="utf-8-sig")
            self.saveLabel.setText(f'檔案已儲存至: {self.save_path}')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = TimeDifferenceApp()
    ex.show()
    sys.exit(app.exec_())
