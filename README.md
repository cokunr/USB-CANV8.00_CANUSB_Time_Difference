# CANUSB 時間差計算工具

## 概述
對USB-CANV8.00生成的.xls檔案，做兩種報文之間的時間差計算
計算方式為第一次出現報文1到第二次出現報文2之間的時間間隔


## 功能
- 將USB-CANV8.00生成的xls檔案轉換為CSV檔案。
- 從 CSV 檔案加載 CAN 數據。
- 從 CSV 檔案中的 "帧ID" 列中選擇報文 ID。
- 根據選擇的 ID 過濾並選擇相應的報文。
- 透過時間差篩選功能，篩掉連點造成的誤差值，方便統計
- 計算選擇報文之間的時間差。
- 以表格格式顯示結果。
- 將結果保存到新的 CSV 檔案。

## 要求
要以Python運行此應用程式，您需要安裝以下內容：
- Python 3.x
- Pandas 庫
- PyQt5 庫
- xlrd 庫


您可以使用 pip 安裝所需的庫：

```sh
pip install pandas PyQt5 xlrd
```

## 使用方法
1. 啟動應用程式。
2. 點擊 "選擇 CSV 檔案" 以加載USB-CANV8.00生成的xls檔案或是CSV 檔案。
3. 加載檔案後，從下拉選單中選擇所需的 ID。
4. 從第二個下拉選單中選擇相應的報文。
5. 可選:可設置最小時間差，篩掉低於該時間差的組別，默認為不使用(時間0)
6. 點擊 "計算時間差" 以計算時間差。
7. 結果將顯示在下方的表格中。
8. 可選: 若有指定存檔路徑可將結果保存到新的 CSV 檔案。

## 注意
確保您的 CSV 檔案包含必要的列，包括 "时间标识" 用於時間戳和 "数据" 用於報文數據。應用程式期望時間戳格式與代碼中指定的格式兼容。

## 授權
此專案採用 MIT 許可證。歡迎修改和分發。
