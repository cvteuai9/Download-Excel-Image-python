# 目標
1. 自動將Excel中每個儲存格內的圖片網址下載下來
2. 將各個圖片加上流水號
****
## downloadImgExcel.py 是用來下載Excel網址中的圖片
## insertExcel.py 將每一張圖片的新名字填入Excel中
****
# 步驟
## 一、 Clone我的檔案，git clone或是直接下載zip檔都行  
![image](https://github.com/cvteuai9/Download-Excel-Image-python/assets/51105037/0cdc0613-e41a-4315-b850-b9f4f776baf3)
## 二、 VScode環境設定
1. 開啟VScode
2. 移動到此檔案所在的資料夾(工作區)
3. 確認已安裝python!!
4. 按下 "__Ctrl__" + "__~__" 鍵開啟終端
5. 輸入指令  
   ``` pip install openpyxl ```
## 三、將你的Excel欄位設置成以下模式  
![image](https://github.com/cvteuai9/Download-Excel-Image-python/assets/51105037/464367f3-d0d5-47a9-add5-c407475a4722)  
## 四、 設定 `downloadImgExcel.py` 參數  
![image](https://github.com/cvteuai9/Download-Excel-Image-python/assets/51105037/aa58518c-3712-472c-836c-ae9fe3043e4f)  
   1. ```python
      excel_path = '{你的excel文件路徑}'
      ```
   2. ```python
      column_letter = '{儲存圖片URL的行}'
      ```  
   3. ```python
      output_folder = '{儲存圖片的目的資料夾}'
      ```  
   4. ```python
      filename_column_letter = ''{儲存圖片名字的行}'
      ```  
## 五、執行
   * 輸入指令執行python檔案
     ```python
     python downloadImgExcel.py
     ```
     ![image](https://github.com/cvteuai9/Download-Excel-Image-python/assets/51105037/a03f448f-c518-4a0e-be68-605a72d15a11)
****
# 附件
* images.xlsx -> 已完成的Excel範例
  ![image](https://github.com/cvteuai9/Download-Excel-Image-python/assets/51105037/ebac8fb6-e6b2-473c-8f02-b553eb85bbbf)

****
# 參考
* ChatGpt
