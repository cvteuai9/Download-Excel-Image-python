# 目標
1. 自動將Excel中每個儲存格內的圖片網址下載下來
2. 將各個圖片加上流水號
****
## downloadImgExcel.py 是用來下載Excel網址中的圖片
## insertExcel.py 將每一張圖片的新名字填入Excel中
****
# 步驟
## 一、 VScode環境設定
1. 開啟VScode
2. 確認已安裝python!!
3. 按下 "__Ctrl__" + "__~__" 鍵開啟終端
4. 輸入指令  
   ``` pip install openpyxl ```
## 二、將你的Excel欄位設置成以下模式  
![image](https://github.com/cvteuai9/Download-Excel-Image-python/assets/51105037/464367f3-d0d5-47a9-add5-c407475a4722)  
## 三、 設定 `downloadImgExcel.py` 參數  
   1. ```python
      excel_path = '{你的excel文件路徑}'
      ```
   2. ```python
      column_letter = '{儲存圖片URL的行}'
      ```  
   3. ```python
      output_folder = '{儲存圖片的目的資料夾}'
      ```  
![image](https://github.com/cvteuai9/Download-Excel-Image-python/assets/51105037/8cecab29-7aad-4581-b659-e14e2019f5c5)
## 四、 設定 `downloadImgExcel.py`
