# 資料品質整理與分類工具（Streamlit 版）

將原本 Google Colab 的資料品質檢查腳本改寫為 **Streamlit Web App**，可在本機執行或部署到 Streamlit Cloud / GitHub。

## 功能

- 上傳 Excel（`.xlsx`）
- 自動檢查資料品質（空值、日期、電話、Email、年齡、薪資、信用分數、就業狀態等）
- 將資料分為「完整」與「不完整」
- 產生問題摘要與統計摘要
- 一鍵下載多工作表 Excel：
  1. `完整資料_Complete`
  2. `不完整資料_Incomplete`
  3. `問題摘要_Issues`
  4. `統計摘要_Summary`

## 本機執行

```bash
# 1. 建立虛擬環境（建議）
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate

# 2. 安裝依賴
pip install -r requirements.txt

# 3. 啟動 App
streamlit run app.py