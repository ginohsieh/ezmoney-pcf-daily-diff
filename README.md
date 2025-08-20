# EZMoney PCF Daily Diff

## 輸出資訊

### 股票清單包含欄位

- 股票代號
- 股票名稱  
- 股數
- 持股權重
- 前日股數
- 前日持股權重

### 比較分析新增欄位

- 股數變化（絕對值）
- 股數變化（%）
- 持股權重變化（%）

### Excel顏色標示

- 🟡 **黃色**: 新增股票
- 🟢 **淺綠色**: 減持股票（股數變化為負數）
- 🌸 **粉紅色**: 增持超過30%的股票

### 分析報告包含

- 總股票數統計
- 新增/清倉股票數量
- 股數增加/減少/不變的股票數量
- 所有持股清單（按當日權重排序）
- 新增股票清單
- 股數變化最大的前5支股票（排除新增/清倉）
- 權重變化最大的前5支股票

## 安裝要求

```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法（使用預設參數）
```bash
python pcf_downloader.py
```
- 預設日期：今天 (2025-08-20)
- 預設基金代碼：49YTW

### 指定參數
```bash
python pcf_downloader.py --date 2025-08-15 --fund-code 49YTW
```

### 參數說明
- `--date`: 指定日期，格式為 YYYY-MM-DD（例如：2025-08-15）
- `--fund-code`: 基金代碼（例如：49YTW）

## 輸出檔案

下載的檔案會儲存在 `downloads/` 資料夾中，最終只會產生比較結果檔案：

- `PCF_Comparison_{基金代碼}_{日期}.xlsx` - 包含顏色標示的比較分析結果

例如：

- `PCF_Comparison_49YTW_20250820.xlsx` - 當日與前一工作日的比較結果

## 範例

```bash
# 下載今天和前一工作日的 49YTW PCF 檔案
python pcf_downloader.py

# 下載 2025-08-15 和前一工作日的 49YTW PCF 檔案
python pcf_downloader.py --date 2025-08-15

# 下載指定日期和基金代碼的PCF檔案
python pcf_downloader.py --date 2025-08-15 --fund-code 00878
```
