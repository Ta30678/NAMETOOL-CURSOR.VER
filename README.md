# ETABS 梁自動編號工具 - 完整解決方案

## 概述
這是一個完整的ETABS梁編號解決方案，包含三種AutoCAD整合方式，讓您能夠將編號結果直接套用到AutoCAD圖面上。

## 功能特色

### 1. 網頁版編號工具
- 讀取ETABS .e2k檔案
- 自動生成梁編號
- 支援鏡像對稱編號
- 視覺化平面圖顯示
- 多種匯出格式

### 2. AutoCAD整合方案

#### 方案A：AutoCAD .NET 外掛
- **檔案位置**: `AutoCAD_BeamLabel_Plugin/`
- **功能**: 直接在AutoCAD中插入梁編號
- **使用方法**:
  1. 編譯外掛專案
  2. 載入到AutoCAD
  3. 執行 `BEAMLABEL` 命令
  4. 選擇編號資料檔案

#### 方案B：DXF檔案匯出
- **檔案位置**: `DXF_Export/`
- **功能**: 匯出DXF格式檔案，可直接在AutoCAD中開啟
- **特色**:
  - 自動分層（大梁、小梁、特殊梁）
  - 正確的文字旋轉角度
  - 保持座標精度

#### 方案C：AutoCAD腳本
- **檔案位置**: `AutoCAD_Scripts/`
- **功能**: 生成AutoCAD腳本檔案
- **使用方法**:
  1. 匯出腳本檔案
  2. 在AutoCAD中執行 `SCRIPT` 命令
  3. 選擇生成的腳本檔案

#### 方案D：AutoLISP腳本
- **檔案位置**: `AutoLISP/`
- **功能**: 生成AutoLISP腳本
- **使用方法**:
  1. 在AutoCAD中載入 `.lsp` 檔案
  2. 執行 `(beam-label-insert)` 命令

## 安裝與使用

### 1. 網頁版工具
```bash
# 開啟 enhanced_html_tool.html
# 在瀏覽器中開啟檔案
```

### 2. AutoCAD .NET 外掛
```bash
# 使用Visual Studio開啟 BeamLabelPlugin.csproj
# 編譯專案
# 將生成的DLL複製到AutoCAD外掛目錄
```

### 3. 腳本檔案
```bash
# 直接使用生成的 .scr 或 .lsp 檔案
# 在AutoCAD中載入並執行
```

## 使用流程

### 步驟1：編號處理
1. 開啟 `enhanced_html_tool.html`
2. 上傳ETABS .e2k檔案
3. 設定編號參數
4. 執行編號處理

### 步驟2：匯出到AutoCAD
選擇以下任一方式：

#### 方式A：DXF檔案
1. 點擊「匯出 DXF 檔案」
2. 在AutoCAD中開啟生成的DXF檔案

#### 方式B：AutoCAD腳本
1. 點擊「匯出 AutoCAD 腳本」
2. 在AutoCAD中執行 `SCRIPT` 命令
3. 選擇生成的腳本檔案

#### 方式C：AutoLISP腳本
1. 點擊「匯出 AutoLISP 腳本」
2. 在AutoCAD中載入 `.lsp` 檔案
3. 執行 `(beam-label-insert)` 命令

#### 方式D：.NET外掛
1. 編譯並安裝外掛
2. 在AutoCAD中執行 `BEAMLABEL` 命令
3. 選擇編號資料檔案

## 檔案結構

```
1002vibe coding test/
├── enhanced_html_tool.html          # 增強版網頁工具
├── AutoCAD_BeamLabel_Plugin/         # .NET外掛
│   ├── BeamLabelPlugin.cs
│   ├── BeamLabelPlugin.csproj
│   └── Properties/
├── DXF_Export/                       # DXF匯出功能
│   └── dxf_writer.js
├── AutoCAD_Scripts/                  # AutoCAD腳本
│   └── beam_label_script.scr
├── AutoLISP/                         # AutoLISP腳本
│   └── beam_label.lsp
└── README.md                         # 說明文件
```

## 技術規格

### 支援的AutoCAD版本
- AutoCAD 2020及以上版本
- 支援.NET Framework 4.8

### 支援的檔案格式
- ETABS .e2k檔案
- Excel .xlsx檔案
- DXF檔案
- AutoCAD腳本檔案
- AutoLISP腳本檔案

### 圖層設定
- `BEAM_MAIN`: 大梁（紅色）
- `BEAM_SECONDARY`: 小梁（綠色）
- `BEAM_SPECIAL`: 特殊梁（洋紅色）
- `BEAM_LABELS`: 標籤（白色）

## 注意事項

1. **座標系統**: 確保AutoCAD圖面與ETABS使用相同的座標系統
2. **文字樣式**: 建議在AutoCAD中預先設定文字樣式
3. **圖層管理**: 使用前建議先執行圖層設定
4. **備份**: 建議在執行前備份原始圖面

## 故障排除

### 常見問題
1. **外掛無法載入**: 檢查AutoCAD版本和.NET Framework版本
2. **座標不正確**: 確認座標系統設定
3. **文字顯示異常**: 檢查文字樣式設定
4. **圖層問題**: 手動創建所需圖層

### 技術支援
如有問題，請檢查：
1. 瀏覽器控制台錯誤訊息
2. AutoCAD命令列錯誤訊息
3. 檔案路徑和權限設定

## 更新日誌

### v1.0.0
- 初始版本發布
- 支援四種AutoCAD整合方式
- 完整的編號功能
- 多種匯出格式

## 授權
本工具僅供學習和研究使用。
