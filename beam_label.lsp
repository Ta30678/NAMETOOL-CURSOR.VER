; AutoLISP腳本 - 自動插入梁編號
; 使用方法：在AutoCAD中載入此檔案，然後執行 (beam-label-insert)

; 全域變數
(setq *beam-data-file* "beam_data.txt")
(setq *text-height* 2.5)
(setq *text-style* "BEAM_TEXT")

; 主函數：插入梁編號
(defun beam-label-insert (/ file data)
  (princ "\n=== ETABS梁編號插入工具 ===")
  
  ; 設定文字樣式
  (beam-setup-text-style)
  
  ; 創建圖層
  (beam-setup-layers)
  
  ; 讀取資料檔案
  (setq data (beam-read-data))
  
  ; 插入標籤
  (if data
    (progn
      (beam-insert-labels data)
      (princ (strcat "\n成功插入 " (itoa (length data)) " 個梁編號"))
    )
    (princ "\n錯誤：無法讀取資料檔案")
  )
  
  (princ)
)

; 設定文字樣式
(defun beam-setup-text-style ()
  (if (not (tblsearch "STYLE" *text-style*))
    (progn
      (command "STYLE" *text-style* "Arial" *text-height* "0.8" "0" "N" "N" "N")
      (princ (strcat "\n創建文字樣式: " *text-style*))
    )
  )
)

; 創建圖層
(defun beam-setup-layers ()
  ; 大梁圖層 (紅色)
  (if (not (tblsearch "LAYER" "BEAM_MAIN"))
    (command "LAYER" "M" "BEAM_MAIN" "C" "1" "BEAM_MAIN" "")
  )
  
  ; 小梁圖層 (綠色)
  (if (not (tblsearch "LAYER" "BEAM_SECONDARY"))
    (command "LAYER" "M" "BEAM_SECONDARY" "C" "3" "BEAM_SECONDARY" "")
  )
  
  ; 特殊梁圖層 (洋紅色)
  (if (not (tblsearch "LAYER" "BEAM_SPECIAL"))
    (command "LAYER" "M" "BEAM_SPECIAL" "C" "6" "BEAM_SPECIAL" "")
  )
  
  ; 標籤圖層 (白色)
  (if (not (tblsearch "LAYER" "BEAM_LABELS"))
    (command "LAYER" "M" "BEAM_LABELS" "C" "7" "BEAM_LABELS" "")
  )
)

; 讀取資料檔案
(defun beam-read-data (/ file line data)
  (setq file (open *beam-data-file* "r"))
  (if file
    (progn
      (setq data '())
      (while (setq line (read-line file))
        (setq data (cons (beam-parse-line line) data))
      )
      (close file)
      (reverse data)
    )
    nil
  )
)

; 解析資料行
(defun beam-parse-line (line / parts)
  (setq parts (beam-split-string line ","))
  (if (>= (length parts) 4)
    (list
      (nth 0 parts)  ; 標籤
      (atof (nth 1 parts))  ; X座標
      (atof (nth 2 parts))  ; Y座標
      (atof (nth 3 parts))  ; 旋轉角度
      (if (> (length parts) 4) (nth 4 parts) "BEAM_LABELS")  ; 圖層
    )
    nil
  )
)

; 字串分割函數
(defun beam-split-string (str delim / pos result)
  (setq result '())
  (while (setq pos (vl-string-search delim str))
    (setq result (cons (substr str 1 pos) result))
    (setq str (substr str (+ pos 2)))
  )
  (if (> (strlen str) 0)
    (setq result (cons str result))
  )
  (reverse result)
)

; 插入標籤
(defun beam-insert-labels (data / item label x y rotation layer)
  (foreach item data
    (if item
      (progn
        (setq label (nth 0 item))
        (setq x (nth 1 item))
        (setq y (nth 2 item))
        (setq rotation (nth 3 item))
        (setq layer (nth 4 item))
        
        ; 設定圖層
        (command "LAYER" "S" layer "")
        
        ; 插入文字
        (command "TEXT" 
          (list x y 0)  ; 位置
          *text-height*  ; 高度
          rotation  ; 旋轉角度
          label  ; 文字內容
        )
      )
    )
  )
)

; 批量插入函數（從Excel資料）
(defun beam-insert-from-excel (/ data)
  (princ "\n請選擇包含梁編號資料的Excel檔案...")
  (setq data (beam-read-excel-data))
  (if data
    (beam-insert-labels data)
    (princ "\n無法讀取Excel資料")
  )
)

; 讀取Excel資料（需要外部工具）
(defun beam-read-excel-data ()
  ; 這裡需要與Excel檔案互動
  ; 可以使用COM介面或其他方法
  ; 暫時返回範例資料
  '(
    ("B1-1" 0.0 0.0 0.0 "BEAM_MAIN")
    ("B1-2" 10.0 0.0 0.0 "BEAM_MAIN")
    ("G1-1" 20.0 0.0 0.0 "BEAM_MAIN")
    ("b1-1" 0.0 10.0 0.0 "BEAM_SECONDARY")
    ("b1-2" 10.0 10.0 0.0 "BEAM_SECONDARY")
  )
)

; 設定資料檔案路徑
(defun beam-set-data-file (filename)
  (setq *beam-data-file* filename)
  (princ (strcat "\n資料檔案設定為: " filename))
)

; 設定文字高度
(defun beam-set-text-height (height)
  (setq *text-height* height)
  (princ (strcat "\n文字高度設定為: " (rtos height)))
)

; 清除所有梁標籤
(defun beam-clear-labels ()
  (princ "\n清除所有梁標籤...")
  (command "ERASE" "ALL" "")
  (princ "\n清除完成")
)

; 顯示使用說明
(defun beam-help ()
  (princ "\n=== ETABS梁編號工具使用說明 ===")
  (princ "\n1. (beam-label-insert) - 插入梁編號")
  (princ "\n2. (beam-set-data-file \"檔案路徑\") - 設定資料檔案")
  (princ "\n3. (beam-set-text-height 高度) - 設定文字高度")
  (princ "\n4. (beam-clear-labels) - 清除所有標籤")
  (princ "\n5. (beam-help) - 顯示此說明")
  (princ "\n")
  (princ "\n資料檔案格式：標籤,X座標,Y座標,旋轉角度,圖層")
  (princ "\n範例：B1-1,0.0,0.0,0.0,BEAM_MAIN")
)

; 載入時顯示說明
(princ "\nETABS梁編號工具已載入！")
(princ "\n輸入 (beam-help) 查看使用說明")
(princ "\n輸入 (beam-label-insert) 開始插入梁編號")
