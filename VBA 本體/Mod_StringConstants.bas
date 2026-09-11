' ==========================================================
' MODULE: Mod_StringConstants (標準模組)
' PURPOSE: 中央全域常數庫。集中管理全系統所有的提示文字、工作表名稱、
'          儲存格格子位址、介面顏色與「莫蘭迪寬幅儀表板」尺寸合約。
' EXPORTS: 全域文字／數值／色彩／格式常數（共 21 個分區，詳見模組內部區塊註記）
' IMPORTS: 無
' FORBIDDEN: 本模組為純文字與數值常數庫，嚴禁寫入任何邏輯運算、條件判斷或狀態變更。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================
Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_StringConstants"
Private Const MODULE_VERSION As String = "1.0.0"

' ==========================================================
' 第一區：前台介面圖形與按鈕名稱
' ==========================================================
Public Const SH_CANVAS_NAME             As String = "Img_Canvas"
Public Const SH_ANCHOR_NAME             As String = "Txt_WatermarkAnchor"
Public Const SH_BLEED_ZONE_NAME         As String = "Line_BleedZone"
Public Const SH_TEMPLATE_CHART_NAME     As String = "ACDS_TemplateChart"

' 五核心控制按鈕物件名稱（橫向雙軌對稱佈局）
Public Const BTN_FOLDER_NAME            As String = "BTN_PHASE1_FOLDER"
Public Const BTN_FILES_NAME             As String = "BTN_PHASE1_FILES"
Public Const BTN_EXPORT_NAME            As String = "BTN_PHASE2_EXPORT"
Public Const BTN_REBUILD_NAME           As String = "BTN_FORCE_REBUILD"
Public Const BTN_CONSOLIDATE_NAME       As String = "BTN_CONSOLIDATE_LOGS"

' 舊版垃圾清理探測名稱
Public Const LGC_BTN_OLD_IMPORT         As String = "BTN_PHASE1_IMPORT"
Public Const LGC_BTN_OLD_EXPORT         As String = "BTN_PHASE2_EXPORT"

' ==========================================================
' 第二區：巨集路由名稱
' ==========================================================
Public Const MACRO_IMPORT_FOLDER        As String = "Macro_Phase1_ImportFolder"
Public Const MACRO_IMPORT_FILES         As String = "Macro_Phase1_ImportFiles"
Public Const MACRO_EXPORT               As String = "Macro_Phase2_Export"
Public Const MACRO_REBUILD_UI           As String = "Macro_ForceRebuildUI"
Public Const MACRO_CONSOLIDATE          As String = "Macro_ConsolidateAllLogs"

' ==========================================================
' 第三區：工作表分頁名稱
' ==========================================================
Public Const SHEET_CONFIG               As String = "ConfigSheet"
Public Const SHEET_STAGING              As String = "Staging_Images"
Public Const SHEET_LOG_EXPORT           As String = "LOG_Export_成果"
Public Const SHEET_MASTER_ARCHIVE       As String = "LOG_Master_Archive"

' ==========================================================
' 第四區：儲存格格子絕對地址合約 (終極校正：標籤 Row 2，選單 Row 3 facts)
' ==========================================================
Public Const ADDR_CFG_SESSION_STATUS    As String = "A1"    ' 頂層狀態橫幅特區
Public Const ADDR_CFG_LBL_CHINESE       As String = "A2"    ' 中文字體標籤 (Row 2)
Public Const ADDR_CFG_FONT_CHINESE      As String = "A3"    ' 中文字體下拉框 (Row 3)
Public Const ADDR_CFG_LBL_ENGLISH       As String = "B2"    ' 英數字體標籤 (Row 2)
Public Const ADDR_CFG_FONT_ENGLISH      As String = "B3"    ' 英數字體下拉框 (Row 3)
Public Const ADDR_CFG_LBL_COLOR         As String = "C2"    ' 印記顏色標籤 (Row 2)
Public Const ADDR_CFG_FONT_COLOR        As String = "C3"    ' 印記顏色下拉框 (Row 3)
Public Const ADDR_CFG_LBL_SIZE          As String = "D2"    ' 字體大小標籤 (Row 2)
Public Const ADDR_CFG_FONT_SIZE         As String = "D3"    ' 字體大小下拉框 (Row 3)
Public Const ADDR_CFG_LBL_TEMPLATE      As String = "E2"    ' 排版範本標籤 (Row 2)
Public Const ADDR_CFG_VAL_TEMPLATE      As String = "E3"    ' 排版範本下拉框 (Row 3)

Public Const ADDR_CFG_LBL_RANGE         As String = "A2:E2" ' 標籤層刷色範圍
Public Const ADDR_CFG_VAL_RANGE         As String = "A3:E3" ' 下拉層刷色範圍
Public Const ADDR_CFG_CTRL_CARD_RANGE   As String = "A2:E3" ' 卡片全域擦洗範圍 (絕對不包含 A1)

' 成果日誌工作表專屬網格地址
Public Const ADDR_LOG_HDR_STATUS        As String = "A1"
Public Const ADDR_LOG_HDR_NAME          As String = "B1"
Public Const ADDR_LOG_HDR_SRC           As String = "C1"
Public Const ADDR_LOG_HDR_PATH          As String = "D1"
Public Const ADDR_LOG_AUDIT_CELL        As String = "E1"
Public Const ADDR_LOG_HEADER_RANGE      As String = "A1:D1"
Public Const ADDR_LOG_COL_ABC           As String = "A:C"
Public Const ADDR_LOG_COL_D             As String = "D"
Public Const ADDR_LOG_COL_E             As String = "E"
Public Const ADDR_LOG_COL_ALL           As String = "A:E"

' ==========================================================
' 第五區：前台 UI 寬幅對稱網格骨架座標
' ==========================================================
Public Const GEO_GRID_CANVAS_ROW        As Long = 4
Public Const GEO_GRID_CANVAS_COL        As Long = 1
Public Const GEO_GRID_BTN_PANEL_COL     As Long = 8         ' H 欄起始

' 【衝突去重】：完全刪除舊版直立式的死結，全面對齊雙軌新制
Public Const GEO_GRID_BTN_ROW_P1        As Long = 6         ' 第一軌：圖片資產匯入區列高起點
Public Const GEO_GRID_BTN_ROW_P2        As Long = 10        ' 第二軌：核心大功能導出區列高起點
Public Const GEO_GRID_BTN_ROW_MAINT     As Long = 14        ' 第三軌：後台系統維護區列高起點

' ==========================================================
' 第六區：介面標籤與按鈕顯示文字（全面優化對齊）
' ==========================================================
Public Const LBL_CONFIG_FONT_CHINESE    As String = "中文字體"
Public Const LBL_CONFIG_FONT_ENGLISH    As String = "英數字體"
Public Const LBL_CONFIG_FONT_COLOR      As String = "印記顏色"
Public Const LBL_CONFIG_FONT_SIZE       As String = "字體大小"
Public Const LBL_CONFIG_TEMPLATE        As String = "排版範本"
Public Const LBL_CANVAS_BANNER          As String = "相片定位虛擬畫布區 (可自由拖拽下方位置控制項)"
Public Const LBL_ANCHOR_TEXT            As String = "【印記】"
Public Const LBL_BTN_FOLDER             As String = "匯入整個資料夾"
Public Const LBL_BTN_FILES              As String = "手選特定圖片"
Public Const LBL_BTN_EXPORT             As String = "啟動批次浮水印導出"
Public Const LBL_BTN_REBUILD            As String = "一鍵重整工作檯"
Public Const LBL_BTN_CONSOLIDATE        As String = "歷史日誌總歸檔"

' ==========================================================
' 第七區：介面字體大小、預設值與「呼吸列高」尺寸
' ==========================================================
Public Const VAL_DEFAULT_FONT_CHINESE   As String = "微軟正黑體"
Public Const VAL_DEFAULT_FONT_ENGLISH   As String = "微軟正黑體"

' 【收容硬編碼】：填補缺失的英文備用字體與預設排版範本 facts
Public Const VAL_FALLBACK_ENG_FONT      As String = "Arial"
Public Const VAL_DEFAULT_TEMPLATE       As String = "日期: {date}\n名稱: {name}\n備註: {remark}"

Public Const GEO_FONT_SIZE_DEFAULT      As Single = 12!
Public Const GEO_FONT_SIZE_LABEL        As Single = 9.5!
Public Const GEO_FONT_SIZE_CANVAS       As Single = 10!
Public Const GEO_FONT_SIZE_BTN          As Single = 11!
Public Const GEO_FONT_SIZE_BTN_SMALL    As Single = 10.5!
Public Const GEO_MAX_SHEET_NAME_LEN     As Long = 31

' 【高級感呼吸列高】：強迫網格高度拉開，產生高階 UI 留白感
Public Const GEO_ROW_HEIGHT_STATUS      As Single = 65!     ' 【新增】專門給 A1 三行狀態橫幅的高容積列高 facts
Public Const GEO_ROW_HEIGHT_HEADER      As Single = 30!     ' 統一所有表頭高度為 30 像素
Public Const GEO_ROW_HEIGHT_DATA        As Single = 24!     ' 統一所有資料列高度為 24 像素

' 大視覺實體幾何尺寸（單位：像素點）
Public Const GEO_CANVAS_WIDTH           As Single = 300!    ' 瘦身寬度 (符合 16:9 比例)
Public Const GEO_CANVAS_HEIGHT          As Single = 180!    ' 瘦身高度
Public Const GEO_BLEED_OFFSET_X         As Single = 16!
Public Const GEO_BLEED_OFFSET_Y         As Single = 12!
Public Const GEO_BLEED_WIDTH            As Single = 268!
Public Const GEO_BLEED_HEIGHT           As Single = 156!
Public Const GEO_ANCHOR_WIDTH           As Single = 80!     ' 精品化縮小
Public Const GEO_ANCHOR_HEIGHT          As Single = 25!     ' 精品化縮低
Public Const GEO_BTN_WIDTH              As Single = 110!
Public Const GEO_BTN_HEIGHT             As Single = 38!
Public Const GEO_BTN_GAP_X              As Single = 10!     ' 按鈕橫向間距

' 成果日誌欄寬與版控裁剪幾何
Public Const GEO_LOG_COL_WIDTH_SHORT    As Single = 18!
Public Const GEO_LOG_COL_WIDTH_LONG     As Single = 50!
Public Const GEO_LOG_COL_WIDTH_AUDIT    As Single = 75!
Public Const GEO_LOG_TRUNCATE_LEN       As Long = 24

' ==========================================================
' 第八區：環境強固化限制常數
' ==========================================================
Public Const GEO_WINDOW_ZOOM_STANDARD   As Long = 100
Public Const NUM_MAX_HYPERLINK_LEN      As Long = 255
Public Const NUM_RETRY_LIMIT            As Long = 3
Public Const NUM_RETRY_DELAY            As Single = 0.1!

' ==========================================================
' 第九區：前台高階莫蘭迪與馬卡龍色彩常數（BGR 十六進位）
' ==========================================================
Public Const COLOR_UI_LABEL_GRAY        As Long = &H707070
Public Const COLOR_UI_CANVAS_BG         As Long = &HFFFFFF
Public Const COLOR_UI_CANVAS_LINE       As Long = &HD0D0D0
Public Const COLOR_UI_CANVAS_TEXT       As Long = &HB4B4B4

' 1. 互動錨點：優雅莫蘭迪乾燥陶土/灰泥色
Public Const COLOR_UI_ANCHOR_BG         As Long = &H9C938C
Public Const COLOR_UI_ANCHOR_LINE       As Long = &H7D7570
Public Const COLOR_UI_ANCHOR_TEXT       As Long = &HFFFFFF

' 2. 下拉選單卡片區底色：法式優雅莫蘭迪暖沙色
Public Const COLOR_UI_CTRL_CARD_BG      As Long = &HDAD5CD

' 3. 寬幅儀表板按鈕色彩語意分流
Public Const COLOR_UI_BTN_P1_BG         As Long = &HE1D9D0      ' 輕柔馬卡龍暖亞麻
Public Const COLOR_UI_BTN_P1_TEXT       As Long = &H4A4A4A
Public Const COLOR_UI_BTN_P2_BG         As Long = &H7A6B5D      ' 核心莫蘭迪深褐青
Public Const COLOR_UI_BTN_P2_TEXT       As Long = &HFFFFFF
Public Const COLOR_UI_BTN_MAINT_BG      As Long = &HEAEAEA      ' 霧面燕麥灰
Public Const COLOR_UI_BTN_MAINT_TEXT    As Long = &H909090

' 4. 日誌網格高階色
Public Const COLOR_LOG_HEADER_BG        As Long = &HEAEAEA      ' 標題列：霧面燕麥灰
Public Const COLOR_UI_ZEBRA_EVEN        As Long = &HFDFBF9

' 5. 狀態落盤語意色（去霓虹化、高雅低飽和度）
Public Const STATUS_IMPORT_OK_BG        As Long = &HDCEAD2      ' 成功：莫蘭迪鼠尾草綠
Public Const STATUS_IMPORT_OK_TXT       As Long = &H3D5A2B      ' 深森林綠字
Public Const STATUS_EXPORT_ERR_BG       As Long = &HCCB3C4      ' 失敗：莫蘭迪乾燥玫瑰紅
Public Const STATUS_EXPORT_ERR_TXT      As Long = &H5C3A46      ' 深桑椹紅字

' ==========================================================
' 第十區：實體檔案系統隔離資料夾與字串基礎常數
' ==========================================================
Public Const DIR_NAME_WORKSPACE_ROOT    As String = "作業資料夾"
Public Const DIR_NAME_IMPORT_FOLDER     As String = "匯入圖片"
Public Const DIR_NAME_RESULT_FOLDER     As String = "成果"
Public Const DIR_NAME_BACKUP_ROOT       As String = "備份"
Public Const DIR_NAME_BACKUP_CHECKPOINT_PREFIX As String = "斷點備份_"
Public Const DIR_NAME_BACKUP_FINAL_PREFIX      As String = "成果備份_"

' 字串基礎常數
Public Const DOT_STR                    As String = "."
Public Const BACKSLASH_STR              As String = "\"
Public Const UNDERLINE_STR              As String = "_"
Public Const SPACE_STR                  As String = " "
Public Const CHAR_ZERO_PAD              As String = "0"
Public Const FALLBACK_BASE_NAME         As String = "IMG_SAFE_BACKUP"
Public Const HASH_FAILED_MARK           As String = "HASH_CALC_FAILED"

' ==========================================================
' 第十一區：日誌器專用字串與改名跳號標籤
' ==========================================================
Public Const SHEET_LOG_PREFIX           As String = "LOG_Export_"
Public Const LBL_RERUN_MARK             As String = "_重跑_"
Public Const LBL_RERUN_DEFAULT          As String = "_重跑_01"
Public Const LBL_RERUN_01               As String = "01"
Public Const LBL_RERUN_02               As String = "02"
Public Const LBL_RERUN_03               As String = "03"
Public Const LBL_RERUN_04               As String = "04"
Public Const LBL_RERUN_05               As String = "05"
Public Const LBL_RERUN_06               As String = "06"
Public Const LBL_RERUN_07               As String = "07"
Public Const LBL_RERUN_08               As String = "08"
Public Const LBL_RERUN_09               As String = "09"
Public Const LBL_RERUN_10               As String = "10"
Public Const LBL_RERUN_XX               As String = "XX"
Public Const LBL_OLD_VERSION_UNDER      As String = "_舊版_"

' ==========================================================
' 第十二區：WIA 影像處理與圖表計算限制 (1.5 字元留白與對齊 facts)
' ==========================================================
Public Const NUM_MAX_CANVAS_DIM         As Long = 2200
Public Const NUM_DEFAULT_WIDTH          As Long = 1123
Public Const NUM_DEFAULT_HEIGHT         As Long = 794
Public Const NUM_SCALE_BASE             As Single = 800!
Public Const NUM_BOX_WIDTH_BASE         As Single = 240!
Public Const NUM_LINE_SPACING_BASE      As Single = 1.2!

' 浮水印自訂排版幾何合約
Public Const NUM_WM_PADDING_CHARS       As Single = 1.5!        ' 上下左右剛性留白：1.5 個字元寬高
Public Const NUM_WM_ALIGN_HORZ          As Long = 1             ' 水平靠左對齊 (msoAlignLeft)
Public Const NUM_WM_ALIGN_VERT          As Long = 3             ' 垂直置中對齊 (msoAnchorMiddle)

' 動態範本權杖標籤識別碼
Public Const TOK_NAME                   As String = "{name}"
Public Const TOK_DATE                   As String = "{date}"
Public Const TOK_REMARK                 As String = "{remark}"

Public Const NUM_CHART_HIDDEN_POS       As Long = -5000
Public Const NUM_MAX_REMARK_LEN         As Long = 30

Public Const FILTER_SCALE               As String = "Scale"
Public Const PROP_MAX_WIDTH             As String = "MaximumWidth"
Public Const PROP_MAX_HEIGHT            As String = "MaximumHeight"

Public Const EXT_TMP                    As String = ".tmp"
Public Const EXT_DOT_JPG                As String = ".jpg"
Public Const ENV_TEMP_VAR               As String = "TEMP"
Public Const TMP_FILE_PREFIX            As String = "\_wm_tmp_"

Public Const LBL_WM_PREFIX_NAME         As String = "名稱: "
Public Const LBL_WM_PREFIX_DATE         As String = "日期: "
Public Const LBL_WM_PREFIX_REMARK       As String = "備註: "
Public Const SUFFIX_TRUNCATE            As String = "..."

' ==========================================================
' 第十三區：資料表欄位名稱與大腦定位索引（工作合約）
' ==========================================================
Public Const COL_NAME_STATUS            As String = "狀態備註"
Public Const COL_NAME_IMG_NAME          As String = "圖片名稱"
Public Const COL_NAME_REMARK            As String = "使用者備註"
Public Const COL_NAME_LAST_MOD          As String = "最新更新日期"
Public Const COL_NAME_CUR_PATH          As String = "圖片現在位置"
Public Const COL_NAME_IMPORT_DATE       As String = "圖片匯入日期"
Public Const COL_NAME_IMPORT_REL        As String = "計畫匯出位置"
Public Const COL_NAME_HASH              As String = "檔案雜湊"
Public Const COL_NAME_EXPORTED          As String = "匯出狀態"

Public Const COL_NAME_LOG_STATUS        As String = "狀態備註"
Public Const COL_NAME_LOG_FILENAME      As String = "圖片名稱"
Public Const COL_NAME_LOG_SRC_SHEET     As String = "匯出工作表名稱"
Public Const COL_NAME_LOG_FINAL_PATH    As String = "最終成品路徑"

Public Const NUM_STG_COL_STATUS         As Long = 1
Public Const NUM_STG_COL_NAME           As Long = 2
Public Const NUM_STG_COL_REMARK         As Long = 3
Public Const NUM_STG_COL_MOD            As Long = 4
Public Const NUM_STG_COL_PATH           As Long = 5
Public Const NUM_STG_COL_DATE           As Long = 6
Public Const NUM_STG_COL_REL            As Long = 7
Public Const NUM_STG_COL_HASH           As Long = 8
Public Const NUM_STG_COL_EXPORTED       As Long = 9

Public Const NUM_LOG_COL_STATUS         As Long = 1
Public Const NUM_LOG_COL_NAME           As Long = 2
Public Const NUM_LOG_COL_SRC            As Long = 3
Public Const NUM_LOG_COL_PATH           As Long = 4

' ==========================================================
' 第十四區：狀態標籤文字
' ==========================================================
Public Const STATUS_IMPORT_OK           As String = "[成功] 匯入成功"
Public Const STATUS_IMPORT_ERR          As String = "[警告] 檔案損毀/無法讀取"
Public Const STATUS_EXPORT_OK           As String = "[成功] 匯出成功"
Public Const STATUS_EXPORT_ERR          As String = "[警告] 壓印失敗"
Public Const STATUS_LOG_VERSIONING      As String = "[版控更替] 舊版資產已安全降級封存"
Public Const STATUS_EXPORT_ERR_LOCKED   As String = "[警告] 實體鎖定失敗"
Public Const STATUS_STG_EXPORTED        As String = "[完成] 已匯出封存"


' ==========================================================
' 第十五區：前台介面驗證下拉選單字串與預設值
' ==========================================================
Public Const VAL_FONT_CHINESE           As String = "微軟正黑體,標楷體,新細明體"
Public Const VAL_FONT_ENGLISH           As String = "Arial,Calibri,Times New Roman"
Public Const VAL_COLORS                 As String = "黑色,白色,紅色,黃色,藍色,綠色"
Public Const VAL_FONT_SIZES             As String = "8,10,12,14,16,18,20"

' 排版範本下拉事實，剛性將標準三行修正為 (日期 -> 名稱 -> 備註) 順序
Public Const VAL_TEMPLATE_LIST          As String = "日期: {date}\n名稱: {name}\n備註: {remark},名稱: {name}\n備註: {remark},{remark}"

Public Const VAL_DEFAULT_COLOR          As String = "黑色"

' ==========================================================
' 第十六區：系統提示與異常阻斷訊息庫
' ==========================================================
Public Const TITLE_SYS_INFO             As String = "系統提示"
Public Const TITLE_P1_DONE              As String = "第一階段完成"
Public Const TITLE_P2_DONE              As String = "第二階段完成"
Public Const TITLE_SESSION_CONFIRM      As String = "工作階段 (Session) 意圖確認"
Public Const TITLE_EXPORT_SPLIT_CONFIRM As String = "匯出模式雙軌分流確認"
Public Const TITLE_EXPORT_DONE          As String = "批次作業完工報告"
Public Const TITLE_DRY_RUN_CONFIRM      As String = "測試模式雙向確認"
Public Const TITLE_REBUILD_SUCCESS      As String = "工作檯還原成功"
Public Const TITLE_CONSOLIDATE_SUCCESS  As String = "歷史總帳打包成功"
Public Const TITLE_ENV_MELTDOWN         As String = "環境熔斷器"
Public Const TITLE_LOG_MELTDOWN         As String = "歸檔熔斷"
Public Const TITLE_WIA_MELTDOWN         As String = "組件缺失熔斷"
Public Const TITLE_SKIP_ALERT           As String = "重複圖片提示"

' 【收容硬編碼】：將原本散落在地基建構師中的狀態判別字串升格為全域常數
Public Const KEYWORD_STATUS_CHECK       As String = "工作階段"

Public Const MSG_SYS_START_IMPORT       As String = "系統提示：開始執行第一階段圖片匯入與清洗作業。"
Public Const MSG_SYS_IMPORT_DONE        As String = "作業完成：圖片已成功匯入暫存區！請在「使用者備註(C欄)」確認或補打文字，完成後再執行第二階段匯出。"
Public Const MSG_SYS_START_EXPORT       As String = "系統提示：開始執行第二階段批次浮水印匯出作業。"
Public Const MSG_SYS_EXPORT_DONE        As String = "作業完成：批次浮水印壓印與匯出已全部結束，請檢視成果頁日誌。"
Public Const MSG_DRY_RUN_WAIT           As String = "【測試模式】環境檢查與路徑模擬已完成，程式即將停下。請確認沒問題後關閉此視窗。"
Public Const MSG_REBUILD_SUCCESS        As String = "前端直角畫布、虛線出血安全框與五核心按鈕面板已強制還原重塑完畢！"
Public Const MSG_CONSOLIDATE_SUCCESS    As String = "全案歷史分期日誌快照已流式合併壓縮併入總帳，舊表已物理清除，系統減重完畢！"

Public Const MSG_SESSION_PROMPT         As String = "偵測到目前作業資料夾已有進行中的照片資產 Facts。" & vbCrLf & vbCrLf & "【意圖：是 (Yes)】：追加累積！保持現況，繼續追加照片資產。" & vbCrLf & "【意圖：否 (No)】：開啟新案！目前成果將整批備份後，清空作業資料夾作為新起點。" & vbCrLf & "【意圖：取消 (Cancel)】：取消本次操作，保留現場安全退出。"
Public Const MSG_EXPORT_MODE_PROMPT     As String = "【偵測到已存在的匯出成品】" & vbCrLf & "系統發現成果資料夾內已有部分實體照片 Facts。" & vbCrLf & vbCrLf & "【是 (Yes)】：啟動「斷點續跑模式」，自動跳過已存在的成品相片。" & vbCrLf & "【否 (No)】：啟動「結束並重印模式」，目前成果將整批備份，清空成果資料夾後，重新壓印全部照片。"
Public Const MSG_DRY_RUN_HEADER         As String = "【系統安全 Dry Run 煞車檢核點】" & vbCrLf & "1. 待處理照片總數："
Public Const MSG_DRY_RUN_BODY           As String = " 張" & vbCrLf & "2. 浮水印幾何定位：X="
Public Const MSG_DRY_RUN_FOOTER         As String = "%, Y="
Public Const MSG_DRY_RUN_END            As String = "%" & vbCrLf & vbCrLf & "是否確認環境 facts 並解除煞車啟動實體壓印？"

Public Const LBL_DONE_MSG_RUN           As String = " (本輪處理/覆寫: "
Public Const LBL_DONE_MSG_UNIT          As String = " 張"
Public Const LBL_DONE_MSG_SKIP          As String = "，續跑跳過: "
Public Const MSG_STATUS_BAR_RUN_PREFIX  As String = "ACDS 調度器：已成功完成 "
Public Const MSG_STATUS_BAR_RUN_SUFFIX  As String = " 張壓印，實施原子性定時存檔..."
Public Const MSG_PICKER_FOLDER_TITLE    As String = "請選取含有原始照片的來源資料夾"
Public Const MSG_PICKER_FILE_TITLE      As String = "請按住 Ctrl 鍵複選您要匯入的圖片資產"
Public Const MSG_NO_DATA_ABORT          As String = "暫存區內無有效圖片資料，終止匯出。"
Public Const MSG_USER_ABORT_EXPORT      As String = "操作員主動中止，未執行 any 影像壓印。"
Public Const LBL_LOG_ARCHIVE_DONE_PREFIX As String = "已完成_"

Public Const MSG_SKIP_PREFIX            As String = "本次匯入偵測到 "
Public Const MSG_SKIP_SUFFIX            As String = " 張內容重複的圖片（雜湊值相同），已自動跳過，不重複匯入。"

' ==========================================================
' 第十七區：外部組件 ProgID 常數
' ==========================================================
Public Const PROGID_FSO                 As String = "Scripting.FileSystemObject"
Public Const PROGID_STREAM              As String = "ADODB.Stream"
Public Const PROGID_MD5                 As String = "System.Security.Cryptography.MD5CryptoServiceProvider"
Public Const PROGID_WIA_IMAGE           As String = "WIA.ImageFile"
Public Const PROGID_WIA_PROCESS         As String = "WIA.ImageProcess"
Public Const PROGID_DICTIONARY          As String = "Scripting.Dictionary"

' ==========================================================
' 第十八區：系統錯誤文字與自癒警告
' ==========================================================
Public Const MSG_HEAL_CONFIG            As String = "安全自癒：偵測到配置設定空白 or 錯誤，系統已自動啟用安全預設值繼續執行。"
Public Const MSG_HEAL_TRUNCATE          As String = "防禦自癒：部分圖片名稱 or 備註字數超過限制，系統已自動截斷處理。"
Public Const MSG_SESSION_CLOSED_PREFIX  As String = "[本工作階段已於 "
Public Const MSG_SESSION_CLOSED_SUFFIX  As String = " 成功匯出結案並歸檔；將於下次工作開始時清零重啟]"
Public Const LBL_LOG_ERR_PREFIX         As String = "錯誤因由: "

Public Const ERR_WIA_NOT_FOUND          As String = "環境斷層：本機缺乏微軟原生的 WIA 影像組件 (WIA.ImageFile)，系統拒絕執行影像壓印！"
Public Const ERR_LOG_STRUCTURE_LOCKED   As String = "環境阻斷：當前活頁簿結構已被鎖定保護，系統無法動態建立 LOG 回報頁！"
Public Const ERR_LOG_CREATE_FAILED      As String = "資源潰堤：實體工作表建立失敗，系統終止回報路由。"
Public Const ERR_UI_BUILD_FAILED_PREFIX As String = "視覺工作檯鑄造失敗，原因："
Public Const ERR_CONSOLIDATE_FAILED_PREFIX As String = "歷史日誌總帳快照壓縮歸檔失敗，原因："
Public Const ERR_CRITICAL_ANCHOR_LOST   As String = "嚴重錯誤：找不到畫布上的控制項或文字方塊，流程強行中斷！"
Public Const ERR_CONTRACT_VIOLATION     As String = "合約檢查失敗：當前工作表不符合 Staging_Images 的標準欄位規格！"
Public Const ERR_DIR_CREATE_FAILED      As String = "環境錯誤：無法建立實體防撞資料夾，請檢查磁碟寫入權限。"
Public Const ERR_ARCHIVE_COPY_FAILED    As String = "封存錯誤：成果資料夾複製備份失敗，可能原因為磁碟空間不足或權限不足，本次匯出未完成封存。"
Public Const ERR_STREAMP_FAILED         As String = "串流中斷：處理單張圖片時發生核心影像損毀，已由異常攔截器阻斷。"
Public Const ERR_HASH_FILE_NOT_FOUND    As String = "雜湊錯誤：找不到指定的實體檔案，無法進行Facts雜湊校對"
Public Const MSG_AUDIT_LINE             As String = "[ACDS 審計安全線：本輪 Session 原始匯入資料夾已實施實體物理鎖定]"

Public Const FILE_FILTER_IMAGES         As String = "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif"
Public Const LBL_PICKER_IMAGE_FILTER_PREFIX As String = "合法圖片資產 ("
Public Const LBL_PICKER_IMAGE_FILTER_SUFFIX As String = ")"

' ==========================================================
' 第十九區：資料格式常數（修正為絕對 6 位數時分秒）
' ==========================================================
Public Const PWD_PROTECT                As String = ""
Public Const FMT_TEXT_CELL              As String = "@"
Public Const FMT_DATE_YMD               As String = "yyyymmdd"
Public Const FMT_COUNTER_00             As String = "00"
Public Const FMT_COUNTER_01             As String = "01"
Public Const FMT_TIME_HMSM              As String = "hhnnss"
Public Const FMT_DATETIME_STD           As String = "yyyy-mm-dd hh:mm:ss"

' ==========================================================
' 第二十區：暫存看板與總帳工作表欄寬排版設定
' ==========================================================
Public Const ADDR_STG_HDR_RANGE         As String = "A1:I1"
Public Const ADDR_STG_FREEZE_CELL       As String = "A2"
Public Const COL_STG_A                  As String = "A"
Public Const COL_STG_B                  As String = "B"
Public Const COL_STG_C                  As String = "C"
Public Const COL_STG_D                  As String = "D"
Public Const COL_STG_E                  As String = "E"
Public Const COL_STG_F_TO_H             As String = "F:H"
Public Const COL_STG_I                  As String = "I"

Public Const WID_STG_A                  As Single = 12!
Public Const WID_STG_B                  As Single = 35!
Public Const WID_STG_C                  As Single = 45!
Public Const WID_STG_D                  As Single = 20!
Public Const WID_STG_E                  As Single = 40!
Public Const WID_STG_F_TO_H             As Single = 15!
Public Const WID_STG_I                  As Single = 15!

Public Const ADDR_LOG_RANGE_A1D1        As String = "A1:D1"
Public Const ADDR_LOG_RANGE_A2D_PREFIX  As String = "A2:D"
Public Const COL_LOG_A                  As String = "A"
Public Const COL_LOG_D                  As String = "D"
Public Const COL_LOG_ABC                As String = "A:C"

Public Const ERR_DATA_DISCONNECT        As String = "資料斷層：無法從暫存看板中動態探測到任何有效的原始資料夾路徑 Facts！"

' ==========================================================
' 第二十一區：DAG 流程追蹤器常數標籤
' ==========================================================
Public Const LOG_DAG_P1_FILES           As String = "第一階段：挑選檔案匯入"
Public Const LOG_DAG_P1_FOLDER          As String = "第一階段：整個資料夾匯入"
Public Const LOG_DAG_ENV_CHECK          As String = "環境與組件預檢"
Public Const LOG_DAG_LOAD_CFG           As String = "加載介面使用者設定"
Public Const LOG_DAG_DIR_GEN            As String = "實體防撞資料夾鑄造"
Public Const LOG_DAG_PICKER             As String = "呼叫 Windows 路徑選擇器"
Public Const LOG_DAG_PIPELINE           As String = "啟動第一階段匯入管線大迴圈"

Public Const LOG_DAG_P2_EXPORT          As String = "第二階段：批次浮水印匯出"
Public Const LOG_DAG_PRE_CHECK          As String = "匯出前置環境校對"
Public Const LOG_DAG_CONTRACT           As String = "檢查 Staging 看板數據合約"
Public Const LOG_DAG_PERCENT            As String = "換算畫布拖挪相對百分比"
Public Const LOG_DAG_RESUME_DETECT      As String = "探測硬碟歷史進度與斷點"
Public Const LOG_DAG_LOG_FRAME          As String = "鑄造導出成果日誌表頭"
Public Const LOG_DAG_DRY_RUN            As String = "彈出安全 Dry Run 煞車檢核單"
Public Const LOG_DAG_LOOP_START         As String = "啟動 second 階段匯出管線大迴圈"

' ==========================================================
' 【終極除垢】：全系統剛性列號、保底幾何與歷史色碼事實化
' ==========================================================
Public Const NUM_ROW_HEADER             As Long = 1         ' 誓死捍衛：第 1 列永遠是剛性表頭
Public Const NUM_ROW_CONFIG_VAL         As Long = 2         ' 設定工作表：第 2 列永遠是下拉數值
Public Const NUM_ROW_DATA_START         As Long = 2         ' 數據工作表：第 2 列永遠是資料起始點

Public Const GEO_POS_MIN_SAFE           As Long = 1         ' 算力引擎：圖表與圖形防崩潰安全起點 (1 像素)

' 歷史遺留過渡色碼（對齊莫蘭迪，若您想保留之前的傳統紅也可以在這裡改）
Public Const COLOR_UI_FAIL_RED_BG       As Long = &HC8C8FF    ' 傳統淡紅底
Public Const COLOR_UI_FAIL_RED_TXT      As Long = &H96        ' 傳統深紅字

' ==========================================================
' 【補票】第一階段匯入管線專用純文字與格式解耦常數
' ==========================================================
Public Const MSG_STATUS_BAR_IMPORT_PREFIX As String = "ACDS 守門員：正在清洗校對實體資產 - "
Public Const LBL_IMPORT_FAIL_REMARK     As String = "實體轉移失敗，請檢查權限"
Public Const FMT_DATE_DASH              As String = "yyyy-mm-dd"
Public Const FMT_DATETIME_DASH          As String = "yyyy-mm-dd hh:mm:ss"
Public Const ERR_TARGET_DIR_MISSING     As String = "目標防撞暫存資料夾未實體落地，拒絕執行複製管線"

' ==========================================================
' 【補票】空白工作表清洗
' ==========================================================
Public Const LBL_PROTECT_MARK       As String = "###"   ' 工作表名稱保護標記前綴，加了此標記者永不被自動清除
Public Const MSG_PURGE_CONFIRM_PREFIX As String = "偵測到以下空白頁面，即將一併清除：" & vbCrLf & vbCrLf
Public Const MSG_PURGE_CONFIRM_SUFFIX As String = vbCrLf & vbCrLf & "若想保留特定頁面，可在頁籤名稱前加上 ### 標記後取消本次操作。確定要清除嗎？"
Public Const TITLE_PURGE_CONFIRM      As String = "空白頁面清理確認"
Public Const MSG_PURGE_NONE_FOUND     As String = "沒有偵測到多餘的空白頁面，檯面已經很乾淨。"
Public Const ERR_STRUCTURE_LOCKED_PURGE As String = "結構阻斷：活頁簿結構目前處於保護狀態，無法刪除工作表，請先解除保護再重試。"

' ==========================================================
' 新增於 Mod_StringConstants.bas（建議併入既有 UI 幾何常數區與訊息文字區）
' ==========================================================
Public Const SH_LICENSE_FOOTER_NAME As String = "Footer_LicenseNotice"
Public Const LBL_LICENSE_FOOTER     As String = "Code: MIT License | Documentation: CC BY-NC-SA 4.0｜Copyright (C) 2026 鄭詩樺" & vbCrLf & "完整條款見 LICENSE"
Public Const GEO_FONT_SIZE_FOOTER   As Single = 7                  ' 小字級，不搶主要按鈕視覺焦點
Public Const GEO_FOOTER_GAP_Y       As Single = -8                 ' 與上方最後一列按鈕的垂直間距
Public Const GEO_FOOTER_HEIGHT      As Single = 32
Public Const COLOR_UI_FOOTER_TEXT   As Long = &HA0A0A0   ' 低調灰 RGB(160,160,160)，與系統維護區按鈕同色調呼應

Public Const SHEET_OLD_CATALOG_MASTER   As String = "遺失檔案舊目錄_總帳"
Public Const MSG_OLD_CATALOG_BATCH_MARK As String = "【封存事件】時間："
Public Const MSG_LEDGER_ORPHANED_HEADER As String = "偵測到帳本與硬碟狀態不一致：暫存表記錄顯示有 "
Public Const MSG_LEDGER_ORPHANED_FOOTER As String = " 筆已匯入的照片資產，但「匯入圖片」資料夾目前是空的（可能被手動刪除或整個資料夾搬移過）。系統已將舊紀錄封存至「舊目錄LOG」分頁存查，並清空目前工作區，請重新匯入這批照片。"
Public Const MSG_RESULT_ORPHANED_HEADER As String = "偵測到帳本與硬碟狀態不一致：暫存表記錄顯示有 "
Public Const MSG_RESULT_ORPHANED_FOOTER As String = " 筆已標記「已匯出」的照片，但「成果」資料夾目前是空的（可能被手動刪除或整個資料夾搬移過）。系統已將這些照片的已匯出狀態重置，稍後會重新壓印。"
