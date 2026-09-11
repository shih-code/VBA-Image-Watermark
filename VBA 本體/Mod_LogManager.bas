' ==========================================================
' MODULE: Mod_LogManager (標準模組)
' PURPOSE: 日誌管家。專門負責處理所有與「成果日誌工作表」相關的工作，
'          包含建立日誌架構、將日誌改名歸檔，以及將所有歷史日誌合併成總帳。
' EXPORTS: CreateLogSheetFrame, ArchiveLogSheet, ConsolidateAllLogs
' IMPORTS: Mod_StringConstants, Mod_UI_Messenger
' FORBIDDEN: 嚴禁在此模組內私自呼叫 MsgBox 以外的介面通訊（一律透過 Mod_UI_Messenger）；
'            ConsolidateAllLogs 合併歷史資料時嚴禁經過 Windows 剪貼簿。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================
Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_LogManager"
Private Const MODULE_VERSION As String = "1.0.0"

' =========================================================================
' 函數名稱：CreateLogSheetFrame
' 功用說明：在活頁簿中準備好一張名為「LOG_Export_成果」的工作表，做為填寫導出結果的畫檯。
' 邏輯說明：
' 1. Fail-Fast 預檢：檢查活頁簿結構有沒有被鎖定保護，如果被鎖定就無法新增工作表，立刻中斷並報錯。
' 2. 快照記錄：記住使用者在執行這個功能前正在看哪一張工作表（wsActiveSnapshot）。
' 3. 尋找或新建：嘗試在活頁簿中抓取日誌工作表。
'    - 如果找不到這張表：不論是不是接續執行，直接在所有分頁的最右邊新增一張新表，並呼叫 WriteLogHeader 寫入表頭。
'    - 如果找到了，且本次執行是「全新執行，不接續上次（isResuming = False）」：強制清空整張表的舊資料與格式，重新寫入標題表頭。
' 4. 焦點還原：處理完畢後，把 Excel 畫面切換回使用者一開始在看的那張工作表，不打擾使用者的視覺。
' =========================================================================
Public Function CreateLogSheetFrame(ByVal isResuming As Boolean) As Worksheet
    Dim wsNewLogFrame As Worksheet
    Dim wsActiveSnapshot As Worksheet
    
    ' 檢查活頁簿結構是否受保護，防止新增工作表時發生系統當機
    If ThisWorkbook.ProtectStructure Then
        Err.Raise 5501, MOD_NAME, ERR_LOG_STRUCTURE_LOCKED
    End If
    
    ' 記錄進入前的使用者畫面快照，以便稍後切換回來
    On Error Resume Next
    Set wsActiveSnapshot = ActiveSheet
    On Error GoTo 0
    
    ' 關閉螢幕更新以加速處理
    Application.ScreenUpdating = False
    
    ' 嘗試抓取既有的目標日誌工作表
    On Error Resume Next
    Set wsNewLogFrame = ThisWorkbook.Worksheets(SHEET_LOG_EXPORT)
    On Error GoTo 0
    
    ' 判斷邏輯：如果找不到這張日誌工作表
    If wsNewLogFrame Is Nothing Then
        On Error Resume Next
        ' 在現有所有工作表的最後面新增一張工作表
        Set wsNewLogFrame = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error GoTo 0
        
        ' 如果新增失敗，拋出嚴重錯誤
        If wsNewLogFrame Is Nothing Then
            Application.ScreenUpdating = True
            Err.Raise 5502, MOD_NAME, ERR_LOG_CREATE_FAILED
        End If
        
        ' 將新工作表命名為規定的名稱，並呼叫程序寫入標題格式
        wsNewLogFrame.Name = SHEET_LOG_EXPORT
        Call WriteLogHeader(wsNewLogFrame)
        
    ' 判斷邏輯：如果表存在，但這次執行是「全新重來（非接續上次進度）」
    ElseIf Not isResuming Then
        ' 擦除工作表上的所有儲存格內容與格式，重新寫入標題
        wsNewLogFrame.Cells.Clear
        Call WriteLogHeader(wsNewLogFrame)
    End If
    ' 備註：如果 isResuming 為 True 且表已存在，程式會保留現場，直接沿用現有內容繼續向下追加
    
    ' 歸還畫面焦點給使用者一開始的工作表快照
    If Not wsActiveSnapshot Is Nothing Then
        On Error Resume Next
        wsActiveSnapshot.Select
        On Error GoTo 0
    End If
    
    ' 重新開啟螢幕更新
    Application.ScreenUpdating = True
    
    ' 將準備好的工作表物件回傳給呼叫的大腦流程
    Set CreateLogSheetFrame = wsNewLogFrame
End Function

' =========================================================================
' 程序名稱：WriteLogHeader (內部私有程序)
' 功用說明：專門用來設定日誌工作表第一列（標題列）的文字、顏色與欄位寬度。
' 邏輯說明：直接指定特定儲存格的數值，並套用中央常數庫規定的字體大小、粗體、背景色與寬度。
' =========================================================================
Private Sub WriteLogHeader(ByVal wsTarget As Worksheet)
    ' 寫入第一列各直欄的標準標題文字
    wsTarget.Range(ADDR_LOG_HDR_STATUS).Value = COL_NAME_LOG_STATUS
    wsTarget.Range(ADDR_LOG_HDR_NAME).Value = COL_NAME_LOG_FILENAME
    wsTarget.Range(ADDR_LOG_HDR_SRC).Value = COL_NAME_LOG_SRC_SHEET
    wsTarget.Range(ADDR_LOG_HDR_PATH).Value = COL_NAME_LOG_FINAL_PATH
    
    ' 寫入 E1 儲存格的審計安全防線文字
    wsTarget.Range(ADDR_LOG_AUDIT_CELL).Value = MSG_AUDIT_LINE
    
    ' 設定標題列範圍（A1:D1）的文字與背景格式
    With wsTarget.Range(ADDR_LOG_HEADER_RANGE)
        .Font.Bold = True
        .Interior.Color = COLOR_LOG_HEADER_BG
        .HorizontalAlignment = xlCenter
        
        '【最小改動 1】：刪除原本複雜的 Left/InStr 切割，直接指定莫蘭迪微軟正黑體常數
        .Font.Name = VAL_DEFAULT_FONT_CHINESE
    End With
    
    ' 設定 E1 審計文字的外觀格式
    With wsTarget.Range(ADDR_LOG_AUDIT_CELL)
        .Font.Color = COLOR_UI_LABEL_GRAY
        .Font.Size = GEO_FONT_SIZE_LABEL
        
        ' 【最小改動 2】：此處同樣直接指定莫蘭迪微軟正黑體常數
        .Font.Name = VAL_DEFAULT_FONT_CHINESE
        .Font.Italic = True
    End With
    wsTarget.Rows(1).RowHeight = GEO_ROW_HEIGHT_HEADER
    ' 設定各直欄的寬度
    wsTarget.Columns(ADDR_LOG_COL_ABC).ColumnWidth = GEO_LOG_COL_WIDTH_SHORT
    wsTarget.Columns(ADDR_LOG_COL_D).ColumnWidth = GEO_LOG_COL_WIDTH_LONG
    wsTarget.Columns(ADDR_LOG_COL_E).ColumnWidth = GEO_LOG_COL_WIDTH_AUDIT
End Sub

' =========================================================================
' 程序名稱：ArchiveLogSheet
' 功用說明：當一整輪的浮水印批次壓印全部完成後，將當前的日誌工作表改名封存，
'           以便下次執行時，系統不會誤判其為斷點，並且能留下完整的歷史執行軌跡。
' 參數說明：suffixLabel 傳入的改名後綴標籤，例如現在的時間戳記。
' 邏輯說明：
' 1. 自動撐開欄寬（AutoFit），確保超連結與路徑文字完全顯現不被擠壓。
' 2. 組合出預計修改的新工作表名稱，並檢查有沒有超過 Excel 規定的 31 個字元限制。
' 3. 進入「防撞跳號自癒迴圈（Do Loop）」：使用 wsProbe 物件去測試該名稱在活頁簿中存不存在。
'    - 如果不存在：代表名稱安全、無人重複，直接跳出迴圈。
'    - 如果存在：代表發生連點或在同秒內重複執行，啟動跳號自癒判定，將名稱後綴
'      從「_重跑_01」依序累加遞增（02, 03... 最多到 XX），直到不重複為止。
' 4. 正式更換工作表名稱。
' =========================================================================
Public Sub ArchiveLogSheet(ByVal suffixLabel As String)
    Dim wsTargetLog As Worksheet
    Dim wsProbe As Worksheet
    Dim strTargetSheetName As String
    
    ' 嘗試抓取目前的日誌工作表
    On Error Resume Next
    Set wsTargetLog = ThisWorkbook.Worksheets(SHEET_LOG_EXPORT)
    On Error GoTo 0
    
    ' 如果找不到日誌工作表，代表無資料可歸檔，直接跳出程序
    If wsTargetLog Is Nothing Then Exit Sub
    
    ' 歷史結案改名前，強行自動調整所有欄寬，確保文字不被擠壓
    wsTargetLog.Columns(ADDR_LOG_COL_ALL).AutoFit
    
    ' 根據傳入的後綴字串組合出預計的工作表新名稱
    If InStr(suffixLabel, "_") > 0 Then
        strTargetSheetName = SHEET_LOG_PREFIX & suffixLabel
    Else
        strTargetSheetName = SHEET_LOG_PREFIX & suffixLabel & "_" & Format(Now, "yyyymmdd_hhnnss")
    End If
    
    ' 限制名稱長度絕對不能超過 Excel 規定的 31 個字元
    If Len(strTargetSheetName) > GEO_MAX_SHEET_NAME_LEN Then
        strTargetSheetName = Left(strTargetSheetName, GEO_MAX_SHEET_NAME_LEN)
    End If
    
    ' 開始進行同秒歸檔名稱碰撞的自癒探測迴圈
    On Error Resume Next
    Do
        Set wsProbe = Nothing
        ' 嘗試使用新名字去抓工作表
        Set wsProbe = ThisWorkbook.Worksheets(strTargetSheetName)
        
        ' 如果 wsProbe 為 Nothing，代表這個新名字目前沒人使用，可以安全過關退出迴圈
        If wsProbe Is Nothing Then Exit Do
        
        ' 如果名字有人用了，檢查目前有沒有加上「_重跑_」的編號標記
        If InStr(strTargetSheetName, LBL_RERUN_MARK) = 0 Then
            ' 第一次發現名稱碰撞，在名稱後面切斷並加上預設的 "_重跑_01"
            strTargetSheetName = Left(strTargetSheetName, GEO_LOG_TRUNCATE_LEN) & LBL_RERUN_DEFAULT
        Else
            ' 如果早就有了重跑編號，直接讀取最末尾的兩位數編碼，依序向後累加
            Select Case Right(strTargetSheetName, 2)
                Case LBL_RERUN_01: strTargetSheetName = Left(strTargetSheetName, Len(strTargetSheetName) - 2) & LBL_RERUN_02
                Case LBL_RERUN_02: strTargetSheetName = Left(strTargetSheetName, Len(strTargetSheetName) - 2) & LBL_RERUN_03
                Case LBL_RERUN_03: strTargetSheetName = Left(strTargetSheetName, Len(strTargetSheetName) - 2) & LBL_RERUN_04
                Case LBL_RERUN_04: strTargetSheetName = Left(strTargetSheetName, Len(strTargetSheetName) - 2) & LBL_RERUN_05
                Case LBL_RERUN_05: strTargetSheetName = Left(strTargetSheetName, Len(strTargetSheetName) - 2) & LBL_RERUN_06
                Case LBL_RERUN_06: strTargetSheetName = Left(strTargetSheetName, Len(strTargetSheetName) - 2) & LBL_RERUN_07
                Case LBL_RERUN_07: strTargetSheetName = Left(strTargetSheetName, Len(strTargetSheetName) - 2) & LBL_RERUN_08
                Case LBL_RERUN_08: strTargetSheetName = Left(strTargetSheetName, Len(strTargetSheetName) - 2) & LBL_RERUN_09
                Case LBL_RERUN_09: strTargetSheetName = Left(strTargetSheetName, Len(strTargetSheetName) - 2) & LBL_RERUN_10
                Case Else: strTargetSheetName = Left(strTargetSheetName, Len(strTargetSheetName) - 2) & LBL_RERUN_XX
            End Select
        End If
        
        ' 每次修改完名字後，都要再次確保沒有超出 31 個字元的長度上限
        If Len(strTargetSheetName) > GEO_MAX_SHEET_NAME_LEN Then strTargetSheetName = Left(strTargetSheetName, GEO_MAX_SHEET_NAME_LEN)
    Loop
    On Error GoTo 0
    
    ' 正式變更日誌工作表名稱，完成不可逆的歷史改名歸檔
    wsTargetLog.Name = strTargetSheetName
    
    ' 安全釋放參照
    Set wsProbe = Nothing
    Set wsTargetLog = Nothing
End Sub

' =========================================================================
' 程序名稱：ConsolidateAllLogs
' 功用說明：將所有各期獨立產生的「歷史已完成日誌表」，全部流式打包、合併到一張名為
'           「LOG_Master_Archive」的總帳表格中，並將舊表物理刪除，達到檔案減重目的。
' 邏輯說明：
' 1. 關閉螢幕更新與關閉 Excel 系統內建的警告視窗（DisplayAlerts = False）。
'    這樣在稍後程式刪除舊工作表分頁時，Excel 就不會一直跳出對話框干擾操作。
' 2. 尋找或新建總帳工作表（SHEET_MASTER_ARCHIVE）。如果新建立，則自動寫上四大標題欄位。
' 3. 使用 For 迴圈，從最後一張工作表開始倒著往前翻找（Count To 1 Step -1）。
' 4. 如果發現某個工作表的名稱開頭符合「LOG_Export_」且「不是」目前正在使用的那一張表：
'    - 讀取該張歷史表的最後一行。
'    - 如果內有資料（列數大於等於2），採用高雅的【剪貼簿直注技術】：
'      直接讓「總帳表.Value = 歷史表.Value」。這種做法不經過 Windows 剪貼簿，
'      不會破壞使用者原本複製的內容，且速度極快。
'    - 複製完成後，直接下令將該張歷史表物理刪除。
' 5. 全部合併完成後，恢復警告開啟狀態，並彈出成功提示。
' =========================================================================
Public Sub ConsolidateAllLogs()
    Dim wsMasterArchive As Worksheet
    Dim wsLog As Worksheet
    Dim idxSheet As Long
    Dim lngLastRow As Long
    Dim cntArchivePointer As Long
    
    On Error GoTo LogErrorHandler
    
    ' 關閉螢幕更新，並強制關閉 Excel 的彈窗警告視窗（防止刪除分頁時跳出確認對話框）
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 嘗試尋找總帳工作表
    On Error Resume Next
    Set wsMasterArchive = ThisWorkbook.Worksheets(SHEET_MASTER_ARCHIVE)
    On Error GoTo LogErrorHandler
    
    ' 如果找不到總帳工作表，代表是第一次歸檔，自動建立一張新的
    If wsMasterArchive Is Nothing Then
        Set wsMasterArchive = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsMasterArchive.Name = SHEET_MASTER_ARCHIVE
        
        ' 寫入總帳表頭欄位文字
        wsMasterArchive.Range(ADDR_LOG_RANGE_A1D1).Cells(1, 1).Value = COL_NAME_LOG_STATUS
        wsMasterArchive.Range(ADDR_LOG_RANGE_A1D1).Cells(1, 2).Value = COL_NAME_LOG_FILENAME
        wsMasterArchive.Range(ADDR_LOG_RANGE_A1D1).Cells(1, 3).Value = COL_NAME_LOG_SRC_SHEET
        wsMasterArchive.Range(ADDR_LOG_RANGE_A1D1).Cells(1, 4).Value = COL_NAME_LOG_FINAL_PATH
        
        ' 設定總帳標題的粗體、背景色與對齊格式
        With wsMasterArchive.Range(ADDR_LOG_RANGE_A1D1)
            .Font.Bold = True
            .Interior.Color = COLOR_LOG_HEADER_BG
            .HorizontalAlignment = xlCenter
            .Font.Name = Left(VAL_FONT_CHINESE, InStr(VAL_FONT_CHINESE, ",") - 1)
        End With
        ' 調整總帳欄寬
        wsMasterArchive.Columns(COL_LOG_ABC).ColumnWidth = GEO_LOG_COL_WIDTH_SHORT
        wsMasterArchive.Columns(COL_LOG_D).ColumnWidth = GEO_LOG_COL_WIDTH_LONG
    End If
    
    ' 算出目前總帳表格最底下有資料的下一行，作為複製寫入點的指針位置
    cntArchivePointer = wsMasterArchive.Cells(wsMasterArchive.Rows.Count, 1).End(xlUp).Row + 1
    
    ' 使用倒序迴圈，從最後一張工作表開始倒著往前尋找，確保刪除分頁時不會造成陣列索引錯亂
    For idxSheet = ThisWorkbook.Sheets.Count To 1 Step -1
        Set wsLog = ThisWorkbook.Sheets(idxSheet)
        
        ' 篩選條件：如果這張表的名稱開頭包含 "LOG_Export_"，且不是目前正在產出的日誌表
        If InStr(wsLog.Name, Left(SHEET_LOG_EXPORT, 11)) > 0 And wsLog.Name <> SHEET_LOG_EXPORT Then
            ' 找出這張舊日誌表的最後一行資料行數
            lngLastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
            
            ' 如果裡面除了標題之外確實有資料存在
            If lngLastRow >= 2 Then
                ' 採用 100% 絕緣剪貼簿的 Value 直注技術，直接將歷史資料倒入總帳的最末尾中
                wsMasterArchive.Range(wsMasterArchive.Cells(cntArchivePointer, NUM_LOG_COL_STATUS), _
                                      wsMasterArchive.Cells(cntArchivePointer + lngLastRow - 2, NUM_LOG_COL_PATH)).Value = _
                    wsLog.Range(ADDR_LOG_RANGE_A2D_PREFIX & lngLastRow).Value
                
                ' 更新總帳接下來的寫入點指針位置
                cntArchivePointer = wsMasterArchive.Cells(wsMasterArchive.Rows.Count, 1).End(xlUp).Row + 1
            End If
            ' 合併完畢後，直接物理刪除這張舊日誌表分頁，為 Excel 檔案減重
            wsLog.Delete
        End If
    Next idxSheet
    
    ' 恢復畫面更新與正常的警告視窗彈出功能
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' 彈出成功提示(取消
    'Call Mod_UI_Messenger.ShowInfo(MSG_CONSOLIDATE_SUCCESS, TITLE_CONSOLIDATE_SUCCESS)
    Exit Sub
    
LogErrorHandler:
    ' 發生異常時的安全善後：確保重新打開警告彈窗，防止 Excel 從此失去警告能力
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Call Mod_UI_Messenger.ShowError(ERR_CONSOLIDATE_FAILED_PREFIX & Err.Description, TITLE_LOG_MELTDOWN)
End Sub

