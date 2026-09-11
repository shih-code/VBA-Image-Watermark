' ==========================================================
' MODULE: Mod_InteractionAdapter (標準模組)
' PURPOSE: 介面互動轉接器。專門處理所有與 Excel 視窗、工作表底層網格讀寫、
'          以及 Windows 檔案選擇視窗的直接互動。本模組為全專案唯一的網格操作實體。
' EXPORTS: VerifySystemWIA, ReleaseUIFocus, PickSingleFolder, PickMultipleFiles,
'          ClearStagingWorkarea, GetStagingRecordCount,
'          PackageStagingTasks, LogExportSuccess,
'          LogExportFailure, TriggerAtomicSave, FinalizeWorkareaVisuals,
'          PurgeBlankOrphanSheets（見 ADR-013）, MarkStagingRowAsExported,
'          ResetAllExportedFlags, ArchiveStaleStagingToOldCatalogLog,
'          GetExportedRecordCount
' IMPORTS: cls_ImageEntity, Mod_StringConstants
' FORBIDDEN: 嚴禁被下層算力引擎（Mod_FileSystem／Mod_Crypto／Mod_ImageWatermark）呼叫，
'            僅供 Mod_WorkflowManager／Mod_ImportPipeline／Mod_ExportPipeline 呼叫。
' DEPENDENCIES: ACDS_ContractRegistry.md, ACDS_DependencyGraph.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================
Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_InteractionAdapter"
Private Const MODULE_VERSION As String = "1.0.0"

' =========================================================================
' 函數名稱：VerifySystemWIA
' 功用說明：檢查這台電腦的 Windows 系統中，有沒有安裝處理圖片所需的 WIA 必要元件。
' =========================================================================
Public Function VerifySystemWIA() As Boolean
    Dim objWiaProbe As Object
    On Error Resume Next
    Set objWiaProbe = CreateObject(PROGID_WIA_IMAGE)
    If Err.Number = 0 Then
        VerifySystemWIA = True
    Else
        VerifySystemWIA = False
    End If
    Set objWiaProbe = Nothing
    On Error GoTo 0
End Function

' =========================================================================
' 程序名稱：ReleaseUIFocus
' 功用說明：強制轉移前台畫面的選取焦點，防止圖形物件死鎖。
' =========================================================================
Public Sub ReleaseUIFocus()
    On Error Resume Next
    ActiveSheet.Range(ADDR_CFG_FONT_CHINESE).Select
    On Error GoTo 0
End Sub

' =========================================================================
' 函數名稱：PickSingleFolder
' 功用說明：彈出 Windows 標準的資料夾選取視窗。
' =========================================================================
Public Function PickSingleFolder() As String
    Dim objFolderPicker As Object
    Set objFolderPicker = Application.FileDialog(4)
    
    objFolderPicker.Title = MSG_PICKER_FOLDER_TITLE
    objFolderPicker.InitialFileName = ThisWorkbook.Path & BACKSLASH_STR
    
    If objFolderPicker.Show = -1 Then
        PickSingleFolder = CStr(objFolderPicker.SelectedItems(1))
    Else
        PickSingleFolder = ""
    End If
    Set objFolderPicker = Nothing
End Function

' =========================================================================
' 函數名稱：PickMultipleFiles
' 功用說明：彈出 Windows 標準的檔案選取視窗（支援複選圖片）。
' =========================================================================
Public Function PickMultipleFiles() As Collection
    Dim objFilePicker As Object
    Dim colSelectedPaths As Collection
    Dim varItem As Variant
    
    Set colSelectedPaths = New Collection
    Set objFilePicker = Application.FileDialog(3)
    
    objFilePicker.AllowMultiSelect = True
    objFilePicker.Title = MSG_PICKER_FILE_TITLE
    objFilePicker.InitialFileName = ThisWorkbook.Path & BACKSLASH_STR
    objFilePicker.Filters.Clear
    objFilePicker.Filters.Add LBL_PICKER_IMAGE_FILTER_PREFIX & FILE_FILTER_IMAGES & LBL_PICKER_IMAGE_FILTER_SUFFIX, FILE_FILTER_IMAGES
    
    If objFilePicker.Show = -1 Then
        For Each varItem In objFilePicker.SelectedItems
            colSelectedPaths.Add CStr(varItem)
        Next varItem
    End If
    
    Set PickMultipleFiles = colSelectedPaths
    Set objFilePicker = Nothing
End Function

' =========================================================================
' 程序名稱：TogglePerformanceMode
' 功用說明：控制 Excel 的畫面更新與計算模式以實施加速。
' =========================================================================
Public Sub TogglePerformanceMode(ByVal isFastMode As Boolean)
    If isFastMode Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
    Else
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub

' =========================================================================
' 程序名稱：NeutralizeZoomState / RestoreZoomState
' 功用說明：強制鎖定 100% 視窗縮放比，防範幾何運算位移?差。
' =========================================================================
Public Sub NeutralizeZoomState(ByVal sysTokens As Object)
    sysTokens.Add "OLD_ZOOM", ActiveWindow.Zoom
    ActiveWindow.Zoom = GEO_WINDOW_ZOOM_STANDARD
End Sub

Public Sub RestoreZoomState(ByVal sysTokens As Object)
    If Not sysTokens Is Nothing Then
        If sysTokens.Exists("OLD_ZOOM") Then
            ActiveWindow.Zoom = sysTokens("OLD_ZOOM")
        End If
    End If
End Sub



' =========================================================================
' 程序名稱：ClearStagingWorkarea
' 功用說明：徹底物理擦除暫存工作表上的所有網格與殘留圖表。
' =========================================================================
Public Sub ClearStagingWorkarea(ByVal wsConfig As Worksheet)
    Dim wsStaging As Worksheet
    On Error Resume Next
    Set wsStaging = ThisWorkbook.Worksheets(SHEET_STAGING)
    If Not wsStaging Is Nothing Then
        ' 舊版致命行：wsStaging.Cells.Clear -> 連帶把表頭文字與格式一起殺光
        ' 修正後防線：只清空第 2 列以下的歷史相片數據，第 1 列剛性表頭毫髮無損
        If wsStaging.FilterMode Then wsStaging.ShowAllData
        wsStaging.Rows(NUM_ROW_DATA_START & ":" & wsStaging.Rows.Count).Clear
        
        wsStaging.ChartObjects.Delete
        wsStaging.Tab.ColorIndex = xlColorIndexNone
    End If
    If Not wsConfig Is Nothing Then
        wsConfig.Range(ADDR_CFG_SESSION_STATUS).Value = ""
    End If
    On Error GoTo 0
End Sub

' =========================================================================
' 函數名稱：GetStagingRecordCount
' 功用說明：計算目前暫存看板上的排隊照片總張數。
' =========================================================================
Public Function GetStagingRecordCount() As Long
    Dim wsStaging As Worksheet
    Set wsStaging = ThisWorkbook.Worksheets(SHEET_STAGING)
    GetStagingRecordCount = wsStaging.Cells(wsStaging.Rows.Count, NUM_STG_COL_NAME).End(xlUp).Row - 1
    If GetStagingRecordCount < 0 Then GetStagingRecordCount = 0
End Function


' =========================================================================
' 【新增解耦核心】：PackageStagingTasks
' 功用說明：專職充當大腦的資料搬運工。把 Staging 表格子裡的原始 facts，打包裝入 Collection 物件箱中。
'          有了它，管線從此不需要知道 Excel 看板網格的存在，徹底解耦！
' =========================================================================
 Public Function PackageStagingTasks(ByVal wsStaging As Worksheet) As Collection
     Dim colTasks As Collection
     Dim lngLastRow As Long
     Dim idx As Long
     Dim objImgEntity As cls_ImageEntity
     
     Set colTasks = New Collection
     lngLastRow = wsStaging.Cells(wsStaging.Rows.Count, NUM_STG_COL_NAME).End(xlUp).Row
     
     For idx = 2 To lngLastRow

        ' 【根治重複檔案修正】：只打包「匯入成功」且「尚未匯出」的列。
        ' 已匯出過的列直接跳過，快取斷點與封存快照才不會一次又一次裝進同一批舊圖片。
        If wsStaging.Cells(idx, NUM_STG_COL_STATUS).Value = STATUS_IMPORT_OK _
           And Trim(CStr(wsStaging.Cells(idx, NUM_STG_COL_EXPORTED).Value)) <> STATUS_STG_EXPORTED Then
             Set objImgEntity = New cls_ImageEntity
             
            ' 剛性對齊：精確傳入 8 個參數，最後一個是這筆資料在 Staging 表的實際列號，
            ' 供匯出成功後 MarkStagingRowAsExported 精準回寫，不靠檔名或雜湊反查
             Call objImgEntity.InitializeEntity( _
                 wsStaging.Cells(idx, NUM_STG_COL_NAME).Value, _
                 wsStaging.Cells(idx, NUM_STG_COL_MOD).Value, _
                 wsStaging.Cells(idx, NUM_STG_COL_PATH).Value, _
                 wsStaging.Cells(idx, NUM_STG_COL_DATE).Value, _
                 wsStaging.Cells(idx, NUM_STG_COL_REL).Value, _
                 wsStaging.Cells(idx, NUM_STG_COL_REMARK).Value, _
                 wsStaging.Cells(idx, NUM_STG_COL_HASH).Value, _
                 idx)
             
             colTasks.Add objImgEntity
         End If
     Next idx
     
     Set PackageStagingTasks = colTasks
 End Function
' =========================================================================
' 【新增解耦核心】：LogExportSuccess
' 功用說明：代勞管線回寫壓印成功的日誌紀錄，並內置 255 字超連結字數防禦。
' =========================================================================
Public Sub LogExportSuccess(ByVal wsLogFrame As Worksheet, ByVal strImageName As String, ByVal strExportFolder As String)
    Dim idxLogNextRow As Long
    Dim strFinalPath As String
    
    idxLogNextRow = wsLogFrame.Cells(wsLogFrame.Rows.Count, 1).End(xlUp).Row + 1
    
    wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_STATUS).Value = STATUS_EXPORT_OK
    wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_NAME).Value = strImageName
    wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_SRC).Value = SHEET_STAGING
    
    strFinalPath = strExportFolder & BACKSLASH_STR & strImageName
    
    ' 超連結 255 長度安全硬體限制自癒防禦
    If Len(strFinalPath) > NUM_MAX_HYPERLINK_LEN Then
        wsLogFrame.Hyperlinks.Add anchor:=wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_PATH), _
                                  Address:=strExportFolder, TextToDisplay:=strFinalPath
    Else
        wsLogFrame.Hyperlinks.Add anchor:=wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_PATH), _
                                  Address:=strFinalPath, TextToDisplay:=strFinalPath
    End If
End Sub

' =========================================================================
' 程序名稱：LogExportFailure
' 功用說明：回寫失敗日誌，並自動染上高雅的莫蘭迪乾燥玫瑰紅
' =========================================================================
Public Sub LogExportFailure(ByVal wsLogFrame As Worksheet, ByVal strImageName As String, ByVal strErrDesc As String)
    Dim idxLogNextRow As Long
    idxLogNextRow = wsLogFrame.Cells(wsLogFrame.Rows.Count, 1).End(xlUp).Row + 1
    
    wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_STATUS).Value = STATUS_EXPORT_ERR
    
    ' 換裝 Facts：一鍵換上莫蘭迪低飽和乾燥玫瑰色，兼顧警示與法式精緻感
    wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_STATUS).Interior.Color = STATUS_EXPORT_ERR_BG
    wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_STATUS).Font.Color = STATUS_EXPORT_ERR_TXT
    wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_STATUS).Font.Name = VAL_DEFAULT_FONT_CHINESE
    
    wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_NAME).Value = strImageName
    wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_SRC).Value = SHEET_STAGING
    wsLogFrame.Cells(idxLogNextRow, NUM_LOG_COL_PATH).Value = LBL_LOG_ERR_PREFIX & strErrDesc
End Sub

' =========================================================================
' 程序名稱：MarkStagingRowAsExported
' 功用說明：單張圖片壓印匯出成功後，用 cls_ImageEntity 攜帶的實際列號，
'           精準回寫 Staging 表對應列的「已匯出」狀態。取代舊有靠快取斷點
'           資料夾殘留檔案判斷續跑的脆弱旁證。
' =========================================================================
Public Sub MarkStagingRowAsExported(ByVal wsStaging As Worksheet, ByVal lngRow As Long)
    If lngRow < NUM_ROW_DATA_START Then Exit Sub
    wsStaging.Cells(lngRow, NUM_STG_COL_EXPORTED).Value = STATUS_STG_EXPORTED
End Sub

' =========================================================================
' 程序名稱：ResetAllExportedFlags
' 功用說明：【v2 新增】把 Staging 表所有資料列的「已匯出」狀態整批清空。
'           供「匯出時選否＝結束並重印」情境呼叫——成果資料夾已經被整批
'           備份並清空，這裡必須同步讓 Staging 表恢復成「全部尚未匯出」，
'           下一次 PackageStagingTasks 才會重新把全部照片納入打包清單。
' =========================================================================
Public Sub ResetAllExportedFlags(ByVal wsStaging As Worksheet)
    Dim lngLastRow As Long
    lngLastRow = wsStaging.Cells(wsStaging.Rows.Count, NUM_STG_COL_NAME).End(xlUp).Row
    If lngLastRow < NUM_ROW_DATA_START Then Exit Sub
    wsStaging.Range(wsStaging.Cells(NUM_ROW_DATA_START, NUM_STG_COL_EXPORTED), _
                     wsStaging.Cells(lngLastRow, NUM_STG_COL_EXPORTED)).ClearContents
End Sub

' =========================================================================
' 程序名稱：ArchiveStaleStagingToOldCatalogLog
' 功用說明：【v2 修正】當偵測到「帳本有紀錄、硬碟卻是空的」不一致狀態時，
'           把目前 Staging 表的幽靈紀錄，用直注技術（不經剪貼簿）整批附加進
'           單一固定的「舊目錄LOG_總帳」分頁，每批資料前插入一行時間戳記標記
'           以利事後追溯是哪一次事件留下的。不再每次另開一張新分頁，
'           自然不會有分頁越積越多、需要事後壓縮合併的問題。
'           【畫面釘選】全程關閉螢幕更新，結束後把畫面釘回操作介面。
' =========================================================================
Public Sub ArchiveStaleStagingToOldCatalogLog(ByVal wsStaging As Worksheet)
    Dim wsMaster As Worksheet
    Dim lngStagingLastRow As Long
    Dim lngMasterNextRow As Long
    Dim lngDataRows As Long
    Dim flagOldScreenState As Boolean
    
    flagOldScreenState = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    lngStagingLastRow = wsStaging.Cells(wsStaging.Rows.Count, NUM_STG_COL_NAME).End(xlUp).Row
    lngDataRows = lngStagingLastRow - NUM_ROW_DATA_START + 1
    If lngDataRows < 1 Then GoTo ArchiveCleanup
    
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Worksheets(SHEET_OLD_CATALOG_MASTER)
    On Error GoTo 0
    
    If wsMaster Is Nothing Then
        Set wsMaster = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsMaster.Name = SHEET_OLD_CATALOG_MASTER
        wsMaster.Tab.Color = COLOR_UI_LABEL_GRAY
        ' 表頭直接沿用 Staging 現有表頭文字，不經剪貼簿、不重新發明欄位定義
        wsMaster.Range(ADDR_STG_HDR_RANGE).Value = wsStaging.Range(ADDR_STG_HDR_RANGE).Value
    End If
    
    lngMasterNextRow = wsMaster.Cells(wsMaster.Rows.Count, NUM_STG_COL_NAME).End(xlUp).Row + 1
    If lngMasterNextRow < NUM_ROW_DATA_START Then lngMasterNextRow = NUM_ROW_DATA_START
    
    ' 封存分隔標記：這批舊帳是哪一次事件、什麼時間點被封存的
    wsMaster.Cells(lngMasterNextRow, 1).Value = MSG_OLD_CATALOG_BATCH_MARK & Format(Now, "yyyy-mm-dd hh:nn:ss")
    lngMasterNextRow = lngMasterNextRow + 1
    
    ' 【直注技術】：整批資料一次性用陣列賦值搬過去，不經剪貼簿，
    ' 不干擾使用者原本複製的內容，速度也遠快於逐列迴圈寫入
    wsMaster.Range(wsMaster.Cells(lngMasterNextRow, 1), _
                    wsMaster.Cells(lngMasterNextRow + lngDataRows - 1, NUM_STG_COL_EXPORTED)).Value = _
        wsStaging.Range(wsStaging.Cells(NUM_ROW_DATA_START, 1), _
                         wsStaging.Cells(lngStagingLastRow, NUM_STG_COL_EXPORTED)).Value

ArchiveCleanup:
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_CONFIG).Activate
    On Error GoTo 0
    Application.ScreenUpdating = flagOldScreenState
    Set wsMaster = Nothing
End Sub
' =========================================================================
' 函數名稱：GetExportedRecordCount
' 功用說明：【防呆新增】計算 Staging 表中，已被標記「已匯出」的照片筆數。
'           供 Phase 2 開場判斷「帳本說已匯出、硬碟卻找不到成品」這種
'           不一致狀態時使用，跟 GetStagingRecordCount 各自服務不同層級的判斷。
' =========================================================================
Public Function GetExportedRecordCount() As Long
    Dim wsStaging As Worksheet
    Dim lngLastRow As Long
    Dim idx As Long
    Dim lngCount As Long
    
    Set wsStaging = ThisWorkbook.Worksheets(SHEET_STAGING)
    lngLastRow = wsStaging.Cells(wsStaging.Rows.Count, NUM_STG_COL_NAME).End(xlUp).Row
    
    For idx = NUM_ROW_DATA_START To lngLastRow
        If Trim(CStr(wsStaging.Cells(idx, NUM_STG_COL_EXPORTED).Value)) = STATUS_STG_EXPORTED Then
            lngCount = lngCount + 1
        End If
    Next idx
    
    GetExportedRecordCount = lngCount
End Function
' =========================================================================
' 【新增解耦核心】：TriggerAtomicSave
' 功用說明：代勞管線發動原子性中斷防禦定時存檔。
' =========================================================================
Public Sub TriggerAtomicSave(ByVal cntSuccess As Long)
    DoEvents
    Application.StatusBar = MSG_STATUS_BAR_RUN_PREFIX & cntSuccess & MSG_STATUS_BAR_RUN_SUFFIX
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
End Sub

' =========================================================================
' 【新增解耦核心】：FinalizeWorkareaVisuals
' 功用說明：代勞管線在全部結束後，重寫設定檔結案狀態字串、清除隱藏製圖 Chart 範本。
' =========================================================================
Public Sub FinalizeWorkareaVisuals(ByVal wsCanvasHost As Worksheet)
    Dim wsConfig As Worksheet
    On Error Resume Next
    
    Set wsConfig = ThisWorkbook.Worksheets(SHEET_CONFIG)
    If Not wsConfig Is Nothing Then
        wsConfig.Range(ADDR_CFG_SESSION_STATUS).Value = MSG_SESSION_CLOSED_PREFIX & Format(Now, "yyyy-mm-dd hh:mm") & MSG_SESSION_CLOSED_SUFFIX
        wsConfig.Range(ADDR_CFG_SESSION_STATUS).Font.Color = COLOR_UI_LABEL_GRAY
        wsConfig.Range(ADDR_CFG_SESSION_STATUS).Font.Size = GEO_FONT_SIZE_CANVAS
        wsConfig.Range(ADDR_CFG_SESSION_STATUS).Font.Name = Left(VAL_FONT_CHINESE, InStr(VAL_FONT_CHINESE, ",") - 1)
    End If
    
    If Not wsCanvasHost Is Nothing Then
        wsCanvasHost.Tab.Color = COLOR_UI_LABEL_GRAY
        wsCanvasHost.ChartObjects(SH_TEMPLATE_CHART_NAME).Delete
    End If
    On Error GoTo 0
End Sub


' =========================================================================
' 程序名稱：PurgeBlankOrphanSheets
' 功用說明：掃描活頁簿中所有「非系統表」且「內容真正空白」的工作表，
'           經使用者確認後一併刪除，供「一鍵重整工作檯」呼叫（見 ADR-013）。
' 判斷邏輯：
'   1. 四張系統表（ConfigSheet／Staging_Images／LOG_Export_成果／
'      LOG_Master_Archive）一律強制排除，不論名稱或內容為何——這是
'      獨立於「空白判斷」之外的第一道防線，避免系統表恰好處於空白
'      狀態的時間點（例如歸檔剛完成、下次匯出前）被誤判刪除。
'   2. 名稱以 ### 開頭者視為使用者手動標記保護，一律排除。
'   3. 通過前兩關後，使用 CountA(UsedRange) = 0 判斷是否真正空白
'      （沒有任何儲存格含有內容，包含公式或格式化但無值的儲存格
'      不列入空白判斷的干擾範圍）。
'   4. 有候選清單才彈出確認視窗列出名稱，使用者按「是」才真正刪除，
'      按「否」則完全不動作，維持刪除前必須確認的專案慣例。
' =========================================================================
Public Sub PurgeBlankOrphanSheets()
    Dim ws              As Worksheet
    Dim colCandidates    As New Collection
    Dim strDisplayList  As String
    Dim v                As Variant
    
    ' 結構保護檢查（比照 Mod_LogManager.CreateLogSheetFrame 既有慣例，見 ADR-013）：
    ' 若活頁簿結構被保護導致無法刪除工作表，立即明確中斷，不自動解鎖代管，
    ' 讓使用者自行決定是否要解除保護後再重試。
    If ThisWorkbook.ProtectStructure Then
        Err.Raise 5504, MOD_NAME, ERR_STRUCTURE_LOCKED_PURGE
    End If
    
    For Each ws In ThisWorkbook.Worksheets
        ' --- 第一道防線：系統表無條件排除，不受空白判斷影響 ---
        If ws.Name = SHEET_CONFIG Or ws.Name = SHEET_STAGING Or _
           ws.Name = SHEET_LOG_EXPORT Or ws.Name = SHEET_MASTER_ARCHIVE Then
            GoTo NextSheetProbe
        End If
        
        ' --- 第二道防線：使用者手動標記保護 ---
        If Left(ws.Name, Len(LBL_PROTECT_MARK)) = LBL_PROTECT_MARK Then
            GoTo NextSheetProbe
        End If
        
        ' --- 空白判斷：整張表沒有任何儲存格含有內容 ---
        If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
            colCandidates.Add ws.Name
        End If
NextSheetProbe:
    Next ws
    
    ' 沒有任何候選頁面，安靜結束，不需要打擾使用者
    If colCandidates.Count = 0 Then Exit Sub
    
    ' 組合確認視窗要顯示的頁面名稱清單
    For Each v In colCandidates
        strDisplayList = strDisplayList & "- " & v & vbCrLf
    Next v
    
    ' 動手刪除前，強制經過使用者確認（預設焦點鎖定在「否」，見 Mod_UI_Messenger 慣例）
    If Mod_UI_Messenger.AskQuestion(MSG_PURGE_CONFIRM_PREFIX & strDisplayList & MSG_PURGE_CONFIRM_SUFFIX, TITLE_PURGE_CONFIRM) = vbYes Then
        Application.DisplayAlerts = False
        For Each v In colCandidates
            On Error Resume Next
            ThisWorkbook.Worksheets(v).Delete
            On Error GoTo 0
        Next v
        Application.DisplayAlerts = True
    End If
End Sub

' =========================================================================
' 函數名稱：LoadExistingHashSet
' 功用說明：讀取 Staging_Images 既有列的所有雜湊值，回傳一個字典（點名簿），
'           供匯入管線做跨批次去重比對。取代原本 Mod_ImportPipeline 內部
'           直接持有 wsStaging 進行讀取的做法。
' =========================================================================
Public Function LoadExistingHashSet() As Object
    Dim wsStaging   As Worksheet
    Dim dicHashes   As Object
    Dim lngLastRow  As Long
    Dim lngRow      As Long
    Dim strHash     As String
    
    Set dicHashes = CreateObject(PROGID_DICTIONARY)
    Set wsStaging = ThisWorkbook.Worksheets(SHEET_STAGING)
    
    lngLastRow = wsStaging.Cells(wsStaging.Rows.Count, NUM_STG_COL_NAME).End(xlUp).Row
    
    If lngLastRow >= NUM_ROW_DATA_START Then
        For lngRow = NUM_ROW_DATA_START To lngLastRow
            strHash = Trim(wsStaging.Cells(lngRow, NUM_STG_COL_HASH).Value)
            If strHash <> "" And strHash <> HASH_FAILED_MARK Then
                If Not dicHashes.Exists(strHash) Then
                    dicHashes.Add strHash, wsStaging.Cells(lngRow, NUM_STG_COL_NAME).Value
                End If
            End If
        Next lngRow
    End If
    
    Set LoadExistingHashSet = dicHashes
    Set wsStaging = Nothing
End Function

' =========================================================================
' 程序名稱：WriteImportResults
' 功用說明：接收 Mod_ImportPipeline 處理完畢的 cls_ImportResult 集合，統一
'           在此一次寫入 Staging_Images 對應列，包含成功／失敗兩種狀態的
'           儲存格內容與標色。寫入完成後自動撐開欄寬。
' =========================================================================
Public Sub WriteImportResults(ByVal colResults As Collection)
    Dim wsStaging   As Worksheet
    Dim lngNextRow  As Long
    Dim objResult   As cls_ImportResult
    
    If colResults Is Nothing Then Exit Sub
    If colResults.Count = 0 Then Exit Sub
    
    Set wsStaging = ThisWorkbook.Worksheets(SHEET_STAGING)
    lngNextRow = wsStaging.Cells(wsStaging.Rows.Count, NUM_STG_COL_NAME).End(xlUp).Row + 1
    If lngNextRow < NUM_ROW_DATA_START Then lngNextRow = NUM_ROW_DATA_START
    
    For Each objResult In colResults
        If objResult.IsSuccess Then
            wsStaging.Cells(lngNextRow, NUM_STG_COL_STATUS).Value = STATUS_IMPORT_OK
            wsStaging.Cells(lngNextRow, NUM_STG_COL_STATUS).Interior.ColorIndex = xlColorIndexNone
            wsStaging.Cells(lngNextRow, NUM_STG_COL_STATUS).Font.ColorIndex = xlColorIndexAutomatic
            
            wsStaging.Cells(lngNextRow, NUM_STG_COL_NAME).Value = objResult.ImageName
            wsStaging.Cells(lngNextRow, NUM_STG_COL_REMARK).Value = objResult.UserRemark
            wsStaging.Cells(lngNextRow, NUM_STG_COL_MOD).Value = objResult.LastModified
            wsStaging.Cells(lngNextRow, NUM_STG_COL_PATH).Value = objResult.CurrentPath
            wsStaging.Cells(lngNextRow, NUM_STG_COL_DATE).Value = objResult.ImportDate
            wsStaging.Cells(lngNextRow, NUM_STG_COL_REL).Value = objResult.ImportRelativePath
            wsStaging.Cells(lngNextRow, NUM_STG_COL_HASH).Value = objResult.FileHash
        Else
            wsStaging.Cells(lngNextRow, NUM_STG_COL_STATUS).Value = STATUS_IMPORT_ERR
            wsStaging.Cells(lngNextRow, NUM_STG_COL_STATUS).Interior.Color = STATUS_EXPORT_ERR_BG
            wsStaging.Cells(lngNextRow, NUM_STG_COL_STATUS).Font.Color = STATUS_EXPORT_ERR_TXT
            wsStaging.Cells(lngNextRow, NUM_STG_COL_STATUS).Font.Name = VAL_DEFAULT_FONT_CHINESE
            
            wsStaging.Cells(lngNextRow, NUM_STG_COL_NAME).Value = objResult.ImageName
            wsStaging.Cells(lngNextRow, NUM_STG_COL_REMARK).Value = ""
            wsStaging.Cells(lngNextRow, NUM_STG_COL_MOD).Value = objResult.FailReason
        End If
        lngNextRow = lngNextRow + 1
    Next objResult
    
    wsStaging.Columns.AutoFit
    Set wsStaging = Nothing
End Sub
