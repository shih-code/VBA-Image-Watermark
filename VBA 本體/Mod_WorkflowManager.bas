' ==========================================================
' MODULE: Mod_WorkflowManager (標準模組)
' PURPOSE: 流程管理大腦中心。負責管控所有階段的執行順序與邏輯劇本，
'          本身不介入任何 Excel 網格讀寫、迴圈運算與系統互動（100% 剛性解耦）。
' EXPORTS: RunPhase1, RunPhase2, RunRebuildUI, RunConsolidateLogs
' IMPORTS: Mod_StringConstants, cls_DAGEngine, cls_Settings, Mod_InteractionAdapter,
'          Mod_SheetSchemaBuilder, Mod_UIRenderer, Mod_FileSystem, Mod_ImportPipeline,
'          Mod_ExportPipeline, Mod_LogManager, View_CanvasManager, Mod_UI_Messenger
' FORBIDDEN: 嚴禁在此模組直接讀寫任何 Excel 儲存格；嚴禁繞過 Mod_InteractionAdapter
'            直接操作工作表物件（見 ADR-004，本模組僅持有跨模組調度順序知識）。
' DEPENDENCIES: ACDS_ContractRegistry.md, ACDS_DependencyGraph.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================
Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_WorkflowManager"
Private Const MODULE_VERSION As String = "1.0.0"

' =========================================================================
' 程序名稱：RunPhase1
' 功用說明：負責指揮第一階段（圖片匯入與清洗）的執行順序與邏輯判斷劇本。
' =========================================================================
Public Sub RunPhase1(ByVal isFileMode As Boolean)
    Dim objDAG As cls_DAGEngine
    Dim objConfig As cls_Settings
    Dim dicSysTokens As Object
    Dim colSelectedPaths As Collection
    Dim strImportTargetFolder As String
    Dim strPickedFolder As String
    Dim wsConfig As Worksheet
    Dim userIntent As VbMsgBoxResult
    Dim lngExistingRecordCount As Long
    
    Application.ScreenUpdating = True
    
    Set objDAG = New cls_DAGEngine
    Set objConfig = New cls_Settings
    
    On Error GoTo ErrorHandler
    
    If isFileMode Then
        objDAG.StartPipeline LOG_DAG_P1_FILES
    Else
        objDAG.StartPipeline LOG_DAG_P1_FOLDER
    End If
    
    If Not Mod_InteractionAdapter.VerifySystemWIA() Then
        objDAG.TerminatePipeline
        Call Mod_UI_Messenger.ShowError(ERR_WIA_NOT_FOUND, TITLE_WIA_MELTDOWN)
        GoTo RunPhase1Cleanup
    End If
    Call Mod_InteractionAdapter.ReleaseUIFocus
    
    objDAG.LogStep LOG_DAG_ENV_CHECK
    Call Mod_SheetSchemaBuilder.BuildSystemSchemas
    Call Mod_UIRenderer.RenderUIElements
    
    objDAG.LogStep LOG_DAG_LOAD_CFG
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets(SHEET_CONFIG)
    On Error GoTo ErrorHandler
    
    If Not wsConfig Is Nothing Then
        objConfig.LoadInterfaceSettings wsConfig
    Else
        objDAG.LogStep MSG_HEAL_CONFIG
    End If
    
    Set dicSysTokens = CreateObject(PROGID_DICTIONARY)
    dicSysTokens.Add "FSO", CreateObject(PROGID_FSO)
    dicSysTokens.Add "STREAM", CreateObject(PROGID_STREAM)
    dicSysTokens.Add "MD5", CreateObject(PROGID_MD5)
    
    ' 5. 【v2 固定拓撲】確保「作業資料夾／匯入圖片／成果／備份」四層結構存在，
    '    取回固定的匯入圖片資料夾路徑，不再需要每次動態跳號計算
    strImportTargetFolder = Mod_FileSystem.GetImportFolderPath(dicSysTokens)
    
    ' 6. 用 Staging 表現有紀錄筆數判斷是否為「進行中的舊工作階段」，決定要不要跳出開新案詢問
    lngExistingRecordCount = Mod_InteractionAdapter.GetStagingRecordCount()
    
    ' 6.5【防呆新增】：帳本有紀錄，但硬碟上的匯入圖片資料夾卻是空的——
    ' 代表資料夾整個被手動刪除或搬移過。這種狀態下不能沿用一般的開新案詢問，
    ' 因為使用者選「是」也沒有實體檔案可以追加，選「否」正常清空反而顯得
    ' 好像什麼都沒發生過。直接封存舊帳本、清空表格，讓使用者能重新匯入。
    If lngExistingRecordCount > 0 Then
        If dicSysTokens("FSO").GetFolder(strImportTargetFolder).Files.Count = 0 Then
            Call Mod_UI_Messenger.ShowInfo( _
                MSG_LEDGER_ORPHANED_HEADER & lngExistingRecordCount & MSG_LEDGER_ORPHANED_FOOTER, _
                TITLE_SYS_INFO)
            Call Mod_InteractionAdapter.ArchiveStaleStagingToOldCatalogLog(ThisWorkbook.Worksheets(SHEET_STAGING))
            Call Mod_InteractionAdapter.ClearStagingWorkarea(wsConfig)
            lngExistingRecordCount = 0
        End If
    End If

    
    If lngExistingRecordCount > 0 Then
        userIntent = Mod_UI_Messenger.AskQuestionCancel(MSG_SESSION_PROMPT, TITLE_SESSION_CONFIRM)
            
        If userIntent = vbYes Then
            ' 點擊「是」：追加模式。固定作業資料夾本來就存在，直接沿用，不需要任何額外動作
        ElseIf userIntent = vbNo Then
            ' 點擊「否」：開新案。【v3 簡化】匯入圖片只是原始照片的工作複製品，
            ' 不具備份價值，直接清空即可；只有「成果」是真正的工作產出，值得備份。
            objDAG.LogStep LOG_DAG_DIR_GEN
            Call Mod_FileSystem.BackupAndClearFolder( _
                Mod_FileSystem.GetResultFolderPath(dicSysTokens), _
                Mod_FileSystem.GetBackupRootPath(dicSysTokens), _
                DIR_NAME_BACKUP_FINAL_PREFIX, dicSysTokens("FSO"))
            Call Mod_FileSystem.ClearFolderContents(strImportTargetFolder, dicSysTokens("FSO"))
            Call Mod_InteractionAdapter.ClearStagingWorkarea(wsConfig)
        Else
            objDAG.TerminatePipeline
            GoTo RunPhase1Cleanup
        End If
    End If
    
    Application.ScreenUpdating = True
    DoEvents
    
    objDAG.LogStep LOG_DAG_PICKER
    If isFileMode Then
        Set colSelectedPaths = Mod_InteractionAdapter.PickMultipleFiles()
        If colSelectedPaths.Count = 0 Then
            objDAG.TerminatePipeline
            GoTo RunPhase1Cleanup
        End If
    Else
        strPickedFolder = Mod_InteractionAdapter.PickSingleFolder()
        If strPickedFolder = "" Then
            objDAG.TerminatePipeline
            GoTo RunPhase1Cleanup
        End If
        Set colSelectedPaths = New Collection
        colSelectedPaths.Add strPickedFolder
    End If
    
    Mod_UI_Messenger.ShowInfo MSG_SYS_START_IMPORT, TITLE_SYS_INFO
    Call Mod_InteractionAdapter.TogglePerformanceMode(True)
    
    objDAG.LogStep LOG_DAG_PIPELINE
    Call Mod_ImportPipeline.ExecuteImportPipeline(colSelectedPaths, strImportTargetFolder, objConfig, dicSysTokens)
    
    Call Mod_InteractionAdapter.TogglePerformanceMode(False)
    objDAG.TerminatePipeline
    Mod_UI_Messenger.ShowInfo MSG_SYS_IMPORT_DONE, TITLE_P1_DONE

RunPhase1Cleanup:
    Application.ScreenUpdating = True
    Set objDAG = Nothing: Set objConfig = Nothing: Set dicSysTokens = Nothing
    Set colSelectedPaths = Nothing: Set wsConfig = Nothing
    Exit Sub

ErrorHandler:
    Call Mod_InteractionAdapter.TogglePerformanceMode(False)
    Application.ScreenUpdating = True
    If Not objDAG Is Nothing Then objDAG.DispatchMeltdown Err.Description
    GoTo RunPhase1Cleanup
End Sub
' =========================================================================
' 程序名稱：RunPhase2
' 功用說明：負責第二階段（圖片壓印與成果匯出）的總指揮劇本。
' =========================================================================
Public Sub RunPhase2()
    Dim objDAG As cls_DAGEngine
    Dim objConfig As cls_Settings
    Dim wsConfig As Worksheet
    Dim wsLogFrame As Worksheet
    Dim wsStagingHost As Worksheet
    
    Dim dicSysTokens As Object
    Dim colTasks As Collection
    Dim sngPctX As Single
    Dim sngPctY As Single
    Dim lngRecordCount As Long
    Dim strExportFolder As String
    Dim flagIsResuming As Boolean
    
    Application.ScreenUpdating = True
    
    Set objDAG = New cls_DAGEngine
    Set objConfig = New cls_Settings
    
    On Error GoTo ErrorHandler
    
    If Not Mod_InteractionAdapter.VerifySystemWIA() Then
        objDAG.TerminatePipeline
        Call Mod_UI_Messenger.ShowError(ERR_WIA_NOT_FOUND, TITLE_WIA_MELTDOWN)
        GoTo RunPhase2Cleanup
    End If
    
    Set dicSysTokens = CreateObject(PROGID_DICTIONARY)
    dicSysTokens.Add "FSO", CreateObject(PROGID_FSO)
    dicSysTokens.Add "STREAM", CreateObject(PROGID_STREAM)
    dicSysTokens.Add "MD5", CreateObject(PROGID_MD5)
    Call Mod_InteractionAdapter.NeutralizeZoomState(dicSysTokens)
    
    objDAG.StartPipeline LOG_DAG_P2_EXPORT
    Mod_UI_Messenger.ShowInfo MSG_SYS_START_EXPORT, TITLE_SYS_INFO
    Call Mod_InteractionAdapter.ReleaseUIFocus
    
    objDAG.LogStep LOG_DAG_PRE_CHECK
    Call Mod_SheetSchemaBuilder.BuildSystemSchemas
    
    objDAG.LogStep LOG_DAG_CONTRACT
    On Error Resume Next
    Set wsStagingHost = ThisWorkbook.Worksheets(SHEET_STAGING)
    Set wsConfig = ThisWorkbook.Worksheets(SHEET_CONFIG)
    On Error GoTo ErrorHandler
    
    If wsStagingHost Is Nothing Then Err.Raise 5501, MOD_NAME, ERR_CONTRACT_VIOLATION
    If Not wsConfig Is Nothing Then objConfig.LoadInterfaceSettings wsConfig
    
    objDAG.LogStep LOG_DAG_PERCENT
    Call View_CanvasManager.GetWatermarkPercentages(sngPctX, sngPctY)
    
    ' 5. 【v2 固定拓撲】直接取得固定的成果資料夾路徑，不再需要從 Staging 表反查匯入路徑
    strExportFolder = Mod_FileSystem.GetResultFolderPath(dicSysTokens)
    
    ' 5.5【防呆新增】：帳本顯示有已匯出紀錄，但成果資料夾卻是空的——代表輸出檔案
    ' 整個被刪除或搬移過。這種狀態下不能讓「已匯出」繼續卡著，否則
    ' PackageStagingTasks 會把這些照片全部濾掉，變成「什麼都不用做」的假象，
    ' 使用者會看到「無資料可執行」的警告，完全摸不著頭緒。
    Dim lngExportedCount As Long
    lngExportedCount = Mod_InteractionAdapter.GetExportedRecordCount()
    If lngExportedCount > 0 Then
        If dicSysTokens("FSO").GetFolder(strExportFolder).Files.Count = 0 Then
            Call Mod_UI_Messenger.ShowInfo( _
                MSG_RESULT_ORPHANED_HEADER & lngExportedCount & MSG_RESULT_ORPHANED_FOOTER, _
                TITLE_SYS_INFO)
            Call Mod_InteractionAdapter.ResetAllExportedFlags(wsStagingHost)
        End If
    End If

    
    ' 6. 判定斷點續跑模式，並且【必須在這裡、打包 colTasks 之前】處理完「否」分支的
    '    備份＋清空＋重置已匯出狀態，下一步打包才會反映最新狀態
    objDAG.LogStep LOG_DAG_RESUME_DETECT
    flagIsResuming = Mod_FileSystem.IsResumingPreviousRun(strExportFolder, dicSysTokens)
    
    If flagIsResuming Then
        If Mod_UI_Messenger.AskQuestion(MSG_EXPORT_MODE_PROMPT, TITLE_EXPORT_SPLIT_CONFIRM) <> vbYes Then
            Call Mod_FileSystem.BackupAndClearFolder( _
                strExportFolder, Mod_FileSystem.GetBackupRootPath(dicSysTokens), _
                DIR_NAME_BACKUP_CHECKPOINT_PREFIX, dicSysTokens("FSO"))
            Call Mod_InteractionAdapter.ResetAllExportedFlags(wsStagingHost)
            flagIsResuming = False
        End If
    End If
    
    ' 7. 大腦不自己數格子，委託轉接器去把 Staging 看板內的成功相片打包
    Set colTasks = Mod_InteractionAdapter.PackageStagingTasks(wsStagingHost)
    lngRecordCount = colTasks.Count
    
    If lngRecordCount = 0 Then
        objDAG.TerminatePipeline
        Mod_UI_Messenger.ShowWarning MSG_NO_DATA_ABORT, TITLE_SYS_INFO
        GoTo RunPhase2Cleanup
    End If
    
    objDAG.LogStep LOG_DAG_LOG_FRAME
    Set wsLogFrame = Mod_LogManager.CreateLogSheetFrame(flagIsResuming)
    
    If Not flagIsResuming Then
        objDAG.LogStep LOG_DAG_DRY_RUN
        Dim strDryRunNotice As String
        strDryRunNotice = MSG_DRY_RUN_HEADER & lngRecordCount & _
                          MSG_DRY_RUN_BODY & Format(sngPctX * 100, "0.0") & _
                          MSG_DRY_RUN_FOOTER & Format(sngPctY * 100, "0.0") & _
                          MSG_DRY_RUN_END
                          
        Application.ScreenUpdating = True
        DoEvents
                          
        If Mod_UI_Messenger.AskQuestion(strDryRunNotice, TITLE_DRY_RUN_CONFIRM) <> vbYes Then
            objDAG.TerminatePipeline
            Mod_UI_Messenger.ShowWarning MSG_USER_ABORT_EXPORT, TITLE_SYS_INFO
            GoTo RunPhase2Cleanup
        End If
    End If
    
    Call Mod_InteractionAdapter.TogglePerformanceMode(True)
    objDAG.LogStep LOG_DAG_LOOP_START
    
    Call Mod_ExportPipeline.ExecuteExportPipeline(colTasks, objConfig, sngPctX, sngPctY, _
                                                  flagIsResuming, strExportFolder, wsStagingHost, _
                                                  wsLogFrame, dicSysTokens)
    
    Call Mod_InteractionAdapter.TogglePerformanceMode(False)
    objDAG.TerminatePipeline

RunPhase2Cleanup:
    Call Mod_InteractionAdapter.RestoreZoomState(dicSysTokens)
    Application.ScreenUpdating = True
    Set dicSysTokens = Nothing: Set objConfig = Nothing: Set objDAG = Nothing
    Set wsStagingHost = Nothing: Set wsConfig = Nothing: Set wsLogFrame = Nothing: Set colTasks = Nothing
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Call Mod_InteractionAdapter.TogglePerformanceMode(False)
    Application.ScreenUpdating = True
    If Not objDAG Is Nothing Then objDAG.DispatchMeltdown Err.Description
    GoTo RunPhase2Cleanup
End Sub

' =========================================================================
' 程序名稱：RunRebuildUI
' =========================================================================
Public Sub RunRebuildUI()
    Dim wsCurrent As Worksheet
    
    ' 1. 開啟最高級別黑幕防線
    Application.ScreenUpdating = False
    
    ' 2. 死鎖原點：記住您點擊按鈕時所在的表單
    On Error Resume Next
    Set wsCurrent = ActiveSheet
    On Error GoTo 0
    
    Call Mod_InteractionAdapter.ReleaseUIFocus
    Call Mod_SheetSchemaBuilder.BuildSystemSchemas
    Call Mod_UIRenderer.RenderUIElements(True)  ' 強制重置版面，恢復原廠設定，見 ADR-012
    Call Mod_InteractionAdapter.PurgeBlankOrphanSheets  ' 清除多餘空白頁面，見 ADR-013
    
    ' 3. 強行拉回：在解開黑幕前，把畫面強制切回您原本的表單
    ' 【防呆修正，見 ADR-014】：wsCurrent 記錄的可能正是 PurgeBlankOrphanSheets
    ' 剛剛才刪除的那張空白頁面（例如全新活頁簿情境下，使用者停留在 Excel
    ' 預設的「工作表1」），此時該物件參照已懸空，直接 .Activate 會拋出物件
    ' 錯誤。改為容錯處理：若原表已不存在，安全退回顯示 ConfigSheet。
    On Error Resume Next
    If Not wsCurrent Is Nothing Then wsCurrent.Activate
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Worksheets(SHEET_CONFIG).Activate
    End If
    On Error GoTo 0
    
    ' 4. 確定各就各位後，原地安全解鎖
    Application.ScreenUpdating = True
    
    Call Mod_UI_Messenger.ShowInfo(MSG_REBUILD_SUCCESS, TITLE_REBUILD_SUCCESS)
End Sub


' =========================================================================
' 程序名稱：RunConsolidateLogs
' =========================================================================
Public Sub RunConsolidateLogs()
    Dim wsCurrent As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 1. 記憶原點
    On Error Resume Next
    Set wsCurrent = ActiveSheet
    On Error GoTo 0
    
    ' 2. 靜默執行歸檔
    Call Mod_LogManager.ConsolidateAllLogs
    
    ' 3. 【關鍵】：趁著黑幕還沒拉開，立刻強行把你傳送回操作台！
    If Not wsCurrent Is Nothing Then wsCurrent.Activate
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' 4. 【最後發言】：已經安全回到原點了，這時才彈出成功視窗！
    Call Mod_UI_Messenger.ShowInfo(MSG_CONSOLIDATE_SUCCESS, TITLE_CONSOLIDATE_SUCCESS)
End Sub

