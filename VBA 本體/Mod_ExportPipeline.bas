' ==========================================================
' MODULE: Mod_ExportPipeline (標準模組)
' PURPOSE: 負責執行第二階段的圖片浮水印批次壓印與導出核心運算管線。
'          本模組已達成 100% 網格解耦，不持有任何 Excel 格子地址與列號合約。
' EXPORTS: ExecuteExportPipeline
' IMPORTS: Mod_ImageWatermark, Mod_InteractionAdapter, Mod_LogManager, Mod_UI_Messenger,
'          Mod_StringConstants, cls_ImageEntity, cls_Settings
' FORBIDDEN: 嚴禁直接讀寫任何 Excel 儲存格（Cells／Range）；日誌寫入與存檔動作一律
'            委託 Mod_InteractionAdapter 執行，本模組只面對 Collection 與物件。
' DEPENDENCIES: ACDS_ContractRegistry.md, ACDS_DependencyGraph.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_ExportPipeline"
Private Const MODULE_VERSION As String = "1.0.0"

' =========================================================================
' 程序名稱：ExecuteExportPipeline
' 功用說明：純粹的批次壓印管線大腦。只負責跑迴圈調度算力引擎，對 Excel 網格完全無知。
' 【完全解耦合約】：
'   - 舊版的 wsStaging、wsConfig、wsLogFrame 等直接讀寫儲存格的變數全部剔除。
'   - 引進 colTasks：這是一個純粹的物件集合箱，裡面裝滿了由轉接器提早封裝好的 cls_ImageEntity。
'   - 每個圖片實體（entity）內部都已經自動帶有該張照片的所有 Facts（包含操作員打的 UserRemark 備註）。
'   - 寫入日誌與原子性存檔的髒活，全部委託給 Mod_InteractionAdapter 執行，管線只專注於排隊壓印。
' =========================================================================
Public Sub ExecuteExportPipeline( _
    ByVal colTasks As Collection, _
    ByVal objSettings As cls_Settings, _
    ByVal sngPctX As Single, _
    ByVal sngPctY As Single, _
    ByVal flagIsResuming As Boolean, _
    ByVal strExportFolder As String, _
    ByVal wsCanvasHost As Object, _
    ByVal wsLogFrame As Object, _
    ByVal dicTokens As Object)
    
    Dim objLocalFSO As Object
    Dim objImgEntity As cls_ImageEntity
    Dim cntSuccess As Long
    Dim cntResumedSkip As Long
    Dim strThisExportedPath As String
    
    On Error GoTo PipelineErrorHandler
    
    Set objLocalFSO = dicTokens("FSO")
    cntSuccess = 0
    cntResumedSkip = 0
    
    For Each objImgEntity In colTasks
        
        strThisExportedPath = objLocalFSO.BuildPath(strExportFolder, objImgEntity.ImageName)
        
        If flagIsResuming Then
            If objLocalFSO.FileExists(strThisExportedPath) Then
                cntResumedSkip = cntResumedSkip + 1
                GoTo ContinueNextTask
            End If
        End If
        
        Application.StatusBar = "ACDS 調度器：正傳送資產至算力核心 - " & objImgEntity.ImageName
        
        On Error Resume Next
        Call Mod_ImageWatermark.ExecuteWatermarkStream( _
            objImgEntity, objSettings, sngPctX, sngPctY, _
            objImgEntity.UserRemark, strExportFolder, wsCanvasHost, dicTokens)
            
        If Err.Number = 0 Then
            Call Mod_InteractionAdapter.LogExportSuccess(wsLogFrame, objImgEntity.ImageName, strExportFolder)
            cntSuccess = cntSuccess + 1
            Call Mod_InteractionAdapter.MarkStagingRowAsExported(wsCanvasHost, objImgEntity.StagingRowNumber)
        Else
            Call Mod_InteractionAdapter.LogExportFailure(wsLogFrame, objImgEntity.ImageName, Err.Description)
            Err.Clear
        End If
        On Error GoTo PipelineErrorHandler
        
        If (cntSuccess Mod 50) = 0 And cntSuccess > 0 Then
            Call Mod_InteractionAdapter.TriggerAtomicSave(cntSuccess)
        End If

ContinueNextTask:
    Next objImgEntity
    
    Application.StatusBar = False
    
    ' 【v2 精簡】：備份／封存已不是本管線的職責，改由 Mod_WorkflowManager 在
    ' 「斷點結束」或「開新案」這兩個真正有意義的時機點主動呼叫
    ' Mod_FileSystem.BackupAndClearFolder。本模組結案時不需要再知道任何
    ' 備份資料夾命名規則。
    Call Mod_LogManager.ArchiveLogSheet(LBL_LOG_ARCHIVE_DONE_PREFIX & Format(Now, FMT_DATE_YMD & "_" & FMT_TIME_HMSM))
    
    Call Mod_InteractionAdapter.FinalizeWorkareaVisuals(wsCanvasHost)
    
    Dim strDoneMsg As String
    strDoneMsg = MSG_SYS_EXPORT_DONE & LBL_DONE_MSG_RUN & cntSuccess & LBL_DONE_MSG_UNIT
    If cntResumedSkip > 0 Then
        strDoneMsg = strDoneMsg & LBL_DONE_MSG_SKIP & cntResumedSkip & LBL_DONE_MSG_UNIT
    End If
    
    Application.ScreenUpdating = True
    DoEvents
    
     Call Mod_UI_Messenger.ShowInfo(strDoneMsg, TITLE_P2_DONE)

     GoTo Cleanup
 
PipelineErrorHandler:
     Application.StatusBar = False
     Dim lngErrNum As Long, strErrDesc As String
     lngErrNum = Err.Number
     strErrDesc = Err.Description

Cleanup:
     Set objImgEntity = Nothing
     Set objLocalFSO = Nothing

     If lngErrNum <> 0 Then
         Err.Raise lngErrNum, MOD_NAME & ".ExecuteExportPipeline", strErrDesc
     End If
 End Sub
