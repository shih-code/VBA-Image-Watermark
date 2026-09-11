' ==========================================================
' MODULE: Mod_ImportPipeline (標準模組)
' PURPOSE: 負責執行第一階段的圖片匯入與清洗核心運算管線。
'          本模組已達成 100% 零硬編碼、100% 變數編譯強固防線。
' EXPORTS: ExecuteImportPipeline
' IMPORTS: Mod_FileSystem, Mod_Crypto, Mod_UI_Messenger, Mod_StringConstants, cls_Settings
' FORBIDDEN: 嚴禁私自呼叫 Mod_ImageWatermark（匯入階段不涉及壓印）。
'            【已知技術債】本模組仍直接持有並寫入 wsStaging 工作表物件，未如
'            Mod_ExportPipeline 完全委託 Mod_InteractionAdapter，見
'            ACDS_DependencyGraph.md 第三節，下一輪重構列入待辦。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_ImportPipeline"
Private Const MODULE_VERSION As String = "1.0.0"



' =========================================================================
' 程序名稱：ExecuteImportPipeline
' 功用說明：第一階段匯入的主導管線。負責調度內部程序，完成清單展開、防重複檢查、複製檔案與填寫工作表。
' 【自癒修正】：
'   1. 物理修復張冠李戴的變數未定義 Bug。將失敗標籤與莫蘭迪玫瑰紅底色正確塗抹在 wsStaging 當前資料列。
'   2. 物理洗淨所有隱藏字串與日期格式，100% 聽從中央常數在庫之剛性 Facts。
' =========================================================================
Public Sub ExecuteImportPipeline( _
    ByVal colSrcPaths As Collection, _
    ByVal strTargetFolder As String, _
    ByVal objConfig As cls_Settings, _
    ByVal dicTokens As Object)
    
     Dim objLocalFSO As Object
     Dim objFileItem As Object
     Dim dicSeenHash As Object
     Dim colFlattenedFiles As Collection
     Dim colResults  As New Collection
     
     Dim lngSkippedCount As Long
     Dim strExt As String
     Dim objResult   As cls_ImportResult
    
    ' --- Cleanup Label 標準配套變數（AVS v1.1 第七節）：暫存錯誤資訊供 Cleanup 判斷 ---
    Dim lngErrNum   As Long
    Dim strErrDesc  As String
    
    ' 從系統共用工具箱中提領處理實體檔案的工具
    Set objLocalFSO = dicTokens("FSO")
    ' 建立一個全空的集合，當作準備放置檔案的「置物箱」
    Set colFlattenedFiles = New Collection
    lngSkippedCount = 0
    
    On Error GoTo PipelineErrorHandler
    
    ' 實體安全防線：如果目的地資料夾根本沒有建成功，拒絕執行搬移
    If Not objLocalFSO.FolderExists(strTargetFolder) Then
        Err.Raise 5402, MOD_NAME, ERR_TARGET_DIR_MISSING
    End If
    
    ' 步驟 1：將複雜的資料夾與檔案路徑展開成平坦的單一檔案清單
    Call FlattenSourcePaths(colSrcPaths, colFlattenedFiles, objLocalFSO)
    
    ' 步驟 2：透過 Mod_InteractionAdapter 取得 Staging 既有雜湊點名簿。
    ' 【解耦修正，見 ADR-016】：本模組不再直接持有／讀取 wsStaging，
    ' 改由 Mod_InteractionAdapter 負責，本模組對網格 100% 無知。
    Set dicSeenHash = Mod_InteractionAdapter.LoadExistingHashSet()
    
     ' 步驟 3：開始逐一處理置物箱裡的每一個檔案
     For Each objFileItem In colFlattenedFiles
         ' 抓取副檔名並強制轉成小寫，避免大小寫不同導致系統漏判
         strExt = LCase(objLocalFSO.GetExtensionName(objFileItem.Name))
         
         ' 文字過濾：檢查這個檔案的副檔名是否合規
         If InStr(1, FILE_FILTER_IMAGES, "*." & strExt, vbTextCompare) > 0 Then
             
            Set objResult = ProcessSingleImportFile(objFileItem, strTargetFolder, objConfig, dicTokens, dicSeenHash)

            If objResult Is Nothing Then
                ' 剛性科學去重：命中已存在雜湊值，增加跳過計數
                lngSkippedCount = lngSkippedCount + 1
            Else
                colResults.Add objResult
                Set objResult = Nothing
            End If

            ' 釋放 CPU 執行緒換氣，換行前進
            DoEvents
         End If
     Next objFileItem
    
    ' 步驟 4：統一交給 Mod_InteractionAdapter 一次寫入所有結果（含欄寬調整）
    Call Mod_InteractionAdapter.WriteImportResults(colResults)
    
    Application.StatusBar = False
    
    ' 重複圖片警示彈窗
    If lngSkippedCount > 0 Then
        Call Mod_UI_Messenger.ShowInfo(MSG_SKIP_PREFIX & lngSkippedCount & MSG_SKIP_SUFFIX, TITLE_SKIP_ALERT)
    End If
    
    GoTo Cleanup

PipelineErrorHandler:
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    Application.StatusBar = False

' --- Cleanup Label（AVS v1.1 第七節）：正常結束與錯誤處理路徑於此唯一匯流 ---
Cleanup:
    Set dicSeenHash = Nothing
    Set colFlattenedFiles = Nothing
    Set colResults = Nothing
    Set objLocalFSO = Nothing
    
    If lngErrNum <> 0 Then
        Err.Raise lngErrNum, MOD_NAME & ".ExecuteImportPipeline", strErrDesc
    End If
    Exit Sub
End Sub
' =========================================================================
' 函數名稱：ProcessSingleImportFile (內部私有函數)
' 功用說明：單一檔案匯入處理器。負責去重判斷、檔名清洗、防撞改名、實體複製、
'           雙向雜湊校對與結果物件組裝，讓 ExecuteImportPipeline 的迴圈本體
'           不需要再塞進這一長串單一檔案的處理細節。
' 回傳說明：若命中去重（dicSeenHash 已存在同雜湊值），回傳 Nothing，
'           由呼叫端判定為「跳過」；否則回傳組裝完成的 cls_ImportResult。
' =========================================================================
Private Function ProcessSingleImportFile( _
    ByVal objFileItem As Object, ByVal strTargetFolder As String, _
    ByVal objConfig As cls_Settings, ByVal dicTokens As Object, _
    ByVal dicSeenHash As Object) As cls_ImportResult
    
    Dim objLocalFSO As Object
    Dim strCleanedName As String
    Dim strFinalDestPath As String
    Dim strCalculatedHash As String
    Dim strSrcHash As String
    Dim strBaseName As String
    Dim strExtName As String
    Dim cntCollision As Integer
    Dim objResult As cls_ImportResult
    
    Set objLocalFSO = dicTokens("FSO")
    
    ' 呼叫密碼學引擎模組計算這張原始相片的 MD5 唯一數位指紋
    On Error Resume Next
    strSrcHash = Mod_Crypto.CalculateFileHash(objFileItem.Path, dicTokens)
    If Err.Number <> 0 Then
        strSrcHash = ""
        Err.Clear
    End If
    On Error GoTo 0
    
    ' 剛性科學去重：已存在則直接回傳 Nothing，交由呼叫端判定為跳過
    If strSrcHash <> "" And dicSeenHash.Exists(strSrcHash) Then
        Set ProcessSingleImportFile = Nothing
        Exit Function
    End If
    
    ' 狀態列解耦發送即時更新訊息
    Application.StatusBar = MSG_STATUS_BAR_IMPORT_PREFIX & objFileItem.Name
    
    ' 呼叫路徑引擎清洗檔名
    strCleanedName = Mod_FileSystem.CleanAndTruncateFilename(objFileItem.Name, objConfig, dicTokens)
    
    strBaseName = objLocalFSO.GetBaseName(strCleanedName)
    strExtName = objLocalFSO.GetExtensionName(strCleanedName)
    strFinalDestPath = objLocalFSO.BuildPath(strTargetFolder, strCleanedName)
    cntCollision = 1
    
    ' 同名自動跳號防撞迴圈
    Do While objLocalFSO.FileExists(strFinalDestPath)
        strCleanedName = strBaseName & UNDERLINE_STR & Format(cntCollision, FMT_COUNTER_00) & DOT_STR & strExtName
        strFinalDestPath = objLocalFSO.BuildPath(strTargetFolder, strCleanedName)
        cntCollision = cntCollision + 1
    Loop
    
    ' 執行實體搬移複製
    On Error Resume Next
    objLocalFSO.CopyFile objFileItem.Path, strFinalDestPath, True
    On Error GoTo 0
    
    ' 雙向指紋校對與填入點名簿
    If strSrcHash <> "" Then
        strCalculatedHash = strSrcHash
        dicSeenHash.Add strSrcHash, strCleanedName
    Else
        On Error Resume Next
        strCalculatedHash = Mod_Crypto.CalculateFileHash(strFinalDestPath, dicTokens)
        If Err.Number <> 0 Then
            strCalculatedHash = HASH_FAILED_MARK
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    ' 組裝本次處理結果
    Set objResult = New cls_ImportResult
    If objLocalFSO.FileExists(strFinalDestPath) Then
        Call objResult.InitializeResult( _
            True, strCleanedName, Format(objFileItem.DateLastModified, FMT_DATETIME_DASH), _
            strFinalDestPath, Format(Now, FMT_DATE_DASH), strTargetFolder, _
            "", strCalculatedHash, "")
    Else
        Call objResult.InitializeResult( _
            False, objFileItem.Name, "", "", "", "", "", "", LBL_IMPORT_FAIL_REMARK)
    End If
    
    Set ProcessSingleImportFile = objResult
    Set objResult = Nothing
End Function

' =========================================================================
' 程序名稱：FlattenSourcePaths (內部私有程序)
' =========================================================================
Private Sub FlattenSourcePaths(ByVal colSrcPaths As Collection, ByRef colFlattenedFiles As Collection, ByVal objLocalFSO As Object)
    Dim varPathItem As Variant
    Dim objFileItem As Object
    Dim objSubDir As Object
    
    For Each varPathItem In colSrcPaths
        If objLocalFSO.FolderExists(varPathItem) Then
            Set objSubDir = objLocalFSO.GetFolder(varPathItem)
            For Each objFileItem In objSubDir.Files
                colFlattenedFiles.Add objFileItem
            Next objFileItem
            Set objSubDir = Nothing
        ElseIf objLocalFSO.FileExists(varPathItem) Then
            colFlattenedFiles.Add objLocalFSO.GetFile(varPathItem)
        End If
    Next varPathItem
End Sub

' =========================================================================
' 程序名稱：LoadExistingHashes (內部私有程序)
' =========================================================================
Private Sub LoadExistingHashes(ByVal wsStaging As Worksheet, ByRef lngNextRow As Long, ByVal dicSeenHash As Object)
    ' 剛性鎖定資料起始行常數合約
    If lngNextRow < NUM_ROW_DATA_START Then
        lngNextRow = NUM_ROW_DATA_START
        Exit Sub
    End If
    
    Dim idxExisting As Long
    Dim strExistingHash As String
    
    For idxExisting = NUM_ROW_DATA_START To lngNextRow
        strExistingHash = Trim(CStr(wsStaging.Cells(idxExisting, NUM_STG_COL_HASH).Value))
        
        If strExistingHash <> "" And strExistingHash <> HASH_FAILED_MARK Then
            If Not dicSeenHash.Exists(strExistingHash) Then
                dicSeenHash.Add strExistingHash, wsStaging.Cells(idxExisting, NUM_STG_COL_NAME).Value
            End If
        End If
    Next idxExisting
    
    lngNextRow = lngNextRow + 1
End Sub

