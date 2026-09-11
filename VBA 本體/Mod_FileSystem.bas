' ==========================================================
' MODULE: Mod_FileSystem (標準模組)
' PURPOSE: 硬碟與路徑引擎。專職固定作業資料夾拓撲之建立與存取、
'          檔名清洗、字數截斷防禦，以及探測斷點續跑與成果/備份搬遷。
' EXPORTS: EnsureWorkspaceTopology, GetImportFolderPath, GetResultFolderPath,
'          GetBackupRootPath, CleanAndTruncateFilename, IsResumingPreviousRun,
'          BackupAndClearFolder, ClearFolderContents
' IMPORTS: Mod_StringConstants, cls_Settings
' FORBIDDEN: 嚴禁私自調用 MsgBox；嚴禁私自 CreateObject 取得 FSO 元件（必須由大腦
'            以 sysTokens 注入）。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 2.0.0 [Stability: Experimental]
' ==========================================================

Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_FileSystem"
Private Const MODULE_VERSION As String = "2.0.0"

' =========================================================================
' 函數名稱：EnsureWorkspaceTopology
' 功用說明：固定作業資料夾拓撲總入口。確保「作業資料夾／匯入圖片／成果／備份」
'           四層結構在硬碟上實體存在，不存在就建立，已存在就直接沿用。
'           回傳作業資料夾本身的絕對路徑。
' =========================================================================
Public Function EnsureWorkspaceTopology(ByVal sysTokens As Object) As String
    Dim localFSO As Object
    Dim strWorkspaceRoot As String
    
    Set localFSO = sysTokens("FSO")
    strWorkspaceRoot = localFSO.BuildPath(ThisWorkbook.Path, DIR_NAME_WORKSPACE_ROOT)
    
    If Not localFSO.FolderExists(strWorkspaceRoot) Then
        localFSO.CreateFolder strWorkspaceRoot
    End If
    
    ' 三個子層一次到位：匯入圖片／成果／備份，缺哪個補哪個
    Call EnsureSingleFolder(localFSO.BuildPath(strWorkspaceRoot, DIR_NAME_IMPORT_FOLDER), localFSO)
    Call EnsureSingleFolder(localFSO.BuildPath(strWorkspaceRoot, DIR_NAME_RESULT_FOLDER), localFSO)
    Call EnsureSingleFolder(localFSO.BuildPath(strWorkspaceRoot, DIR_NAME_BACKUP_ROOT), localFSO)
    
    EnsureWorkspaceTopology = strWorkspaceRoot
    Set localFSO = Nothing
End Function

' =========================================================================
' 函數名稱：GetImportFolderPath / GetResultFolderPath / GetBackupRootPath
' 功用說明：三個固定子路徑的便捷取用介面，內部都會確保拓撲存在後才回傳，
'           呼叫端不需要自己拼路徑字串，杜絕路徑拼接邏輯散落各處。
' =========================================================================
Public Function GetImportFolderPath(ByVal sysTokens As Object) As String
    Dim strRoot As String
    strRoot = EnsureWorkspaceTopology(sysTokens)
    GetImportFolderPath = sysTokens("FSO").BuildPath(strRoot, DIR_NAME_IMPORT_FOLDER)
End Function

Public Function GetResultFolderPath(ByVal sysTokens As Object) As String
    Dim strRoot As String
    strRoot = EnsureWorkspaceTopology(sysTokens)
    GetResultFolderPath = sysTokens("FSO").BuildPath(strRoot, DIR_NAME_RESULT_FOLDER)
End Function

Public Function GetBackupRootPath(ByVal sysTokens As Object) As String
    Dim strRoot As String
    strRoot = EnsureWorkspaceTopology(sysTokens)
    GetBackupRootPath = sysTokens("FSO").BuildPath(strRoot, DIR_NAME_BACKUP_ROOT)
End Function

' =========================================================================
' 函數名稱：IsResumingPreviousRun
' 功用說明：探測「成果」資料夾裡是不是已經有東西，供 Phase 2 判斷要不要
'           跳出「斷點續跑／覆寫」詢問視窗。邏輯與 v1 相同，只是現在餵給它
'           的路徑永遠是同一個固定「成果」資料夾，不再是每次新算出來的。
' =========================================================================
Public Function IsResumingPreviousRun(ByVal exportFolder As String, ByVal sysTokens As Object) As Boolean
    Dim localFSO As Object
    Dim objTargetFolder As Object
    
    Set localFSO = sysTokens("FSO")
    IsResumingPreviousRun = False
    
    If localFSO.FolderExists(exportFolder) Then
        Set objTargetFolder = localFSO.GetFolder(exportFolder)
        If objTargetFolder.Files.Count > 0 Then
            IsResumingPreviousRun = True
        End If
        Set objTargetFolder = Nothing
    End If
    
    Set localFSO = Nothing
End Function

' =========================================================================
' 函數名稱：BackupAndClearFolder
' 功用說明：【v2 核心】把來源資料夾「當下的全部內容」整批複製進備份區一個
'           帶時間戳記的新子資料夾，複製完驗證數量一致後，才清空來源資料夾
'           （只清內容物，資料夾本身留著）。適用於「成果」資料夾在
'           開新案／斷點覆寫兩種情境下的「確認結束」快照備份。
'           若來源資料夾根本沒有任何檔案，代表沒有東西好備份，直接跳過，
'           不建立空殼備份資料夾，回傳空字串。
' =========================================================================
Public Function BackupAndClearFolder(ByVal strSourceFolder As String, ByVal strBackupRoot As String, ByVal strBackupPrefix As String, ByVal objFSO As Object) As String
    Dim objSrcFolder As Object
    Dim objFileItem As Object
    Dim colFilesToBackup As Collection
    Dim varPathItem As Variant
    Dim strFinalBackupFolder As String
    Dim lngBackedUpCount As Long
    
    Dim lngErrNum As Long
    Dim strErrDesc As String
    
    If Not objFSO.FolderExists(strSourceFolder) Then
        BackupAndClearFolder = ""
        Exit Function
    End If
    
    Set objSrcFolder = objFSO.GetFolder(strSourceFolder)
    If objSrcFolder.Files.Count = 0 Then
        ' 沒有東西好備份，安靜結束，不留空殼資料夾
        BackupAndClearFolder = ""
        Set objSrcFolder = Nothing
        Exit Function
    End If
    
    On Error GoTo BackupErrorHandler
    
    strFinalBackupFolder = objFSO.BuildPath(strBackupRoot, strBackupPrefix & Format(Now, FMT_DATE_YMD & "_" & FMT_TIME_HMSM))
    objFSO.CreateFolder strFinalBackupFolder
    
    ' 先收集完整清單再動作，避免邊列舉 Files 集合邊搬移造成跳號漏漏
    Set colFilesToBackup = New Collection
    For Each objFileItem In objSrcFolder.Files
        colFilesToBackup.Add objFileItem.Name
    Next objFileItem
    
    ' 逐一複製，每一份當場驗證確實落地才算數（P0 複製優先於刪除的精神）
    For Each varPathItem In colFilesToBackup
        objFSO.CopyFile objFSO.BuildPath(strSourceFolder, CStr(varPathItem)), _
                        objFSO.BuildPath(strFinalBackupFolder, CStr(varPathItem)), True
        If Not objFSO.FileExists(objFSO.BuildPath(strFinalBackupFolder, CStr(varPathItem))) Then
            Err.Raise 5403, MOD_NAME & ".BackupAndClearFolder", ERR_ARCHIVE_COPY_FAILED
        End If
        lngBackedUpCount = lngBackedUpCount + 1
    Next varPathItem
    
    ' 驗證通過（複製數量與清單數量一致）：清空來源資料夾內容物，資料夾本身保留
    If lngBackedUpCount = colFilesToBackup.Count Then
        For Each varPathItem In colFilesToBackup
            objFSO.DeleteFile objFSO.BuildPath(strSourceFolder, CStr(varPathItem)), True
        Next varPathItem
    Else
        Err.Raise 5403, MOD_NAME & ".BackupAndClearFolder", ERR_ARCHIVE_COPY_FAILED
    End If
    
    BackupAndClearFolder = strFinalBackupFolder
    GoTo Cleanup

BackupErrorHandler:
    lngErrNum = Err.Number
    strErrDesc = Err.Description

Cleanup:
    Set objSrcFolder = Nothing
    Set objFileItem = Nothing
    Set colFilesToBackup = Nothing
    
    If lngErrNum <> 0 Then
        Err.Raise lngErrNum, MOD_NAME & ".BackupAndClearFolder", strErrDesc
    End If
End Function

' =========================================================================
' 程序名稱：ClearFolderContents
' 功用說明：單純清空一個資料夾內的所有檔案，不備份、不驗證數量、不留痕跡。
'           專供「匯入圖片」這種純工作複製品使用——開新案時直接丟棄即可，
'           它從來不是正本，不需要走 BackupAndClearFolder 那套驗證流程。
' =========================================================================
Public Sub ClearFolderContents(ByVal strFolder As String, ByVal objFSO As Object)
    Dim objFileItem As Object
    Dim colFilesToDelete As Collection
    Dim varPathItem As Variant
    
    If Not objFSO.FolderExists(strFolder) Then Exit Sub
    
    Set colFilesToDelete = New Collection
    For Each objFileItem In objFSO.GetFolder(strFolder).Files
        colFilesToDelete.Add objFileItem.Path
    Next objFileItem
    
    For Each varPathItem In colFilesToDelete
        On Error Resume Next
        objFSO.DeleteFile varPathItem, True
        On Error GoTo 0
    Next varPathItem
    
    Set colFilesToDelete = Nothing
End Sub

' =========================================================================
' 函數名稱：EnsureSingleFolder (內部私有函數)
' 功用說明：單一資料夾存在性檢查與建立，供 EnsureWorkspaceTopology 重複呼叫，
'           避免同一段「不存在就建立」的判斷邏輯在同一個函式裡寫三次。
' =========================================================================
Private Sub EnsureSingleFolder(ByVal strPath As String, ByVal objFSO As Object)
    If Not objFSO.FolderExists(strPath) Then
        objFSO.CreateFolder strPath
    End If
End Sub

' =========================================================================
' 函數名稱：CleanAndTruncateFilename
' 功用說明：清洗圖片的檔案名稱。洗掉 Windows 系統不允許的特殊符號，以免建立 Excel 超連結時斷線，並限制最大長度。
' =========================================================================
Public Function CleanAndTruncateFilename(ByVal rawFilename As String, ByVal cfg As cls_Settings, ByVal sysTokens As Object) As String
    Dim localFSO As Object
    Dim strBaseName As String
    Dim strExtName As String
    Dim varCharItem As Variant
    
    Set localFSO = sysTokens("FSO")
    
    strBaseName = localFSO.GetBaseName(rawFilename)
    strExtName = localFSO.GetExtensionName(rawFilename)
    
    For Each varCharItem In cfg.ForbiddenChars
        strBaseName = Replace(strBaseName, CStr(varCharItem), UNDERLINE_STR)
    Next varCharItem
    
    strBaseName = Replace(strBaseName, SPACE_STR, UNDERLINE_STR)
    
    If Len(strBaseName) > cfg.MaxFilenameLength Then
        strBaseName = Left(strBaseName, cfg.MaxFilenameLength)
    End If
    
    If Trim(strBaseName) = "" Then
        strBaseName = FALLBACK_BASE_NAME
    End If
    
    CleanAndTruncateFilename = strBaseName & DOT_STR & strExtName
    
    Set localFSO = Nothing
End Function
