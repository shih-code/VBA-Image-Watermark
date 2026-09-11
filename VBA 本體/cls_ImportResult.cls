' ==========================================================
' MODULE: cls_ImportResult (類別模組)
' PURPOSE: 匯入結果資料模型。承載單一檔案匯入處理完畢後、準備寫入 Staging_Images
'          一列的完整資料，與 cls_ImageEntity（匯出用唯讀模型）對稱設計，讓
'          Mod_ImportPipeline 得以比照 Mod_ExportPipeline 的解耦模式，完全不
'          直接碰觸 Excel 網格（見 ADR-016）。
' EXPORTS: InitializeResult；唯讀屬性 IsSuccess, ImageName, LastModified,
'          CurrentPath, ImportDate, ImportRelativePath, UserRemark, FileHash,
'          FailReason
' IMPORTS: 無
' FORBIDDEN: 嚴禁在 InitializeResult 以外的地方修改內部私有變數；一經初始化
'            鎖定即不可二次寫入。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit

' --- 模組識別碼（第七章 SSOT） ---
Private Const MOD_NAME As String = "cls_ImportResult"

Private p_IsSuccess          As Boolean
Private p_ImageName          As String
Private p_LastModified       As String
Private p_CurrentPath        As String
Private p_ImportDate         As String
Private p_ImportRelativePath As String
Private p_UserRemark         As String
Private p_FileHash           As String
Private p_FailReason         As String
Private p_IsInitialized      As Boolean

' =========================================================================
' 程序名稱：InitializeResult
' 功用說明：一次性寫入單一檔案的匯入處理結果，並鎖定。
' isSuccess = True 時代表成功匯入，failReason 可留空；
' isSuccess = False 時代表匯入失敗，其餘資料欄位可能為空，僅 imageName／
' failReason 有意義（供錯誤列顯示原始檔名與失敗原因）。
' =========================================================================
Public Sub InitializeResult( _
    ByVal IsSuccess As Boolean, _
    ByVal ImageName As String, _
    ByVal lastMod As String, _
    ByVal curPath As String, _
    ByVal impDate As String, _
    ByVal impRelPath As String, _
    ByVal UserRemark As String, _
    ByVal hashVal As String, _
    ByVal FailReason As String)
    
    If p_IsInitialized Then
        Err.Raise 58, MOD_NAME, "安全防禦：此匯入結果物件已鎖定，不允許重複初始化！"
    End If
    
    p_IsSuccess = IsSuccess
    p_ImageName = Trim(ImageName)
    p_LastModified = Trim(lastMod)
    p_CurrentPath = Trim(curPath)
    p_ImportDate = Trim(impDate)
    p_ImportRelativePath = Trim(impRelPath)
    p_UserRemark = Trim(UserRemark)
    p_FileHash = Trim(hashVal)
    p_FailReason = Trim(FailReason)
    
    p_IsInitialized = True
End Sub

Public Property Get IsSuccess() As Boolean
    IsSuccess = p_IsSuccess
End Property

Public Property Get ImageName() As String
    ImageName = p_ImageName
End Property

Public Property Get LastModified() As String
    LastModified = p_LastModified
End Property

Public Property Get CurrentPath() As String
    CurrentPath = p_CurrentPath
End Property

Public Property Get ImportDate() As String
    ImportDate = p_ImportDate
End Property

Public Property Get ImportRelativePath() As String
    ImportRelativePath = p_ImportRelativePath
End Property

Public Property Get UserRemark() As String
    UserRemark = p_UserRemark
End Property

Public Property Get FileHash() As String
    FileHash = p_FileHash
End Property

Public Property Get FailReason() As String
    FailReason = p_FailReason
End Property


