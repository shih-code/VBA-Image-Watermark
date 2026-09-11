' ==========================================================
' MODULE: cls_ImageEntity (類別模組)
' PURPOSE: 資料物件模型，專門存放與封裝單張圖片的所有屬性資訊。
' EXPORTS: InitializeEntity；唯讀屬性 ImageName, LastModified, CurrentPath,
'          ImportDate, ImportRelativePath, FileHash, UserRemark, StagingRowNumber
' IMPORTS: 無
' FORBIDDEN: 嚴禁在 InitializeEntity 以外的地方修改內部私有變數；資料一經初始化
'            鎖定（p_IsInitialized = True）即不可二次寫入。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit

Private Const MOD_NAME As String = "cls_ImageEntity"
Private Const MODULE_VERSION As String = "1.0.0"

' --- 內部私有變數（僅供本模組內部儲存使用，外部無法直接存取） ---
Private p_ImageName As String          ' 儲存圖片檔案名稱
Private p_LastModified As String       ' 儲存圖片實體檔案的最後修改時間
Private p_CurrentPath As String        ' 儲存圖片目前的絕對路徑
Private p_ImportDate As String         ' 儲存圖片匯入到系統的日期
Private p_ImportRelativePath As String ' 儲存計畫匯出的目標資料夾路徑
Private p_FileHash As String           ' 儲存圖片的 MD5 雜湊值
Private p_UserRemark As String         ' 【補票 1】儲存使用者手打的備註文字
Private p_StagingRowNumber As Long     ' 儲存這筆資料在 Staging_Images 工作表上的實際列號，供匯出完成後精準回寫「已匯出」狀態，不靠檔名或雜湊猜位置
Private p_IsInitialized As Boolean     ' 鎖定旗標，True 代表這張圖片的資料已經填寫過並鎖定

' =========================================================================
' 程序名稱：InitializeEntity
' 功用說明：將一張圖片的所有屬性資料一次性填入這個容器物件中，並執行資料清理與鎖定。
' 邏輯說明：
' 1. 檢查 p_IsInitialized 旗標，如果已經是 True，代表已經設定過，立刻報錯阻斷。
' 2. 使用 Replace 函數將圖片名稱與路徑中的所有空白字元（包含中間的空白）徹底刪除。
' 3. 檢查關鍵欄位（名稱與路徑）是否為空字串，如果其中一個是空的，代表資料不合法，報錯阻斷。
' 4. 通過檢查後，將資料寫入對應的私有變數中。
' 5. 將 p_IsInitialized 改為 True，從此鎖死這個物件。
' =========================================================================
Public Sub InitializeEntity( _
    ByVal imgName As String, _
    ByVal lastMod As String, _
    ByVal curPath As String, _
    ByVal impDate As String, _
    ByVal impRelPath As String, _
    ByVal UserRemark As String, _
    ByVal hashVal As String, _
    ByVal stagingRow As Long)
    
    
    ' 1. 安全檢查：如果這張圖片已經填過資料，就拒絕再次寫入，防止資料中途被覆蓋
    If p_IsInitialized Then
        Err.Raise 58, MOD_NAME, "安全防禦：此圖片物件已鎖定，不允許重複初始化！"
    End If
    
    ' 2. 資料清理：利用 Replace 刪除所有半形空白，再利用 Trim 移除去頭去尾的字元
    '    注意：Replace 會將文字中間的空白也一併刪除
    'imgName = Trim(Replace(imgName, " ", ""))
    'curPath = Trim(Replace(curPath, " ", ""))
    imgName = Trim(imgName)
    curPath = Trim(curPath)
    hashVal = Trim(hashVal)
    
    ' 3. 關鍵資料檢查：圖片名稱與路徑是系統運作的必要條件，如果是空字串就代表資料錯誤，拒絕建立
    If imgName = "" Or curPath = "" Then
        Err.Raise 5801, MOD_NAME, "資料斷層：圖片核心欄位(名稱或路徑)不允許為空值，拒絕建立實體 facts！"
    End If
    
    ' 列號防禦：Staging 表資料列從 NUM_ROW_DATA_START（=2）開始，小於這個值代表呼叫端傳錯，
    ' 若讓它悄悄過關，後面回寫「已匯出」狀態時就會寫到表頭或不存在的列，直接攔下比事後除錯划算
    If stagingRow < NUM_ROW_DATA_START Then
        Err.Raise 5802, MOD_NAME, "資料斷層：傳入的 Staging 列號不合法（小於資料起始列），拒絕建立實體 facts！"
    End If
    
    ' 4. 將清理乾淨的資料寫入內部的私有儲存區
    p_ImageName = imgName
    p_LastModified = Trim(lastMod)
    p_CurrentPath = curPath
    p_ImportDate = Trim(impDate)
    p_ImportRelativePath = Trim(impRelPath)
    p_UserRemark = Trim(UserRemark)    ' 【補票 2】將傳入的備註 Facts 正式落盤
    p_FileHash = hashVal
    p_StagingRowNumber = stagingRow
    
    ' 5. 變更鎖定旗標，正式鎖死這張照片的資料
    p_IsInitialized = True
End Sub

' =========================================================================
' 外部唯讀屬性接口 (Property Get)
' 功用說明：提供公開的管道，讓外部的其他程式（例如浮水印管線）可以讀取這些資料。
'           因為這裡只有寫 Property Get（讀取），沒有寫 Property Let（寫入），
'           所以外部程式只能看資料，無法修改資料。
' =========================================================================

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

Public Property Get FileHash() As String
    FileHash = p_FileHash
End Property

Public Property Get StagingRowNumber() As Long
    StagingRowNumber = p_StagingRowNumber
End Property
' =========================================================================
' 外部唯讀屬性接口 (Property Get)
' 功用說明：僅開放外部管線與轉接器「自由讀取」備註資訊。
'           徹底刪除 Property Let，從語言結構上封死任何外部程式竄改資料的可能。
' =========================================================================
Public Property Get UserRemark() As String
    UserRemark = p_UserRemark
End Property


