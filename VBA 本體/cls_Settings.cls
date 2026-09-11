' ==========================================================
' MODULE: cls_Settings (類別模組)
' PURPOSE: 參數配置中心。負責儲存浮水印的字體、顏色、大小等設定值，
'          並將 Excel 工作表上的中文設定翻譯成系統適用的數值。記錄使用者的設定。
' EXPORTS: LoadInterfaceSettings；唯讀屬性 FontNameChinese, FontNameEnglish,
'          FontColorDefault, FontSizeDefault, BgColorDefault, BgTransparency,
'          LineVisible, MaxFilenameLength, ForbiddenChars
' IMPORTS: Mod_StringConstants
' FORBIDDEN: 內部設定值一經 LoadInterfaceSettings 載入即透過 Property Get 唯讀對外
'            開放，嚴禁提供 Property Let 讓外部任意寫入修改。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit

Private Const MOD_NAME As String = "cls_Settings"
Private Const MODULE_VERSION As String = "1.0.0"

' --- 內部私有儲存區（不開放給外部直接修改，確保資料安全） ---
Private p_FontNameChinese As String    ' 儲存中文字體名稱（例如：微軟正黑體）
Private p_FontNameEnglish As String    ' 儲存英數字體名稱（例如：Arial）
Private p_FontColorDefault As Long     ' 儲存字體顏色碼（數字格式，如 &H0）
Private p_FontSizeDefault As Single    ' 儲存字體大小（例如：12）
Private p_BgColorDefault As Long       ' 儲存背景底板顏色碼（數字格式）
Private p_BgTransparency As Single     ' 儲存背景半透明度（數值介於 0 到 1 之間）
Private p_LineVisible As Boolean       ' 儲存是否顯示邊框（True 代表顯示，False 代表隱藏）
Private p_MaxFilenameLength As Integer ' 儲存限制的檔名最大長度（例如：20 個字）
Private p_ForbiddenChars As Variant    ' 儲存系統禁用的特殊字元清單（陣列格式）

' =========================================================================
' 內建程序：Class_Initialize (物件建立時自動執行)
' 功用說明：當主程式一建立這個設定物件時，這裡會自動幫所有設定填上「安全預設值」。
'           如果使用者沒有在 Excel 畫面上填寫設定，系統就會採用這裡的數值。
' 邏輯說明：直接將固定的預設數值指派給內部的私有變數。
' =========================================================================
Private Sub Class_Initialize()
    p_FontNameChinese = "微軟正黑體"
    p_FontNameEnglish = "Arial"
    p_FontColorDefault = &H0       ' &H0 在電腦系統中代表黑色
    p_FontSizeDefault = 12         ' 預設字體大小為 12 級字
    p_BgColorDefault = &HFFFFFF    ' &HFFFFFF 在電腦系統中代表白色
    p_BgTransparency = 0.5         ' 0.5 代表 50% 的透明度
    p_LineVisible = False          ' 預設不顯示文字方塊的外框線
    p_MaxFilenameLength = 20       ' 預設檔名最多 20 個字，超過會被截斷
    
    ' 定義 Windows 作業系統不允許出現在檔案名稱中的特殊符號，用來清洗檔名
    p_ForbiddenChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|", "[", "]", "#")
End Sub

' =========================================================================
' 程序名稱：LoadInterfaceSettings
' 功用說明：從指定的 Excel 設定工作表（ConfigSheet）中讀取使用者自訂的數值，並更新內部設定。
' 邏輯說明：
' 1. 檢查傳進來的工作表存不存在，不存在就直接結束，維持原本的預設值。
' 2. 從指定儲存格讀取中、英文字體名稱，如果不是空白就更新內部設定。
' 3. 讀取中文顏色名稱，透過 Select Case 分類判定，將中文轉換為系統專用的十六進位數值。
' 4. 讀取字體大小，確認是數字且大於零才進行更新。
' =========================================================================
Public Sub LoadInterfaceSettings(ByVal wsConfig As Worksheet)
    ' 檢查防禦：如果主程式沒有傳入正確的工作表，就直接退出，不更新任何設定
    If wsConfig Is Nothing Then Exit Sub
    
    ' 宣告各別獨立、專屬的暫存變數，用來裝剛從 Excel 格子裡讀出來的原始文字
    Dim strChineseFont As String
    Dim strEnglishFont As String
    Dim strColorName As String
    
    ' 如果讀取過程中發生非預期的個別儲存格錯誤（例如格子裡有損毀的公式），先跳過該儲存格，確保其他設定能繼續加載
    On Error Resume Next
    
    ' 1. 讀取中文字體（位置：A1儲存格，由常數 ADDR_CFG_FONT_CHINESE 控制）
    ' CStr 是將內容強制轉成文字，Trim 是把文字前後不小心打出來的空白空白刪除
    strChineseFont = Trim(CStr(wsConfig.Range(ADDR_CFG_FONT_CHINESE).Value))
    If Err.Number <> 0 Then Debug.Print "[DEBUG] " & MOD_NAME & " 讀取中文字體儲存格失敗(" & Err.Number & ")，沿用預設值": Err.Clear
    If strChineseFont <> "" Then
        p_FontNameChinese = strChineseFont
    End If
    
    ' 2. 讀取英數字體（位置：B1儲存格，由常數 ADDR_CFG_FONT_ENGLISH 控制）
    strEnglishFont = Trim(CStr(wsConfig.Range(ADDR_CFG_FONT_ENGLISH).Value))
    If Err.Number <> 0 Then Debug.Print "[DEBUG] " & MOD_NAME & " 讀取英數字體儲存格失敗(" & Err.Number & ")，沿用預設值": Err.Clear
    If strEnglishFont <> "" Then
        p_FontNameEnglish = strEnglishFont
    End If
    
    ' 3. 讀取字體顏色名稱並翻譯為系統色碼（位置：C1儲存格，由常數 ADDR_CFG_FONT_COLOR 控制）
    strColorName = Trim(CStr(wsConfig.Range(ADDR_CFG_FONT_COLOR).Value))
    If Err.Number <> 0 Then Debug.Print "[DEBUG] " & MOD_NAME & " 讀取字體顏色儲存格失敗(" & Err.Number & ")，沿用預設值": Err.Clear
    
    ' 根據讀到的中文名稱，改寫電腦專用的長整型（Long）顏色碼
    Select Case strColorName
        Case "黑色"
            p_FontColorDefault = &H0
        Case "白色"
            p_FontColorDefault = &HFFFFFF
        Case "紅色"
            p_FontColorDefault = &HFF        ' VBA 採用的數值格式為藍、綠、紅（BGR）
        Case "黃色"
            p_FontColorDefault = &HFFFF
        Case "藍色"
            p_FontColorDefault = &HFF0000
        Case "綠色"
            p_FontColorDefault = &HFF00
        Case Else
            p_FontColorDefault = &H0        ' 安全底線：如果使用者空著沒填或打錯字，一律強制轉為黑色
    End Select

    ' 4. 讀取字體大小（位置：D1儲存格，由常數 ADDR_CFG_FONT_SIZE 控制）
    ' IsNumeric 是用來檢查該儲存格輸入的是不是數字，防止使用者在格子裡輸入英文字母導致程式當機
    If IsNumeric(wsConfig.Range(ADDR_CFG_FONT_SIZE).Value) Then
        ' 確保字體大小必須大於 0 才可以採用
        If wsConfig.Range(ADDR_CFG_FONT_SIZE).Value > 0 Then
            p_FontSizeDefault = CSng(wsConfig.Range(ADDR_CFG_FONT_SIZE).Value)
        End If
    ElseIf Err.Number <> 0 Then
         Debug.Print "[DEBUG] " & MOD_NAME & " 讀取字體大小儲存格失敗(" & Err.Number & ")，沿用預設值": Err.Clear
     End If
    
    ' 結束恢復正常的錯誤檢查機制
    On Error GoTo 0
End Sub

' =========================================================================
' 外部唯讀屬性接口 (Property Get)
' 功用說明：這些程序是專門開放給外部模組（例如浮水印影像引擎）使用的。
'           因為這裡只有寫 Property Get（代表只能讀取），沒有寫 Property Let（代表寫入），
'           所以外部程式只能呼叫「Settings.FontNameChinese」來得知設定值，但絕對無法修改它。
' =========================================================================

Public Property Get FontNameChinese() As String
    FontNameChinese = p_FontNameChinese
End Property

Public Property Get FontNameEnglish() As String
    FontNameEnglish = p_FontNameEnglish
End Property

Public Property Get FontColorDefault() As Long
    FontColorDefault = p_FontColorDefault
End Property

Public Property Get FontSizeDefault() As Single
    FontSizeDefault = p_FontSizeDefault
End Property

Public Property Get BgColorDefault() As Long
    BgColorDefault = p_BgColorDefault
End Property

Public Property Get BgTransparency() As Single
    BgTransparency = p_BgTransparency
End Property

Public Property Get LineVisible() As Boolean
    LineVisible = p_LineVisible
End Property

Public Property Get MaxFilenameLength() As Integer
    MaxFilenameLength = p_MaxFilenameLength
End Property

' 以 Variant 陣列形式回傳禁用字元清單，方便外部進行 For Each 迴圈比對清洗
Public Property Get ForbiddenChars() As Variant
    ForbiddenChars = p_ForbiddenChars
End Property
