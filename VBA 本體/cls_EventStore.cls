' ==========================================================
' MODULE: cls_EventStore (類別模組)
' PURPOSE: 執行日誌儲存器。負責在記憶體中動態記錄系統執行的每一個步驟，
'          在系統出錯時，可以提供最後一個成功的執行斷點紀錄。
' EXPORTS: AddEvent, ClearStore；唯讀屬性 GetLatestEvent, EventCount
' IMPORTS: 無
' FORBIDDEN: 嚴禁外部繞過 AddEvent 直接操作內部 Collection；本模組僅負責記憶體暫存
'            （服務斷點續跑／運行期步驟追蹤，語意見 AVS v1.1 第四節），嚴禁寫入硬碟。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit

Private Const MOD_NAME As String = "cls_EventStore"
Private Const MODULE_VERSION As String = "1.0.0"


' --- 內部私有狀態（外部無法直接存取，只能透過公開方法操作） ---
Private p_Events As Collection ' 用來存放每一條步驟文字紀錄的集合容器

' =========================================================================
' 內建程序：Class_Initialize (物件建立時自動執行)
' 功用說明：當這個紀錄器被建立時，自動初始化內部的集合容器（New Collection）。
' =========================================================================
Private Sub Class_Initialize()
    Set p_Events = New Collection
End Sub

' =========================================================================
' 內建程序：Class_Terminate (物件銷毀時自動執行)
' 功用說明：當程式結束、這個紀錄器準備離開記憶體時，強制清空內部集合，釋放電腦記憶體。
' =========================================================================
Private Sub Class_Terminate()
    Set p_Events = Nothing
End Sub

' =========================================================================
' 程序名稱：AddEvent
' 功用說明：在目前的紀錄清單最後面，追加寫入一筆新的步驟描述。
' 邏輯說明：
' 1. 使用 Trim 檢查傳進來的文字是不是空的（例如只有空白鍵）。
' 2. 如果是空文字，就直接退出不做紀錄。
' 3. 如果是有意義的文字，就放入 p_Events 集合中保存。
' =========================================================================
Public Sub AddEvent(ByVal eventDescription As String)
    ' 安全檢查：防止寫入無意義的空白紀錄
    If Trim(eventDescription) = "" Then
        Exit Sub
    End If
    
    ' 將文字壓入集合容器的最末尾
    p_Events.Add eventDescription
End Sub

' =========================================================================
' 程序名稱：ClearStore
' 功用說明：把白板擦乾淨，清除之前留下的所有歷史執行紀錄。
' 邏輯說明：直接將 p_Events 重新設定為一組全新且全空的 Collection 物件。
' =========================================================================
Public Sub ClearStore()
    Set p_Events = New Collection
End Sub

' =========================================================================
' 外部唯讀屬性接口 (Property Get)
' =========================================================================

' =========================================================================
' 屬性名稱：GetLatestEvent
' 功用說明：抓出目前寫在白板上的最後一行紀錄。當程式當機時，可以用來得知死在哪一步。
' 邏輯說明：檢查集合內有沒有資料，有的話就傳回最後一筆（索引值等於總計數）；沒有的話就傳回初始化提示。
' =========================================================================
Public Property Get GetLatestEvent() As String
    If p_Events.Count > 0 Then
        ' 傳回集合中的最後一筆文字紀錄
        GetLatestEvent = p_Events(p_Events.Count)
    Else
        GetLatestEvent = "系統處於初始狀態，尚無執行紀錄。"
    End If
End Property

' =========================================================================
' 屬性名稱：EventCount
' 功用說明：得知目前總共記錄了多少個執行步驟。
' 邏輯說明：直接傳回集合容器的 Count 數值。
' 備註：返回值已由 Integer 優化為 Long，防止大批次處理時因數字破萬而溢位損毀。
' =========================================================================
Public Property Get EventCount() As Long
    EventCount = p_Events.Count
End Property
