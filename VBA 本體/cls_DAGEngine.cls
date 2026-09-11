' ==========================================================
' MODULE: cls_DAGEngine (類別模組)
' PURPOSE: 系統級生命週期守衛。控管單一管線的啟動、逐步紀錄、正常終止與異常熔斷，
'          不涉及任何節點依賴解析或動態命令排程（語意定案見 ADR-001）。
' EXPORTS: 狀態方法 (StartPipeline, LogStep, TerminatePipeline, DispatchMeltdown)，
'          唯讀屬性 (IsRunning)
' IMPORTS: cls_EventStore
' FORBIDDEN: 本引擎只負責流程順序的驗證與狀態攔截，嚴禁塞入具體影像處理或實體 IO 動作。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit

Private Const MOD_NAME As String = "cls_DAGEngine"
Private Const MODULE_VERSION As String = "1.0.0"


' --- 內部私有狀態 ---
Private p_Store As cls_EventStore   ' 宣告內部綁定的黑盒子紀錄器物件
Private p_IsRunning As Boolean      ' 紀錄管線目前是否正在執行中的開關旗標（True/False）

' =========================================================================
' 內建程序：Class_Initialize (物件建立時自動執行)
' 功用說明：當大腦建立此引擎時，自動在內部配備一個全新、專屬的文字紀錄器。
' =========================================================================
Private Sub Class_Initialize()
    Set p_Store = New cls_EventStore
    p_IsRunning = False
End Sub

' =========================================================================
' 內建程序：Class_Terminate (物件銷毀時自動執行)
' 功用說明：安全解構防禦。當引擎被釋放時，確保命令內部的紀錄器也把記憶體清空。
' =========================================================================
Private Sub Class_Terminate()
    If Not p_Store Is Nothing Then
        On Error Resume Next
        p_Store.ClearStore
        On Error GoTo 0
        Set p_Store = Nothing
    End If
End Sub

' =========================================================================
' 程序名稱：StartPipeline
' 功用說明：正式開啟一條工作流水線（任務開始）。
' 邏輯說明：
' 1. 清理傳進來的任務名稱，移去所有空白。如果沒填名字就給予備用名稱。
' 2. 檢查 p_IsRunning 旗標。如果已經是 True，代表上個任務還沒做完，立刻報錯阻斷（Fail-Fast）。
' 3. 通過檢查後，將狀態改為 True，並清空紀錄器，寫下任務啟動的第一筆紀錄。
' =========================================================================
Public Sub StartPipeline(ByVal pipelineName As String)
    ' 移除字串中所有的空白（注意：包含文字中間的空白也會被移除）
    pipelineName = Trim(Replace(pipelineName, " ", ""))
    If pipelineName = "" Then pipelineName = "未命名安全管線"
    
    ' 安全防禦：如果發現目前已有管線在跑，不允許重複啟動，直接拋出錯誤停止程式
    If p_IsRunning Then
        Err.Raise 513, MOD_NAME, "管線錯誤：當前已有管線正在執行中，拒絕重複啟動！"
    End If
    
    ' 正式變更狀態為正在執行，並寫入初始化紀錄
    p_IsRunning = True
    p_Store.ClearStore
    p_Store.AddEvent "管線 [" & pipelineName & "] 正式啟動。"
End Sub

' =========================================================================
' 程序名稱：LogStep
' 功用說明：在任務執行過程中，隨時回報目前做到哪一個小步驟。
' 邏輯說明：
' 1. 檢查 p_IsRunning 是否為 True。如果任務根本沒啟動就想寫紀錄，代表程式順序寫錯，立刻報錯。
' 2. 清理傳入的步驟說明文字，若不是空文字，就寫入文字紀錄器中。
' =========================================================================
Public Sub LogStep(ByVal stepDescription As String)
    ' 狀態檢查：如果管線根本沒開，不允許執行紀錄動作
    If Not p_IsRunning Then
        Err.Raise 514, MOD_NAME, "防禦熔斷：管線未啟動，無法記錄步驟！"
    End If
    
    ' 清理文字空白
    stepDescription = Trim(Replace(stepDescription, " ", ""))
    If stepDescription = "" Then Exit Sub ' 如果是空字串，直接跳過不理會
    
    ' 將本步驟步驟寫入文字紀錄器中保存
    p_Store.AddEvent stepDescription
End Sub

' =========================================================================
' 程序名稱：TerminatePipeline
' 功用說明：正常且安全地結束目前的工作流水線（任務順利完成）。
' 邏輯說明：在紀錄器寫下結束通知，並將 p_IsRunning 狀態開關調回 False，釋放執行權限。
' =========================================================================
Public Sub TerminatePipeline()
    If p_IsRunning Then
        p_Store.AddEvent "管線安全結束，釋放調度權限。"
        p_IsRunning = False
    End If
End Sub

' =========================================================================
' 程序名稱：DispatchMeltdown
' 功用說明：異常阻斷熔斷器。當外部（如管線或大腦）抓到不可預期的錯誤時，呼叫這裡進行系統熔斷。
' 邏輯說明：
' 1. 將執行狀態開關強制作為 False（停止運轉）。
' 2. 從文字紀錄器中，撈出最後一個成功寫入的步驟文字。
' 3. 檢查外部傳進來的錯誤原因文字，若空白則補上預設說明。
' 4. 組裝出一份極為詳細的報錯面板文字（包含最後成功斷點、異常因由）。
' 5. 清空文字紀錄器，確保記憶體乾淨（不留半成品狀態）。
' 6. 正式拋出 515 錯誤代碼，強行阻斷 Excel 的所有後續動作。
' =========================================================================
Public Sub DispatchMeltdown(ByVal errorMessage As String)
    ' 立即解除執行權限
    p_IsRunning = False
    
    Dim lastSuccess As String
    On Error Resume Next
    ' 向紀錄器索取最後一筆成功執行的紀錄
    lastSuccess = p_Store.GetLatestEvent
    On Error GoTo 0
    
    ' 自癒修正：如果紀錄器裡剛好沒資料，給予預設說明
    If Trim(lastSuccess) = "" Then lastSuccess = "未知初始化斷點"
    
    ' 清理外部傳入的錯誤說明文字
    errorMessage = Trim(errorMessage)
    If errorMessage = "" Then errorMessage = "外部發生未知的非預期執行階段崩潰"
    
    ' 組合出最終的 Fail-Fast 熔斷警告報告文字
    Dim finalReport As String
    finalReport = "【核心安全熔斷器啟動】" & vbCrLf & _
                  "系統已強行中斷，未留半成品。" & vbCrLf & _
                  "最後成功斷點：" & lastSuccess & vbCrLf & _
                  "異常原因：" & errorMessage
                  
    ' 原子性保護：清空文字紀錄器，不留下任何殘留歷史進度，防止下次執行誤判
    p_Store.ClearStore
                  
    ' 正式拋出系統熔斷錯誤，強制中斷 VBA 執行緒
    Err.Raise 515, MOD_NAME, finalReport
End Sub

' =========================================================================
' 外部唯讀屬性接口 (Property Get)
' 屬性名稱：IsRunning
' 功用說明：讓外部的大腦或按鈕可以隨時偵測，目前系統是不是正忙碌於某個任務中。
' =========================================================================
Public Property Get IsRunning() As Boolean
    IsRunning = p_IsRunning
End Property

