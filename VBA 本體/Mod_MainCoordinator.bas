' ==========================================================
' MODULE: Mod_MainCoordinator (標準模組)
' PURPOSE: 按鈕點擊接收器。只負責接收前台按鈕的點擊指令，並將工作轉交給流程管理模組。
' EXPORTS: Macro_Phase1_ImportFolder, Macro_Phase1_ImportFiles, Macro_Phase2_Export,
'          Macro_ForceRebuildUI, Macro_ConsolidateAllLogs（以上五支皆帶 Optional Dummy
'          參數，刻意從 Alt+F8 巨集清單隱藏，只能透過按鈕觸發，見 ADR-010）；
'          重繪操作台（Macro_ForceRebuildUI 的中文別名，刻意不加 Dummy、保留於巨集
'          清單中，作為前台視覺元件損毀時的緊急救援手動入口，見 ADR-010）
' IMPORTS: Mod_WorkflowManager
' FORBIDDEN: 嚴禁在此模組塞入任何 If／迴圈／環境檢查邏輯，每個入口維持單行委派
'            （見 ADR-004），確保操作人員無法透過巨集清單繞過大腦的環境檢查。
'            本模組是全專案唯一 OnAction／巨集入口，嚴禁另立模組承接按鈕綁定。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================
Option Explicit

Private Const MOD_NAME As String = "Mod_MainCoordinator"
Private Const MODULE_VERSION As String = "1.0.0"

' =========================================================================
' 巨集名稱：Macro_Phase1_ImportFolder
' 功用說明：當使用者點擊 Excel 設定畫面上的「匯入資料夾」按鈕時，系統會自動觸發此程序。
' 邏輯說明：直接呼叫大腦的 RunPhase1 流程，並傳入參數 False，代表本次是要處理「整個資料夾」。
' =========================================================================
Public Sub Macro_Phase1_ImportFolder(Optional Dummy As Byte = 0)
    Call Mod_WorkflowManager.RunPhase1(False)
End Sub

' =========================================================================
' 巨集名稱：Macro_Phase1_ImportFiles
' 功用說明：當使用者點擊 Excel 設定畫面上的「挑選圖片」按鈕時，系統會自動觸發此程序。
' 邏輯說明：直接呼叫大腦的 RunPhase1 流程，並傳入參數 True，代表本次是要處理「使用者手動複選的個別檔案」。
' =========================================================================
Public Sub Macro_Phase1_ImportFiles(Optional Dummy As Byte = 0)
    Call Mod_WorkflowManager.RunPhase1(True)
End Sub

' =========================================================================
' 巨集名稱：Macro_Phase2_Export
' 功用說明：當使用者點擊 Excel 設定畫面上的「匯出浮水印」按鈕時，系統會自動觸發此程序。
' 邏輯說明：直接呼叫大腦的 RunPhase2 流程，啟動批次浮水印壓印與成品導出管線。
' =========================================================================
Public Sub Macro_Phase2_Export(Optional Dummy As Byte = 0)
    Call Mod_WorkflowManager.RunPhase2
End Sub

' =========================================================================
' 巨集名稱：Macro_ForceRebuildUI
' 功用說明：當使用者點擊 Excel 設定畫面上的「一鍵還原檯面」按鈕時，系統會自動觸發此程序。
' 邏輯說明：直接呼召大腦的 RunRebuildUI 流程，強制刪除畫面上被弄亂的圖形，並重新繪製標準按鈕與畫布。
' =========================================================================
Public Sub Macro_ForceRebuildUI(Optional Dummy As Byte = 0)
    Call Mod_WorkflowManager.RunRebuildUI
End Sub

' =========================================================================
' 巨集名稱：重繪操作台
' 功用說明：使用者可從 Alt+F8 進入使用。
' 邏輯說明：同 Macro_ForceRebuildUI，直接呼召大腦的 RunRebuildUI 流程，強制刪除畫面上被弄亂的圖形，並重新繪製標準按鈕與畫布。
' =========================================================================
Public Sub 重繪操作台()
    Call Mod_WorkflowManager.RunRebuildUI
End Sub

' =========================================================================
' 巨集名稱：Macro_ConsolidateAllLogs
' 功用說明：當使用者點擊 Excel 設定畫面上的「歷史日誌歸檔」按鈕時，系統會自動觸發此程序。
' 邏輯說明：直接呼叫大腦的 RunConsolidateLogs 流程，下令日誌管家執行歷史日誌回收與檔案減重。
' =========================================================================
Public Sub Macro_ConsolidateAllLogs(Optional Dummy As Byte = 0)
    Call Mod_WorkflowManager.RunConsolidateLogs
End Sub
