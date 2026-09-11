' ==========================================================
' MODULE: Mod_UI_Messenger (標準模組)
' PURPOSE: 全系統唯一具備彈出視窗 (MsgBox) 權限的模組，徹底與業務邏輯解耦。
' EXPORTS: AskQuestion, AskQuestionCancel, AskCritical, ShowInfo, ShowWarning, ShowError
' IMPORTS: 無
' FORBIDDEN: 本模組為純介面通訊器，嚴禁寫入任何核心運算、資料庫或工作表讀寫邏輯。
' DEPENDENCIES: 無
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_UI_Messenger"
Private Const MODULE_VERSION As String = "1.0.0"

' =========================================================================
' 函數名稱：AskQuestion
' 功用說明：彈出一個帶有「問號圖示」與「是(Yes) / 否(No)」按鈕的標準詢問對話框。
' 邏輯說明：
' 1. 接收大腦傳來的提示內文與標題。
' 2. 使用 vbQuestion（問號圖示）與 vbYesNo（顯示是與否按鈕）。
' 3. 特別加上 vbDefaultButton2：這會把預設的藍色高亮框鎖定在「否」按鈕上。
'    如果使用者太緊張，沒看清楚就不小心連按了鍵盤的 Enter 鍵，系統會自動選擇安全
'    的「否」，以防止使用者誤觸高風險的覆寫操作。
' =========================================================================
Public Function AskQuestion(ByVal promptText As String, ByVal titleText As String) As VbMsgBoxResult
    AskQuestion = MsgBox(promptText, vbQuestion + vbYesNo + vbDefaultButton2, titleText)
End Function

' =========================================================================
' 函數名稱：AskQuestionCancel
' 功用說明：彈出一個帶有「問號圖示」且擁有「是(Yes) / 否(No) / 取消(Cancel)」三個按鈕的對話框。
' 邏輯說明：主要用於匯入階段。當大腦發現有舊資料時，提供三個按鈕讓使用者選擇。
'           此處使用 vbDefaultButton1，將預設焦點放在第一個按鈕「是」上面。
' =========================================================================
Public Function AskQuestionCancel(ByVal promptText As String, ByVal titleText As String) As VbMsgBoxResult
    AskQuestionCancel = MsgBox(promptText, vbQuestion + vbYesNoCancel + vbDefaultButton1, titleText)
End Function

' =========================================================================
' 函數名稱：AskCritical
' 功用說明：彈出一個帶有「紅色叉叉（嚴重警告）圖示」與「是(Yes) / 否(No)」按鈕的對話框。
' 邏輯說明：用於極度危險或即將發生系統變更的操作。同樣強制將預設焦點鎖定在
'           vbDefaultButton2（否）上面，落實防禦性安全編程。
' =========================================================================
Public Function AskCritical(ByVal promptText As String, ByVal titleText As String) As VbMsgBoxResult
    AskCritical = MsgBox(promptText, vbCritical + vbYesNo + vbDefaultButton2, titleText)
End Function

' =========================================================================
' 程序名稱：ShowInfo
' 功用說明：彈出一個帶有「藍色驚嘆號圖示」的純通知視窗（只有一個確定按鈕）。
' 邏輯說明：用於告訴使用者工作已經順利完成。
' =========================================================================
Public Sub ShowInfo(ByVal promptText As String, ByVal titleText As String)
    MsgBox promptText, vbInformation, titleText
End Sub

' =========================================================================
' 程序名稱：ShowWarning
' 功用說明：彈出一個帶有「黃色三角形驚嘆號圖示」的警告視窗（只有一個確定按鈕）。
' 邏輯說明：用於提醒使用者系統發生了不影響執行的缺陷（例如設定檔打錯字，系統啟動自動自癒修復）。
' =========================================================================
Public Sub ShowWarning(ByVal promptText As String, ByVal titleText As String)
    MsgBox promptText, vbExclamation, titleText
End Sub

' =========================================================================
' 程序名稱：ShowError
' 功用說明：彈出一個帶有「紅色叉叉圖示」的致命錯誤死機視窗（只有一個確定按鈕）。
' 邏輯說明：當系統觸發 Fail-Fast 安全熔斷機制（例如缺乏必要組件或找不到檔案）而無法繼續執行時，
'           由大腦調用此程序向使用者回報最終的崩潰報告。
' =========================================================================
Public Sub ShowError(ByVal promptText As String, ByVal titleText As String)
    MsgBox promptText, vbCritical, titleText
End Sub
