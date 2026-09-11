' ==========================================================
' MODULE: View_CanvasManager (標準模組)
' PURPOSE: 負責讀取 Excel 畫布與控制項方塊的相對位置，換算成浮水印的幾何百分比座標。
' EXPORTS: GetWatermarkPercentages
' IMPORTS: Mod_StringConstants
' FORBIDDEN: 本模組僅准負責幾何位置計算，嚴禁直接變更檔案系統或進行任何影像壓印 IO 動作。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit
Option Private Module

' --- 模組識別碼（第七章 SSOT：僅此一處手打類別名稱字面值，其餘全部引用此常數） ---
Private Const MOD_NAME           As String = "View_CanvasManager"
Private Const MODULE_VERSION As String = "1.0.0"

' --- 內部私有狀態 ---
Private Const ERR_CONFIG_MISSING As String = "介面錯誤：找不到配置工作表"


' ==========================================================
' 程序名稱：GetWatermarkPercentages
' 功用說明：本模組的公開入口。負責去設定工作表上找出「畫布圖形」與「錨點方塊」，
'            並呼叫下方的計算工具換算出相對的百分比位置。
' 參數說明：
'   - outPctX: (ByRef 傳址參數) 用來將計算好的水平百分比數值（0.0 到 1.0）傳回給呼叫大腦。
'   - outPctY: (ByRef 傳址參數) 用來將計算好的垂直百分比數值（0.0 到 1.0）傳回給呼叫大腦。
' ==========================================================
Public Sub GetWatermarkPercentages(ByRef outPctX As Single, ByRef outPctY As Single)
    Dim ws          As Worksheet
    Dim canvasShape As Shape
    Dim anchorShape As Shape
    
    ' 1. 嘗試抓取指定的設定工作表
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    On Error GoTo 0
    
    ' 如果工作表不存在，拋出錯誤中斷
    If ws Is Nothing Then
        Err.Raise 5301, MOD_NAME, ERR_CONFIG_MISSING
    End If
    
    ' 2. 從工作表中讀取實體控制項圖形物件
    On Error Resume Next
    Set canvasShape = ws.Shapes(SH_CANVAS_NAME)
    Set anchorShape = ws.Shapes(SH_ANCHOR_NAME)
    On Error GoTo 0
    
    ' 3. Fail-Fast 安全熔斷：如果任何一個控制項不見了，立刻停止，防止後續影像引擎算錯座標
    If canvasShape Is Nothing Or anchorShape Is Nothing Then
        Err.Raise 5302, MOD_NAME, ERR_CRITICAL_ANCHOR_LOST
    End If
    
    ' 4. 圖形物件確認齊全，丟入下方專職計算的純函數中進行數學換算
    CalculateRelativePosition canvasShape, anchorShape, outPctX, outPctY
End Sub

' =========================================================================
' 程序名稱：CalculateRelativePosition (內部私有程序)
' 功用說明：專職處理幾何座標換算的數學計算，並包含四向邊界自癒磁吸防禦。
' 參數說明：
'   - canvas: 傳入的畫布圖形物件。
'   - anchor: 傳入的文字錨點圖形物件（使用者可以在畫布上拖挪位置的那個小方塊）。
'   - pctX: (ByRef 傳出) 計算完畢後，文字錨點佔畫布水平寬度的相對百分比。
'   - pctY: (ByRef 傳出) 計算完畢後，文字錨點佔畫布垂直高度的相對百分比。
' =========================================================================
Private Sub CalculateRelativePosition( _
    ByVal canvas As Shape, _
    ByVal anchor As Shape, _
    ByRef pctX As Single, _
    ByRef pctY As Single)
    
    Dim relativeLeft   As Single
    Dim relativeTop    As Single
    Dim maxAllowedLeft As Single
    Dim maxAllowedTop  As Single
    
    ' 1. 計算水平與垂直移動的最大極限 Facts 限制（畫布總長扣掉方塊本身長度）
    maxAllowedLeft = canvas.Width - anchor.Width
    maxAllowedTop = canvas.Height - anchor.Height
    
    ' 自癒防禦：萬一方塊尺寸不小心被使用者拉到比畫布還大，將極限上限強行鎖死為 0
    If maxAllowedLeft < 0! Then maxAllowedLeft = 0!
    If maxAllowedTop < 0! Then maxAllowedTop = 0!
    
    ' 2. 計算方塊相對於畫布左上角的原始相對位移距離
    relativeLeft = anchor.Left - canvas.Left
    relativeTop = anchor.Top - canvas.Top
    
    ' 3. 四向磁吸自癒防禦（全方位預防使用者將控制項方塊亂拖出邊界）
    ' 左方邊界與上方邊界阻斷：如果位置變成負數，代表被拉出左方或上方外面，強制作為 0
    If relativeLeft < 0! Then relativeLeft = 0!
    If relativeTop < 0! Then relativeTop = 0!
    
    ' 右方邊界與下方邊界阻斷：如果超過最大極限，代表被拉出右方或下方外面，強字作為最大允許上限
    If relativeLeft > maxAllowedLeft Then relativeLeft = maxAllowedLeft
    If relativeTop > maxAllowedTop Then relativeTop = maxAllowedTop
    
    ' 4. 排除畫布尺寸為 0 的除以零風險，並換算為精確的相對百分比
    If canvas.Width > 0! Then
        pctX = relativeLeft / canvas.Width
    Else
        pctX = 0!
    End If
    
    If canvas.Height > 0! Then
        pctY = relativeTop / canvas.Height
    Else
        pctY = 0!
    End If
    
    ' 5. 極限底線防禦（雙重保險鎖死，確保最終傳回的百分比數值絕對在 0.0 ~ 1.0 的安全合法區間內）
    If pctX < 0! Then pctX = 0!
    If pctX > 1! Then pctX = 1!
    If pctY < 0! Then pctY = 0!
    If pctY > 1! Then pctY = 1!
End Sub
