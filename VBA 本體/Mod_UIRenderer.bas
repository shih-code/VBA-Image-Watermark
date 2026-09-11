' ==========================================================
' MODULE: Mod_UIRenderer (標準模組)
' PURPOSE: 視覺與互動介面渲染器。專職負責 ConfigSheet 前台直角畫布、
'          出血線安全框、互動錨點與「寬幅儀表板佈局」五核心按鈕面板的實體鑄造與渲染。
' EXPORTS: RenderUIElements
' IMPORTS: Mod_UI_Messenger, Mod_StringConstants
' FORBIDDEN: 嚴禁在此模組直接呼叫任何檔案系統或影像處理函式。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_UIRenderer"
Private Const MODULE_VERSION As String = "1.0.0"

' =========================================================================
' 程序名稱：RenderUIElements
' 功用說明：前端視覺互動組件的總鑄造入口。
' 參數說明：
'   - forceResetLayout：預設 False（沿用），保留使用者拖曳過的浮水印錨點位置，
'                       供「匯入資料夾」／「挑選圖片」等日常操作呼叫，不打擾版面。
'                       傳入 True 則強制重置畫布／出血框／錨點回預設座標，
'                       僅供「一鍵重整工作檯」這類明確要求恢復原廠設定的救援
'                       動作呼叫。見 ADR-012（延伸 ADR-011）。
' 【自癒修正】：引進 wsCurrent 錨點記憶，物理封殺瞬移與畫面閃爍！
' =========================================================================
Public Sub RenderUIElements(Optional ByVal forceResetLayout As Boolean = False)
    Dim wsConfig  As Worksheet
    Dim wsCurrent As Worksheet ' 追加位置記憶錨點
    
    On Error Resume Next
    ' 【修復防線 1】：在黑幕拉上之前，用最快速度記住您現在正看著哪一張工作表 Facts
    Set wsCurrent = ActiveSheet
    Set wsConfig = ThisWorkbook.Worksheets(SHEET_CONFIG)
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        Err.Raise 5401, MOD_NAME, "環境阻斷：找不到配置表 (ConfigSheet)，視覺渲染器拒絕發車！"
    End If
    
    ' 【修復防線 2】：第一步立刻拉上黑幕！不讓微軟底層的任何重繪畫面漏出，徹底消除閃爍
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' 解除保護並切換至後台施工
    wsConfig.Unprotect Password:=PWD_PROTECT
    wsConfig.Activate
    
    ' 物理擦除 Excel 內建的灰色死板網格線
    ActiveWindow.DisplayGridlines = False
    
    ' 流式調用私有渲染程序（在後台純白地基上默默施工）
    Call ClearAllExistingShapes(wsConfig)
    Call RenderCanvasAndAnchor(wsConfig, forceResetLayout)
    Call RenderControlPanel(wsConfig)
    
    ' 完工上鎖
    wsConfig.Protect Password:=PWD_PROTECT, DrawingObjects:=True, Contents:=False, UserInterfaceOnly:=True
    
    ' 【修復防線 3】：神不知鬼不覺！在拉開黑幕前，強行把畫面跳轉回操作員原本點擊按鈕的那張表
    If Not wsCurrent Is Nothing Then wsCurrent.Activate
    
    ' 完工，安全拉開黑幕，前台視覺體驗將如同行雲流水般順暢
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' 發生異常時，同樣要確保先把畫面送回原地、並扯開黑幕，防止 Excel 介面集體癱瘓
    If Not wsCurrent Is Nothing Then wsCurrent.Activate
    Application.ScreenUpdating = True
    Call Mod_UI_Messenger.ShowError(ERR_UI_BUILD_FAILED_PREFIX & Err.Description, TITLE_ENV_MELTDOWN)
End Sub

' =========================================================================
' 程序名稱：ClearAllExistingShapes (內部私有程序)
' 功用說明：物理清除檯面上舊版遺留與按鈕面板組件。
' 【位置保留修復，見 ADR-011】：SH_CANVAS_NAME／SH_BLEED_ZONE_NAME／SH_ANCHOR_NAME
'                   三者不在本清單內無條件刪除，是否重置改由 RenderCanvasAndAnchor
'                   依 forceReset 參數決定（見 ADR-012）。
'                   按鈕（BTN_*）本身不承載使用者自訂狀態，維持每次刪除重建無妨。
' =========================================================================
Private Sub ClearAllExistingShapes(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Shapes(LGC_BTN_OLD_IMPORT).Delete
    ws.Shapes(LGC_BTN_OLD_EXPORT).Delete
    ws.Shapes(BTN_FOLDER_NAME).Delete
    ws.Shapes(BTN_FILES_NAME).Delete
    ws.Shapes(BTN_EXPORT_NAME).Delete
    ws.Shapes(BTN_REBUILD_NAME).Delete
    ws.Shapes(BTN_CONSOLIDATE_NAME).Delete
    ws.Shapes(SH_LICENSE_FOOTER_NAME).Delete
    On Error GoTo 0
End Sub

' =========================================================================
' 程序名稱：RenderCanvasAndAnchor (內部私有程序)
' 功用說明：繪製左側直角模擬畫布、出血安全框與互動錨點。
' 參數說明：
'   - forceReset：預設 False（沿用既有位置，見 ADR-011）。傳入 True 時，會先行
'                 物理刪除既有的畫布／出血框／錨點，強制回到預設座標重新鑄造，
'                 僅供「一鍵重整工作檯」呼叫，作為版面異常時的救援手段（見 ADR-012）。
' =========================================================================
Private Sub RenderCanvasAndAnchor(ByVal ws As Worksheet, Optional ByVal forceReset As Boolean = False)
    Dim shpMainCanvas      As Shape
    Dim shpBleedZone       As Shape
    Dim shpWatermarkAnchor As Shape
    Dim strDefaultFont     As String
    
    ' 鎖定微軟正黑體
    strDefaultFont = VAL_DEFAULT_FONT_CHINESE
    
    ' 【重整UI專用】：若明確要求強制重置，先物理刪除三者，讓下方探測邏輯視為
    ' 不存在，回歸依預設座標重新鑄造，達成「恢復原廠設定」的救援效果。
    If forceReset Then
        On Error Resume Next
        ws.Shapes(SH_CANVAS_NAME).Delete
        ws.Shapes(SH_BLEED_ZONE_NAME).Delete
        ws.Shapes(SH_ANCHOR_NAME).Delete
        On Error GoTo 0
    End If
    
    ' --- 探測既有畫布是否已存在 ---
    On Error Resume Next
    Set shpMainCanvas = ws.Shapes(SH_CANVAS_NAME)
    On Error GoTo 0
    
    ' 1. 主畫布：只有真的不存在時才建立【對齊修正】：Y軸加 1 列，避開上方 3 列的儲存格表頭
    If shpMainCanvas Is Nothing Then
        Set shpMainCanvas = ws.Shapes.AddShape(msoShapeRectangle, _
                            ws.Cells(GEO_GRID_CANVAS_ROW + 1, GEO_GRID_CANVAS_COL).Left + 5, _
                            ws.Cells(GEO_GRID_CANVAS_ROW + 1, GEO_GRID_CANVAS_COL).Top + 5, _
                            GEO_CANVAS_WIDTH, GEO_CANVAS_HEIGHT)
        With shpMainCanvas
            .Name = SH_CANVAS_NAME
            .Fill.Solid
            .Fill.ForeColor.RGB = COLOR_UI_CANVAS_BG
            .Line.ForeColor.RGB = COLOR_UI_CANVAS_LINE
            .Line.Weight = 1#
            With .TextFrame
                .Characters.Text = LBL_CANVAS_BANNER
                .Characters.Font.Name = strDefaultFont
                .Characters.Font.Size = GEO_FONT_SIZE_CANVAS
                .Characters.Font.Color = COLOR_UI_CANVAS_TEXT
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End With
    End If
    
    ' --- 探測既有出血框是否已存在 ---
    On Error Resume Next
    Set shpBleedZone = ws.Shapes(SH_BLEED_ZONE_NAME)
    On Error GoTo 0
    
    ' 2. 出血安全框：只有真的不存在時才建立，位置相對於畫布計算
    If shpBleedZone Is Nothing Then
        Set shpBleedZone = ws.Shapes.AddShape(msoShapeRectangle, _
                          shpMainCanvas.Left + GEO_BLEED_OFFSET_X, _
                          shpMainCanvas.Top + GEO_BLEED_OFFSET_Y, _
                          GEO_BLEED_WIDTH, GEO_BLEED_HEIGHT)
        With shpBleedZone
            .Name = SH_BLEED_ZONE_NAME
            .Fill.Visible = msoFalse
            .Line.ForeColor.RGB = COLOR_UI_LABEL_GRAY
            .Line.DashStyle = msoLineDash
            .Line.Weight = 1#
        End With
    End If
    
    ' --- 探測既有互動錨點是否已存在（本專案唯一承載使用者自訂拖曳狀態的物件）---
    On Error Resume Next
    Set shpWatermarkAnchor = ws.Shapes(SH_ANCHOR_NAME)
    On Error GoTo 0
    
    ' 3. 互動錨點：只有真的不存在時才建立於預設位置；存在則完全不動，
    '    保留使用者上次拖曳過的實際 Left／Top 座標（forceReset=True 時已於上方
    '    被物理刪除，此處會視為不存在並依預設位置重建）。
    If shpWatermarkAnchor Is Nothing Then
        Set shpWatermarkAnchor = ws.Shapes.AddShape(msoShapeRectangle, _
                                shpMainCanvas.Left + GEO_BLEED_OFFSET_X, _
                                shpMainCanvas.Top + shpMainCanvas.Height - (GEO_ANCHOR_HEIGHT + GEO_BLEED_OFFSET_Y), _
                                GEO_ANCHOR_WIDTH, GEO_ANCHOR_HEIGHT)
        With shpWatermarkAnchor
            .Name = SH_ANCHOR_NAME
            .Fill.Solid
            .Fill.ForeColor.RGB = COLOR_UI_ANCHOR_BG      ' 載入莫蘭迪藍底
            .Line.ForeColor.RGB = COLOR_UI_ANCHOR_LINE    ' 載入細深藍線
            .Line.Weight = 1#
            With .TextFrame
                .Characters.Text = LBL_ANCHOR_TEXT
                .Characters.Font.Name = strDefaultFont      ' 鎖定微軟正黑體
                .Characters.Font.Size = GEO_FONT_SIZE_LABEL
                .Characters.Font.Bold = True
                .Characters.Font.Color = COLOR_UI_ANCHOR_TEXT ' 純白字
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            .Locked = msoFalse ' 剛性保留可拖拽事實
        End With
    End If
    
    Set shpMainCanvas = Nothing: Set shpBleedZone = Nothing: Set shpWatermarkAnchor = Nothing
End Sub
' =========================================================================
' 函數名稱：CreateStyledButton (內部私有函數)
' 功用說明：按鈕鑄造工廠。統一承接圓角矩形按鈕的建立、填色、文字樣式設定，
'           讓 RenderControlPanel 不需要為每顆按鈕重複同一套 AddShape／TextFrame 樣板。
' =========================================================================
Private Function CreateStyledButton( _
    ByVal ws As Worksheet, ByVal sngLeft As Single, ByVal sngTop As Single, _
    ByVal sngWidth As Single, ByVal sngHeight As Single, _
    ByVal strShapeName As String, ByVal strOnAction As String, _
    ByVal lngFillColor As Long, ByVal lngTextColor As Long, _
    ByVal strLabel As String, ByVal strFontName As String, ByVal sngFontSize As Single) As Shape
    
    Dim shpNewButton As Shape
    Set shpNewButton = ws.Shapes.AddShape(msoShapeRoundedRectangle, sngLeft, sngTop, sngWidth, sngHeight)
    
    With shpNewButton
        .Name = strShapeName
        .OnAction = strOnAction
        .Fill.Solid
        .Fill.ForeColor.RGB = lngFillColor
        .Line.Visible = msoFalse
        With .TextFrame
            .Characters.Text = strLabel
            .Characters.Font.Name = strFontName
            .Characters.Font.Size = sngFontSize
            .Characters.Font.Bold = True
            .Characters.Font.Color = lngTextColor
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End With
    
    Set CreateStyledButton = shpNewButton
    Set shpNewButton = Nothing
End Function
' =========================================================================
' 程序名稱：RenderControlPanel (內部私有程序)
' 功用說明：寬幅對稱控制面板佈局。0% 硬編碼，完全聽從中央常數在庫的剛性Facts。
' =========================================================================
Private Sub RenderControlPanel(ByVal ws As Worksheet)
    Dim btnPhase1Folder    As Shape
    Dim btnPhase1Files     As Shape
    Dim btnPhase2Export    As Shape
    Dim btnForceRebuild    As Shape
    Dim btnConsolidateLogs As Shape
    Dim strDefaultFont     As String
    
    Dim sngBaseLeft        As Single
    Dim sngHeroWidth       As Single ' 【致命型態修正】：將 String 改為 Single，避免報錯
    
    strDefaultFont = Left(VAL_FONT_CHINESE, InStr(VAL_FONT_CHINESE, ",") - 1)
    
    ' 【對齊修正】：按鈕左緣對齊第 5 欄 (E欄)，Y軸全部加 1 列以齒合畫布與表頭高度
    sngBaseLeft = ws.Cells(GEO_GRID_BTN_ROW_P1 + 1, 5).Left
    sngHeroWidth = (GEO_BTN_WIDTH * 2) + GEO_BTN_GAP_X
    
    ' ---------------------------------------------------------
    ' 【第一軌】：圖片資產匯入區（雙軌並排、柔和冷灰藍常數）
    ' ---------------------------------------------------------
    ' 1. 匯入資料夾（靠左）
    Set btnPhase1Folder = CreateStyledButton(ws, sngBaseLeft, ws.Cells(GEO_GRID_BTN_ROW_P1 + 1, 5).Top, _
         GEO_BTN_WIDTH, GEO_BTN_HEIGHT, BTN_FOLDER_NAME, MACRO_IMPORT_FOLDER, _
         COLOR_UI_BTN_P1_BG, COLOR_UI_BTN_P1_TEXT, LBL_BTN_FOLDER, strDefaultFont, GEO_FONT_SIZE_BTN_SMALL)

     ' 2. 挑選圖片（向右偏移）
     Set btnPhase1Files = CreateStyledButton(ws, sngBaseLeft + GEO_BTN_WIDTH + GEO_BTN_GAP_X, ws.Cells(GEO_GRID_BTN_ROW_P1 + 1, 5).Top, _
         GEO_BTN_WIDTH, GEO_BTN_HEIGHT, BTN_FILES_NAME, MACRO_IMPORT_FILES, _
         COLOR_UI_BTN_P1_BG, COLOR_UI_BTN_P1_TEXT, LBL_BTN_FILES, strDefaultFont, GEO_FONT_SIZE_BTN_SMALL)
    
    ' ---------------------------------------------------------
    ' 【第二軌】：核心主導輸出區（滿幅 Hero Button、權威皇家藍）
    ' ---------------------------------------------------------
    ' 3. 匯出浮水印（獨佔整列，只有核心 Hero 按鈕加大加厚，展現主宰感）
    Set btnPhase2Export = CreateStyledButton(ws, sngBaseLeft, ws.Cells(GEO_GRID_BTN_ROW_P2 + 1, 5).Top, _
        sngHeroWidth, GEO_BTN_HEIGHT + 4, BTN_EXPORT_NAME, MACRO_EXPORT, _
        COLOR_UI_BTN_P2_BG, COLOR_UI_CANVAS_BG, LBL_BTN_EXPORT, strDefaultFont, GEO_FONT_SIZE_BTN)
    
    ' ---------------------------------------------------------
    ' 【第三軌】：系統維護工具區（雙軌並排、內斂極簡低調灰）
    ' ---------------------------------------------------------
    ' 4. 一鍵還原檯面（靠左）
    Set btnForceRebuild = CreateStyledButton(ws, sngBaseLeft, ws.Cells(GEO_GRID_BTN_ROW_MAINT + 1, 5).Top, _
        GEO_BTN_WIDTH, GEO_BTN_HEIGHT, BTN_REBUILD_NAME, MACRO_REBUILD_UI, _
        COLOR_UI_BTN_MAINT_BG, COLOR_UI_BTN_MAINT_TEXT, LBL_BTN_REBUILD, strDefaultFont, GEO_FONT_SIZE_BTN_SMALL)

    ' 5. 歷史日誌歸檔（向右偏移）
    Set btnConsolidateLogs = CreateStyledButton(ws, sngBaseLeft + GEO_BTN_WIDTH + GEO_BTN_GAP_X, ws.Cells(GEO_GRID_BTN_ROW_MAINT + 1, 5).Top, _
        GEO_BTN_WIDTH, GEO_BTN_HEIGHT, BTN_CONSOLIDATE_NAME, MACRO_CONSOLIDATE, _
        COLOR_UI_BTN_MAINT_BG, COLOR_UI_BTN_MAINT_TEXT, LBL_BTN_CONSOLIDATE, strDefaultFont, GEO_FONT_SIZE_BTN_SMALL)
    
    ' ---------------------------------------------------------
    ' 【版權頁尾，修正版】：低調小字版權聲明，與系統維護工具區呼應同色調。
    ' 【寬度修正】：原本沿用 sngHeroWidth（僅兩顆按鈕併排寬），文字在 7pt
    ' 字級下仍會自動換行，但方塊高度只有一行，第二行被垂直裁切看起來像
    ' 截斷。改為從畫布左緣（GEO_GRID_CANVAS_COL 所在欄）橫跨至按鈕面板
    ' 右緣，涵蓋整個工作表可視寬度。V17 起文字本身改為兩行（授權種類／
    ' 版權聲明各一行），GEO_FOOTER_HEIGHT 同步加倍，避免重演此處原本
    ' 就踩過的裁切問題。
    ' 與 ThisWorkbook.cls 內完整授權全文互為表裡：此處為精簡可視版本，
    ' ThisWorkbook 為完整法律條款存放處，兩者不重複維護同一份全文。
    ' ---------------------------------------------------------
    Dim shpLicenseFooter As Shape
    Dim sngFooterLeft     As Single
    Dim sngFooterWidth    As Single
    
    sngFooterLeft = ws.Cells(GEO_GRID_CANVAS_ROW + 1, GEO_GRID_CANVAS_COL).Left + 5
    sngFooterWidth = (sngBaseLeft + sngHeroWidth) - sngFooterLeft
    
    Set shpLicenseFooter = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                          sngFooterLeft, _
                          ws.Cells(GEO_GRID_BTN_ROW_MAINT + 1, 5).Top + GEO_BTN_HEIGHT + GEO_FOOTER_GAP_Y, _
                          sngFooterWidth, GEO_FOOTER_HEIGHT)
    With shpLicenseFooter
        .Name = SH_LICENSE_FOOTER_NAME
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame
            .Characters.Text = LBL_LICENSE_FOOTER
            .Characters.Font.Name = strDefaultFont
            .Characters.Font.Size = GEO_FONT_SIZE_FOOTER
            .Characters.Font.Bold = False
            .Characters.Font.Color = COLOR_UI_FOOTER_TEXT
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    End With
    Set shpLicenseFooter = Nothing

    ' 解除所有物件參照
    Set btnPhase1Folder = Nothing: Set btnPhase1Files = Nothing: Set btnPhase2Export = Nothing
    Set btnForceRebuild = Nothing: Set btnConsolidateLogs = Nothing
End Sub

