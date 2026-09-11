' ==========================================================
' MODULE: Mod_SheetSchemaBuilder (標準模組)
' PURPOSE: 底層數據網格建構師。專職負責 Config 頁莫蘭迪暖沙卡片、Staging 看板、
'          以及日誌追蹤頁的呼吸感列高、正黑體格式剛性鑄造。
' EXPORTS: BuildSystemSchemas
' IMPORTS: Mod_StringConstants
' FORBIDDEN: 嚴禁寫入任何影像處理或批次迴圈邏輯，僅負責工作表結構與格式鑄造。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_SheetSchemaBuilder"
Private Const MODULE_VERSION As String = "1.0.0"

' ==========================================================
' 程序名稱：BuildSystemSchemas
' 功用說明：數據地基總調度入口。
' ==========================================================
Public Sub BuildSystemSchemas()
    Dim wbActive  As Workbook
    Dim wsConfig  As Worksheet
    Dim wsStaging As Worksheet
    
    Set wbActive = ThisWorkbook
    
    Set wsConfig = FetchOrCreateSheet(wbActive, SHEET_CONFIG)
    Call BuildSubConfigSchema(wsConfig)
    
    Set wsStaging = FetchOrCreateSheet(wbActive, SHEET_STAGING)
    Call BuildSubStagingSchema(wsStaging)
End Sub

' =========================================================================
' 程序名稱：BuildSubConfigSchema
' 功用說明：配置表網格架構再生引擎。徹底清洗 Row 2 與 Row 3，防禦箭頭交叉感染。
' =========================================================================
Private Sub BuildSubConfigSchema(ByVal ws As Worksheet)
    ' 施工前先解鎖
    ws.Unprotect Password:=PWD_PROTECT
    
    ' 【終極橫向攤開防線】：確保 A1:E1 絕對貫通合併，並嚴禁文字換行！
    Application.DisplayAlerts = False
    ws.Range("A1:E1").Merge
    ws.Range("A1:E1").WrapText = False
    Application.DisplayAlerts = True
    
    ' 1. 【物理超渡 facts】：連根拔除 A2:E3 區間內所有錯置選單，誓死捍衛表頭純淨
    On Error Resume Next
    ws.Range(ADDR_CFG_CTRL_CARD_RANGE).Validation.Delete
    On Error GoTo 0

    ' 2. 第二列 (Row 2) 剛性寫入純文字表頭 Facts
    ws.Range(ADDR_CFG_LBL_CHINESE).Value = LBL_CONFIG_FONT_CHINESE
    ws.Range(ADDR_CFG_LBL_ENGLISH).Value = LBL_CONFIG_FONT_ENGLISH
    ws.Range(ADDR_CFG_LBL_COLOR).Value = LBL_CONFIG_FONT_COLOR
    ws.Range(ADDR_CFG_LBL_SIZE).Value = LBL_CONFIG_FONT_SIZE
    ws.Range(ADDR_CFG_LBL_TEMPLATE).Value = LBL_CONFIG_TEMPLATE
    
    ' 標題層卡片樣式美化 (Row 2) ? 莫蘭迪暖沙色覆蓋
    With ws.Range(ADDR_CFG_LBL_RANGE)
        .Font.Name = VAL_DEFAULT_FONT_CHINESE
        .Font.Size = GEO_FONT_SIZE_LABEL
        .Font.Color = COLOR_UI_LABEL_GRAY
        .Font.Bold = True
        .Interior.Color = COLOR_UI_CTRL_CARD_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' 3. 第三列 (Row 3) 數據選數層預防性自癒 (0% 硬編碼，完全呼叫常數 facts)
    If Trim(ws.Range(ADDR_CFG_FONT_CHINESE).Value) = "" Or ws.Range(ADDR_CFG_FONT_CHINESE).Value = LBL_CONFIG_FONT_CHINESE Then
        ws.Range(ADDR_CFG_FONT_CHINESE).Value = VAL_DEFAULT_FONT_CHINESE
    End If
    If Trim(ws.Range(ADDR_CFG_FONT_ENGLISH).Value) = "" Or ws.Range(ADDR_CFG_FONT_ENGLISH).Value = LBL_CONFIG_FONT_ENGLISH Then
        ws.Range(ADDR_CFG_FONT_ENGLISH).Value = VAL_FALLBACK_ENG_FONT
    End If
    If Trim(ws.Range(ADDR_CFG_FONT_COLOR).Value) = "" Or ws.Range(ADDR_CFG_FONT_COLOR).Value = LBL_CONFIG_FONT_COLOR Then
        ws.Range(ADDR_CFG_FONT_COLOR).Value = VAL_DEFAULT_COLOR
    End If
    If Val(ws.Range(ADDR_CFG_FONT_SIZE).Value) = 0 Then
        ws.Range(ADDR_CFG_FONT_SIZE).Value = GEO_FONT_SIZE_DEFAULT
    End If
    If Trim(CStr(ws.Range(ADDR_CFG_VAL_TEMPLATE).Value)) = "" Then
        ws.Range(ADDR_CFG_VAL_TEMPLATE).Value = VAL_DEFAULT_TEMPLATE
    End If
    
    ' 數值選數層樣式美化 (Row 3) ? 高質感純白底色 facts
    With ws.Range(ADDR_CFG_VAL_RANGE)
        .Interior.Color = vbWhite
        .Font.Name = VAL_DEFAULT_FONT_CHINESE
        .Font.Size = GEO_FONT_SIZE_DEFAULT
        .Font.Color = COLOR_UI_BTN_P1_TEXT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' 4. 剛性架設：在第三列 (Row 3) 精準綁定全新的資料驗證選單物件
    On Error Resume Next
    With ws.Range(ADDR_CFG_FONT_CHINESE).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=VAL_FONT_CHINESE
        .IgnoreBlank = True: .InCellDropdown = True
    End With
    With ws.Range(ADDR_CFG_FONT_ENGLISH).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=VAL_FONT_ENGLISH
        .IgnoreBlank = True: .InCellDropdown = True
    End With
    With ws.Range(ADDR_CFG_FONT_COLOR).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=VAL_COLORS
        .IgnoreBlank = True: .InCellDropdown = True
    End With
    With ws.Range(ADDR_CFG_FONT_SIZE).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=VAL_FONT_SIZES
        .IgnoreBlank = True: .InCellDropdown = True
    End With
    With ws.Range(ADDR_CFG_VAL_TEMPLATE).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=VAL_TEMPLATE_LIST
        .IgnoreBlank = True: .InCellDropdown = True: .ShowInput = True: .ShowError = True
    End With
    On Error GoTo 0
    
    ' 5. 高級感垂直呼吸高度對齊
    ws.Rows(1).RowHeight = GEO_ROW_HEIGHT_HEADER ' 頂層橫幅列 (回歸單行 30 像素 facts)
    ws.Rows(2).RowHeight = GEO_ROW_HEIGHT_HEADER ' 標籤層
    ws.Rows(3).RowHeight = GEO_ROW_HEIGHT_DATA   ' 選數層
    ws.Rows(4).RowHeight = 15                    ' 留白呼吸層

    ' 6. 【剛性擴張欄寬】撐開 A~D 容納畫布，E 欄退讓給按鈕 facts
    ws.Columns("A").ColumnWidth = 15
    ws.Columns("B").ColumnWidth = 18 ' 英文字體長，給多一點
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 45 ' 範本欄位與按鈕起始區

    ' 【核心防禦 7】：退場精密上鎖 facts
    ws.Protect Password:=PWD_PROTECT, DrawingObjects:=True, Contents:=False, UserInterfaceOnly:=True
End Sub

' =========================================================================
' 程序名稱：BuildSubStagingSchema
' =========================================================================
Private Sub BuildSubStagingSchema(ByVal ws As Worksheet)
    ws.Unprotect Password:=PWD_PROTECT
    
    ws.Cells(NUM_ROW_HEADER, NUM_STG_COL_STATUS).Value = COL_NAME_STATUS
    ws.Cells(NUM_ROW_HEADER, NUM_STG_COL_NAME).Value = COL_NAME_IMG_NAME
    ws.Cells(NUM_ROW_HEADER, NUM_STG_COL_REMARK).Value = COL_NAME_REMARK
    ws.Cells(NUM_ROW_HEADER, NUM_STG_COL_MOD).Value = COL_NAME_LAST_MOD
    ws.Cells(NUM_ROW_HEADER, NUM_STG_COL_PATH).Value = COL_NAME_CUR_PATH
    ws.Cells(NUM_ROW_HEADER, NUM_STG_COL_DATE).Value = COL_NAME_IMPORT_DATE
    ws.Cells(NUM_ROW_HEADER, NUM_STG_COL_REL).Value = COL_NAME_IMPORT_REL
    ws.Cells(NUM_ROW_HEADER, NUM_STG_COL_HASH).Value = COL_NAME_HASH
    ws.Cells(NUM_ROW_HEADER, NUM_STG_COL_EXPORTED).Value = COL_NAME_EXPORTED
    
    With ws.Range(ADDR_STG_HDR_RANGE)
        .Font.Name = VAL_DEFAULT_FONT_CHINESE
        .Font.Bold = True
        .Font.Size = GEO_FONT_SIZE_BTN_SMALL
        .Font.Color = COLOR_UI_BTN_P1_TEXT
        .Interior.Color = COLOR_LOG_HEADER_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Rows(NUM_ROW_HEADER).RowHeight = GEO_ROW_HEIGHT_HEADER
    
    ws.Columns(COL_STG_A).ColumnWidth = WID_STG_A
    ws.Columns(COL_STG_B).ColumnWidth = WID_STG_B
    ws.Columns(COL_STG_C).ColumnWidth = WID_STG_C
    ws.Columns(COL_STG_D).ColumnWidth = WID_STG_D
    ws.Columns(COL_STG_E).ColumnWidth = WID_STG_E
    ws.Columns(COL_STG_F_TO_H).ColumnWidth = WID_STG_F_TO_H
    ws.Columns(COL_STG_I).ColumnWidth = WID_STG_I
    
    ws.Cells.NumberFormat = FMT_TEXT_CELL
End Sub

' =========================================================================
' 函數名稱：FetchOrCreateSheet (內部私有函數)
' =========================================================================
Private Function FetchOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set FetchOrCreateSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If FetchOrCreateSheet Is Nothing Then
        If wb.ProtectStructure Then Err.Raise 5503, MOD_NAME, ERR_LOG_STRUCTURE_LOCKED
        Set FetchOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        FetchOrCreateSheet.Name = sheetName
    End If
End Function
