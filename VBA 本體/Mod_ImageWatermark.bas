' ==========================================================
' MODULE: Mod_ImageWatermark (標準模組)
' PURPOSE: 影像算力引擎。專職調度 WIA 壓縮巨型相片、調用 Excel Chart 渲染
'          浮水印底板與文字，並匯出 JPG 成品。
' EXPORTS: ExecuteWatermarkStream
' IMPORTS: Mod_StringConstants, cls_Settings, cls_ImageEntity
' FORBIDDEN: 嚴禁在本模組內部私自讀取任何特定名稱之工作表，亦嚴禁私自調用
'            Application.ScreenUpdating 黑幕開關（見 ADR-005，主權歸屬 Mod_WorkflowManager）。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]
' ==========================================================

Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_ImageWatermark"
Private Const MODULE_VERSION As String = "1.0.0"

' =========================================================================
' 程序名稱：ExecuteWatermarkStream
' 功用說明：影像處理核心管線。本程序已達成 100% 介面解耦，只專注於純粹的影像運算。
' =========================================================================
Public Sub ExecuteWatermarkStream( _
    ByVal entity As cls_ImageEntity, _
    ByVal cfg As cls_Settings, _
    ByVal pctX As Single, _
    ByVal pctY As Single, _
    ByVal rawRemark As String, _
    ByVal destFolder As String, _
    ByVal wsCanvasHost As Worksheet, _
    ByVal sysTokens As Object)
    
    Dim localFSO              As Object
    Dim strFinalWatermarkText As String
    Dim strChartSourceImagePath As String
    Dim strTargetJpgPath      As String
    Dim strTempExportPath     As String
    Dim strTempWiaResizedPath As String
    
    Dim lngWorkingWidth       As Long
    Dim lngWorkingHeight      As Long
    Dim strTargetBaseName     As String
    Dim strTargetExtName      As String
    Dim strVersionedPath      As String
    Dim idxVersionCounter     As Long
    
    ' --- Cleanup Label 標準配套變數（AVS v1.1 第七節）：暫存錯誤資訊供 Cleanup 判斷 ---
    Dim lngErrNum   As Long
    Dim strErrDesc  As String
    
    On Error GoTo ErrorHandler

    Set localFSO = sysTokens("FSO")
    
    ' ---------------------------------------------------------
    ' 步驟 1. 組合路徑與加工備註文字 (呼叫標籤替換引擎)
    ' ---------------------------------------------------------
    strTargetJpgPath = destFolder & BACKSLASH_STR & entity.ImageName
    strTempExportPath = strTargetJpgPath & EXT_TMP
    
    strFinalWatermarkText = PrepareWatermarkText(entity, rawRemark)
    
    ' ---------------------------------------------------------
    ' 步驟 2. WIA 影像處理：尺寸探測與等比例縮小
    ' ---------------------------------------------------------
    strChartSourceImagePath = entity.CurrentPath
    strTempWiaResizedPath = HandleWiaResizing(entity.CurrentPath, lngWorkingWidth, lngWorkingHeight, sysTokens)
    
    If strTempWiaResizedPath <> "" Then
        strChartSourceImagePath = strTempWiaResizedPath
    End If
    
    ' ---------------------------------------------------------
    ' 步驟 3. 圖表鑄造與浮水印壓印（傳遞至幾何排版工廠）
    ' ---------------------------------------------------------
    Call RenderWatermarkOnChart(strChartSourceImagePath, strTempExportPath, strFinalWatermarkText, _
                                lngWorkingWidth, lngWorkingHeight, pctX, pctY, wsCanvasHost, cfg)
    
    ' 清理 WIA 降維暫存檔
    If strTempWiaResizedPath <> "" Then
        On Error Resume Next
        Kill strTempWiaResizedPath
        On Error GoTo ErrorHandler
        strTempWiaResizedPath = ""
    End If
    
    ' 剛性檢查成品暫存檔是否存在
    If Not localFSO.FileExists(strTempExportPath) Then
        Err.Raise 5405, MOD_NAME, ERR_STREAMP_FAILED
    End If
    
    ' ---------------------------------------------------------
    ' 步驟 4. 目的地舊檔案版控降級處理（改名備份）
    ' ---------------------------------------------------------
    If localFSO.FileExists(strTargetJpgPath) Then
        strTargetBaseName = localFSO.GetBaseName(strTargetJpgPath)
        strTargetExtName = localFSO.GetExtensionName(strTargetJpgPath)
        
        idxVersionCounter = 1
        Do
            strVersionedPath = destFolder & BACKSLASH_STR & strTargetBaseName & LBL_OLD_VERSION_UNDER & Format(idxVersionCounter, FMT_COUNTER_00) & DOT_STR & strTargetExtName
            If Not localFSO.FileExists(strVersionedPath) Then Exit Do
            idxVersionCounter = idxVersionCounter + 1
        Loop
        
        Call TryRenameWithRetry(strTargetJpgPath, strVersionedPath, False)
    End If
    
    ' ---------------------------------------------------------
    ' 步驟 5. 將 .tmp 成品正名為最終的 .jpg 成品
    ' ---------------------------------------------------------
    Call TryRenameWithRetry(strTempExportPath, strTargetJpgPath, True)

    GoTo Cleanup

ErrorHandler:
    ' 【黑幕越權防線】：嚴格禁止在此私自開啟畫面！主權歸於大腦！
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    
    On Error Resume Next
    If strTempExportPath <> "" Then
        If Dir(strTempExportPath) <> "" Then Kill strTempExportPath
    End If
    If strTempWiaResizedPath <> "" Then
        If Dir(strTempWiaResizedPath) <> "" Then Kill strTempWiaResizedPath
    End If
    On Error GoTo 0

' --- Cleanup Label（AVS v1.1 第七節）：正常結束與錯誤處理路徑於此唯一匯流 ---
Cleanup:
    Set localFSO = Nothing
    
    If lngErrNum <> 0 Then
        Err.Raise lngErrNum, MOD_NAME & DOT_STR & "ExecuteWatermarkStream", strErrDesc
    End If
    Exit Sub
End Sub

' =========================================================================
' 函數名稱：PrepareWatermarkText (標籤化範本替換引擎)
' 功用說明：將剛性拼接升格為「標籤替換」。支援操作員在備註欄自訂多行格式。
' =========================================================================
Private Function PrepareWatermarkText(ByVal entity As cls_ImageEntity, ByVal rawRemark As String) As String
    Dim strTemplate As String
    Dim strResult   As String
    Dim arrLines()  As String
    Dim i           As Long
    Dim strFinal    As String
    
    ' 1. 定義預設範本 (三行式)
    strTemplate = LBL_WM_PREFIX_NAME & TOK_NAME & vbCrLf & _
                  LBL_WM_PREFIX_DATE & TOK_DATE & vbCrLf & _
                  LBL_WM_PREFIX_REMARK & TOK_REMARK
                  
    ' 【神級 UX 攔截】：若使用者在備註欄直接輸入了 {} 標籤，則直接將備註當作「畫布範本」！
    ' 這樣可以達成 "萬一我想變成純備註，或新增多行自訂資料" 的需求
    If InStr(rawRemark, "{") > 0 And InStr(rawRemark, "}") > 0 Then
        strTemplate = rawRemark
        rawRemark = "" ' 已經轉職為範本，清空實體備註避免重複替換
    End If
    
    ' 2. 執行標籤流式替換 (Token Replacement Engine)
    strResult = strTemplate
    strResult = Replace(strResult, TOK_NAME, entity.ImageName)
    strResult = Replace(strResult, TOK_DATE, entity.LastModified)
    strResult = Replace(strResult, TOK_REMARK, rawRemark)
    
    ' 3. 智慧清洗空行 (防止未填寫備註時留下一行空白)
    arrLines = Split(strResult, vbCrLf)
    For i = LBound(arrLines) To UBound(arrLines)
        ' 若該行不是全空，且不是僅剩下 "備註: " 這種空殼，則予以保留
        If Trim(arrLines(i)) <> "" And Trim(arrLines(i)) <> LBL_WM_PREFIX_REMARK Then
            If strFinal = "" Then
                strFinal = arrLines(i)
            Else
                strFinal = strFinal & vbCrLf & arrLines(i)
            End If
        End If
    Next i
    
    PrepareWatermarkText = strFinal
End Function

' =========================================================================
' 函數名稱：HandleWiaResizing (內部私有函數)
' 功用說明：影像降維工廠。
' =========================================================================
Private Function HandleWiaResizing(ByVal srcPath As String, ByRef outWidth As Long, ByRef outHeight As Long, ByVal sysTokens As Object) As String
    Dim objWiaImage     As Object
    Dim objWiaProcess   As Object
    Dim sngScaleRatio   As Single
    Dim lngScaledWidth  As Long
    Dim lngScaledHeight As Long
    Dim strTempPath     As String
    
    HandleWiaResizing = ""
    
    On Error Resume Next
    Set objWiaImage = CreateObject(PROGID_WIA_IMAGE)
    objWiaImage.LoadFile srcPath
    outWidth = objWiaImage.Width
    outHeight = objWiaImage.Height
    On Error GoTo 0
    
    If outWidth = 0 Or outHeight = 0 Then
        outWidth = NUM_DEFAULT_WIDTH
        outHeight = NUM_DEFAULT_HEIGHT
        Exit Function
    End If
    
    If outWidth > NUM_MAX_CANVAS_DIM Or outHeight > NUM_MAX_CANVAS_DIM Then
        If outWidth >= outHeight Then
            sngScaleRatio = NUM_MAX_CANVAS_DIM / outWidth
        Else
            sngScaleRatio = NUM_MAX_CANVAS_DIM / outHeight
        End If
        lngScaledWidth = CLng(outWidth * sngScaleRatio)
        lngScaledHeight = CLng(outHeight * sngScaleRatio)
        
        On Error Resume Next
        Set objWiaProcess = CreateObject(PROGID_WIA_PROCESS)
        objWiaProcess.Filters.Add objWiaProcess.FilterInfos(FILTER_SCALE).FilterID
        objWiaProcess.Filters(1).Properties(PROP_MAX_WIDTH) = lngScaledWidth
        objWiaProcess.Filters(1).Properties(PROP_MAX_HEIGHT) = lngScaledHeight
        Set objWiaImage = objWiaProcess.Apply(objWiaImage)
        On Error GoTo 0
        
        If Not objWiaImage Is Nothing Then
            strTempPath = Environ(ENV_TEMP_VAR) & TMP_FILE_PREFIX & Format(Now, FMT_TIME_HMSM) & EXT_DOT_JPG
            On Error Resume Next
            If Dir(strTempPath) <> "" Then Kill strTempPath
            objWiaImage.SaveFile strTempPath
            On Error GoTo 0
            
            If Dir(strTempPath) <> "" Then
                HandleWiaResizing = strTempPath
                outWidth = lngScaledWidth
                outHeight = lngScaledHeight
            End If
        End If
    End If
    
    Set objWiaProcess = Nothing
    Set objWiaImage = Nothing
End Function

' =========================================================================
' 程序名稱：RenderWatermarkOnChart (內部私有程序)
' 功用說明：圖表壓印工廠 (1.5倍留白精準排版、單一圖形鎔鑄版)
' 【黑幕防禦】：徹底拔除首尾兩端的 ScreenUpdating 越權開關
' =========================================================================
Private Sub RenderWatermarkOnChart( _
    ByVal srcImagePath As String, _
    ByVal exportPath As String, _
    ByVal watermarkText As String, _
    ByVal imgWidth As Long, _
    ByVal imgHeight As Long, _
    ByVal pctX As Single, _
    ByVal pctY As Single, _
    ByVal wsCanvasHost As Worksheet, _
    ByVal cfg As cls_Settings)
    
    Dim objChartContainer As ChartObject
    Dim objChart          As Chart
    Dim shpWatermark      As Shape  ' 【優化】：背景與文字合一
    
    Dim sngGeometryScale  As Single
    Dim sngPosX As Single, sngPosY As Single
    Dim sngBoxWidth As Single, sngBoxHeight As Single
    Dim sngPadPoints      As Single
    
    Dim arrTextLines()    As String
    Dim lngLineCount      As Long
    Dim i                 As Long
    Dim lngMaxLen         As Long
    Dim lngCurLen         As Long
    
    ' 【黑幕越權防線】：嚴格禁止在此私自開啟黑暗！主權歸於大腦！
    
    On Error Resume Next
    Set objChartContainer = wsCanvasHost.ChartObjects(SH_TEMPLATE_CHART_NAME)
    On Error GoTo 0
    
    If objChartContainer Is Nothing Then
        Set objChartContainer = wsCanvasHost.ChartObjects.Add(NUM_CHART_HIDDEN_POS, NUM_CHART_HIDDEN_POS, imgWidth, imgHeight)
        objChartContainer.Name = SH_TEMPLATE_CHART_NAME
    Else
        objChartContainer.Width = imgWidth
        objChartContainer.Height = imgHeight
    End If
    
    Set objChart = objChartContainer.Chart
    
    Dim idxShape As Long
    For idxShape = objChart.Shapes.Count To 1 Step -1
        On Error Resume Next
        objChart.Shapes(idxShape).Delete
        On Error GoTo 0
    Next idxShape
    
    objChart.ChartArea.Format.Fill.UserPicture srcImagePath
    
    sngGeometryScale = imgWidth / NUM_SCALE_BASE
    If sngGeometryScale < 1! Then sngGeometryScale = 1!
    
    arrTextLines = Split(watermarkText, vbCrLf)
    lngLineCount = UBound(arrTextLines) - LBound(arrTextLines) + 1
    
    ' ---------------------------------------------------------
    ' 步驟 A：幾何算力流 - 動態偵測最長字串寬度與 1.5 倍點數換算
    ' ---------------------------------------------------------
    lngMaxLen = 0
    For i = LBound(arrTextLines) To UBound(arrTextLines)
        lngCurLen = Len(arrTextLines(i))
        If lngCurLen > lngMaxLen Then lngMaxLen = lngCurLen
    Next i
    If lngMaxLen < 2 Then lngMaxLen = 2
    
    ' 將 1.5 字元寬度 (NUM_WM_PADDING_CHARS) 換算為精準的實體點數 (Points)
    sngPadPoints = NUM_WM_PADDING_CHARS * cfg.FontSizeDefault * sngGeometryScale
    
    ' 精算總寬與總高：文字實際尺寸 + 兩側留白
    sngBoxWidth = (lngMaxLen * cfg.FontSizeDefault * sngGeometryScale) + (sngPadPoints * 2)
    sngBoxHeight = (lngLineCount * cfg.FontSizeDefault * NUM_LINE_SPACING_BASE * sngGeometryScale) + (sngPadPoints * 2)
    
    sngPosX = imgWidth * pctX
    sngPosY = imgHeight * pctY
    
    ' ---------------------------------------------------------
    ' 步驟 B：邊界自癒流 - 畫布邊界安全煞車防線
    ' ---------------------------------------------------------
    If sngPosY + sngBoxHeight > imgHeight Then sngPosY = imgHeight - sngBoxHeight
    If sngPosX + sngBoxWidth > imgWidth Then sngPosX = imgWidth - sngBoxWidth
    If sngPosX < 0! Then sngPosX = 0!
    If sngPosY < 0! Then sngPosY = 0!
    
    ' ---------------------------------------------------------
    ' 步驟 C：單一圖形鎔鑄 - 同時賦予底板色彩、1.5 倍留白邊距與對齊常數
    ' ---------------------------------------------------------
    Set shpWatermark = objChart.Shapes.AddTextbox(1, sngPosX, sngPosY, sngBoxWidth, sngBoxHeight)
    
    With shpWatermark
        ' 背景渲染
        .Fill.Solid
        .Fill.ForeColor.RGB = cfg.BgColorDefault
        .Fill.Transparency = cfg.BgTransparency
        .Line.Visible = cfg.LineVisible
        
        With .TextFrame2
            ' 1.5 倍四周精準留白注入
            .MarginLeft = sngPadPoints
            .MarginRight = sngPadPoints
            .MarginTop = sngPadPoints
            .MarginBottom = sngPadPoints
            
            ' 垂直置中合約注入 (msoAnchorMiddle)
            .VerticalAnchor = NUM_WM_ALIGN_VERT
            
            ' 文字渲染
            .TextRange.Text = watermarkText
            .TextRange.Font.NameFarEast = cfg.FontNameChinese
            .TextRange.Font.Name = cfg.FontNameEnglish
            .TextRange.Font.Size = cfg.FontSizeDefault * sngGeometryScale
            .TextRange.Font.Fill.ForeColor.RGB = cfg.FontColorDefault
            
            ' 水平靠左合約注入 (msoAlignLeft)
            .TextRange.ParagraphFormat.Alignment = NUM_WM_ALIGN_HORZ
            
            ' 關閉自動換行，強制尊重我們算好的動態寬度防禦線
            .WordWrap = msoFalse
        End With
    End With
    
    objChart.Export Filename:=exportPath, FilterName:="JPG"
    
    Set shpWatermark = Nothing
    Set objChart = Nothing
    Set objChartContainer = Nothing
    
    ' 【黑幕越權防線】：嚴格禁止在此私自開啟畫面！主權歸於大腦！
End Sub

' =========================================================================
' 程序名稱：TryRenameWithRetry (內部共用私有程序)
' 功用說明：強固型改名工具，具備午夜跨日防護與延時避讓機制。
' =========================================================================
 Private Sub TryRenameWithRetry(ByVal sourcePath As String, ByVal destPath As String, ByVal raiseOnError As Boolean)
     Dim cntLockRetry    As Long
     Dim sngTimeSnapshot As Single
     Dim lngCapturedErr  As Long
     Dim strCapturedDesc As String
     
     cntLockRetry = 0
     Do
         On Error Resume Next
         Name sourcePath As destPath
         lngCapturedErr = Err.Number
         strCapturedDesc = Err.Description
         On Error GoTo 0
         
         If lngCapturedErr = 0 Then
             Exit Do
         ElseIf lngCapturedErr = 70 Or lngCapturedErr = 75 Or lngCapturedErr = 58 Then
             cntLockRetry = cntLockRetry + 1
             
             If cntLockRetry >= NUM_RETRY_LIMIT Then
                 If raiseOnError Then
                     Err.Raise 70, MOD_NAME, STATUS_EXPORT_ERR_LOCKED
                 End If
                 Exit Do
             End If
             
             sngTimeSnapshot = Timer
             Do
                 DoEvents
                 ' 午夜跨日防禦：Timer 歸零時強行重置錨點
                 If Timer < sngTimeSnapshot Then sngTimeSnapshot = sngTimeSnapshot - 86400
             Loop While Timer < sngTimeSnapshot + NUM_RETRY_DELAY
         Else
             ' 非鎖定類錯誤（例如目的路徑不存在、磁碟已滿）：不得靜默放棄，
             ' 依 raiseOnError 決定是否往上拋出，讓呼叫端知道改名並未真正成功
             If raiseOnError Then
                 Err.Raise lngCapturedErr, MOD_NAME & ".TryRenameWithRetry", strCapturedDesc
             End If
             Exit Do
         End If
     Loop
 End Sub
