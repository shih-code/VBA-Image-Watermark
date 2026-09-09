' ==========================================================
' MODULE: Mod_Crypto (標準模組)
' PURPOSE: 密碼學與雜湊校對引擎。專職計算實體檔案的 MD5 唯一指紋雜湊值，
'          提供過濾重複圖片的剛性科學依據。
' EXPORTS: CalculateFileHash
' IMPORTS: Mod_StringConstants
' FORBIDDEN: 嚴禁私自 CreateObject 取得 FSO／Stream／MD5 元件，一律透過呼叫端傳入之
'            sysTokens 取用（見 ADR-002）。
' DEPENDENCIES: ACDS_ContractRegistry.md
' VERSION: 1.0.0 [Stability: Stable]  (審計完成版：補上初學者二進位與十六進位轉譯詳細註解)
' ==========================================================
Option Explicit
Option Private Module

Private Const MOD_NAME As String = "Mod_Crypto"
Private Const MODULE_VERSION As String = "1.0.0"

' =========================================================================
' 函數名稱：CalculateFileHash
' 功用說明：讀取一個實體圖片檔案，並計算出該檔案獨一無二的 MD5 數位指紋識別碼。
' 參數說明：
'   - filePath: 要計算的實體檔案的完整絕對路徑（例如 C:\圖片\測試.jpg）。
'   - sysTokens: 主程式大腦傳遞過來的全系統共用工具箱字典。
' 邏輯說明：
' 1. 提領工具：從工具箱中拿出 FSO 工具。
' 2. 預檢防線：檢查 filePath 路徑是否存在，若在硬碟上找不到該檔案，拋出錯誤中斷。
' 3. 二進位讀取：從工具箱提領 STREAM（ADODB.Stream）串流工具，設定為電腦最原始的 Type = 1（二進位模式），
'    把整張相片檔案的位元組（Bytes）完整吸入記憶體，存入 arrFileBytes 陣列。
' 4. 雜湊運算：提領 MD5（.NET 密碼學服務組件），呼叫 ComputeHash_2 方法，對位元組進行核心數學計算，
'    傳回一組包含 16 個數字的結果陣列。
' 5. 字串格式正規化：使用 For 迴圈將 16 個數字一轉成 16 進位的文字（如 255 轉成 "ff"），
'    並內置「兩位數補零防禦」，最後拼接成標準的 32 位元數位指紋字串。
' =========================================================================

Public Function CalculateFileHash(ByVal filePath As String, ByVal sysTokens As Object) As String
    ' 宣告區域變數，明確標示每一個變數的用途
    Dim localFSO As Object          ' 處理硬碟檔案的工具
    Dim localStream As Object       ' 讀取檔案內容資料的工具
    Dim localMD5Provider As Object  ' 專門計算 MD5 識別碼的密碼學工具
    
    Dim arrFileBytes() As Byte      ' 用來裝相片原始二進位資料的陣列（Byte 型態）
    Dim arrHashBytes() As Byte      ' 用來裝密碼學工具算完後的 16 位元組結果陣列
    Dim idxByte As Long             ' 迴圈在讀取陣列時使用的計數器
    Dim strFinalHash As String      ' 最終要回傳的 32 碼完整識別碼字串
    Dim strHexFragment As String    ' 暫存每一次轉換出來的個別十六進位字元
    
    ' --- Cleanup Label 標準配套變數（AVS v1.1 第七節）：暫存錯誤資訊供 Cleanup 判斷 ---
    Dim lngErrNum   As Long
    Dim strErrDesc  As String
    
    ' 1. 從大腦傳進來的工具箱中，提領出處理硬碟檔案的 FSO 檔案工具
    Set localFSO = sysTokens("FSO")
    
    ' 2. 邊界安全檢查：確認要計算的相片在硬碟上是否真的存在。如果找不到檔案，立刻報錯阻斷
    If Not localFSO.FileExists(filePath) Then
        Err.Raise 5402, MOD_NAME, ERR_HASH_FILE_NOT_FOUND
    End If
    
    ' 3. 準備開啟串流讀取實體檔案內容
    '    如果接下來在讀取檔案或計算時發生非預期錯誤，程式會自動跳到最下方的錯誤處理區善後
    On Error GoTo HashStreamErrorHandler
    
    ' 從工具箱中提領出資料串流工具
    Set localStream = sysTokens("STREAM")
    
    ' 設定讀取模式：1 代表 Binary（二進位模式），這是最剛性、不受文字語系影響的原始檔案解讀格式
    localStream.Type = 1
    localStream.Open
    ' 將實體檔案的二進位數據整包下載到記憶體中
    localStream.LoadFromFile filePath
    ' 將數據指派給 Byte 位元組陣列
    arrFileBytes = localStream.Read
    ' 讀取完畢，立刻關閉串流通道，釋放實體檔案鎖定
    localStream.Close
    
    ' 4. 從工具箱中提領出 MD5 密碼學算力工具，並將位元組陣列丟入運算，得到結果位元組
    Set localMD5Provider = sysTokens("MD5")
    arrHashBytes = localMD5Provider.ComputeHash_2(arrFileBytes)
    
    ' 5. 字串格式化：將電腦內部的二進位陣列轉換成人類看得懂的十六進位字串
    strFinalHash = ""
    
    ' 使用 For 迴圈，把算出來的 16 個位元組逐一取出
    For idxByte = LBound(arrHashBytes) To UBound(arrHashBytes)
        ' 使用 Hex 函數將數字轉為十六進位文字（例如：數字 255 會被轉換成字串 "FF"）
        strHexFragment = Hex(arrHashBytes(idxByte))
        
        ' 安全格式修正防線：標準的十六進位數位指紋，每一個位元組必須填滿兩位數。
        ' 如果遇到較小的數值，Hex 轉換出來可能只有一個字元（例如：數字 10 只轉換出一個 "A"），
        ' 此時必須強行在前方補上一個 "0"（使其變成 "0A"），否則最後拼出來的字串會遺失位數。
        ' CHAR_ZERO_PAD 定義於中央常數庫，固定為文字 "0"
        If Len(strHexFragment) = 1 Then
            strHexFragment = CHAR_ZERO_PAD & strHexFragment
        End If
        
        ' 將各段轉換好且補完零的十六進位文字，一段一段向後接合
        strFinalHash = strFinalHash & strHexFragment
    Next idxByte
    
    ' 6. 將最終組合出來的 32 位元字串全部強制轉換為標準小寫字母，正式回傳給呼叫的管線
    CalculateFileHash = LCase(strFinalHash)
    
    GoTo Cleanup

' =========================================================================
' 錯誤處理區塊：只有在讀取實體檔案損毀、或運算中斷時才會執行此區善後
' =========================================================================
HashStreamErrorHandler:
    lngErrNum = 5405                 ' 統一對外呈現為既有的核心影像損毀異常代碼
    strErrDesc = ERR_STREAMP_FAILED
    
    ' 安全解鎖防線：如果串流工具在中途當機、且通道還開著，強制將它關閉，防止該照片一直被 Excel 佔用死鎖
    If Not localStream Is Nothing Then
        On Error Resume Next
        localStream.Close
        On Error GoTo 0
    End If

' --- Cleanup Label（AVS v1.1 第七節）：正常結束與錯誤處理路徑於此唯一匯流 ---
Cleanup:
    Set localFSO = Nothing
    Set localStream = Nothing
    Set localMD5Provider = Nothing
    
    If lngErrNum <> 0 Then
        Err.Raise lngErrNum, MOD_NAME, strErrDesc
    End If
End Function
