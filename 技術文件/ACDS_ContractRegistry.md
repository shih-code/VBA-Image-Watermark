# ACDS_ContractRegistry.md - 中央合約註冊表

**專案名稱**：Excel 驅動型圖片兩階段批次浮水印處理系統（純模組化強固版）
**對應原始碼版本**：V17（2026-09 系列，見 ADR-017、ADR-018；V13→V17 全模組 VERSION 沿用 1.0.0，`Mod_WorkflowManager`／`Mod_FileSystem`／`Mod_InteractionAdapter`／`Mod_ExportPipeline` 因本輪重構降級為 Experimental，待真實資料實測後依第十一章晉升回 Stable）
**本文件狀態**：以實際 VBA 原始碼逐行核對後回寫（reverse-engineered from ground truth），非事前規劃文件。
**核對方法**：`oletools.olevba` 抽取 V17 xlsm 全部業務模組原始碼，逐一比對函式簽章、欄位索引、常數命名。歷經 ADR-001 至 ADR-018 共 18 輪修正後之狀態。

> 本文件取代先前所有版本（V0 的 ACDS_ContractRegistry.md、`合約.txt`、`合約與依賴關係圖.txt` 中的合約段落、`規格計畫.txt`／`規格書.txt` 中的變數清單）。上述文件之間彼此有欄位順序、常數命名等矛盾之處，一律以本文件為準，因為本文件是唯一直接對照原始碼產生的版本。

---

## 一、全域常數（Variables Specification）

全系統唯一的常數 SSOT 是 `Mod_StringConstants.bas`（現行 VERSION 1.0.0，內容規模承襲自 V12 時期
的 4.4.0 版），內部以「區」為單位分類（第一區至第二十一區），涵蓋：控制項名稱、巨集路由名稱、
工作表名稱、儲存格絕對地址、UI 幾何尺寸、莫蘭迪色彩碼、資料夾命名前綴、欄位名稱與欄位索引、
狀態燈號文字、下拉選單選項、系統提示訊息庫、外部元件 ProgID、資料格式字串、DAG 追蹤標籤，
加上 V13 新增的空白頁面清理相關常數，共 21+ 個分類區塊。

**本註冊表不重複列出全部常數字面值**（避免與 `Mod_StringConstants` 產生第二個事實落點，違反 SSOT 原則），僅登記以下對其他模組具有「合約」意義、會被跨模組依賴的關鍵常數群組：

| 常數群組 | 代表常數 | 合約意義 |
|---|---|---|
| 工作表名稱 | `SHEET_CONFIG`, `SHEET_STAGING`, `SHEET_LOG_EXPORT`, `SHEET_MASTER_ARCHIVE` | 全系統工作表定位唯一依據 |
| Staging 欄位索引 | `NUM_STG_COL_STATUS`(1) ~ `NUM_STG_COL_EXPORTED`(9) | 見下方「二、工作表資料合約」，V14 起由 8 欄擴為 9 欄，見 ADR-017 |
| LOG 欄位索引 | `NUM_LOG_COL_STATUS`(1) ~ `NUM_LOG_COL_PATH`(4) | 見下方「二、工作表資料合約」 |
| Config 儲存格地址 | `ADDR_CFG_FONT_CHINESE`("A3") 等 5 組 | ConfigSheet 讀寫唯一依據 |
| 外部元件 ProgID | `PROGID_FSO`, `PROGID_STREAM`, `PROGID_MD5`, `PROGID_WIA_IMAGE`, `PROGID_WIA_PROCESS`, `PROGID_DICTIONARY` | 見下方「四、共用工具箱（sysTokens）合約」 |
| 資料夾命名 | `DIR_NAME_WORKSPACE_ROOT`, `DIR_NAME_IMPORT_FOLDER`, `DIR_NAME_RESULT_FOLDER`, `DIR_NAME_BACKUP_ROOT`, `DIR_NAME_BACKUP_CHECKPOINT_PREFIX`, `DIR_NAME_BACKUP_FINAL_PREFIX` | 見 ADR-017，取代 ADR-006 原始設計（固定拓撲＋事件式整批備份） |
| 舊目錄總帳 | `SHEET_OLD_CATALOG_MASTER`, `MSG_OLD_CATALOG_BATCH_MARK` | 見 ADR-018 |
| 狀態燈號 | `STATUS_IMPORT_OK/ERR`, `STATUS_EXPORT_OK/ERR`, `STATUS_EXPORT_ERR_LOCKED` | 跨模組狀態判斷唯一依據 |
| 環境限制 | `NUM_RETRY_LIMIT`(3), `NUM_RETRY_DELAY`(0.1), `NUM_MAX_HYPERLINK_LEN`(255), `NUM_MAX_CANVAS_DIM`(2200) | 見 ADR-006、第十八章 Fail-Fast Focus |
| 空白頁面清理 | `LBL_PROTECT_MARK`("###"), `MSG_PURGE_CONFIRM_PREFIX/SUFFIX`, `TITLE_PURGE_CONFIRM`, `MSG_PURGE_NONE_FOUND`, `ERR_STRUCTURE_LOCKED_PURGE` | V13 新增，見 ADR-013 |
| 版權聲明頁尾 | `SH_LICENSE_FOOTER_NAME`, `LBL_LICENSE_FOOTER`, `GEO_FONT_SIZE_FOOTER`, `GEO_FOOTER_GAP_Y`, `GEO_FOOTER_HEIGHT`, `COLOR_UI_FOOTER_TEXT` | V13 新增，見 ADR-015 |

**已移除的舊版常數**（V0/V1.1.0 遺留，V12 已不存在，僅供歷史對照）：
- `cfg_*` 系列（改由 `cls_Settings` 屬性承接）
- `sh_CanvasName`／`sh_AnchorName`（改為 `SH_CANVAS_NAME`／`SH_ANCHOR_NAME`，全大寫命名對齊）
- `DIR_NAME_EXPORT_PREFIX`（V1.1.0 已移除，改為固定命名 `DIR_NAME_EXPORT_FIXED`，見 ADR-006）
- `DIR_NAME_MAIN_PREFIX`／`DIR_NAME_IMPORT_PREFIX`／`DIR_NAME_EXPORT_FIXED`／`DIR_NAME_EXPORT_ARCHIVE_PREFIX`（V14 隨日期跳號機制一併廢除，見 ADR-017）
- `MSG_STALE_DIR_REBUILD`（V14 隨固定拓撲重構移除，資料夾遺失改為 `EnsureWorkspaceTopology` 靜默重建，不再彈窗打斷使用者，見 ADR-017）
- `LBL_DONE_MSG_ARCHIVE`／`LBL_DONE_MSG_FOOTER`（V14 隨備份時機點搬離 `Mod_ExportPipeline` 一併移除，見 ADR-017）

---

## 二、工作表資料合約（Data Schema Contract）

### 1. 配置頁：`ConfigSheet`

| 儲存格 | 常數 | 內容 |
|---|---|---|
| A1 | `ADDR_CFG_SESSION_STATUS` | 工作階段狀態橫幅（結案後由系統寫入時間戳記提示） |
| A2 / A3 | `ADDR_CFG_LBL_CHINESE` / `ADDR_CFG_FONT_CHINESE` | 中文字體標籤／下拉選單 |
| B2 / B3 | `ADDR_CFG_LBL_ENGLISH` / `ADDR_CFG_FONT_ENGLISH` | 英數字體標籤／下拉選單 |
| C2 / C3 | `ADDR_CFG_LBL_COLOR` / `ADDR_CFG_FONT_COLOR` | 印記顏色標籤／下拉選單 |
| D2 / D3 | `ADDR_CFG_LBL_SIZE` / `ADDR_CFG_FONT_SIZE` | 字體大小標籤／下拉選單 |
| E2 / E3 | `ADDR_CFG_LBL_TEMPLATE` / `ADDR_CFG_VAL_TEMPLATE` | 排版範本標籤／下拉選單（支援 `{name}`/`{date}`/`{remark}` 標籤替換） |

畫布控制項：`Img_Canvas`（`SH_CANVAS_NAME`）、`Txt_WatermarkAnchor`（`SH_ANCHOR_NAME`）、出血安全框 `Line_BleedZone`（`SH_BLEED_ZONE_NAME`）。

按鈕控制項（五顆，`Mod_UIRenderer.RenderControlPanel` 鑄造）：

| 按鈕物件名稱 | 綁定巨集 | 顯示文字 |
|---|---|---|
| `BTN_PHASE1_FOLDER` | `Macro_Phase1_ImportFolder` | 匯入整個資料夾 |
| `BTN_PHASE1_FILES` | `Macro_Phase1_ImportFiles` | 手選特定圖片 |
| `BTN_PHASE2_EXPORT` | `Macro_Phase2_Export` | 啟動批次浮水印導出 |
| `BTN_FORCE_REBUILD` | `Macro_ForceRebuildUI` | 一鍵重整工作檯 |
| `BTN_CONSOLIDATE_LOGS` | `Macro_ConsolidateAllLogs` | 歷史日誌總歸檔 |

### 2. 暫存頁：`Staging_Images`（嚴格鎖死 9 欄順序，**使用者備註為 C 欄，非 V0 的 E 欄**；I 欄為 V14 新增，見 ADR-017）

| 欄位 | 索引常數 | 內容 |
|---|---|---|
| A | `NUM_STG_COL_STATUS` = 1 | 狀態備註（`STATUS_IMPORT_OK`/`STATUS_IMPORT_ERR`） |
| B | `NUM_STG_COL_NAME` = 2 | 清洗後安全圖片名稱 |
| C | `NUM_STG_COL_REMARK` = 3 | **使用者備註（唯一開放編輯區）** |
| D | `NUM_STG_COL_MOD` = 4 | 原始檔案最後修改時間 |
| E | `NUM_STG_COL_PATH` = 5 | 圖片現在位置（暫存區絕對路徑） |
| F | `NUM_STG_COL_DATE` = 6 | 圖片匯入日期 |
| G | `NUM_STG_COL_REL` = 7 | 圖片匯入位置（V14 起不再供路徑反推使用，僅作記錄，見 ADR-017） |
| H | `NUM_STG_COL_HASH` = 8 | 檔案雜湊（32 位小寫 MD5） |
| I | `NUM_STG_COL_EXPORTED` = 9 | 已匯出狀態（空白＝未匯出，`STATUS_STG_EXPORTED`＝已匯出；`PackageStagingTasks` 主要過濾依據，見 ADR-017） |

欄寬：A=12, B=35, C=45, D=20, E=40, F:H=15, I=15（`Mod_SheetSchemaBuilder.BuildSubStagingSchema`）。

### 3. 成果頁：`LOG_Export_成果`

| 欄位 | 索引常數 | 內容 |
|---|---|---|
| A | `NUM_LOG_COL_STATUS` = 1 | 狀態備註 |
| B | `NUM_LOG_COL_NAME` = 2 | 圖片名稱 |
| C | `NUM_LOG_COL_SRC` = 3 | 匯出工作表名稱（固定 `SHEET_STAGING`） |
| D | `NUM_LOG_COL_PATH` = 4 | 最終成品路徑（超連結，超過 255 字元自動降級錨定至上層資料夾） |
| E | — | 審計安全線文字（`MSG_AUDIT_LINE`，僅表頭列） |

### 4. 總帳頁：`LOG_Master_Archive`

由 `Mod_LogManager.ConsolidateAllLogs` 建立，欄位結構與 `LOG_Export_成果` 的 A~D 欄相同，透過 Range.Value 直注合併所有歷史 `LOG_Export_*` 分頁後將分頁物理刪除。

### 5. 舊目錄總帳頁：`舊目錄LOG_總帳`（`SHEET_OLD_CATALOG_MASTER`，V14 新增，見 ADR-018）

首次觸發「帳本與硬碟狀態不一致」防呆時自動建立，欄位結構完全沿用 `Staging_Images` 的 9 欄表頭，透過 Range.Value 直注技術（不經剪貼簿）附加寫入，每批資料前插入一行時間戳記分隔標記（`MSG_OLD_CATALOG_BATCH_MARK`）。與 `LOG_Master_Archive` 刻意分開存放——兩者欄位語意不相容（一個是「匯出成品路徑」，一個是「使用者備註／匯入位置」），不共用同一張表（見 ADR-018）。

---

## 三、核心模組函式簽章合約（Function Signatures）

### `Mod_MainCoordinator`（按鈕接收層，全部單行委派，不含業務邏輯）
```vba
Public Sub Macro_Phase1_ImportFolder(Optional Dummy As Byte = 0)
Public Sub Macro_Phase1_ImportFiles(Optional Dummy As Byte = 0)
Public Sub Macro_Phase2_Export(Optional Dummy As Byte = 0)
Public Sub Macro_ForceRebuildUI(Optional Dummy As Byte = 0)
Public Sub 重繪操作台()
Public Sub Macro_ConsolidateAllLogs(Optional Dummy As Byte = 0)
```
`Optional Dummy As Byte = 0` 是刻意讓該巨集從 Excel `Alt+F8` 巨集清單隱藏的手法，僅能透過按鈕
`.OnAction` 觸發；`重繪操作台` 是 `Macro_ForceRebuildUI` 的中文別名，刻意不加 `Dummy`、保留於
巨集清單中，作為前台視覺元件損毀時的緊急救援手動入口。詳見 ADR-010。

### `Mod_WorkflowManager`（流程大腦，唯一持有跨模組調度順序知識）
```vba
Public Sub RunPhase1(ByVal isFileMode As Boolean)
Public Sub RunPhase2()
Public Sub RunRebuildUI()
Public Sub RunConsolidateLogs()
```

### `Mod_FileSystem`（V14 全面重寫，見 ADR-017）
```vba
Public Function EnsureWorkspaceTopology(ByVal sysTokens As Object) As String
Public Function GetImportFolderPath(ByVal sysTokens As Object) As String
Public Function GetResultFolderPath(ByVal sysTokens As Object) As String
Public Function GetBackupRootPath(ByVal sysTokens As Object) As String
Public Function IsResumingPreviousRun(ByVal exportFolder As String, ByVal sysTokens As Object) As Boolean
Public Function BackupAndClearFolder(ByVal strSourceFolder As String, ByVal strBackupRoot As String, ByVal strBackupPrefix As String, ByVal objFSO As Object) As String
Public Sub ClearFolderContents(ByVal strFolder As String, ByVal objFSO As Object)
Public Function CleanAndTruncateFilename(ByVal rawFilename As String, ByVal cfg As cls_Settings, ByVal sysTokens As Object) As String
```
`EnsureWorkspaceTopology` 統一確保「作業資料夾／匯入圖片／成果／備份」四層固定拓撲存在，`GetImportFolderPath`／`GetResultFolderPath`／`GetBackupRootPath` 為便捷取用介面，內部皆呼叫前者。`BackupAndClearFolder` 供 Phase 1「開新案」與 Phase 2「結束並重印」共用，兩種情境皆只備份「成果」後清空；`ClearFolderContents` 用於「匯入圖片」的直接清空，不經備份。完整設計脈絡與放棄的替代方案見 ADR-017。

### `Mod_Crypto`
```vba
Public Function CalculateFileHash(ByVal filePath As String, ByVal sysTokens As Object) As String
```
內部實作：`ADODB.Stream`（Type=1 二進位）讀取位元組陣列 → `System.Security.Cryptography.MD5CryptoServiceProvider.ComputeHash_2` → Hex 補零組字串。**不再透過 `certutil` shell-out**（見 ADR-002）。

### `Mod_ImportPipeline`
```vba
Public Sub ExecuteImportPipeline(ByVal colSrcPaths As Collection, ByVal strTargetFolder As String, ByVal objConfig As cls_Settings, ByVal dicTokens As Object)
```
去重範圍涵蓋跨批次歷史資料（V0 僅同批次內去重，見 ADR-007）；去重點名簿透過
`Mod_InteractionAdapter.LoadExistingHashSet()` 取得。**V13 起本模組對 Excel 網格 100% 無知**：
處理結果累積為 `cls_ImportResult` 集合，迴圈結束後一次交給
`Mod_InteractionAdapter.WriteImportResults()` 寫入，不再直接持有或操作 `wsStaging`（見 ADR-016，
與 `Mod_ExportPipeline` 對稱）。

### `Mod_ImageWatermark`
```vba
Public Sub ExecuteWatermarkStream(ByVal entity As cls_ImageEntity, ByVal cfg As cls_Settings, ByVal pctX As Single, ByVal pctY As Single, ByVal rawRemark As String, ByVal destFolder As String, ByVal wsCanvasHost As Worksheet, ByVal sysTokens As Object)
```
本模組 Manifest 明文禁止私自調用 `Application.ScreenUpdating`（螢幕更新主權完全收攏至 `Mod_InteractionAdapter`，見 ADR-005）。

### `Mod_ExportPipeline`
```vba
Public Sub ExecuteExportPipeline(ByVal colTasks As Collection, ByVal objSettings As cls_Settings, ByVal sngPctX As Single, ByVal sngPctY As Single, ByVal flagIsResuming As Boolean, ByVal strExportFolder As String, ByVal wsCanvasHost As Object, ByVal wsLogFrame As Object, ByVal dicTokens As Object)
```
本模組對 Excel 網格 100% 無知，只接受已封裝好的 `cls_ImageEntity` 集合（`colTasks`），由 `Mod_InteractionAdapter.PackageStagingTasks` 負責打包。**V14 起本模組不再負責任何封存／備份邏輯**（原 `ArchiveIncrementalFiles` 已移除），該職責收攏至 `Mod_WorkflowManager` 於「開新案」「結束並重印」兩個有意義的事件邊界主動觸發 `Mod_FileSystem.BackupAndClearFolder`，見 ADR-017。

### `Mod_InteractionAdapter`（全專案唯一的網格讀寫與 Windows 對話框實體）
```vba
Public Function VerifySystemWIA() As Boolean
Public Sub ReleaseUIFocus()
Public Function PickSingleFolder() As String
Public Function PickMultipleFiles() As Collection
Public Sub TogglePerformanceMode(ByVal isFastMode As Boolean)
Public Sub NeutralizeZoomState(ByVal sysTokens As Object)
Public Sub RestoreZoomState(ByVal sysTokens As Object)
Public Sub ClearStagingWorkarea(ByVal wsConfig As Worksheet)
Public Function GetStagingRecordCount() As Long
Public Function PackageStagingTasks(ByVal wsStaging As Worksheet) As Collection
Public Sub LogExportSuccess(ByVal wsLogFrame As Worksheet, ByVal strImageName As String, ByVal strExportFolder As String)
Public Sub LogExportFailure(ByVal wsLogFrame As Worksheet, ByVal strImageName As String, ByVal strErrDesc As String)
Public Sub TriggerAtomicSave(ByVal cntSuccess As Long)
Public Sub FinalizeWorkareaVisuals(ByVal wsCanvasHost As Worksheet)
Public Sub PurgeBlankOrphanSheets()
Public Function LoadExistingHashSet() As Object
Public Sub WriteImportResults(ByVal colResults As Collection)
Public Sub MarkStagingRowAsExported(ByVal wsStaging As Worksheet, ByVal lngRow As Long)
Public Sub ResetAllExportedFlags(ByVal wsStaging As Worksheet)
Public Sub ArchiveStaleStagingToOldCatalogLog(ByVal wsStaging As Worksheet)
Public Function GetExportedRecordCount() As Long
```
`GetExistingStagingFolder`／`DiscoverImportFolder`（V13 版本）於 V14 隨固定拓撲重構移除，見
ADR-017：路徑不再需要從 Staging 表反查，改由 `Mod_FileSystem.GetImportFolderPath`／
`GetResultFolderPath` 直接提供固定路徑。
`LoadExistingHashSet`／`WriteImportResults`（V13 新增，見 ADR-016）：分別承接原本
`Mod_ImportPipeline` 直接讀寫 `wsStaging` 的動作，讓匯入管線比照匯出管線達成對網格 100% 無知。
`PurgeBlankOrphanSheets`（V13 新增，見 ADR-013）：掃描並清除非系統表、無 `###` 保護標記、且真正
空白的工作表，執行前經 `Mod_UI_Messenger.AskQuestion` 確認。偵測到 `ThisWorkbook.ProtectStructure`
為真時明確 `Err.Raise`（比照 `Mod_LogManager.CreateLogSheetFrame` 既有慣例，不自動解鎖代管）。
`MarkStagingRowAsExported`／`ResetAllExportedFlags`（V14 新增，見 ADR-017）：前者供匯出成功後
用列號精準回寫單列「已匯出」狀態；後者供 Phase 2「結束並重印」分支整批重置，讓下一次
`PackageStagingTasks` 重新納入全部照片。
`ArchiveStaleStagingToOldCatalogLog`／`GetExportedRecordCount`（V14 新增，見 ADR-018）：前者
偵測到匯入端帳本與硬碟不一致時，把 Staging 表整批直注附加進「舊目錄LOG_總帳」單一固定分頁後
清空；後者供 Phase 2 判斷成果端帳本與硬碟是否一致。

### `Mod_SheetSchemaBuilder`
```vba
Public Sub BuildSystemSchemas()
```

### `Mod_UIRenderer`
```vba
Public Sub RenderUIElements(Optional ByVal forceResetLayout As Boolean = False)
```
`forceResetLayout`（V13 新增，見 ADR-011／ADR-012）：預設 `False`，畫布／出血框／浮水印錨點
「存在則沿用」，保留使用者拖曳過的位置，供 `RunPhase1`（匯入）呼叫；傳入 `True` 則強制刪除三者
回歸預設座標，僅供 `RunRebuildUI`（一鍵重整工作檯）呼叫，作為恢復原廠設定的救援動作。

本模組同時負責渲染 `Footer_LicenseNotice`（版權聲明頁尾，見 ADR-015）：低調小字文字方塊，橫跨
畫布左緣至按鈕面板右緣、靠左對齊，內容為精簡版授權摘要，與 `ThisWorkbook.cls` 內完整法律全文
互為表裡（此處摘要、`ThisWorkbook` 全文，不重複維護同一份內容）。歸入 `ClearAllExistingShapes`
每次刪除重建清單，不比照畫布／錨點的位置保留邏輯，因其不承載使用者自訂狀態。

### `Mod_LogManager`
```vba
Public Function CreateLogSheetFrame(ByVal isResuming As Boolean) As Worksheet
Public Sub ArchiveLogSheet(ByVal suffixLabel As String)
Public Sub ConsolidateAllLogs()
```

### `View_CanvasManager`
```vba
Public Sub GetWatermarkPercentages(ByRef outPctX As Single, ByRef outPctY As Single)
```

### `Mod_UI_Messenger`
```vba
Public Function AskQuestion(ByVal promptText As String, ByVal titleText As String) As VbMsgBoxResult
Public Function AskQuestionCancel(ByVal promptText As String, ByVal titleText As String) As VbMsgBoxResult
Public Function AskCritical(ByVal promptText As String, ByVal titleText As String) As VbMsgBoxResult
Public Sub ShowInfo(ByVal promptText As String, ByVal titleText As String)
Public Sub ShowWarning(ByVal promptText As String, ByVal titleText As String)
Public Sub ShowError(ByVal promptText As String, ByVal titleText As String)
```

---

## 四、共用工具箱（sysTokens）合約

V8 重構後新增的跨模組資源傳遞機制：`Mod_WorkflowManager` 在流程開頭建立一個 `Scripting.Dictionary`（`dicSysTokens`），統一裝入以下單例物件，向下以參數形式流式注入給所有下層模組，取代各模組各自 `CreateObject`：

| Key | ProgID 常數 | 用途 |
|---|---|---|
| `"FSO"` | `PROGID_FSO` = `Scripting.FileSystemObject` | 檔案系統操作 |
| `"STREAM"` | `PROGID_STREAM` = `ADODB.Stream` | MD5 二進位讀取 |
| `"MD5"` | `PROGID_MD5` = `System.Security.Cryptography.MD5CryptoServiceProvider` | 雜湊運算 |
| `"OLD_ZOOM"` | （執行期動態寫入，非固定 ProgID） | 暫存縮放比，供 `RestoreZoomState` 讀回 |

**合約規則**：`Mod_FileSystem`／`Mod_Crypto`／`Mod_ImageWatermark`／`Mod_ImportPipeline` 等下層模組一律透過 `sysTokens` 取用這些物件，嚴禁在下層模組內部私自 `CreateObject`。

---

## 五、資料物件模型合約

### `cls_ImageEntity.InitializeEntity`（8 參數，一次鎖死唯讀；V14 新增第 8 參數，見 ADR-017）
```vba
Public Sub InitializeEntity( _
    ByVal imgName As String, ByVal lastMod As String, ByVal curPath As String, _
    ByVal impDate As String, ByVal impRelPath As String, _
    ByVal UserRemark As String, ByVal hashVal As String, ByVal stagingRow As Long)
```
**與 V0 的差異**：新增 `UserRemark` 參數（V0 為 6 參數，無備註）；且 V12 移除了 V0 原本會清除字串「中間空白」的 `Replace(x, " ", "")` 邏輯，只保留 `Trim`（去頭尾空白）。這是刻意的行為變更，原因待你確認後補入 ADR（合理推測：Windows 路徑與備註合法含有中間空白，例如「我的 圖片」、「台南 安定區」，強制清除中間空白反而會製造路徑或備註內容錯誤）。

**V14 新增 `stagingRow`**（對應唯讀屬性 `StagingRowNumber`，見 ADR-017）：這筆資料在 Staging 表的實際列號，供匯出成功後 `Mod_InteractionAdapter.MarkStagingRowAsExported` 精準回寫「已匯出」狀態，取代原本靠檔名或雜湊反查列號的脆弱做法。內部新增防禦：`stagingRow < NUM_ROW_DATA_START` 直接 `Err.Raise`，避免呼叫端傳錯值導致回寫到表頭或不存在的列。

### `cls_Settings`（V0 的 `cls_Config` 更名而來，見第十八章命名沿革）
屬性：`FontNameChinese`、`FontNameEnglish`（V0 無此欄，V12 新增中英雙軌字體）、`FontColorDefault`、`FontSizeDefault`、`BgColorDefault`、`BgTransparency`、`LineVisible`、`MaxFilenameLength`、`ForbiddenChars`。

`ForbiddenChars`：`\ / : * ? " < > | [ ] #`，共 12 字元，`CleanAndTruncateFilename` 逐一替換為底線。此處登記具體字元清單，屬於本文件對 `cls_Settings.cls` 內部陣列字面值的唯一對照複本；若日後黑名單異動，兩處須同步更新。`#` 新增原因見 ADR-017 附記。

### `cls_ImportResult`（V13 新增，見 ADR-016，與 `cls_ImageEntity` 對稱設計）
```vba
Public Sub InitializeResult( _
    ByVal isSuccess As Boolean, ByVal imageName As String, ByVal lastMod As String, _
    ByVal curPath As String, ByVal impDate As String, ByVal impRelPath As String, _
    ByVal userRemark As String, ByVal hashVal As String, ByVal failReason As String)
```
承載單一檔案匯入處理結果（成功／失敗兩種狀態），供 `Mod_ImportPipeline` 累積成 `Collection` 後交給
`Mod_InteractionAdapter.WriteImportResults()` 統一寫入，使匯入管線得以對 Excel 網格 100% 無知。

---

## 六、模組清單註冊（V13 最終定案，共 19 支有效模組，全部 VERSION 1.0.0）

> **異動說明**：`ICommand.cls` 已於 V13 正式撤除（見 ADR-009）。V2~V13 期間該介面
> 定義 `Public Sub Execute(ByVal target As Object)` 標準接口，但連續 11 個子版本
> 從未有任何 `cmd*` 類別實作它。經 Shih 裁示確認本專案規模與流程複雜度（四個固定
> 線性流程，無動態組合／替換／排程需求）不符合 AVS v1.1 第五節套用 Command Pattern
> 之判準，不再列入模組清單，不需以任何形式保留。

| 模組名稱 | 型態 | 職責 |
|---|---|---|
| `Mod_StringConstants` | 標準模組 | 中央常數 SSOT |
| `cls_Settings` | 類別 | 參數配置中心 |
| `cls_ImageEntity` | 類別 | 唯讀資料物件模型 |
| `cls_ImportResult` | 類別 | 匯入結果資料模型，與 `cls_ImageEntity` 對稱（見 ADR-016） |
| `cls_EventStore` | 類別 | 記憶體執行軌跡記錄器 |
| `cls_DAGEngine` | 類別 | 系統級生命週期守衛（非真正 DAG，見 ADR-001） |
| `View_CanvasManager` | 標準模組 | 幾何百分比換算 |
| `Mod_ImageWatermark` | 標準模組 | WIA 降維＋Chart 壓印算力引擎 |
| `Mod_WorkflowManager` | 標準模組 | 流程大腦（見 ADR-004） |
| `Mod_UI_Messenger` | 標準模組 | 唯一 MsgBox 出口 |
| `Mod_MainCoordinator` | 標準模組 | 按鈕接收層（見 ADR-004） |
| `Mod_FileSystem` | 標準模組 | 硬碟與路徑引擎 |
| `Mod_ExportPipeline` | 標準模組 | 匯出批次壓印管線 |
| `Mod_SheetSchemaBuilder` | 標準模組 | 工作表結構地基建構師 |
| `Mod_UIRenderer` | 標準模組 | 前台視覺渲染器 |
| `Mod_Crypto` | 標準模組 | MD5 雜湊引擎（見 ADR-002） |
| `Mod_ImportPipeline` | 標準模組 | 匯入批次清洗管線 |
| `Mod_InteractionAdapter` | 標準模組 | 全專案唯一網格讀寫／對話框實體 |
| `Mod_LogManager` | 標準模組 | 日誌建立／歸檔／總帳合併 |

**V0 已不存在的舊模組**：`Mod_ActionExecutor`、`Mod_EnvironmentSetup`、`Mod_CreateLogSheet`（三支已於 V8 拆解為上述 12 支專職模組，見 ADR-003）。
