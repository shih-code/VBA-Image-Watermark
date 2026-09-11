# ACDS_ADR_合輯.md — 架構決策紀錄合輯

**專案名稱**：Excel 驅動型圖片兩階段批次浮水印處理系統（純模組化強固版）
**對應原始碼版本**：V17（ADR-001～016 對應 V12/V13 定案版；ADR-017、ADR-018 對應 V14→V17 過渡期間之重構，詳見內文）
**合併說明**：ACDS v2.4 第十五章原規定 ADR 應各自獨立存放於 `ACDS_ADR/ADR-{編號}.md`，
本檔案是應使用者要求、為避免小型專案下 ADR 檔案數量分散導致管理與蒐集困難，改採
單一檔案彙整而成的例外處理。**本檔案是這批 ADR 內容目前的唯一存放位置**（原本拆開
的 7 個獨立檔案已停用，避免同一份決策內容出現兩個事實落點）。若之後要正式把「小型
專案允許合併 ADR」這條例外收進 ACDS 本體，需另案處理，本檔案本身不構成 ACDS 修訂。

## 目錄

| 編號 | 標題 | 狀態 |
|---|---|---|
| ADR-001 | cls_DAGEngine 語意正名：系統級生命週期守衛，非動態依賴排程引擎 | Accepted |
| ADR-002 | MD5 雜湊計算由 certutil Shell Out 改為 ADODB.Stream + .NET MD5CryptoServiceProvider | Accepted |
| ADR-003 | 上帝物件肢解重構（V8「切除危機」）：變數一物多用問題的根因與修復 | Accepted |
| ADR-004 | Mod_MainCoordinator／Mod_WorkflowManager 職責分離：按鈕接收層與流程大腦分家 | Accepted |
| ADR-005 | Application.ScreenUpdating／Calculation 主權集中至流程大腦層 | Accepted |
| ADR-006 | 匯出目的地撞名處理：由「直接覆寫／單純改名」改為「版本化退位備份」 | Accepted |
| ADR-007 | 圖片去重範圍由「僅同批次內」擴大為「涵蓋 Staging 既有歷史列」 | Accepted |
| ADR-008 | V13 模組識別資訊收斂修正：Stability 詞彙統一、Private Const MOD_NAME 補齊、AUTHOR 欄位收攏至 ThisWorkbook | Accepted |
| ADR-009 | 撤除 ICommand.cls：過早抽象化（Premature Abstraction）案例 | Accepted |
| ADR-010 | 按鈕入口巨集雙軌可見度設計：Optional Dummy 隱藏機制 + 唯一功能雙入口例外 | Accepted |
| ADR-011 | RenderCanvasAndAnchor 位置保留修復：借鑑 V1_CL 分支經驗，畫布與錨點改為「存在則沿用」 | Accepted |
| ADR-012 | RenderCanvasAndAnchor 位置保留策略再區分：匯入沿用、重整UI強制重置（延伸 ADR-011） | Accepted |
| ADR-013 | 一鍵重整工作檯新增清除空白頁面功能：### 保護標記機制 | Accepted |
| ADR-014 | RunRebuildUI 修復 wsCurrent 懸空物件錯誤：PurgeBlankOrphanSheets 與既有畫面還原邏輯的組合型缺陷 | Accepted |
| ADR-015 | UI 控制面板新增版權聲明頁尾：修正「僅存放於 ThisWorkbook」的先前建議 | Accepted |
| ADR-016 | Mod_ImportPipeline 完全解耦：比照 Mod_ExportPipeline 達成對 Excel 網格 100% 無知 | Accepted |
| ADR-017 | 廢除日期跳號機制，改採固定工作資料夾拓撲＋事件式整批備份，根治匯出重複檔案 | Accepted |
| ADR-018 | 帳本與硬碟狀態不一致防呆：匯入端幽靈紀錄封存、成果端已匯出旗標重置 | Accepted |

---

============================================================
【ADR-001】cls_DAGEngine 語意正名：系統級生命週期守衛，非動態依賴排程引擎
- 日期：2026-07（V0 即存在，V12 沿用未變更，本 ADR 為回溯定案）
- 狀態：Accepted
- 背景：
  類別名稱 `cls_DAGEngine` 中的「DAG」（Directed Acyclic Graph，有向無環圖）一詞，
  在一般軟體工程語境下暗示具備「節點依賴排程、拓撲排序、動態組合執行順序」的能力。
  但實際檢視 V0 至 V12 全部版本的原始碼，`cls_DAGEngine` 的公開方法始終只有
  `StartPipeline`／`LogStep`／`TerminatePipeline`／`DispatchMeltdown`／`IsRunning`
  五個成員，行為是單一管線的線性生命週期控管（啟動 → 逐步紀錄 → 正常終止或熔斷），
  不具備任何節點依賴圖、拓撲排序或動態排程能力。這與 AVS v1.1 第三節定義的
  「系統級生命週期守衛」（Pipeline 角色）完全吻合，而非文獻上一般理解的 DAG Engine。
- 決策：
  本專案正式定調 `cls_DAGEngine` 之語意為「系統級生命週期守衛」，對應 AVS v1.1
  第三節的 Pipeline 角色。沿用舊類別名稱不更動（理由見下方放棄的替代方案），但本
  ADR 作為權威解釋來源：未來任何人（含 AI）閱讀此類別時，不得將其誤認為具備動態
  依賴排程能力；若本專案未來真的出現「功能之間需要動態組合、彼此獨立替換、或依賴
  關係需要排程」的需求，須依 AVS v1.1 第五節判準另行建立新元件（`ICommand` 的
  `cmd*` 具體實作 + `Mod_CommandFactory`），不得擴充 `cls_DAGEngine` 承擔新語意。
- 放棄的替代方案：
  (a) 更名為 `cls_PipelineGuard` 等貼合實際行為的名稱——放棄理由：V12 已有 19 支
      模組、跨越 12 個版本迭代皆使用此名稱，且三份既有規格文件（`合約.txt`、
      `拓樸_思考型.txt`、`規格書.txt`）皆已將其列入模組清單並附上「流程調度引擎」
      的職責描述，重構成本（含所有呼叫端與文件同步修改）大於語意精確帶來的效益，
      列為技術債，待下次架構性大重構（Major 版本躍升）時一併處理，不單獨立案。
  (b) 真的補上依賴圖排程能力，讓名稱名符其實——放棄理由：本專案流程固定為
      Phase1（匯入）→ Phase2（匯出）兩步線性執行，加上 Rebuild UI／Consolidate
      Logs 兩個獨立維護動作，全部都不需要動態組合或替換，違反 YAGNI 原則。
- 影響：
  未來任何稽核者（AI 或 Shih 本人）在檢視此類別時，應以本 ADR 定義的語意為準，
  不應被類別名稱誤導去猜測或新增依賴排程相關方法。第十八章專案規格書「Domain Map」
  一欄需明確標註「Pipeline（生命週期守衛）」而非「DAG（依賴排程）」，避免文件層級
  再次產生誤導性描述。
============================================================


============================================================
【ADR-002】MD5 雜湊計算由 certutil Shell Out 改為 ADODB.Stream + .NET MD5CryptoServiceProvider
- 日期：2026-07（V0→V8 之間的重構期間，確切子版本不可考，V8 起穩定沿用至 V12）
- 狀態：Accepted
- 背景：
  V0 版本的 `CalculateFileHash` 透過 `WScript.Shell.Exec` 呼叫 Windows 內建
  `certutil -hashfile <path> MD5`，輪詢子行程 Status 並解析 StdOut 文字輸出取得
  雜湊值。這個做法有三個實際痛點：
  1. 每次呼叫都會短暫彈出 CMD 黑視窗（即使背景執行仍可能被使用者注意到），
     不符合「簡單、無腦、安全、美觀」的產品標準中「美觀」一項。
  2. 效能成本高：每張圖片都要開一次子行程，大量圖片批次處理時累積延遲明顯。
  3. 依賴外部命令列工具存在，若企業 IT 政策封鎖 `certutil.exe`（部分資安敏感環境
     會限制此類系統管理工具），會直接導致雜湊功能失效。
- 決策：
  改用 `ADODB.Stream`（Type=1 二進位模式）將檔案完整讀入記憶體位元組陣列，再交給
  `System.Security.Cryptography.MD5CryptoServiceProvider`（透過 VBA 晚期繫結呼叫
  .NET CLR 物件）的 `ComputeHash_2` 方法運算，最後以 `Hex()` 函式逐位元組轉換並
  補零組成 32 碼小寫字串。全程在記憶體中完成，不開子行程、不產生 CMD 視窗。
- 放棄的替代方案：
  (a) 維持 certutil shell-out——放棄理由：見背景三點痛點，尤其黑視窗閃爍問題
      在多次使用者回饋中被明確提及（見 `合約.txt`「0% CMD 閃爍」的強調用詞）。
  (b) VBA Declare 呼叫 Win32 CryptoAPI（`CryptAcquireContext`/`CryptHashData`）
      ——放棄理由：64 位元 Office 的 Declare 語法複雜、32/64 位元相容性風險高，
      對於「讀 Python 比讀 VBA 順」的非工程背景維護者而言除錯困難，且需要額外
      處理 Handle 生命週期釋放，增加 Cleanup Label 複雜度。
  (c) 第三方 DLL 或外部 COM 元件——放棄理由：需要額外部署或註冊，違反「開啟巨集
      即可用」的部署簡便性目標。
- 影響：
  本函式依賴 `MD5CryptoServiceProvider` 這個 .NET 元件透過 COM 晚期繫結可被建立
  （多數 Windows 10/11 內建 .NET Framework 環境下皆可正常運作），若目標機器完全
  沒有 .NET Framework（極罕見），會在 `CreateObject` 階段直接失敗並被
  `HashStreamErrorHandler` 攔截、拋出 5405 錯誤，符合 Fail-Fast 精神，但屬於環境
  依賴變更，應寫入第十八章「Fail-Fast Focus」與使用說明書「系統使用限制」章節。
  此決策與 ADR-006（sysTokens 共用工具箱）互為前提：`ADODB.Stream` 與
  `MD5CryptoServiceProvider` 物件透過 `Mod_WorkflowManager` 建立一次、以
  `sysTokens` 字典向下注入複用，避免每張圖片重複建立 COM 物件的額外開銷。
============================================================


============================================================
【ADR-003】上帝物件肢解重構（V8「切除危機」）：變數一物多用問題的根因與修復
- 日期：2026-07（V7 → V8 過渡）
- 狀態：Accepted
- 背景：
  V2 至 V7（連續 6 個子版本）期間，系統核心運算全部集中在單一模組
  `Mod_ActionExecutor`（V0 起即存在，V7 版本標記已累積至 3.5.0），身兼防撞資料夾
  建立、檔名清洗、MD5 雜湊計算、WIA 影像壓印、Excel Chart 渲染、批次迴圈調度數項
  完全不同的職責，違反 ACDS 第二章 SRP 原則。與此同時，`Mod_EnvironmentSetup`
  也身兼工作表結構建置與前台視覺渲染兩種職責。
  這個時期的根本原因是：**當時的 ACDS 本體與本專案的第十八章規格書尚未把「模組
  該怎麼切」「變數的生命週期與作用範圍該怎麼界定」寫清楚**，導致 AI（Gemini）在
  持續疊代過程中，為了少改動既有函式簽章，傾向讓同一個變數在不同執行階段承擔不同
  語意（例如同一個路徑變數在 Phase 1 語境代表「匯入來源」、在 Phase 2 語境被重新
  賦值代表「匯出目的地」），或讓同一個常數同時扮演「顯示標籤」與「內部路徑判斷
  依據」兩種角色。這類「變數一物多用」問題不會立即造成程式崩潰（不會被 Fail-Fast
  攔截），但會在需求變動時難以追蹤某個變數當下究竟代表什麼，是典型的隱性技術債。
  本專案作為「用來檢驗 ACDS 本身是否完善」的早期練習專案，這正是它被拿來驗證、
  進而回饋修正 ACDS 本體的實際案例之一。
- 決策：
  V8（檔名標記「切除危機」）進行一次性大重構：將 `Mod_ActionExecutor` 依職責拆解
  為 `Mod_FileSystem`（路徑與檔名）、`Mod_Crypto`（雜湊運算）、
  `Mod_ImageWatermark`（影像壓印）三支專職模組；`Mod_EnvironmentSetup` 拆解為
  `Mod_SheetSchemaBuilder`（工作表結構）與 `Mod_UIRenderer`（視覺渲染）兩支專職
  模組；新增 `Mod_WorkflowManager`（流程大腦，見 ADR-004）與
  `Mod_InteractionAdapter`（網格讀寫與對話框轉接器）承接原本散落各處的協調邏輯；
  `Mod_CreateLogSheet` 更名擴充為 `Mod_LogManager`。模組數量由 14~15 支一口氣增至
  21 支。重構同時在新拆出的模組 Manifest 中明文寫入
  `FORBIDDEN: 嚴禁碰觸表格網格合約與欄位值；嚴禁任何變數一物多用。`
  作為往後的自我約束條款，防止同類問題再次發生。
- 放棄的替代方案：
  (a) 局部修補：只修正當下發現的變數混用個案，不做整體模組拆分——放棄理由：
      問題根因是職責邊界不清，局部修補只會不斷冒出新的變數混用個案，治標不治本。
  (b) 保留單一大模組但補齊註解與命名規範——放棄理由：ACDS 第八章「超出安全線」
      熔斷條件本就規定單一積木超過 100~150 行或包含兩種以上職責應停止繼續累加，
      `Mod_ActionExecutor` 到 V7 時已遠超此規模，補註解無法解決架構層級的耦合問題。
- 影響：
  這次重構是本專案至今唯一一次 Major 等級的破壞性架構調整，V8 之後所有版本
  （V9~V12）的模組邊界基本沿用此次拆分結果，僅有內部函式簽章微調（詳見
  ContractRegistry「已移除的舊版常數」與「V0 已不存在的舊模組」章節）。
  此案例應收錄進 ACDS 本體附錄 C 案例庫，作為「規則本身留有空隙導致系統性技術債
  累積」的第二個實例，與既有 C-2「識別資訊漂移」案例性質相近但成因略有不同：
  C-2 是命名前綴規則有空隙，本案例是「模組職責邊界與變數生命週期界定」規則有空隙，
  建議 ACDS 下次改版時納入附錄 C 作為 C-5 案例，暫定案例代號留待 Shih 裁示。
============================================================


============================================================
【ADR-004】Mod_MainCoordinator／Mod_WorkflowManager 職責分離：按鈕接收層與流程大腦分家
- 日期：2026-07（V8 重構的一部分，隨 ADR-003 一併定案）
- 狀態：Accepted
- 背景：
  V0~V7 期間，`Mod_MainCoordinator` 同時扮演「Excel 按鈕點擊入口」與「流程順序
  調度大腦」兩種角色，`Macro_Phase1_Import`／`Macro_Phase2_Export` 兩個 Sub 內部
  直接寫了完整的 DAG 啟動、環境檢查、對話框詢問、迴圈調度等全部邏輯，導致這兩支
  程序長度膨脹（V7 版本已超過 150 行），且因為是 `Public Sub` 會出現在 Excel
  的巨集清單（Alt+F8）中，操作人員理論上可以繞過正常的按鈕點擊流程，在任何時候
  手動執行這些高風險程序，跳過大腦原本該做的環境檢查與確認彈窗。
- 決策：
  拆分為兩層：`Mod_MainCoordinator` 只保留五個「一行委派」的 `Public Sub`
  （對應五顆按鈕），本身不含任何 If／迴圈／環境檢查邏輯，純粹作為 Excel 巨集清單
  的最外層入口；真正的流程順序知識與跨模組調度邏輯，全部移入新建立的
  `Mod_WorkflowManager`，並將其宣告為 `Option Private Module`（不出現在巨集清單
  中，操作人員無法繞過大腦直接執行內部流程片段）。
- 放棄的替代方案：
  (a) 維持單一模組，僅靠命名慣例（如加上底線前綴）暗示「不要直接執行」——放棄
      理由：命名慣例無法從語言層級強制禁止，操作人員仍可能誤觸；`Option Private
      Module` 是 VBA 語言本身提供的機制保證，比命名約定更可靠。
  (b) 把所有邏輯都收進 `Mod_MainCoordinator` 但改為 `Private Sub`——放棄理由：
      如此一來 Excel 按鈕就無法綁定 `OnAction` 到這些程序（VBA 的
      `Shape.OnAction` 只能指向 `Public` 程序），與「按鈕點擊必須觸發」的硬性
      需求衝突，因此必須拆成兩層而非合併成一層全私有。
- 影響：
  這個分層此後成為本專案新增功能的固定模式：新增一顆按鈕，一律先在
  `Mod_MainCoordinator` 加一行委派，真正邏輯寫在 `Mod_WorkflowManager`
  （或其委派的下層管線模組）。V9~V12 新增的「一鍵重整工作檯」「歷史日誌總歸檔」
  兩個功能皆遵循此模式，未再出現大腦邏輯混入按鈕入口層的情況，可視為此決策已通過
  後續三次疊代的實測驗證。
============================================================


============================================================
【ADR-005】Application.ScreenUpdating／Calculation 主權集中至流程大腦層
- 日期：2026-07（V8 重構期間定案，V9~V12 持續強化，Mod_ImageWatermark 3.4.0
  版本記載為「黑幕清零版：徹底拔除所有 ScreenUpdating 越權干涉」）
- 狀態：Accepted
- 背景：
  V0~V7 期間，`Application.ScreenUpdating`／`Application.Calculation` 的開關散落
  在多個模組各自呼叫（例如影像壓印函式自己開關一次、匯入清洗函式又自己開關一次），
  當多層函式呼叫巢狀發生時，容易出現「內層函式提早把畫面重新打開，外層函式還沒執行
  完」的競態問題，導致批次處理中途畫面閃爍、甚至偶發性的圖形物件錯位或半殘留。
- 決策：
  螢幕更新與計算模式的開關主權，唯一收攏在 `Mod_WorkflowManager`（流程大腦），
  透過統一委派 `Mod_InteractionAdapter.TogglePerformanceMode(isFastMode)` 執行；
  下層算力引擎模組（`Mod_ImageWatermark`、`Mod_ExportPipeline`、`Mod_FileSystem`
  等）的 Manifest 明文加入 `FORBIDDEN` 條款，禁止私自調用
  `Application.ScreenUpdating`。少數需要「彈窗前必須先確保畫面能重繪」的例外情境
  （例如 Dry Run 確認彈窗、結案完工彈窗），由大腦在呼叫 `Mod_UI_Messenger` 前
  明確插入 `Application.ScreenUpdating = True` + `DoEvents` 兩行，並在原始碼中
  以「【大腦收攏對話框呼吸權】」「【破圖終結防線】」等註解標明這是刻意的例外，
  不是遺漏的黑幕控制。
- 放棄的替代方案：
  (a) 每個函式各自負責自己進出場時的畫面狀態（各自 Push/Pop）——放棄理由：VBA
      沒有原生的狀態堆疊機制，巢狀呼叫時容易漏寫或重複寫，且無法從語言層級強制
      檢查配對是否正確，實務上就是 V0~V7 發生問題的根源。
  (b) 完全不管理畫面更新，接受閃爍——放棄理由：不符合「簡單、無腦、安全、美觀」
      中「美觀」與使用者體驗要求，且大量圖片批次處理時的畫面重繪會拖慢實際效能。
- 影響：
  未來任何新增的下層算力模組，若發現自己需要控制 `ScreenUpdating`，這本身就是
  一個訊號：代表該邏輯可能該搬到 `Mod_WorkflowManager` 或至少透過
  `Mod_InteractionAdapter` 統一管道處理，不應該自行開關。稽核時可用「grep 整個
  專案原始碼搜尋 `ScreenUpdating`」快速檢查是否有下層模組違規直接操作，如果搜尋
  結果只出現在 `Mod_WorkflowManager`、`Mod_InteractionAdapter`、`Mod_UIRenderer`、
  `Mod_LogManager.ConsolidateAllLogs` 這幾支模組內，即視為合規。
============================================================


============================================================
【ADR-006】匯出目的地撞名處理：由「直接覆寫／單純改名」改為「版本化退位備份」
- 日期：2026-07（V1.1.0 先有固定命名斷點夾雛形，V8~V9 期間補上版本化改名邏輯）
- 狀態：Accepted
- 背景：
  V0 版本 `ExecuteWatermarkStream` 若發現目的地已有同名正式檔案，做法是先
  `fso.DeleteFile` 刪除舊檔，再把 `.tmp` 暫存檔改名頂上——這代表舊的成品在被
  新成品覆蓋前，有一個「舊檔已被實體刪除、新檔案還沒改名完成」的極短暫真空期，
  若這個瞬間發生中斷（斷電、Excel 崩潰），會同時失去新舊兩份成品，違反 ACDS 第
  三章 P0 第 2 條「操作順序：複製優先於刪除或移動」與第 3 條「覆寫規則：不可恢復
  的覆寫才是紅線」。
  同時，「匯出成果_斷點」改為固定命名（非跳號）本身也是一個獨立決策（V1.1.0
  時期已定案）：因為 Phase 2 需要靠「這個固定路徑是否已存在且內部有實體檔案」來
  判斷是否為斷點續跑，若沿用 Phase 1 的跳號防撞邏輯，每次重新執行 Phase 2 都會
  產生一個新的空資料夾，系統將永遠無法判斷「這是接續」還是「這是全新」。
- 決策：
  `Mod_ImageWatermark.ExecuteWatermarkStream` 若偵測到目的地已有同名正式檔案，
  改為呼叫內部私有程序 `TryRenameWithRetry`，先將舊檔案改名為
  `{原檔名}_舊版_{兩位數跳號}.{副檔名}`（例如 `IMG_001_舊版_01.jpg`），確認改名
  成功後，才將新產出的 `.tmp` 暫存檔改名為正式檔名。整個過程沒有任何一個時間點
  同時「新檔案不存在」且「舊檔案已消失」——舊檔案要嘛还在原名、要嘛已改名為
  `_舊版_NN` 版本但依然實體存在於硬碟上，從未被真正刪除。
  改名操作本身透過 `TryRenameWithRetry` 包一層 Error 70/75/58（檔案鎖定）重試
  防禦，具備最多 3 次、每次間隔 0.1 秒的延時重試（含跨日 Timer Rollover 防護），
  三次皆失敗才對外拋出 `STATUS_EXPORT_ERR_LOCKED` 熔斷。
- 放棄的替代方案：
  (a) 沿用 V0 的「先刪除再改名」——放棄理由：見背景，違反 P0 第 2、3 條。
  (b) 完全禁止覆寫，撞名時直接跳過該張圖片並警告使用者——放棄理由：斷點續跑模式
      下「重新覆寫」是使用者主動選擇的正常操作路徑之一（`MSG_EXPORT_MODE_PROMPT`
      提供「否 (No)：啟動重新覆寫模式」選項），完全禁止會讓這個選項失去意義。
  (c) 匯出資料夾維持跳號機制，另外用獨立 state 檔（如 .json）追蹤斷點進度——
      放棄理由：多一個需要與實體資料夾內容維持同步的持久化檔案，若使用者手動
      搬移或刪除資料夾，state 檔容易與實際檔案狀態失真，兩者互為第二個事實落點。
- 影響：
  匯出資料夾內看到的 `_舊版_NN` 檔案，是系統刻意保留的版本歷史，不是殘留垃圾，
  第十八章使用說明應向使用者說明此行為，避免使用者誤刪或誤以為是程式錯誤產生的
  多餘檔案。這個「跳號防撞」（Phase 1 匯入夾）與「固定命名＋存在性斷點判斷＋版本
  化備份」（Phase 2 匯出夾）是刻意並存的兩種不同機制，分別對應「同一批資料互不
  覆蓋」與「同一任務接續辨識＋安全覆寫」兩種不同需求，未來修改任一套資料夾命名
  邏輯前，須先確認其對應的是哪一種需求。
============================================================


============================================================
【ADR-007】圖片去重範圍由「僅同批次內」擴大為「涵蓋 Staging 既有歷史列」
- 日期：2026-07（V8~V9 期間，隨大重構一併補上）
- 狀態：Accepted
- 背景：
  V0 版本 `BatchImportAndCleanProcess` 的去重邏輯（`dictSeenHash`）只在單次匯入
  執行的當下建立，只能防止「同一次點擊匯入按鈕、選取的來源資料夾內部」有重複圖片；
  如果使用者分兩次匯入（例如今天匯入一批，明天又指到同一個或有重疊檔案的來源
  資料夾再匯入一次），V0 的邏輯完全無法偵測到這是重複圖片，會被當作全新資料
  重複寫入 Staging 看板。這與系統標榜的「重跑時如果內容重複會自動跳過」（見
  `photo_watermark_manual.md`、`正式說明書.md`）實際不符，是文件宣稱與早期實作
  之間的落差之一。
- 決策：
  `Mod_ImportPipeline.ExecuteImportPipeline` 內新增私有程序 `LoadExistingHashes`，
  在正式開始處理新檔案之前，先掃描 Staging_Images 目前已存在的所有列
  （`NUM_ROW_DATA_START` 到目前最後一列），把每一列非空、非
  `HASH_FAILED_MARK` 的雜湊值預先加入 `dicSeenHash` 點名簿，再才開始逐一比對
  新檔案的雜湊值。如此一來，去重範圍從「僅同批次」擴大為「同批次 + 目前 Staging
  看板上所有歷史列」。
- 放棄的替代方案：
  (a) 維持僅同批次去重，把「跨批次可能重複」寫進使用說明書當作已知限制——
      放棄理由：這與既有對外文件的宣稱不符，且技術上補強成本不高（多一次表格
      掃描，對典型使用量級的效能影響可忽略），沒有理由維持這個限制。
  (b) 去重範圍再往外擴大到「跨 Session（即使 Staging 已被清空重來也要記得）」
      ——放棄理由：`Mod_InteractionAdapter.ClearStagingWorkarea` 被呼叫（使用者
      選擇「否：清空重來」）時，Staging 歷史列本來就會被物理清除，代表使用者
      明確表達「這是全新的一輪工作，不延續舊資料」的意圖，此時就不應該再去比對
      已經被使用者主動清空的舊資料，維持「以當前 Staging 看板上實際存在的資料
      為準」是與使用者意圖一致的設計，不需要額外做跨 Session 持久化的雜湊資料庫。
- 影響：
  這個決策隱含一個前提：Staging_Images 的雜湊值欄位（H 欄）本身就是去重比對的
  唯一依據，若使用者手動修改或刪除 H 欄內容，會影響去重判斷的正確性——這點應在
  `photo_watermark_manual.md`「應避免的操作行為」章節中明確提及「嚴禁手動更動
  結構配置」時一併涵蓋雜湊欄位，目前該文件只提到「欄位標題與橫向順序」，未明確
  提及欄位內容本身也不應被手動更動。
============================================================


============================================================
【ADR-008】V13 模組識別資訊收斂修正：Stability 詞彙統一、Private Const MOD_NAME 補齊、AUTHOR 欄位收攏至 ThisWorkbook
- 日期：2026-07（V12 → V13 過渡）
- 狀態：Accepted
- 背景：
  V12 稽核（詳見對話紀錄與 ADR-001~007）發現三項第七章「模組識別資訊單一化」相關
  的殘留缺口：
  1. `cls_DAGEngine`／`cls_ImageEntity` 兩支類別模組的 `Err.Raise` 內重複手打類別
     名稱字面值（如 `Err.Raise 513, "cls_DAGEngine", ...`），未如其餘 `Mod_*` 標準
     模組補上 `Private Const MOD_NAME` 並改為引用；`cls_ImageEntity` 另外殘留一行
     開發過程中忘記刪除的 AI 提示指令註解（`【請在 cls_ImageEntity 類別模組內追加
     此行合約】`）。
  2. 全專案 12 支 `Mod_*` 標準模組的 `Stability` 標記寫的是 `Production-Ready`，
     這不是 ACDS v2.4 第十一章定義的官方四階生命週期詞彙
     （Experimental/Stable/Frozen/Deprecated），屬自創詞彙污染。經 Shih 本人確認
     這批模組已同時滿足第十一章「Stability 晉升條件」兩項門檻（AI 稽核通過 + 
     使用者本人於真實資料上實測驗證），具備改標為 `Stable` 的資格。
  3. 19 支模組的 `AUTHOR` 欄位（ACDS 官方 Manifest 格式本未定義此欄位，屬本專案
     早期自行擴充）重複出現在每一支模組頂部，Shih 決議既然活頁簿另有版權頁功能，
     不需要在每支模組各自重複記載，只需留一個唯一落點。
- 決策：
  1. `cls_DAGEngine`／`cls_ImageEntity` 補上 `Private Const MOD_NAME`，並將全部
     `Err.Raise` 呼叫改為引用該常數；`cls_ImageEntity` 殘留註解一併刪除。
     `cls_EventStore` 經核對內部完全無 `Err.Raise`，不存在重複手打問題，不需修改。
  2. 12 支 `Mod_*` 模組之 `Stability: Production-Ready` 統一改為
     `Stability: Stable`，版本說明文字本體不動。
  3. `AUTHOR` 欄位從全部 19 支模組的 Manifest 中移除，唯一保留於 `ThisWorkbook.cls`
     頂部（`ThisWorkbook` 是活頁簿層級唯一、不會重複的實體，作為全案版權資訊的
     SSOT 落點）。
  4. 同一輪內，Shih 另行於各模組補上 `Private Const MODULE_VERSION`，作為版本號
     的輔助常數。此常數與 Manifest 抬頭的版本號文字之間**無法透過語言機制強制
     同步**（VBA 的 `'` 註解無法讀取常數值），兩者屬於「概念上的單一事實、
     人工同步維護」關係：往後任何一次改版本號，Manifest 抬頭與 `MODULE_VERSION`
     常數須視為同一個動作一起修改，不得只改其中一處。本輪稽核中發現
     `cls_DAGEngine` 曾一度出現 Manifest 抬頭寫 `1.3.0`、常數寫 `1.0.0` 的不同步
     情形，已一併修正為兩處皆為 `1.0.0`。
  5. 常數命名維持專案既有慣例 `MOD_NAME`，不比照姊妹專案（影像自動化歸檔中樞
     系統）的 `MODULE_NAME` 命名。跨專案命名差異屬 AVS v1.1 定義範圍內的正常
     現象，本專案內部一致即可，不需要為了與其他專案對齊而產生不必要的異動。
- 放棄的替代方案：
  (a) 將版本號完全收斂進 `MODULE_VERSION` 常數、Manifest 抬頭改為不寫具體數字
      （例如寫成「版本見下方常數」）——放棄理由：ACDS 第七章範例本身即允許
      Manifest 抬頭與常數並存（Python 範例的 `__VERSION__` 也是先在 Manifest
      寫一次數字，常數再宣告一次同樣的值），且 Manifest 抬頭若不寫實際數字，
      開啟 VBE 逐一瀏覽模組時無法一眼看到版本，可讀性下降，不採用。
  (b) `MOD_NAME` 改名為 `MODULE_NAME` 以對齊姊妹專案——放棄理由：純粹為了跨專案
      命名一致而產生的異動不具備功能效益，屬於不必要的重工，違反 YAGNI；AVS
      本身也不要求跨專案常數命名一致，只要求同一專案內部一致。
  (c) `AUTHOR` 資訊完全刪除，不保留於任何位置——放棄理由：版權與作者資訊仍有
      保留價值（例如未來需要證明開發歸屬時），只是不需要在 19 支模組內重複，
      收攏至 `ThisWorkbook` 保留單一落點是更好的折衷方案。
- 影響：
  往後新增或修改任何模組時，Definition of Done 檢查應包含：(1) 若模組內有
  `Err.Raise`，須引用 `MOD_NAME` 常數，不得手打模組名稱字面值；(2) 新模組不需要
  也不應該加上 `AUTHOR` 欄位；(3) 若修改版本號，Manifest 抬頭與 `MODULE_VERSION`
  常數須同步更新，稽核時應兩處對照檢查，不可只看其中一處就判定版本正確。
============================================================
============================================================
【ADR-009】撤除 ICommand.cls：過早抽象化（Premature Abstraction）案例
- 日期：2026-07（V13 → V14 過渡）
- 狀態：Accepted
- 背景：
  `ICommand.cls` 自 V2 起即存在，定義 `Public Sub Execute(ByVal target As Object)`
  標準接口，`ACDS_DependencyGraph.md` 各版本亦持續將其列為節點。但橫跨 V2 至 V13
  共 11 個子版本、含 V8「切除危機」大重構在內，全專案從未出現任何 `cmd*` 具體類別
  `Implements ICommand`，`Mod_WorkflowManager`（流程大腦）也始終直接呼叫
  `Mod_ImportPipeline`／`Mod_ExportPipeline` 等管線模組的具體實作，未曾透過此介面
  間接呼叫。
  V12/V13 稽核時多次重新提出「這個要不要用」的疑問，Shih 最終以 AVS v1.1 第五節
  判準逐條檢視：本專案四個功能（匯入／匯出／重整工作檯／日誌歸檔）皆為固定線性
  流程，不存在「動態組合、彼此獨立替換、或依賴關係需要排程」的真實需求，確認
  即使補齊 Command Pattern 完整實作，也只是在既有管線外多包一層轉發殼，不會新增
  任何行為，純屬規格過度（over-engineering）。
- 決策：
  V14 正式撤除 `ICommand.cls`，從專案模組清單、`ACDS_ContractRegistry.md`、
  `ACDS_DependencyGraph.md` 中一併移除。本專案不套用 Interface First 原則於管線
  調度層級，`Mod_WorkflowManager` 繼續直接依賴各管線模組的具體實作。
- 放棄的替代方案：
  (a) 保留作為預留擴充點，Stability 標註為 Experimental——放棄理由：VBA 建立新
      類別檔案的成本極低，不存在「以後真的需要時來不及建立」的風險，「先留著
      保險」不構成有效理由；反而每一輪稽核都要重新解釋一次「為何存在但未使用」，
      累積的認知負擔已超過保留它的任何潛在效益。這正是本 ADR 標題所指的「過早
      抽象化」教訓：抽象層的建立必須有明確可預見的替換需求撐著（第二章第九條），
      「未來可能用得到」不構成明確可預見的需求。
  (b) 真正補齊完整實作（`cmd*` 具體類別 + `Mod_CommandFactory` + 將 `cls_DAGEngine`
      升級為真正的依賴排程引擎）——放棄理由：升級 `cls_DAGEngine` 語意會直接推翻
      ADR-001 剛定案的「系統級生命週期守衛」角色，而升級的動機不是來自真實需求，
      是為了讓已經蓋好的介面有東西可用，屬於本末倒置；且姊妹專案（影像自動化
      歸檔中樞系統）已示範了 Command Pattern 在有真實動態排程需求（局部重排、
      兩階段重正規化）時的正確落地方式，本專案不具備同等複雜度，沒有理由套用
      同一套重型架構。
- 影響：
  本案例應收錄進 ACDS 本體附錄 C 案例庫，作為與 C-2（識別資訊漂移）、
  ADR-003 所述之上帝物件技術債（規則空隙導致實作跑偏）性質不同的第三種模式：
  **規則本身清楚（第二章第九條、AVS 第五節判準皆已明文），但實作端仍在早期
  未經充分驗證需求前，先行建立了抽象層**，並非因為判準不清楚而誤用，而是建立
  當下對「是否真的需要」的判斷過於樂觀。建議收錄為 C-6「過早抽象化」案例，核心
  原則：抽象層的建立成本雖低，但長期擱置未使用的抽象層會持續消耗稽核與維護
  注意力，其成本不是零，「以防萬一」不應作為建立抽象層的充分理由。
============================================================
============================================================
【ADR-010】按鈕入口巨集雙軌可見度設計：Optional Dummy 隱藏機制 + 唯一功能雙入口例外
- 日期：2026-07（V13 → V14 過渡）
- 狀態：Accepted
- 背景：
  V13 稽核發現 `Mod_MainCoordinator` 五支按鈕入口巨集中，四支
  （`Macro_Phase1_ImportFolder`／`Macro_Phase1_ImportFiles`／`Macro_Phase2_Export`／
  `Macro_ConsolidateAllLogs`）被加上 `Optional Dummy As Byte = 0` 參數，第五支
  `Macro_ForceRebuildUI` 則被改名為中文 `一鍵還原檯面()` 且未加此參數，同時舊名稱
  `Macro_ForceRebuildUI` 完全消失。這個改動一度造成實際功能斷裂：
  `Mod_StringConstants.MACRO_REBUILD_UI` 常數與 `Mod_UIRenderer` 按鈕
  `.OnAction` 綁定皆仍引用字面值 `"Macro_ForceRebuildUI"`，該巨集改名後
  Excel 已找不到對應程序，使用者點擊「一鍵重整工作檯」按鈕會直接報錯無反應。
  經與 Shih 確認，這批改動背後有兩個獨立、刻意的設計意圖：
  1. `Optional Dummy As Byte = 0` 是刻意讓該巨集從 Excel `Alt+F8` 巨集清單中
     隱藏（VBA 巨集清單不列出帶有 Optional 參數的 Public Sub），避免操作人員
     繞過按鈕直接從巨集清單手動觸發，落實「只能透過按鈕正常流程進入」的邊界。
  2. 唯獨「重整工作檯」功能需要一個不依賴按鈕本身的「緊急救援入口」——因為
     `Macro_ForceRebuildUI` 呼叫的 `RunRebuildUI` 正是負責重建畫布與按鈕視覺
     元件的功能，若前台視覺元件本身損毀，操作人員將無法點擊任何按鈕（包含
     「重整工作檯」按鈕本身）來自救，此時需要一個在 `Alt+F8` 巨集清單中依然
     可見、可手動執行的中文別名入口。其餘四支功能不存在「按鈕壞掉時只能靠
     該功能自己修復按鈕」的情境，不需要此救援設計。
- 決策：
  1. 修復功能斷裂：`一鍵還原檯面()` 與 `Macro_ForceRebuildUI(Optional Dummy As
     Byte = 0)` 兩支巨集**同時保留**於 `Mod_MainCoordinator`，皆為一行委派至
     `Mod_WorkflowManager.RunRebuildUI`，內容完全相同，維持「唯一入口層只做
     委派」原則（見 ADR-004）不變——重複的是呼叫方式，底層邏輯仍只有一處。
     `Macro_ForceRebuildUI` 保留 `Optional Dummy`（供按鈕正常呼叫且從清單隱藏），
     `一鍵還原檯面` 不加 `Optional Dummy`（刻意在 `Alt+F8` 清單中可見，作為
     視覺元件損毀時的手動救援出口）。
  2. `Macro_Phase1_ImportFolder`／`Macro_Phase1_ImportFiles`／`Macro_Phase2_Export`／
     `Macro_ConsolidateAllLogs` 四支維持現狀（有 `Optional Dummy`，巨集清單隱藏，
     無中文別名救援版本），因為這四個功能不存在「功能本身損毀導致無法透過任何
     UI 觸發」的情境，不需要救援入口。
- 放棄的替代方案：
  (a) 全部五支統一加上 `Optional Dummy`，不保留任何巨集清單可見入口——放棄
      理由：會讓「重整工作檯」功能在前台視覺元件損毀時完全無法被觸發，形成
      單點故障（沒有任何救援路徑可以修復壞掉的按鈕本身）。
  (b) 為救援入口另建一支獨立模組（例如 `Mod_EmergencyRecovery`）——放棄理由：
      違反 ACDS_ContractRegistry.md 明定「`Mod_MainCoordinator` 是全專案唯一
      OnAction／巨集入口」的合約，會製造第二個入口層，破壞 ADR-004 的保證
      （不論從哪個管道觸發，最終都收斂到同一層做委派）。兩個名稱、同一支模組、
      同一個底層呼叫，才是唯一不違反現有合約又能達成救援需求的做法。
  (c) 救援入口沿用英文名 `Macro_ForceRebuildUI_Manual` 等變體——放棄理由：
      中文別名 `一鍵還原檯面` 讓非工程背景操作人員在 `Alt+F8` 清單中能直接
      辨識用途，不需要猜測英文命名代表什麼，更符合「簡單、無腦」的產品標準。
- 影響：
  往後任何新增功能若要評估是否需要比照「重整工作檯」設計雙入口，判準是：
  **這個功能是否是唯一能修復「操作人員完全無法透過任何按鈕觸發功能」這個故障
  狀態的手段**？若是，才需要保留巨集清單可見的救援別名；若否，一律採用
  `Optional Dummy As Byte = 0` 隱藏於巨集清單之外，維持只能透過按鈕正常流程
  觸發的邊界。Manifest 的 `EXPORTS` 欄位若列出此類巨集，應註記其巨集清單可見度
  與用途差異，避免未來稽核者誤判為忘記清除的殘留而擅自刪除。
============================================================
============================================================
【ADR-011】RenderCanvasAndAnchor 位置保留修復：借鑑 V1_CL 分支經驗，畫布與錨點改為「存在則沿用」
- 日期：2026-07（V13 → V14 過渡）
- 狀態：Accepted
- 背景：
  Shih 另外保留了一支獨立分支 `V1_CL`（於 V0 之後不久分岔，使用另一個 AI 協作、
  僅開發到 `Mod_ActionExecutor`／`Mod_EnvironmentSetup` 尚未拆分的階段即未再繼續，
  未併入 V2 起的主線）。稽核該分支時發現，其 `Mod_EnvironmentSetup`（對應主線後來
  拆分出的 `Mod_UIRenderer`）版本註記明確記載一項主線從未修復的缺陷：
  `Mod_UIRenderer.ClearAllExistingShapes` 每次執行皆無條件 `Delete` 掉
  `SH_CANVAS_NAME`／`SH_BLEED_ZONE_NAME`／`SH_ANCHOR_NAME` 三個視覺物件，
  `RenderCanvasAndAnchor` 再無條件於固定預設座標重新建立。而 `RenderUIElements`
  （內部呼叫上述兩者）於 `Mod_WorkflowManager.RunPhase1`（即「匯入資料夾」／
  「挑選圖片」兩顆按鈕的底層流程）與 `RunRebuildUI`（「重整工作檯」按鈕）皆會
  觸發。三個物件中，只有 `SH_ANCHOR_NAME`（浮水印位置錨點）是使用者會手動拖曳、
  承載自訂狀態的物件——這代表**使用者每次執行匯入，先前拖曳過的浮水印位置都會
  被靜默重置回預設角落，且系統不會給予任何提示**，需重新拖曳才能恢復預期位置。
  V1_CL 分支的版本註記寫著：「修復每次執行都刪除重建畫布，導致浮水印位置被強制
  歸零的問題」，證實此問題在該分支獨立測試時已被發現並修復，但因該分支後續未
  併入主線、V8 大重構是延續主線（未修復版本）進行，此修復從未進入 V2 之後的
  正式演進路徑，主線一路到 V13 皆帶有此缺陷。
- 決策：
  比照 V1_CL 分支的修復原則，將 `Mod_UIRenderer.ClearAllExistingShapes` 與
  `RenderCanvasAndAnchor` 改為「探測既有物件是否存在，存在則完全沿用（不刪除、
  不重新定位），不存在才依預設座標建立」。三個物件（畫布、出血框、錨點）皆採
  此邏輯，統一以 `On Error Resume Next` 探測 `ws.Shapes(常數名稱)` 取代原本的
  無條件 `Delete`。按鈕（`BTN_*`）維持原本「每次刪除重建」邏輯不變，因為按鈕
  不承載使用者自訂狀態，重建成本低且無副作用。實作對齊 V13 現行的莫蘭迪配色
  常數與命名慣例重寫，非直接搬移 V1_CL 分支的舊程式碼。
- 放棄的替代方案：
  (a) 只修 `SH_ANCHOR_NAME`，維持畫布與出血框每次刪除重建——放棄理由：畫布與
      出血框的位置若被重新定位，會與已保留原位的錨點之間的相對座標關係跑掉
      （`View_CanvasManager` 計算的是錨點相對於畫布的百分比位置，若畫布重新
      定位、錨點不變，相對百分比會失真）；三者要嘛都沿用、要嘛都重建，不可
      只保留其中一個而讓相對關係不一致。
  (b) 額外把使用者拖曳過的錨點座標另外寫入 Staging 或 Config 工作表儲存格，
      每次啟動時再讀回寫入——放棄理由：Excel Shape 物件本身若未被刪除，其
      `Left`／`Top` 屬性天生就會保留在活頁簿檔案中，不需要額外一份持久化
      副本；多開一份座標記錄只會製造新的 SSOT 落點與同步風險，不符 KISS 原則。
- 影響：
  這是本專案第二個「兩個分支各自獨立演化、其中一支修復的問題未被帶回主線」的
  真實案例（第一個是 ADR-003 描述的上帝物件重構期間發生的變數一物多用）。
  建議 Shih 未來若同時維護多個分支進行 AI 協作實驗，在確定某一分支不再繼續開發
  前，先過一次「這支分支相較主線是否修過主線還沒修的問題」的檢查，避免有價值
  的修復隨分支一起被擱置遺忘。此案例收錄進 ACDS 附錄 C 建議編號 C-7
  「分支修復遺失（Branch Fix Loss）」，核心原則：分岔的分支即使不繼續開發，
  仍應在正式歸檔前檢查是否有可回饋主線的修復，不應僅因分支本身停止演進就整批
  視為無價值。
============================================================
============================================================
【ADR-012】RenderCanvasAndAnchor 位置保留策略再區分：匯入沿用、重整UI強制重置（延伸 ADR-011）
- 日期：2026-07（V13 → V14 過渡，ADR-011 落地測試後的即時修正）
- 狀態：Accepted
- 背景：
  ADR-011 將 `RenderCanvasAndAnchor` 改為「畫布／出血框／錨點三者一律存在則
  沿用」，解決了匯入時位置被強制歸零的問題。但 ADR-011 的實作未區分呼叫來源：
  `Mod_WorkflowManager.RunPhase1`（匯入）與 `RunRebuildUI`（一鍵重整工作檯）
  皆呼叫同一支不帶參數的 `RenderUIElements`，導致「重整工作檯」這個原本設計
  用來救援畫面異常、應該恢復原廠預設版面的按鈕，現在也變成沿用既有位置、
  不會真的重置。Shih 確認這不符合預期：**匯入不該動版面，但重整UI本來就該
  回到原始設定**，兩者是不同語意，不該共用同一套無差別行為。
- 決策：
  `RenderUIElements` 新增 `Optional ByVal forceResetLayout As Boolean = False`
  參數，向下傳遞給 `RenderCanvasAndAnchor` 的 `Optional ByVal forceReset As
  Boolean = False`。當 `forceReset = True` 時，函式開頭先物理 `Delete` 既有的
  畫布／出血框／錨點三者，讓後續「探測是否存在」的邏輯視為不存在，依預設座標
  重新鑄造，達成真正的「恢復原廠設定」效果。
  呼叫端調整：`Mod_WorkflowManager.RunPhase1` 維持呼叫
  `Mod_UIRenderer.RenderUIElements`（不帶參數，沿用預設值 False，行為不變，
  延續 ADR-011）；`RunRebuildUI` 改為明確呼叫
  `Mod_UIRenderer.RenderUIElements(True)`，確保這顆救援按鈕的行為與其名稱
  「一鍵重整工作檯」／「重繪操作台」名符其實。
- 放棄的替代方案：
  (a) 另外新建一支獨立的 `ForceResetCanvasAndAnchor` 副本函式，與既有的
      「存在則沿用」版本並存——放棄理由：兩份幾乎相同的建立邏輯（畫布／出血
      框／錨點的座標計算與外觀設定）會產生雙重維護點，未來若要調整外觀樣式，
      需要同步修改兩處，違反 DRY 原則。用參數區分行為分支，核心建立邏輯只寫
      一份，才是正確做法。
  (b) 讓 `RunRebuildUI` 呼叫端自行在呼叫 `RenderUIElements` 前手動
      `ws.Shapes(...).Delete` 三個物件——放棄理由：這會讓「如何重置版面」的
      知識從 `Mod_UIRenderer`（視覺渲染職責歸屬模組）外洩到
      `Mod_WorkflowManager`（流程大腦），違反單一職責；`Mod_WorkflowManager`
      不該知道 UI 元件叫什麼名字、該怎麼重建。
- 影響：
  往後任何呼叫 `RenderUIElements` 的地方，預設行為都是「保留使用者版面」，
  只有明確需要救援重置的情境才需要顯式傳入 `True`，這個預設值方向選擇是刻意
  的：讓「安全、不打擾使用者」成為預設路徑，「破壞性重置」必須是呼叫端主動、
  明確表達的意圖，不能是意外發生的副作用。此設計原則（危險操作需顯式傳參啟用，
  安全行為為預設值）可作為本專案未來新增類似雙模式函式時的參考範本。
============================================================
============================================================
【ADR-013】一鍵重整工作檯新增清除空白頁面功能：### 保護標記機制
- 日期：2026-07（V13 → V14 過渡）
- 狀態：Accepted
- 背景：
  Excel 活頁簿在日常使用中容易累積多餘空白頁面（例如新增活頁簿時 Excel 自動
  帶出的 `工作表1`／`Sheet1`，或使用者操作時不小心手滑建立的空白分頁）。既然
  「一鍵重整工作檯」本來就是使用者主動觸發、期望恢復乾淨檯面的救援動作
  （見 ADR-011、ADR-012），Shih 要求一併清除這類多餘空白頁面。
- 決策：
  新增 `Mod_InteractionAdapter.PurgeBlankOrphanSheets`，於 `RunRebuildUI` 流程
  中、`RenderUIElements(True)` 之後呼叫。判斷邏輯採兩道防線＋一道空白判斷：
  1. 四張系統表（`ConfigSheet`／`Staging_Images`／`LOG_Export_成果`／
     `LOG_Master_Archive`）依名稱無條件排除，獨立於空白判斷之外，避免系統表
     恰好處於空白狀態的時間點被誤刪。
  2. 名稱以 `###` 開頭的工作表視為使用者手動標記保護，無條件排除——這是一個
     開放式的保護機制，使用者可對任何想保留的頁面（不限於系統表）自行標記，
     不需要每次都回頭修改程式碼裡的白名單常數。
  3. 通過前兩關後，以 `CountA(UsedRange) = 0` 判斷是否真正空白，只有真正空白
     的頁面才列入候選刪除清單。
  4. 有候選清單才彈出確認視窗列出頁面名稱，使用者按「是」才真正執行刪除，
     維持本專案「破壞性操作前必須經過使用者確認」的一貫慣例（比照撞名版控
     降位、斷點續跑詢問）。
  刪除本身不套用「複製優先於刪除」（P0 第 2 條），因為判定為候選的頁面定義上
  就是不含任何內容，沒有需要保護的資料，備份一張空白頁面沒有意義。
- 放棄的替代方案：
  (a) 寫死一份允許清單常數，只有清單內名稱才保護——放棄理由：使用者提出的
      `###` 前綴機制更有彈性，任何時候想保留一張新的頁面，直接改名即可，不
      需要每次都回頭修改 `Mod_StringConstants` 常數並重新走一次版本異動流程。
  (b) 靜默清除，不彈出確認視窗——放棄理由：即使判定為空白，刪除工作表仍是
      不可逆操作（活頁簿層級沒有資源回收桶），且與本專案一路以來「刪除前先
      確認」的既有慣例不一致，維持確認視窗以求一致性與安全邊際。
  (c) 將此功能放進 `Mod_WorkflowManager` 直接操作 `ThisWorkbook.Worksheets`
      ——放棄理由：直接違反 `ACDS_DependencyGraph.md` 明定的紅線（大腦嚴禁
      直接操作工作表物件，一律透過 `Mod_InteractionAdapter`），故放進
      `Mod_InteractionAdapter`，維持既有分層鐵律。
  (d) 偵測到 `ThisWorkbook.ProtectStructure` 為 True 時，自動 `Unprotect` 執行
      刪除、完成後再 `Protect` 回去（另一專案之既有範例採此做法）——放棄理由：
      本專案 `Mod_LogManager.CreateLogSheetFrame` 已針對同一情境立下先例，
      偵測到結構受保護時明確 `Err.Raise` 中斷，不自動繞過保護。若
      `PurgeBlankOrphanSheets` 改採自動解鎖做法，會讓同一專案內「遇到結構
      保護」這件事出現兩種不同反應，屬於行為層級的識別資訊漂移，故遵循
      既有 Fail-Fast 慣例，不引入新模式。
- 影響：
  往後若使用者想保留任何非系統表的工作表不被「重整工作檯」清除，只需要在
  頁籤名稱前手動加上 `###` 即可，不需要走程式異動流程。稽核時應留意
  `PurgeBlankOrphanSheets` 的兩道排除防線（系統表白名單、`###` 標記）必須
  同時存在、缺一不可，若未來有人誤以為 `###` 機制可以完全取代系統表白名單而
  將其移除，會使四張核心表格失去「不受空白判斷影響」的保障。
============================================================
============================================================
【ADR-014】RunRebuildUI 修復 wsCurrent 懸空物件錯誤：PurgeBlankOrphanSheets 與既有畫面還原邏輯的組合型缺陷
- 日期：2026-07（V13 → V14 過渡，ADR-013 落地實測後發現）
- 狀態：Accepted
- 背景：
  Shih 實際測試「一鍵重整工作檯」時回報：**在全新方案（新活頁簿、尚未建立
  ConfigSheet／Staging_Images 之前）執行會觸發物件錯誤**。
  追查後確認為組合型缺陷，非任一單一函式獨立的邏輯錯誤：
  `RunRebuildUI` 早於 ADR-013 即存在的既有模式是「記住使用者點擊按鈕當下所在
  的工作表（`wsCurrent`），流程跑完後再切回去」，此模式隱性假設 `wsCurrent`
  在整段流程執行期間必然持續存在。ADR-013 新增的 `PurgeBlankOrphanSheets` 會
  真的刪除被判定為空白的工作表，且未將「目前使用者停留的作用中工作表」納入
  保護範圍。全新活頁簿情境下，使用者操作當下的作用中工作表極可能正是 Excel
  自動帶出的空白預設頁（如「工作表1」），這張表同時滿足
  `PurgeBlankOrphanSheets` 的刪除判定條件，也剛好是被 `wsCurrent` 記錄的那張
  表——遭刪除後，`wsCurrent` 物件參照懸空，執行到 `wsCurrent.Activate` 時
  VBA 對已銷毀 COM 物件呼叫方法／屬性，拋出物件錯誤，整支程序中斷，連帶使
  `Application.ScreenUpdating` 未能恢復為 `True`（因錯誤發生在該行之前），
  介面會卡在凍結狀態。
  這類問題本質上是「兩個各自獨立合理的邏輯，組合後才暴露出未被考慮到的邊界
  情況」，屬於本專案第三種需要留意的協作模式（前兩種為 ADR-003 上帝物件債、
  ADR-011/V1_CL 分支修復遺失）。
- 決策：
  將「切回原表」這一步改為容錯處理：以 `On Error Resume Next` 包裹
  `wsCurrent.Activate`，若觸發錯誤（代表 `wsCurrent` 已在流程中途被刪除），
  改為安全退回顯示 `ConfigSheet`（本專案的主操作介面，使用者重整工作檯後
  理應落在這裡也是合理的預設行為），而非任由錯誤中斷整支程序。
- 放棄的替代方案：
  (a) 讓 `PurgeBlankOrphanSheets` 額外接收 `wsCurrent` 作為參數，內部主動排除
      「目前作用中工作表」不列入刪除候選——放棄理由：這會讓
      `PurgeBlankOrphanSheets` 的職責範圍擴大到需要理解「呼叫端關心哪張表」，
      而它的核心職責應該單純是「掃描並清除真正空白的頁面」；作用中工作表
      未必等同於「使用者想保留的表」（使用者可能就是想清掉自己正站著的那張
      空白頁），用 `###` 標記機制才是使用者表達「我要保留」意圖的正確管道，
      不該用「剛好游標在上面」這種偶然狀態當作保護依據。
  (b) 在 `RunRebuildUI` 呼叫 `PurgeBlankOrphanSheets` 之前，先強制把畫面切到
      `ConfigSheet`，確保 `wsCurrent` 追蹤的表不會被刪除——放棄理由：這樣一來
      `wsCurrent` 記錄的意義就被架空了（永遠等於 `ConfigSheet`，記錄這個動作
      失去實質作用），且會在畫面上多一次不必要的分頁切換閃爍，不如直接讓
      「切回原表」這步本身具備容錯能力來得乾淨。
- 影響：
  本專案任何「記住畫面狀態、流程跑完後還原」的既有模式（本例為 `wsCurrent`），
  只要流程中間插入了會刪除工作表的新動作，都應該重新檢視「原本記住的物件是否
  可能在中途失效」，不能預設物件參照在整段流程期間必然有效。未來新增任何會
  刪除工作表／活頁簿層級物件的功能時，應主動檢查是否與既有的「畫面狀態記憶與
  還原」模式產生類似的組合型風險。
============================================================
============================================================
【ADR-015】UI 控制面板新增版權聲明頁尾：修正「僅存放於 ThisWorkbook」的先前建議
- 日期：2026-07（V13 → V14 過渡）
- 狀態：Accepted
- 背景：
  Claude 於前次對話中建議授權聲明僅存放於 `ThisWorkbook.cls` 註解，理由是
  避免在精心排版的操作面板上堆疊法律文字、且與 ADR-008 的 `AUTHOR` 收攏決策
  維持一致。Shih 提出反例：另一專案（影像自動化歸檔中樞系統）已將授權聲明
  以低調小字放在控制面板頁尾，並要求本專案跟進同一模式。
  重新檢視後，先前建議忽略了一個關鍵事實：**`ThisWorkbook.cls` 的 VBA 註解
  對一般 Excel 使用者是不可見的**，除非主動按 `Alt+F11` 開啟 VBA 編輯器。而
  本專案已確認此 `.xlsm` 檔案有可能脫離 Git repo、單獨流通給不熟悉技術的
  使用者。這類使用者幾乎不可能自行開啟 VBA 編輯器查看授權條款，若授權聲明
  只存在於 `ThisWorkbook`，對這類使用者而言等同於不存在。先前建議只考慮了
  「架構純粹性」（不與 ADR-008 的收攏原則衝突）與「介面美觀」，未考慮「授權
  聲明存在的實際目的是要讓人看到」這件事，屬於判斷不完整，非架構原則本身
  有誤。
- 決策：
  採取雙軌並存，非二選一：
  1. `ThisWorkbook.cls` 保留完整法律全文（MIT 完整條款 + CC BY-NC-SA 4.0 三大
     限制條件摘要與官方連結），作為需要查證完整條款時的權威來源。
  2. `Mod_UIRenderer.RenderControlPanel` 新增低調小字頁尾（7pt 灰色文字，
     `Footer_LicenseNotice`），放置於「系統維護工具區」按鈕列下方，文案
     與另一專案完全一致（`Code: MIT License | Documentation: CC BY-NC-SA 4.0
     | 完整版權聲明請見 LICENSE 檔案`），確保任何開啟此活頁簿的人，不需要
     懂 VBA 就能看到授權資訊存在。
  兩者不重複維護同一份全文——UI 頁尾僅放精簡摘要與指引文字，`ThisWorkbook`
  才是完整條款的唯一存放處，符合 SSOT 精神：UI 頁尾是「指標」，
  `ThisWorkbook` 才是「內容本體」。
  `Footer_LicenseNotice` 歸入 `ClearAllExistingShapes` 每次刪除重建清單
  （比照按鈕，非比照畫布／錨點的位置保留邏輯），因其不承載使用者自訂狀態。
- 放棄的替代方案：
  (a) 維持僅存放於 `ThisWorkbook`，不修改 UI——放棄理由：見背景，無法達成
      「單獨流通情境下讓非技術使用者也能看到授權資訊」的實際目的。
  (b) 只在 UI 放頁尾小字，`ThisWorkbook` 不寫完整全文，僅寫一行指向 UI
      頁尾——放棄理由：UI 頁尾字級極小（7pt）且空間有限，無法容納完整
      MIT 條款全文；若真的發生法律爭議需要查證完整條款文字，仍需要一個
      承載得下全文的地方，`ThisWorkbook` 註解沒有這個空間限制。
- 影響：
  這是本次協作過程中一個值得記錄的自我修正案例：AI 前次建議並非規則層級
  的錯誤，而是「只從架構一致性與美觀角度考慮，遺漏了『這份資訊要給誰看、
  對方看不看得到』這個更根本的實用性問題」。往後任何「這則資訊該放哪裡」
  的類似決策，除了問「跟現有 SSOT 原則一不一致」，也應該多問一句「目標
  讀者實際上有沒有辦法接觸到這個位置」，兩者都通過才算是完整的判斷。
============================================================
============================================================
【ADR-016】Mod_ImportPipeline 完全解耦：比照 Mod_ExportPipeline 達成對 Excel 網格 100% 無知
- 日期：2026-07（V13 → V14 過渡）
- 狀態：Accepted
- 背景：
  自 V8 大重構（ADR-003）起，`Mod_ImportPipeline`／`Mod_ExportPipeline` 作為
  同一層級、職責對稱的兩支批次管線模組，理應遵循相同的解耦程度。但
  `Mod_ExportPipeline` 早已透過 `Mod_InteractionAdapter.PackageStagingTasks`
  在管線外先打包成 `cls_ImageEntity` 集合，管線本體完全不碰 Excel 網格；
  `Mod_ImportPipeline` 卻始終直接持有 `wsStaging` 工作表物件，內部
  `LoadExistingHashes`／逐列寫入 `Cells(...)` 皆為直接網格操作，此不對稱已於
  V13 稽核時列為「已知技術債」誠實記載於 `ACDS_DependencyGraph.md`、Manifest
  與 README。Shih 認為兩支管線行為不對稱本身即是一個值得修正的訊號，即使沒有
  外部功能需求驅動，仍決定投入這輪重構補齊。
- 決策：
  新增 `cls_ImportResult` 類別，與既有 `cls_ImageEntity`（匯出用唯讀模型）
  對稱設計，承載單一檔案「匯入結果」（成功／失敗兩種狀態、對應 Staging 一列
  的完整欄位資料）。`Mod_InteractionAdapter` 新增兩支函式：
  `LoadExistingHashSet()`（讀取 Staging 既有雜湊點名簿，取代原本
  `Mod_ImportPipeline` 內部直接讀取）、`WriteImportResults(colResults)`
  （接收處理完畢的 `cls_ImportResult` 集合，統一在此一次寫入所有列並調整欄寬）。
  `Mod_ImportPipeline.ExecuteImportPipeline` 改為只累積 `Collection`，迴圈
  結束後才一次呼叫 `WriteImportResults`，全程不再持有或操作 `wsStaging`。
  `Application.StatusBar` 更新維持留在管線內——此為 Excel 應用程式層級狀態列，
  非工作表網格物件，不落入「網格讀寫」紅線範圍，`Mod_ExportPipeline` 亦無此
  類操作可供對照，故不強制搬移。
- 放棄的替代方案：
  (a) 維持現狀，僅在文件誠實記載此技術債，不實際修復——放棄理由：此為前一輪
      （ADR 前）的暫定立場，Shih 這輪明確表示「解耦到一半、不對齊」的狀態本身
      造成困擾，即使沒有具體外部需求，模組職責邊界的內部一致性本身即構成
      修正理由，決定投入處理。
  (b) 讓 `Mod_ImportPipeline` 直接重用 `cls_ImageEntity`（匯出用模型）承載
      匯入結果，不另建新類別——放棄理由：`cls_ImageEntity` 語意上代表「已確認
      存在、可供匯出處理的圖片」，不包含「處理失敗」狀態；若強行複用，
      `IsSuccess`／`FailReason` 這類欄位會讓匯出情境下的呼叫端困惑為何存在
      不會用到的屬性，違反 Interface 語意單一性，故另建 `cls_ImportResult`。
  (c) `LoadExistingHashSet` 與 `WriteImportResults` 合併成一支函式——放棄
      理由：兩者時機不同（一個在批次處理前讀取、一個在批次處理後寫入），
      合併會讓 `Mod_ImportPipeline` 必須把整個批次迴圈邏輯內嵌進
      `Mod_InteractionAdapter` 才能在同一函式內完成讀取與寫入，這會讓
      `Mod_InteractionAdapter` 承擔本該屬於管線層的業務邏輯，違反 SRP。
- 影響：
  `Mod_ImportPipeline`／`Mod_ExportPipeline` 現在對稱地達成「對 Excel 網格
  100% 無知」，`ACDS_DependencyGraph.md`「已知的拓撲不完美之處」章節中原本
  記載的此項技術債應予移除；README〈已知限制〉的對應描述亦應同步移除。
  往後任何新增的批次管線模組，若需要讀寫 Staging 或其他工作表，應直接遵循
  本 ADR 建立的模式（管線只處理領域物件集合，實際網格讀寫委託
  `Mod_InteractionAdapter`），不應重蹈 `Mod_ImportPipeline` 早期直接持有
  工作表物件的舊模式。
============================================================

【ADR-017】廢除日期跳號機制，改採固定工作資料夾拓撲＋事件式整批備份，根治匯出重複檔案
- 日期：2026-09（V13 → V17 過渡）
- 狀態：Accepted
- 背景：
  V13 版本每次「開新案」都由 Mod_FileSystem.CreatePreventCollisionDir 依日期
  跳號建立新的「整理圖片_NN」資料夾，Phase 2 匯出目標則是該資料夾下固定命名
  的「快取斷點」子資料夾。稽核發現兩個疊加問題：(1) 每次 Phase 2 結案時，
  ArchiveExportDirectory 會把「快取斷點」當下累積的全部內容整包複製進一個新
  的時間戳備份資料夾，但從未清空「快取斷點」本身——同一批照片因此在每次結案
  時被重複封存一次，跨多輪執行後產生大量實體重複檔案。(2) UI 文案
  （MSG_SESSION_CLOSED_SUFFIX）承諾「將於下次工作開始時清零重啟」，但程式碼
  從未真正清零，文案與實作脫鉤。首次嘗試（僅在封存後清空快取斷點）解決了
  重複檔案，但破壞了「斷點續跑」偵測的前提（該偵測靠資料夾內殘留檔案判斷），
  且與規格書原本定義「快取斷點＝最終成品區」的角色矛盾，判定為不完整修正。
- 決策：
  廢除日期跳號機制，改採固定四層拓撲：來源資料夾下永遠只有一個「作業資料夾」，
  內含「匯入圖片」「成果」「備份」三個固定子資料夾（Mod_FileSystem.
  EnsureWorkspaceTopology 統一確保存在）。「成果」資料夾恢復規格書原意，
  永久累積、不再清空，斷點續跑偵測（IsResumingPreviousRun）因此重新正確運作。
  備份改為「事件式整批備份」：只在兩個真正有意義的邊界觸發，且兩者皆只備份
  「成果」——① Phase 1「開新案」選「否」時，「成果」整批備份（Mod_FileSystem.
  BackupAndClearFolder）後清空；「匯入圖片」定位為純工作複製品，不具備份
  價值，直接以 ClearFolderContents 清空，不進入備份流程（見下方「放棄的替代
  方案 (e)」，此為本 ADR 定案前的中繼反覆）。② Phase 2 偵測到「成果」有殘留、
  使用者選擇「結束並重印」時，「成果」整批備份（Mod_FileSystem.
  BackupAndClearFolder）後清空，並同步呼叫 Mod_InteractionAdapter.
  ResetAllExportedFlags 重置 Staging 表全部「已匯出」狀態，讓下一次
  PackageStagingTasks 能重新納入全部照片
  （此為必要的排序修正：RunPhase2 內「打包 colTasks」的時機點，從原本「先打包
  再問斷點續跑」調整為「先問斷點續跑、處理完覆寫分支後才打包」，否則清空的
  成果資料夾會配上一個空的 colTasks，變成「什麼都不重印」的假象）。
  斷點續跑之外，另搭配 Staging_Images 新增第 9 欄「已匯出狀態」
  （NUM_STG_COL_EXPORTED），PackageStagingTasks 過濾條件同時檢查匯入成功與
  尚未匯出，取代原本單純依賴檔案是否存在的脆弱判斷。cls_ImageEntity.
  InitializeEntity 因此從 7 參數擴為 8 參數，新增 StagingRowNumber 唯讀屬性，
  供匯出成功後 Mod_InteractionAdapter.MarkStagingRowAsExported 精準回寫列號，
  不再依賴檔名或雜湊反查。
- 放棄的替代方案：
  (a) 僅清空快取斷點、不動斷點續跑邏輯——放棄理由：見上方背景段，會讓「斷點
      續跑」永遠偵測不到殘留檔案，該提示視窗形同永久失效，且與規格書「快取
      斷點＝最終成品區」定位矛盾。
  (b) 增量備份（每次只備份本次新產生的檔案，快取斷點维持不清空）——曾短暫
      採用（過渡版本中期），放棄理由：Shih 提出「如果備份資料夾只是備份、
      沒有其他用途，反而增加無意義的混亂」，且增量快照分散在多個時間戳資料夾，
      單一備份夾不再是「當下全部成品」的完整副本，不利於「整批壓縮帶走」的
      實際使用情境。
  (c) 用 NTFS 硬連結（mklink /H）取代真實複製，讓整包快照零硬碟成本——放棄
      理由：VBA FSO 原生不支援硬連結，須改用 Shell 呼叫外部命令列，引入本專案
      過去不存在的新依賴與失敗模式（跨磁碟分割區會直接失敗），Shih 選擇技術
      風險更低的事件式整批備份方案，未採用此方案。
  (d) 維持日期跳號機制，只在其內部修正備份邏輯——放棄理由：Shih 主動提出
      「乾脆變成永遠只有一個作業資料夾」的扁平化訴求，跳號機制本身即為
      「開新案」語意下唯一需要跳號的原因，拿掉開新案時的物理跳號後，跳號
      機制已無存在理由。
  (e) 開新案時「匯入圖片」與「成果」合併備份進同一個時間戳資料夾（新增
      Mod_FileSystem.MergeBackupInto／BuildBackupFolderName，各自落在子目錄
      下湊成一份可直接壓縮帶走的完整舊目錄）——曾短暫採用（本 ADR 定案前的
      中繼版本），放棄理由：此設計是為了模擬「日期跳號機制下，開新案自然
      會把匯入複製品與成果湊在同一個資料夾、方便整包壓縮帶走」的舊有使用
      體驗，但回頭檢視「匯入圖片」的本質——它只是外部原始照片的清洗過
      複製品，原始照片仍在使用者自行管理的外部位置，「匯入圖片」弄丟了
      大不了重新匯入一次，不構成「弄丟了就真的沒了」的備份價值。為了保留
      一個操作習慣（整包壓縮）而去備份一份本不具保護價值的複製品，是把
      「操作便利性」與「資料是否值得備份」兩個不同問題混為一談，判定為
      過度設計，予以撤回。`MergeBackupInto`／`BuildBackupFolderName` 兩支
      函式隨之整支移除，不留死碼。
- 影響：
  `DIR_NAME_MAIN_PREFIX`／`DIR_NAME_IMPORT_PREFIX`／`DIR_NAME_EXPORT_FIXED`／
  `DIR_NAME_EXPORT_ARCHIVE_PREFIX` 四個常數與 `Mod_FileSystem.
  CreatePreventCollisionDir`／`PrepareAndDeriveExportFolder`、
  `Mod_InteractionAdapter.GetExistingStagingFolder`／`DiscoverImportFolder`
  全數移除，全專案無殘留呼叫端。`Mod_FileSystem.MergeBackupInto`／
  `BuildBackupFolderName` 於中繼版本新增後又撤回（見放棄方案 (e)），同樣
  全數移除，無殘留呼叫端。ACDS_ContractRegistry.md 第二節 Staging 欄位
  合約由 8 欄擴為 9 欄；第三節 Mod_FileSystem／Mod_InteractionAdapter／
  cls_ImageEntity 函式簽章需同步更新。MSG_SESSION_PROMPT／
  MSG_EXPORT_MODE_PROMPT 兩則使用者提示文字措辭配合新邏輯調整。
- 附記（V17，非獨立架構決策，記於此處避免散落）：
  稽核發現 `cls_Settings.ForbiddenChars` 黑名單缺漏 `#`。本專案將清洗後檔名
  寫入 `LOG_Export_成果` D 欄超連結，而 `#` 在 Excel HYPERLINK 位址中為錨點
  分隔符保留字元，檔名含 `#` 會導致超連結指向錯誤子位置或失效。黑名單新增
  `#`，`CleanAndTruncateFilename` 源頭清洗，不需要另立 ADR（屬第十章
  Definition of Done 範圍內的輸入驗證補正，非新架構決策，無放棄的替代方案
  需要記錄）。
============================================================

【ADR-018】帳本與硬碟狀態不一致防呆：匯入端幽靈紀錄封存、成果端已匯出旗標重置
- 日期：2026-09（V13 → V17 過渡，緊接 ADR-017 之後）
- 狀態：Accepted
- 背景：
  ADR-017 讓 Staging_Images 的「已匯出」欄位成為 PackageStagingTasks 過濾邏輯
  的事實來源，但 Staging 表（帳本）與硬碟實體檔案是兩個獨立系統，使用者若
  手動刪除「匯入圖片」或「成果」資料夾（或整個搬移），帳本紀錄不會自動同步
  更新，導致兩種卡死情境：① 帳本顯示已有匯入紀錄，但「匯入圖片」資料夾實體
  是空的——MD5 去重邏輯會把使用者想重新匯入的照片誤判為已存在而擋下；
  ② 帳本顯示已有「已匯出」紀錄，但「成果」資料夾是空的——PackageStagingTasks
  會把這些照片全部濾掉，使用者清空成果資料夾後重跑 Phase 2，卻得到「無資料
  可執行」的警告，兩種情境使用者皆無法自行排除。
- 決策：
  兩處新增偵測，皆以「帳本筆數 > 0 但對應資料夾實體檔案數 = 0」為判準：
  ① RunPhase1 開新案判斷之前，偵測到匯入端不一致，呼叫新增之
  Mod_InteractionAdapter.ArchiveStaleStagingToOldCatalogLog 把目前 Staging
  表整批（不經剪貼簿、直注技術）附加進單一固定分頁「舊目錄LOG_總帳」
  （每批資料前插入時間戳記分隔標記），再清空 Staging 表，使用者可重新匯入。
  ② RunPhase2 判斷斷點續跑之前，偵測到成果端不一致，呼叫新增之
  Mod_InteractionAdapter.ResetAllExportedFlags 把 Staging 表全部「已匯出」
  狀態清空，不需要封存（原始照片複製品通常仍在「匯入圖片」，不構成資訊
  損失），讓照片自然於本次執行重新被打包重印。
  「舊目錄LOG_總帳」的設計比照既有 Mod_LogManager 的「單一固定分頁、直注技術
  累加」精神（不比照 ArchiveLogSheet 逐次另開新分頁的模式），刻意不使用
  wsStaging.Copy（該方法會自動切換使用者畫面焦點到新分頁，且逐次拷貝會使
  分頁數量無限累積，需要之後再手動壓縮合併，與本次重構「不製造需要事後清理
  的債務」的精神矛盾）。函式全程關閉 ScreenUpdating，結束後主動 Activate
  回 ConfigSheet，避免畫面跳動。
- 放棄的替代方案：
  (a) 匯入端與成果端比照辦理，都封存到獨立分頁——放棄理由：成果端的照片
      來源複製品通常還在「匯入圖片」，資訊沒有真正遺失，封存分頁只是製造
      不必要的稽核噪音；只有匯入端是真正的「帳本紀錄即將作廢」情境，才需要
      封存存查。
  (b) 用 wsStaging.Copy 逐次複製整張表格為新分頁封存——放棄理由：見決策段，
      會造成畫面跳動且分頁無限累積，需要之後再合併，違反本次重構初衷。
  (c) 讓幽靈紀錄直接合併進既有 LOG_Master_Archive 總帳——放棄理由：兩者
      欄位語意不相容（LOG_Master_Archive 的 C／D 欄是「來源工作表」「最終
      成品路徑」，Staging 的對應欄位是「使用者備註」「最新更新日期」），
      強行合併會讓總帳同一張表混雜兩種不同語意的資料，日後無法分辨某一列
      究竟是匯出紀錄還是幽靈匯入紀錄，違反 SSOT 精神。
- 影響：
  新增 Mod_InteractionAdapter.ArchiveStaleStagingToOldCatalogLog／
  ResetAllExportedFlags／GetExportedRecordCount 三支函式，`SHEET_OLD_CATALOG_
  MASTER`／`MSG_OLD_CATALOG_BATCH_MARK`／`MSG_LEDGER_ORPHANED_HEADER／FOOTER`／
  `MSG_RESULT_ORPHANED_HEADER／FOOTER` 等新常數。活頁簿新增一張固定分頁
  「舊目錄LOG_總帳」（首次觸發時自動建立，欄位結構沿用 Staging_Images 表頭）。
============================================================
