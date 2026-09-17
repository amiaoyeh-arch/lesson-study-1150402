---
name: lesson-study-deck
description: 學習共同體（SLC）公開課課例研究與觀課議課分享簡報全套設計技能。引導從課堂真實教學實踐（說課、觀課焦點、卡點診斷、滾動修正、同儕鷹架、高光頓悟）提煉課例故事，自動產出三合一交付成果：1. RWD 響應式流體排版與 JavaScript 智慧適配 HTML 互動簡報、2. 16:9 微軟圓體/正黑體 PPTX 簡報、3. 高解析 PDF 檔，並內建日系動漫繪本風格課堂插圖生成規範與自動排版工具。當使用者說「製作課例研究簡報」、「做課例分享簡報」、「觀課議課簡報」、「公開課成果報告」、「課例探究簡報」時載入。
---

# 學習共同體課例研究簡報技能 (lesson-study-deck)

本技能提供標準化、自動化且具深厚教育哲學底蘊的「學習共同體（Study of Learning Community, SLC）」公開課課例研究與議課成果簡報製作工作流。

---

## 🧭 核心教育哲學與視角
1. **學習事實的描述、詮釋與反思**：聚焦學生在課堂上的真實語言、眼神、肢體、猶豫與沉默，而非教師單向的教學技術演繹。
2. **跳躍任務（Jumping Task）與認知衝突**：呈現挑戰性任務如何引發學生的探索、卡關與幾何/概念逆推。
3. **同儕鷹架與流動的善意**：記錄小組共學中同伴如何相互傾聽、串聯、提示與協同解題。
4. **教學演進歷程（Design Iterations）**：對比公開課第 1 次實踐（發現卡點與迷思）到第 2 次滾動修正（鷹架重構與突破）的質變。

---

## 🎯 三合一標準交付格式

每次執行課例研究簡報任務，均產出以下三項成果：

1. 🌐 **響應式互動 HTML 簡報 (`*_v3.html` / `slides.html`)**：
   - **RWD 流體排版**：採用 CSS `clamp()` 與 Flexbox 佈局，完美適應任何螢幕解析度與縮放比率。
   - **智慧高度適配引擎（JavaScript Auto-Fit Engine）**：即時動態測量文字高度與卡片高度，自動微調字級（11px ~ 20px）與行距，**保證文字高度最多佔欄位 85%~90%**，垂直居中、不溢出、免捲動、永不裁切。
   - **互動操控**：鍵盤左右鍵/空白鍵翻頁、全螢幕切換（F 鍵）、快速目錄跳轉、夜間護眼主題。
2. 📊 **PowerPoint 簡報檔 (`*_v3.pptx` / `slides.pptx`)**：
   - 16:9 寬螢幕比例（13.333" × 7.5"）。
   - 中文字體預設使用 Windows/Office 內建之**微軟圓體風格（或微軟正黑體/微軟雅黑）**。
   - 左右雙欄對稱黃金比例，內建文字自動階梯縮放不溢出。
3. 📄 **高解析 PDF 匯出檔 (`*_v3.pdf` / `slides.pdf`)**。

---

## 📚 課例簡報五大標準架構（15-20 頁標準範本）

| 單元章節 | 建議頁數 | 核心內容與投影片焦點 |
| :--- | :--- | :--- |
| **一、 緣起與設計** | 1 ~ 4 頁 | • 封面與研究背景（年級、單元、執教教師、課堂演進時間線）<br>• 課例研究核心問題（Research Questions）<br>• UbD 理解為先 ✕ 學習共同體核心理念<br>• 跳躍任務（Jumping Task）設計情境與背後的學科密碼 |
| **二、 第一次實踐與診斷** | 5 ~ 7 頁 | • 公開課課堂現場氛圍與學生卡點（抓頭苦思、試錯盤旋）<br>• 臨床診斷：三大典型認知迷思分析<br>• 觀課與議課教師/輔導團反思洞察（如輔助線干擾、語言斷層） |
| **三、 滾動修正與重構** | 8 ~ 11 頁 | • 兩次教學設計對比（教案、學習單、情境圖卡修改）<br>• 關鍵轉化策略一：去蕪存菁與動態視覺解構<br>• 關鍵轉化策略二：教師提問與語言鷹架精準化<br>• 關鍵轉化策略三：小組共學任務細緻化與科技輔助 |
| **四、 第二次實踐與高光時刻** | 12 ~ 16 頁 | • 第二次課堂現場氛圍的質變（從小組沈默走向熱烈思辨）<br>• 課堂精彩對話實錄（學生語言、同儕引導、突破瞬間）<br>• 學習高光時刻：學生指著黑板恍然大悟（Aha! 頓悟瞬間）<br>• 同儕鷹架與流動的善意（共學互助與情感連結） |
| **五、 總結與專業反思** | 17 ~ 18 頁 | • 學習共同體課堂的質性轉變（課堂風景前後對比）<br>• 教師的專業成長與教學範式轉移（從「教」轉向「學」）<br>• 結語與致謝 |

---

## 🎨 AI 動漫與繪本風格插圖生成規範

課例簡報中每頁插圖應統一採用 **日系清新水彩／動漫繪本風格（Anime & Watercolor Illustration）**，呈現溫馨、專注的學習課堂氛圍。

### 核心 Prompt 模板

1. **封面與概覽**：
   > `Warm, inviting Japanese anime watercolor style illustration of a modern elementary math classroom in Taiwan. An inspiring female teacher with short dark hair in neat professional attire smiling gently. Elementary school students sitting in small groups, looking engaged. Soft daylight streaming through classroom windows, clean blackboard, pastel green and warm earth tones, 16:9 aspect ratio.`
2. **學生卡關苦思（迷思與挫折）**：
   > `Japanese anime watercolor style illustration of a 5th grade student sitting at a wooden desk with compass, ruler and graph paper, scratching their head in deep puzzled thought, looking at geometric circles. Classroom background softly blurred, empathetic mood, natural lighting, high detail.`
3. **學生指著黑板恍然大悟（Aha! 頓悟時刻）**：
   > `Vibrant and touching Japanese anime watercolor style illustration of a 5th grade student standing in front of a classroom blackboard, passionately pointing with a chalk or pointer at a geometric diagram. A bright realization (eureka / aha moment) spark in student's eyes and an excited happy smile. Fellow classmates looking on with admiration. Chalkboard shows clean geometric flower patterns. Warm uplifting sunlight.`
4. **小組圍坐共學與傾聽**：
   > `Japanese anime watercolor style illustration of four 5th grade elementary students (boys and girls) sitting close together around a desk in a collaborative group learning session. One student explains pointing to a worksheet, other three listening with genuine attention and friendly smiles. Natural warm classroom lighting, touching atmosphere of mutual care and peer scaffolding.`

---

## 🛠️ 自動化腳本工具庫 (`scripts/`)

本技能提供全自動化構建工具 `scripts/build_lesson_study_deck.py`：

```bash
# 從課例設定 JSON 檔自動編譯生成 HTML、PPTX 與 PDF
python scripts/build_lesson_study_deck.py --config case_study_data.json --output-dir output/
```

### 支援參數：
- `--config`：包含各頁標題、副標、標籤、內文要點（bullets）、講稿（notes）與插圖路徑之 JSON 設定檔。
- `--output-dir`：產出目錄（自動產生 `*_v3.html`、`*_v3.pptx` 與 `*_v3.pdf`）。
- `--theme`：配色主題（預設 `learning` 溫潤有機茶白/松針綠、可選 `pastel`、`blue`、`modern`）。

---

## 📋 開工與執行 Checklists

當教師提出新的課例研究需求時，依序執行：
1. **收集基本課例素材**：執教教師、公開課日期/班級、單元主題、核心任務、兩次實踐對比紀錄、觀課照片或逐字稿。
2. **提煉 5 大架構 15-18 頁內容大綱**：編寫每頁標題、內文重點、講稿與視覺畫面構想。
3. **生圖與配圖**：利用 `generate_image` 生成日系動漫風格插圖或處理實體照片。
4. **執行腳本產出全套交付成果**：一鍵產生 RWD HTML、16:9 微軟圓體 PPTX 與 PDF。
5. **完整性驗證**：確認 HTML 各頁文字 100% 完整收納不裁切、字級自適應良好。