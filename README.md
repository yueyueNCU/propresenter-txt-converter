# ProPresenter 簡報文字轉換工具

這個工作區用來把教會簡報資料轉成 ProPresenter 可匯入的 UTF-8 `.txt` 檔案。

目前支援：

- 新聖詩
- 啟應文

轉出的 TXT 會放在 [transform_data](transform_data) 底下，並以空白行分隔每張投影片，方便 ProPresenter 文字匯入。

## 資料夾結構

```text
raw_data/
  啟應文/              # 啟應文 PPTX 原始檔
  新聖詩/
    Docs/             # 新聖詩完整 DOCX 歌詞檔
    PPT/              # 新聖詩 PPTX 原始檔
transform_data/
  啟應文/              # 轉出的啟應文 TXT
  新聖詩/              # 轉出的新聖詩 TXT
```

## 執行環境

此專案使用 Python 標準函式庫處理 `.docx` / `.pptx`，不需要另外安裝套件。

目前可用的 Python 指令：

```bash
/Users/enyew/.local/bin/python3.14
```

如果在其他電腦執行，也可以改用：

```bash
python3
```

## 新聖詩校對說明

本專案的[新聖詩 TXT](transform_data/新聖詩) 已在原始文件轉換後，參考原始簡報字形、網路公開聖詩譜頁及字型資料，進一步校對。此次校對以缺字修復及新增台語版本核對為主：

- 將轉換後顯示異常的缺字、私用區字元還原為對應的台語或客語漢字，例如 `` → `𣍐`，以及「莿𲼊」中的「𲼊（草字頭＋帕）」。
- 清除原文件留下的非歌詞排版字元。
- 第 404、464、571 首保留原有客語歌詞並修復缺字，另新增檔名標示「（台語）」的台語漢字版本。其中第 571 首的台語漢字由官方台語白話字轉寫。

校對來源、版本差異與轉寫依據詳見[台語版本與缺字校對](raw_data/新聖詩/台語版本與缺字校對.md)。

現有新聖詩 TXT 已包含上述校對成果，可直接匯入 ProPresenter。這些修正尚未納入轉換腳本；如需重新從原始文件轉換，請透過 `--output-dir` 指定其他資料夾，以保留已校對的 TXT。

## 轉換新聖詩

腳本：[convert_hymns_to_propresenter_txt.py](convert_hymns_to_propresenter_txt.py)

預設使用完整歌詞 DOCX：

- [raw_data/新聖詩/Docs/新聖詩歌詞1-300.docx](raw_data/新聖詩/Docs/新聖詩歌詞1-300.docx)
- [raw_data/新聖詩/Docs/新聖詩歌詞301-650.docx](raw_data/新聖詩/Docs/新聖詩歌詞301-650.docx)

執行：

```bash
/Users/enyew/.local/bin/python3.14 convert_hymns_to_propresenter_txt.py --source docs --clean-output --overwrite
```

輸出到：

[transform_data/新聖詩](transform_data/新聖詩)

輸出格式：

```text
新聖詩1首.txt
新聖詩2首.txt
新聖詩650首.txt
```

每個 TXT 第一行會是相同標題，例如：

```text
新聖詩1首
```

### 新聖詩參數

```bash
/Users/enyew/.local/bin/python3.14 convert_hymns_to_propresenter_txt.py --help
```

常用參數：

- `--source docs`：從 [raw_data/新聖詩/Docs](raw_data/新聖詩/Docs) 的完整 DOCX 歌詞轉換。
- `--source ppt`：從 [raw_data/新聖詩/PPT](raw_data/新聖詩/PPT) 的 PPTX 轉換。
- `--clean-output`：轉換前清掉輸出資料夾內既有的 `新聖詩*.txt`。
- `--overwrite`：允許覆蓋既有 TXT。
- `--output-dir`：指定輸出資料夾。

使用 PPTX 來源時：

```bash
/Users/enyew/.local/bin/python3.14 convert_hymns_to_propresenter_txt.py --source ppt --clean-output --overwrite
```

## 轉換啟應文

腳本：[convert_responsive_readings_to_propresenter_txt.py](convert_responsive_readings_to_propresenter_txt.py)

輸入來源：

[raw_data/啟應文](raw_data/啟應文)

執行：

```bash
/Users/enyew/.local/bin/python3.14 convert_responsive_readings_to_propresenter_txt.py --clean-output --overwrite
```

輸出到：

[transform_data/啟應文](transform_data/啟應文)

輸出格式：

```text
啟應文1.txt
啟應文2.txt
啟應文66.txt
```

每個 TXT 第一行會是相同標題，例如：

```text
啟應文1
```

腳本會自動移除 PPT 裡的頁首，例如：

```text
啟應文1(1/4)
```

### 啟應文參數

```bash
/Users/enyew/.local/bin/python3.14 convert_responsive_readings_to_propresenter_txt.py --help
```

常用參數：

- `--input-dir`：指定啟應文 PPTX 原始檔資料夾。
- `--output-dir`：指定輸出資料夾。
- `--clean-output`：轉換前清掉輸出資料夾內既有的 `啟應文*.txt`。
- `--overwrite`：允許覆蓋既有 TXT。

## ProPresenter 匯入方式

1. 準備要匯入的 TXT；使用本專案現有的新聖詩 TXT 時，已包含校對成果，可直接進行下一步。若需重新轉換，請先依上方說明執行腳本。
2. 打開 ProPresenter。
3. 匯入 [transform_data/新聖詩](transform_data/新聖詩) 或 [transform_data/啟應文](transform_data/啟應文) 裡的 TXT。
4. 匯入文字時，使用空白行分隔投影片。

## 建議字型

新聖詩包含台語及客語的罕用漢字。建議使用 **HanaMin（花園明朝）** 或 **Jigmo（字雲）**，並安裝各系列的完整字型檔案。

### 下載來源與版本

- **HanaMin**：[官方下載頁](https://glyphwiki.org/hanazono/)／[2017-09-04 版 ZIP](https://glyphwiki.org/hanazono/hanazono-20170904.zip)。內含 `HanaMinA.ttf`、`HanaMinB.ttf`；字型選單也可能顯示「花園明朝A」「花園明朝B」。
- **Jigmo**：[官方下載頁](https://kamichikoichi.github.io/jigmo/)／[2025-09-12 版 ZIP](https://kamichikoichi.github.io/jigmo/Jigmo-20250912.zip)。請選擇 **2025-09-12 或較新版**，並安裝 `Jigmo.ttf`、`Jigmo2.ttf`、`Jigmo3.ttf`，以支援本專案使用的 Unicode 17 新字。

HanaMin 與 Jigmo 可以同時安裝。兩套都提供免費使用的授權，詳情見各下載包的授權文件。

### 各字型的用途

| 字型名稱 | 主要用途 | 範例 |
| --- | --- | --- |
| **HanaMinA／花園明朝A** | 常用漢字與部分擴充字，適合作為一般歌詞字型 | 我、主、愛 |
| **HanaMinB／花園明朝B** | 擴充區 B–F 的罕用漢字，可補充許多台語用字 | 𣍐、𪜶、𠢕 |
| **Jigmo** | 常用漢字與擴充區 A，適合作為一般歌詞字型 | 我、主、愛 |
| **Jigmo2** | 擴充區 B、C、D、E、F、I 的漢字，可補充許多台語用字 | 𣍐、𪜶、𠢕、𫢶 |
| **Jigmo3** | 擴充區 G、H、J 的漢字，包含較新的罕用字 | 𲼊（草字頭＋帕） |

使用者已確認：在目前的 ProPresenter 環境中，選用 **Jigmo2** 時，「我、主、愛」仍可正常顯示。Jigmo 的三個檔案分別收錄不同範圍；經檢查，2025-09-12 版的 Jigmo2、Jigmo3 字型檔本身未收錄這三個字，因此這種顯示結果可能來自程式或系統自動使用其他字型補字（font fallback）。字型選單顯示 Jigmo2，不代表每個字都由 Jigmo2 提供字形；目前尚未確認該環境實際使用的補字字型。

本專案已核對：2025-09-12 版的 Jigmo、Jigmo2、Jigmo3 合計涵蓋目前新聖詩 TXT 使用的全部漢字。這項核對確認字型有收錄相應文字，實際投影前仍應檢查 ProPresenter 的顯示結果。

### 安裝與 ProPresenter 使用建議

1. 下載並解壓縮字型包，安裝需要的 `.ttf` 檔案。Windows 可在字型檔上按右鍵選擇「安裝」；macOS 可透過「字體簿」安裝。
2. 更新同名字型時，確認安裝的是新版，完成後重新開啟 ProPresenter／Word，讓程式重新載入字型。
3. 若目前在 ProPresenter 選用 **Jigmo2** 時，歌詞已能正常顯示，可繼續使用此設定。初次設定也可先使用 **Jigmo** 或 **HanaMinA**；若有缺字，再依上表選用收錄該字的字型。
4. 「莿𲼊」中的 **𲼊（草字頭＋帕，U+32F0A）** 請使用 **2025-09-12 或較新版的 Jigmo3**。HanaMin 2017 版及 Jigmo3 2023-08-16 版均未收錄這個字。

安裝後可用 `我、主、愛、𣍐、𪜶、𠢕、𫢶、𲼊` 檢查一般漢字與罕用字的顯示。更多字形核對資料見[台語版本與缺字校對](raw_data/新聖詩/台語版本與缺字校對.md)。

## 注意事項

- 檔名和 TXT 第一行標題都會去掉數字前面的 `0`。
- 目前啟應文原始資料缺少 `新4.pptx`，所以輸出中不會有 `啟應文4.txt`。
- 新聖詩 DOCX 原始資料目前看起來 `639` 後直接跳到 `645`，所以 `640–644` 可能不存在於原始 DOCX 或格式不同。
