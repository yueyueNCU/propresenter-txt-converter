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

1. 執行轉換腳本。
2. 打開 ProPresenter。
3. 匯入 [transform_data/新聖詩](transform_data/新聖詩) 或 [transform_data/啟應文](transform_data/啟應文) 裡的 TXT。
4. 匯入文字時，使用空白行分隔投影片。

## 注意事項

- 檔名和 TXT 第一行標題都會去掉數字前面的 `0`。
- 目前啟應文原始資料缺少 `新4.pptx`，所以輸出中不會有 `啟應文4.txt`。
- 新聖詩 DOCX 原始資料目前看起來 `639` 後直接跳到 `645`，所以 `640–644` 可能不存在於原始 DOCX 或格式不同。
