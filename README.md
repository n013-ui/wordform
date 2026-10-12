# 專題寫作排版工具（wordform）

把 Gemini「專題寫作教練」Gem 給的「排版用文字」貼上，直接在瀏覽器產生排好版的 Word 檔（.docx）。

- 版面：A4、上下邊界 1.2 公分、左右邊界 2 公分、標楷體 12 號、固定行高 20 點（可改 18 點）
- 「一、」「（一）」「1.」自動凸排縮排，小標粗體
- 表格：表頭淺藍底、第一欄淺灰底、整個表格不跨頁
- Napkin 概念圖預留虛線框
- 依固定行高估算每題頁數，並給出可以直接貼回 Gemini 的修改提示句

所有處理都在瀏覽器完成，不會上傳任何資料。

使用網址：https://n013-ui.github.io/wordform/

Word 檔產生使用 [docx](https://github.com/dolanmiu/docx) 9.7.1（MIT License，見 LICENSE-docx.txt）。

## 簡報範本填入工具（ppt/）

把 Gem 產生的簡報內容（JSON）填進學生自己的 PPT 範本，保留範本格式，講稿放進備忘稿。使用網址：https://n013-ui.github.io/wordform/ppt/

在瀏覽器中以 Pyodide 0.26.4 執行 python-pptx 1.0.2（皆為開放原始碼授權），所有檔案都在本機處理。
