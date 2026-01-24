# 專案狀態（給 AI 快速理解用 0124 1044）

> 精簡、好維護；每天更新 1～2 行即可。

---

## 1. 專案一句話說明

Excel 教學用表單：學生輸入學號後，在瀏覽器裡編輯 Model Data / Period Data 多個工作表，可上傳 xlsx、下載、自動存到本機。純前端，無後端。

---

## 2. 目前檔案結構

```
Model_Period_WebApp/
├── index.html
├── css/
│   └── style.css
├── js/
│   ├── app.js           # 主邏輯、state、renderTable、儲存、上傳下載
│   ├── table-selection.js   # 拖曳選取、Delete、複製/剪下/貼上（TSV）
│   └── template-data.js     # 工作表結構、欄位、預設資料（TEMPLATE_DATA）
```

---

## 3. 已完成的功能

- [x] 學號輸入 modal，依學號存到 `localStorage`
- [x] Model Data / Period Data 切換，分組 sheet 導覽
- [x] 表格動態產生（`renderTable`），每個 cell 為 `<td><input></td>`
- [x] 必填驗證、紅框、驗證摘要
- [x] 新增行、刪除行
- [x] 自動儲存（debounce）
- [x] 上傳 xlsx、下載 xlsx、Reset（含確認 modal）
- [x] 拖曳選取、Delete 清空選取格、複製/剪下/貼上（Excel 相容 TSV）

---

## 4. 目前卡住 / 問題

1. （目前無；有再填）
2. —
3. —

---

## 5. 下一步只做的一件事

（例如：補齊某個 sheet 的必填規則 / 改善 Upload 失敗時的錯誤訊息 / … 只寫 1 件）

---

## 6. AI 協作規則

- **優先最小修改**：能改少就不要改多。
- **不隨便新增 JS 檔**：新功能先放進 `app.js` 或 `table-selection.js`、`template-data.js`，真的結構變大再考慮拆分。
- **修改前要說明**：改哪個檔案、哪個區塊（函式名或行數）、為什麼。

---

*最後更新：可以寫日期或「無」*
