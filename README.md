打包說明

產物位置
- 可執行檔: dist\vendor_order_parser.exe

放置外部資料（不打包）
- 請將「廠商名單.xlsx」放在與 exe 同一資料夾，程式會從該檔讀取寄件廠商對照表。
- 若要使用 Google Sheets 上傳功能，請將 service account 的 JSON（金鑰）放在同一資料夾，並確保您已將該 service account 的 email 加為目標試算表的編輯者。

執行方式
- 直接執行 dist\vendor_order_parser.exe（雙擊或在命令列執行）。
- 若出現 Google Sheets 權限或 API 問題，程式會自動回退成輸出 Excel 檔案（保存在目前工作目錄下）。

備註
- 我已按您的要求把測試程式碼與開發用資料保留在原路徑，不會打包這些測試檔案。
- 打包時已使用您的圖示: ChatGPTImage_cat.png（已內嵌於 exe）。

聯絡/後續
- 如要我把 exe 與 README 壓成一個 zip（方便發給同事），回覆「請建立 zip」即可，我會把 dist\vendor_order_parser.exe 與 README 打包成 dist\vendor_order_parser_package.zip 並回報結果。
