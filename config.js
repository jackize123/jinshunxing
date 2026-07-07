/* ============================================================
   今順興訂購系統 前台設定檔
   ------------------------------------------------------------
   ORDER_API_URL:Google Apps Script 部署後的「網頁應用程式網址」
   (以 https://script.google.com/macros/s/......./exec 結尾)
   設定完成後,訂單會自動寫入 Google 試算表,
   並寄送確認信給店家與訂購人。

   尚未設定時(留空字串),送出訂單會改用開啟郵件軟體的備援方式。
   設定教學請看 repo 內的「後台設定說明.md」。
   ============================================================ */
window.ORDER_API_URL = "https://script.google.com/macros/s/AKfycbyZnTvT7bm7tl41ob0r8XwG7rmzTgZtdxh4d8ORVtTgBtWCUxDygt5tc2fuXyB08A-hVA/exec";

/* 店家接收訂單通知的信箱(備援 mailto 用;Apps Script 內另有設定) */
window.OWNER_EMAIL = "claude04@ydlo.ylc.edu.tw";
