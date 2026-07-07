/**
 * ============================================================
 * 嘉義今順興 年菜線上訂購 — 後台程式(Google Apps Script)
 * ------------------------------------------------------------
 * 功能:
 *  1. 接收網站訂單 → 自動寫入 Google 試算表「訂單」工作表
 *  2. 同步寫入「明細」工作表(每道菜一列,方便統計)
 *  3. 自動建立「品項統計」工作表(各品項總數量/總金額)
 *  4. 寄送訂購確認信給「客人」與「店家」雙方
 *
 * 安裝方式請看「後台設定說明.md」
 * ============================================================
 */

// ★ 店家接收訂單通知的信箱(可改成任何信箱)
const OWNER_EMAIL = 'jackize123@gmail.com';
const SHOP_NAME   = '嘉義今順興';
const SHOP_PHONE  = '0928-712056';
const SHOP_ADDR   = '嘉義縣民雄鄉秀林村東義11-58號';

/** 測試用:瀏覽器直接開啟部署網址會看到這行字 */
function doGet() {
  return ContentService.createTextOutput('✅ 今順興訂購系統後台運作中');
}

/** 接收前台訂單 */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---- 訂單總表 ----
    let sh = ss.getSheetByName('訂單');
    if (!sh) {
      sh = ss.insertSheet('訂單', 0);
      sh.appendRow(['訂單編號','訂購時間','姓名','電話','Email','配送方式',
                    '門市/地址','付款方式','訂購明細','商品小計','運費','總金額',
                    '預估箱數','備註','狀態']);
      sh.setFrozenRows(1);
      sh.getRange('1:1').setFontWeight('bold').setBackground('#8f1616').setFontColor('#ffffff');
    }
    sh.appendRow([
      d.orderId, new Date(), d.name, "'" + d.phone, d.email,
      d.shipLabel, d.store || '', d.payLabel, d.detail,
      d.subtotal, d.fee, d.total, d.boxes || '', d.note || '', '新訂單'
    ]);

    // ---- 明細表(每道菜一列) ----
    let dt = ss.getSheetByName('明細');
    if (!dt) {
      dt = ss.insertSheet('明細');
      dt.appendRow(['訂單編號','訂購時間','品項','數量','金額']);
      dt.setFrozenRows(1);
      dt.getRange('1:1').setFontWeight('bold').setBackground('#b3872a').setFontColor('#ffffff');
    }
    (d.items || []).forEach(function(it) {
      dt.appendRow([d.orderId, new Date(), it.name, it.qty, it.amount]);
    });

    // ---- 品項統計表(自動彙整,可直接匯出) ----
    let st = ss.getSheetByName('品項統計');
    if (!st) {
      st = ss.insertSheet('品項統計');
      st.getRange('A1').setFormula(
        '=QUERY(明細!A:E,"select C, sum(D), sum(E) where C is not null group by C order by sum(D) desc label C \'品項\', sum(D) \'總數量\', sum(E) \'總金額\'",1)'
      );
    }

    // ---- 寄信給客人 ----
    const rows = (d.items || []).map(function(it) {
      return '<tr><td style="padding:6px 12px;border-bottom:1px solid #eee">' + it.name +
             '</td><td style="padding:6px 12px;border-bottom:1px solid #eee;text-align:center">' + it.qty +
             '</td><td style="padding:6px 12px;border-bottom:1px solid #eee;text-align:right">NT$ ' + it.amount.toLocaleString() + '</td></tr>';
    }).join('');
    const tableHtml =
      '<table style="border-collapse:collapse;width:100%;max-width:480px;font-size:14px">' +
      '<tr style="background:#8f1616;color:#fff"><th style="padding:8px 12px">品項</th><th style="padding:8px 12px">數量</th><th style="padding:8px 12px">金額</th></tr>' +
      rows +
      '<tr><td colspan="2" style="padding:6px 12px;text-align:right">商品小計</td><td style="padding:6px 12px;text-align:right">NT$ ' + d.subtotal.toLocaleString() + '</td></tr>' +
      '<tr><td colspan="2" style="padding:6px 12px;text-align:right">運費</td><td style="padding:6px 12px;text-align:right">NT$ ' + d.fee.toLocaleString() + '</td></tr>' +
      '<tr style="font-weight:bold;color:#8f1616"><td colspan="2" style="padding:8px 12px;text-align:right">總金額</td><td style="padding:8px 12px;text-align:right">NT$ ' + d.total.toLocaleString() + '</td></tr>' +
      '</table>';

    const custHtml =
      '<div style="font-family:Microsoft JhengHei,sans-serif;color:#3a2413">' +
      '<h2 style="color:#8f1616">' + SHOP_NAME + ' 訂購成功確認</h2>' +
      '<p>' + d.name + ' 您好,感謝您的訂購!以下是您的訂單內容:</p>' +
      '<p><b>訂單編號:</b>' + d.orderId + '</p>' + tableHtml +
      '<p><b>配送方式:</b>' + d.shipLabel + (d.store ? '<br><b>門市/地址:</b>' + d.store : '') +
      '<br><b>付款方式:</b>' + d.payLabel +
      (d.note ? '<br><b>備註:</b>' + d.note : '') + '</p>' +
      '<p>店家將盡快與您電話確認取貨時間與付款細節。<br>' +
      '如需大量訂購或辦理外燴,歡迎致電/LINE:<b>' + SHOP_PHONE + '</b></p>' +
      '<p style="color:#888;font-size:12px">取貨地址:' + SHOP_ADDR + '</p></div>';

    if (d.email) {
      GmailApp.sendEmail(d.email, '【' + SHOP_NAME + '】訂購成功確認 - 訂單 ' + d.orderId,
        '您的訂單 ' + d.orderId + ' 已成立,總金額 NT$' + d.total,
        { htmlBody: custHtml, name: SHOP_NAME });
    }

    // ---- 寄信給店家 ----
    const ownerHtml =
      '<div style="font-family:Microsoft JhengHei,sans-serif;color:#3a2413">' +
      '<h2 style="color:#8f1616">新訂單 ' + d.orderId + '</h2>' +
      '<p><b>姓名:</b>' + d.name + '<br><b>電話:</b>' + d.phone +
      '<br><b>Email:</b>' + d.email + '</p>' + tableHtml +
      '<p><b>配送:</b>' + d.shipLabel + (d.store ? '<br><b>門市/地址:</b>' + d.store : '') +
      (d.boxes ? '<br><b>預估箱數:</b>' + d.boxes : '') +
      '<br><b>付款:</b>' + d.payLabel +
      (d.note ? '<br><b>備註:</b>' + d.note : '') + '</p>' +
      '<p><a href="' + ss.getUrl() + '" style="color:#8f1616;font-weight:bold">開啟訂單試算表查看完整資料</a></p></div>';

    GmailApp.sendEmail(OWNER_EMAIL,
      '【新訂單】' + d.orderId + ' ' + d.name + ' NT$' + d.total.toLocaleString(),
      '新訂單 ' + d.orderId + ',請開啟試算表查看',
      { htmlBody: ownerHtml, name: SHOP_NAME + ' 訂購系統' });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, orderId: d.orderId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
