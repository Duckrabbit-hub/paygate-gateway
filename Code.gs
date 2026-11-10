// ==================== GOOGLE APPS SCRIPT CONFIGURATION ====================
const CONFIG = {
  SHEET_NAME: "Transactions",
  SLIP_FOLDER_NAME: "Paygate_Slips",
  ADMIN_USERNAME: "admin",
  ADMIN_PASSWORD: "PaygatePassword123!",

  // 1. ระบบแจ้งเตือน Telegram (ใส่ค่าว่าง "" หากไม่ใช้งาน)
  TELEGRAM_BOT_TOKEN: "", // ตัวอย่าง: "123456789:ABCdefGhIJKlmNoPQRsTUVwxyZ"
  TELEGRAM_CHAT_ID: "",   // ตัวอย่าง: "987654321"

  // 2. ระบบแจ้งเตือน LINE Messaging API (ใส่ค่าว่าง "" หากไม่ใช้งาน)
  LINE_CHANNEL_ACCESS_TOKEN: "", // Channel Access Token จาก LINE Developers
  LINE_USER_ID: ""               // Your User ID จาก LINE Developers (Uxxxxxxxx...)
};

function getFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.SLIP_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.SLIP_FOLDER_NAME);
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    const data = ss.getDataRange().getValues();
    if (data.length <= 1) return jsonResponse({ status: "success", data: [] });

    const headers = data[0];
    const rows = data.slice(1);
    const result = rows.map((row, index) => {
      let obj = { rowIndex: index + 2 };
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
    return jsonResponse({ status: "success", data: result.reverse() });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;

    // ระบบ Login
    if (action === "login") {
      if (postData.username === CONFIG.ADMIN_USERNAME && postData.password === CONFIG.ADMIN_PASSWORD) {
        return jsonResponse({ status: "success", token: "AUTH_" + Date.now() });
      }
      return jsonResponse({ status: "error", message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);

    // อัปเดตสถานะ (อนุมัติ / ปฏิเสธ)
    if (action === "update_status") {
      ss.getRange(postData.rowIndex, 7).setValue(postData.status);
      return jsonResponse({ status: "success" });
    }

    // สร้างรายการธุรกรรมใหม่และบันทึกสลิป
    if (action === "create_transaction") {
      let slipUrl = "";
      if (postData.slip_base64 && postData.slip_filename) {
        try {
          const folder = getFolder();
          const contentType = postData.slip_base64.substring(5, postData.slip_base64.indexOf(";"));
          const bytes = Utilities.base64Decode(postData.slip_base64.split(",")[1]);
          const blob = Utilities.newBlob(bytes, contentType, `SLIP_${Date.now()}_${postData.slip_filename}`);
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          slipUrl = `https://lh3.googleusercontent.com/d/${file.getId()}`;
        } catch (e) {
          slipUrl = "";
        }
      }

      const newId = "TXN" + Date.now().toString().slice(-6);
      const time = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
      const amount = Number(postData.amount) || 0;
      const user = postData.user_id || "Customer";

      ss.appendRow([newId, time, user, amount, postData.bank_account || "PromptPay QR", slipUrl, "Pending", ""]);

      // ส่งสัญญาณแจ้งเตือนอัตโนมัติ
      sendAlert(newId, user, amount, time, slipUrl);

      return jsonResponse({ status: "success", id: newId });
    }

    return jsonResponse({ status: "error", message: "Invalid action" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function sendAlert(txnId, user, amount, time, slipUrl) {
  const msg = `🔔 มีรายการแจ้งโอนเงินใหม่!
` +
              `🆔 รหัส: ${txnId}
` +
              `👤 ผู้โอน: ${user}
` +
              `💰 ยอดเงิน: ฿${amount.toLocaleString()} บาท
` +
              `⏰ เวลา: ${time}
` +
              `📎 สลิป: ${slipUrl || 'ไม่มีรูปภาพ'}`;

  // 1. ส่งเข้า Telegram
  if (CONFIG.TELEGRAM_BOT_TOKEN && CONFIG.TELEGRAM_CHAT_ID) {
    try {
      UrlFetchApp.fetch(`https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ chat_id: CONFIG.TELEGRAM_CHAT_ID, text: msg }),
        muteHttpExceptions: true
      });
    } catch (e) {}
  }

  // 2. ส่งเข้า LINE Messaging API
  if (CONFIG.LINE_CHANNEL_ACCESS_TOKEN && CONFIG.LINE_USER_ID) {
    try {
      UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
        method: "post",
        headers: {
          "Content-Type": "application/json",
          "Authorization": "Bearer " + CONFIG.LINE_CHANNEL_ACCESS_TOKEN
        },
        payload: JSON.stringify({
          to: CONFIG.LINE_USER_ID,
          messages: [{ type: "text", text: msg }]
        }),
        muteHttpExceptions: true
      });
    } catch (e) {}
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
