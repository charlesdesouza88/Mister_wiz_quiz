// ============================================================
//  MISTER WIZ – Quiz Lead Capture + Email Notification
//  Google Apps Script  |  Cole em script.google.com
//  Notificação via Gmail (gratuito, sem credenciais extras)
// ============================================================

// ── CONFIGURAÇÕES ── (edite apenas aqui)
const CONFIG = {
  SHEET_NAME:    "Leads Quiz",
  NOTIFY_EMAIL:  "misterwizcaratinga@gmail.com",
};

// ── ENTRY POINT ── chamado pelo quiz via fetch POST
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    saveToSheet(data);
    if (data.finished) sendEmail(data);
    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ── SALVA NA PLANILHA ──
function saveToSheet(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow([
      "Data/Hora", "Nome", "Idade", "Cidade", "WhatsApp",
      "Acertos", "Resultado", "Código Voucher"
    ]);
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold")
         .setBackground("#752481").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  const now    = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");
  const result = data.score == 5 ? "✅ GANHOU" : `❌ ${data.score}/5`;

  sheet.appendRow([
    now,
    data.name  || "",
    data.age   || "",
    data.city  || "",
    data.phone || "",
    data.score ?? "",
    result,
    data.code  || ""
  ]);

  const lastRow = sheet.getLastRow();
  const rowColor = (lastRow % 2 === 0) ? "#F0E6F3" : "#FFFFFF";
  sheet.getRange(lastRow, 1, 1, 8).setBackground(rowColor);
}

// ── MANDA EMAIL VIA GMAIL (gratuito) ──
function sendEmail(data) {
  const ganhou  = data.score == 5;
  const result  = ganhou ? "🏆 GANHOU o prêmio!" : `❌ ${data.score}/5 acertos`;
  const voucher = data.code ? `<br><strong>🎟 Voucher:</strong> ${data.code}` : "";
  const hora    = Utilities.formatDate(new Date(), "America/Sao_Paulo", "HH:mm 'de' dd/MM/yyyy");
  const waLink  = `https://wa.me/55${data.phone.replace(/\D/g,"")}`;

  const subject = ganhou
    ? `🏆 [Mister Wiz Quiz] ${data.name} GANHOU! – ${hora}`
    : `📋 [Mister Wiz Quiz] Novo lead: ${data.name} – ${hora}`;

  const body = `
    <div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;border:1px solid #ddd;border-radius:12px;overflow:hidden;">
      <div style="background:#752481;padding:20px 24px;">
        <h2 style="color:#fff;margin:0;font-size:20px;">🎯 Novo Lead – Quiz Mister Wiz</h2>
        <p style="color:#F0A500;margin:4px 0 0;font-size:13px;">${hora}</p>
      </div>
      <div style="padding:20px 24px;background:#fff;">
        <table style="width:100%;border-collapse:collapse;font-size:15px;">
          <tr><td style="padding:6px 0;color:#888;">👤 Nome</td><td style="padding:6px 0;font-weight:bold;">${data.name}</td></tr>
          <tr><td style="padding:6px 0;color:#888;">🎂 Idade</td><td style="padding:6px 0;">${data.age} anos</td></tr>
          <tr><td style="padding:6px 0;color:#888;">📍 Cidade</td><td style="padding:6px 0;">${data.city}</td></tr>
          <tr><td style="padding:6px 0;color:#888;">📱 WhatsApp</td><td style="padding:6px 0;">${data.phone}</td></tr>
          <tr><td style="padding:6px 0;color:#888;">📊 Resultado</td><td style="padding:6px 0;">${result}${voucher}</td></tr>
        </table>
        <div style="margin-top:20px;">
          <a href="${waLink}" style="display:inline-block;background:#25D366;color:#fff;padding:10px 20px;border-radius:8px;text-decoration:none;font-weight:bold;">
            💬 Responder no WhatsApp
          </a>
        </div>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to:       CONFIG.NOTIFY_EMAIL,
    subject:  subject,
    htmlBody: body
  });
}

// ── UTILITÁRIO ──
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
