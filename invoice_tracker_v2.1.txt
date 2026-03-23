// ============================================================
// 212 INVOICE TRACKER v3.0 — MASTERPLAN
// CFO-Grade Fatura Takip Sistemi
// Gemini AI + 5 Katmanlı Koruma + VIP + Dashboard
// ============================================================
//
// KURULUM:
// 1. Google Sheet aç → Extensions → Apps Script
// 2. Bu kodu yapıştır, Ctrl+S
// 3. Sol menü → Services → Drive API → Add
// 4. Script Properties → GEMINI_API_KEY ekle
//    (https://aistudio.google.com/apikey)
// 5. setupTriggers() çalıştır → yetki ver
// 6. scanInvoices() çalıştır (ilk tarama)
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
var CONFIG = {
  SPREADSHEET_ID: "1F4CSD6AQnPaLfNr9dOxeIC8caXVkLiwZXIuz3e2YHrg",
  NOTIFICATION_EMAIL: "dogukan@212.vc",
  APPROVAL_EMAIL_TO: "",
  MY_SIGNATURE: "Doğukan",

  SCAN_DAYS: 7,
  PDF_LIMIT: 30,
  MAX_GEMINI_PER_SCAN: 12,
  GEMINI_DELAY_MS: 4500,
  THREAD_LIMIT: 75,
  EXECUTION_LIMIT: 240000,

  // Sheet isimleri
  SHEET_FOUND: "📧 Faturalar",
  SHEET_CROSSCHECK: "🔍 Cross-Check",
  SHEET_SUSPICIOUS: "⚠️ Şüpheli",
  SHEET_KNOWN: "📋 Bilinen Faturalar",
  SHEET_VENDORS: "🏷️ Eksik Vendorlar",
  SHEET_DASHBOARD: "📊 Dashboard",
  SHEET_LOG: "📝 Log",

  // Fatura anahtar kelimeleri
  INVOICE_KEYWORDS: [
    "invoice","fatura","rechnung","facture","payment due",
    "amount due","fee note","debit note","tax invoice",
    "proforma","credit note","ödeme","hesap özeti","borç dekontu",
    "receipt","makbuz","billing","subscription"
  ],

  // Denetim kelimeleri (Gemini "fatura değil" dese bile bu varsa şüpheli)
  AUDIT_KEYWORDS: [
    "total","toplam","amount due","vat","kdv","iban",
    "swift","fatura no","invoice no","bank account","wire transfer"
  ],

  // Ödenmiş fatura tespiti
  ALREADY_PAID_INDICATORS: [
    "payment made","balance due: $0","balance due: 0",
    "already paid","paid in full"
  ],

  // Domain exclude — tüm @212.vc mailleri (VIP hariç)
  DOMAIN_EXCLUDE: ["@212.vc"],

  // Ek exclude
  EXCLUDE_SENDERS: ["noreply","stripe","iyzico"],
  EXCLUDE_SUBJECTS: [
    "expense report","toplantı daveti","meeting invite",
    "calendar:","accepted:","declined:","tentative:"
  ],
  EXCLUDE_KEYWORDS: [
    "newsletter","unsubscribe","marketing","promotion",
    "webinar","event invitation","happy hour","out of office"
  ],

  // VIP — bu kişiler HİÇBİR filtreye takılmaz
  VIP_EMAILS: ["numan@212.vc","onur@212.vc","ali@212.vc"],

  // Bilinen vendor email parçaları
  VENDOR_DOMAINS: [
    "hawksford","amstone","tamimi","vistra","mazars",
    "eurotraduc","ustaxfs","investeurope","informa",
    "dechert","fladgate","orrick","vestbee","vauban",
    "deloitte","ey.com","kpmg","audit","schoenherr",
    "loyens","moroglu","dryocopus","mevca","ebury"
  ]
};

// ── AYLIK VENDOR PATTERNLERİ ────────────────────────────────
var MONTHLY_VENDOR_PATTERN = {
  1: ["ACSE","Actifit","Affinity","Amstone","Bunch Capital","CSSF","Erta SMMM A.Ş.","Hawksford","Igniters Tech Consulting LLC.","Intabulis SCSp","Kolay İK","Mazars","Schönherr","Tamimi","The Audit Company"],
  2: ["Amstone","Dryocopus LLC","Erta SMMM A.Ş.","Girişim Akademi","Hawksford","Kyocera","Moroğlu Arseven","Natus İletişim","Talent Melon","Tamimi","The Audit Company","Vistra"],
  3: ["Amstone","Erta SMMM A.Ş.","Hawksford","INAM","Invest Europe","Kyocera","Natus İletişim","Talent Melon","Tamimi","USTAXFS","Vistra"],
  4: ["Doğan Burda","Dryocopus LLC","Elmira Bayraslı","Erta SMMM A.Ş.","Hawksford","Kyocera","Natus İletişim","Tamimi"],
  5: ["Amstone","Endeavor Derneği","Erta SMMM A.Ş.","Hawksford","Kyocera","Natus İletişim","Tamimi","Vistra"],
  6: ["Erta SMMM A.Ş.","Hawksford","Juniper","Kyocera","Moroğlu Arseven","Talent Melon","Tamimi"],
  7: ["Chamber of Commerce Luxembourg","Erta SMMM A.Ş.","Hawksford","Intabulis SCSp","Kyocera","Moroğlu Arseven","Tamimi","Van Campen Liem"],
  8: ["Erta SMMM A.Ş.","Hawksford","Kyocera","Talent Melon","Tamimi","Vira Yatçılık"],
  9: ["DHL","Erta SMMM A.Ş.","Hawksford","Kyocera","Specter","Tamimi"],
  10: ["Administration des Contributions Directes","Bunch Capital","Erta SMMM A.Ş.","Hawksford","Kyocera","Mazars","Tamimi"],
  11: ["Erta SMMM A.Ş.","Eurtraduc","Hawksford","Moroğlu Arseven","Talent Melon","Tamimi","Venture360"],
  12: ["Erta SMMM A.Ş.","Hawksford","Tamimi"]
};

// Tüm bilinen vendor'lar (deduplicated)
var ALL_KNOWN_VENDORS = (function() {
  var all = [], seen = {};
  var months = Object.keys(MONTHLY_VENDOR_PATTERN);
  for (var m = 0; m < months.length; m++) {
    var vendors = MONTHLY_VENDOR_PATTERN[months[m]];
    for (var v = 0; v < vendors.length; v++) {
      if (!seen[vendors[v]]) { seen[vendors[v]] = true; all.push(vendors[v]); }
    }
  }
  return all;
})();

// ══════════════════════════════════════════════════════════════
// ANA FONKSİYON
// ══════════════════════════════════════════════════════════════
function scanInvoices() {
  var ss = getOrCreateSpreadsheet();
  var startTime = Date.now();
  log(ss, "🚀 Tarama başladı");

  try {
    // 1. Drive Excel sync (değişmişse)
    try { syncFromDriveExcel(ss); } catch(e) {
      log(ss, "⚠️ Excel sync atlandı: " + e.message);
    }

    // 2. Gmail tara
    var scanResult = scanGmail(ss, startTime);
    log(ss, "📧 " + scanResult.invoices.length + " fatura, " +
        scanResult.suspicious.length + " şüpheli, " +
        scanResult.unreadablePdfs.length + " okunamayan PDF");

    // 3. Sheet'lere yaz
    writeFoundInvoices(ss, scanResult.invoices);
    writeSuspicious(ss, scanResult.suspicious);

    // 4. Cross-check
    try { crossCheck(ss, scanResult.invoices); } catch(e) {
      log(ss, "⚠️ Cross-check hatası: " + e.message);
    }

    // 5. Eksik vendor kontrolü
    var missingVendors = [];
    try { missingVendors = checkMissingVendors(ss, scanResult.invoices); } catch(e) {
      log(ss, "⚠️ Vendor kontrol hatası: " + e.message);
    }

    // 6. Dashboard güncelle
    try { updateDashboard(ss); } catch(e) {
      log(ss, "⚠️ Dashboard hatası: " + e.message);
    }

    // 7. CFO Raporu gönder
    sendCfoReport(ss, {
      invoices: scanResult.invoices,
      suspicious: scanResult.suspicious,
      unreadablePdfs: scanResult.unreadablePdfs,
      missingVendors: missingVendors
    });

    var duration = ((Date.now() - startTime) / 1000).toFixed(1);
    log(ss, "✅ Tarama tamamlandı (" + duration + "s)");

  } catch(e) {
    log(ss, "❌ KRİTİK HATA: " + e.message);
    try {
      MailApp.sendEmail(CONFIG.NOTIFICATION_EMAIL,
        "🔴 Invoice Tracker HATA", e.message + "\n\n" + e.stack);
    } catch(me) {}
  }
}

// ══════════════════════════════════════════════════════════════
// GMAIL TARAMA
// ══════════════════════════════════════════════════════════════
function scanGmail(ss, startTime) {
  var invoices = [];
  var suspicious = [];
  var unreadablePdfs = [];
  var geminiCalls = 0;
  var pdfCount = 0;

  // İşlenmiş ID'leri yükle
  var processedIds = loadProcessedIds();
  var newProcessedIds = [];

  // Bilinen faturalar (Gemini prompt için, sadece Overdue)
  var knownForGemini = [];
  try { knownForGemini = getOverdueInvoices(ss); } catch(e) {}

  // Tarih limiti
  var afterDate = new Date();
  afterDate.setDate(afterDate.getDate() - CONFIG.SCAN_DAYS);
  var dateStr = Utilities.formatDate(afterDate, Session.getScriptTimeZone(), "yyyy/MM/dd");

  // Gmail sorguları
  var kwQuery = CONFIG.INVOICE_KEYWORDS.map(function(k) { return '"' + k + '"'; }).join(" OR ");
  var queries = [
    "after:" + dateStr + " (" + kwQuery + ")",
    "after:" + dateStr + " has:attachment filename:pdf"
  ];

  var seenMsgIds = {};

  for (var q = 0; q < queries.length; q++) {
    if ((Date.now() - startTime) > CONFIG.EXECUTION_LIMIT) break;

    var threads;
    try {
      threads = GmailApp.search(queries[q], 0, CONFIG.THREAD_LIMIT);
    } catch(e) {
      log(ss, "⚠️ Arama hatası: " + e.message);
      continue;
    }

    for (var ti = 0; ti < threads.length; ti++) {
      if ((Date.now() - startTime) > CONFIG.EXECUTION_LIMIT) {
        log(ss, "⏱️ Süre limiti — kalan mailler sonraki taramada");
        break;
      }

      var messages = threads[ti].getMessages();
      for (var mi = 0; mi < messages.length; mi++) {
        var msg = messages[mi];
        var msgId = msg.getId();
        if (seenMsgIds[msgId] || processedIds[msgId]) continue;
        if (msg.getDate() < afterDate) continue;
        seenMsgIds[msgId] = true;

        var from = msg.getFrom() || "";
        var subject = msg.getSubject() || "";
        var body = msg.getPlainBody() || "";
        var lowerFrom = from.toLowerCase();
        var lowerSubject = subject.toLowerCase();
        var lowerAll = (subject + " " + body).toLowerCase();

        // ── VIP KONTROLÜ ──
        var isVip = false;
        for (var vi = 0; vi < CONFIG.VIP_EMAILS.length; vi++) {
          if (lowerFrom.indexOf(CONFIG.VIP_EMAILS[vi]) >= 0) { isVip = true; break; }
        }

        // ── EXCLUDE KURALLARI (VIP bypass) ──
        if (!isVip) {
          var skipSender = false;
          for (var si = 0; si < CONFIG.DOMAIN_EXCLUDE.length; si++) {
            if (lowerFrom.indexOf(CONFIG.DOMAIN_EXCLUDE[si]) >= 0) { skipSender = true; break; }
          }
          if (!skipSender) {
            for (var si2 = 0; si2 < CONFIG.EXCLUDE_SENDERS.length; si2++) {
              if (lowerFrom.indexOf(CONFIG.EXCLUDE_SENDERS[si2]) >= 0) { skipSender = true; break; }
            }
          }
          if (skipSender) { newProcessedIds.push(msgId); continue; }

          var skipSubject = false;
          for (var sj = 0; sj < CONFIG.EXCLUDE_SUBJECTS.length; sj++) {
            if (lowerSubject.indexOf(CONFIG.EXCLUDE_SUBJECTS[sj]) >= 0) { skipSubject = true; break; }
          }
          if (skipSubject) { newProcessedIds.push(msgId); continue; }

          var skipKw = false;
          for (var sk = 0; sk < CONFIG.EXCLUDE_KEYWORDS.length; sk++) {
            if (lowerAll.indexOf(CONFIG.EXCLUDE_KEYWORDS[sk]) >= 0) { skipKw = true; break; }
          }
          if (skipKw) { newProcessedIds.push(msgId); continue; }
        }

        // ── ÖDENMİŞ FATURA KONTROLÜ (Gemini'den ÖNCE) ──
        var isAlreadyPaid = false;
        for (var ap = 0; ap < CONFIG.ALREADY_PAID_INDICATORS.length; ap++) {
          if (lowerAll.indexOf(CONFIG.ALREADY_PAID_INDICATORS[ap]) >= 0) { isAlreadyPaid = true; break; }
        }
        if (isAlreadyPaid) { newProcessedIds.push(msgId); continue; }

        // ── PDF ANALİZİ ──
        var pdfResults = [];
        var attachments = msg.getAttachments();
        for (var ai = 0; ai < attachments.length; ai++) {
          var att = attachments[ai];
          if (att.getContentType() === "application/pdf" || att.getName().toLowerCase().endsWith(".pdf")) {
            if (pdfCount < CONFIG.PDF_LIMIT) {
              var pdfAnalysis = analyzePdf(att);
              pdfResults.push(pdfAnalysis);
              pdfCount++;
              if (pdfAnalysis.failed) {
                unreadablePdfs.push({
                  date: msg.getDate(), from: from, subject: subject,
                  fileName: att.getName(), error: pdfAnalysis.error,
                  link: "https://mail.google.com/mail/u/0/#inbox/" + threads[ti].getId()
                });
              }
            }
          }
        }

        // ── SKOR HESAPLAMA ──
        var score = calculateScore(subject, body, from, pdfResults);
        if (isVip) score += 3;

        if (score < 2 && !isVip) { newProcessedIds.push(msgId); continue; }

        // ── REGEX ÇIKARIM ──
        var vendorGuess = guessVendor(from);
        var amountGuess = extractAmount(body, pdfResults);
        var currencyGuess = extractCurrency(body, pdfResults);
        var invoiceNoGuess = extractInvoiceNumber(subject, body, pdfResults);
        var geminiNotes = "";

        // ── GEMINI AI ANALİZ ──
        if (geminiCalls < CONFIG.MAX_GEMINI_PER_SCAN && (pdfResults.length > 0 || score >= 5 || isVip)) {
          if (geminiCalls > 0) Utilities.sleep(CONFIG.GEMINI_DELAY_MS);
          try {
            var pdfTexts = [];
            for (var pt = 0; pt < pdfResults.length; pt++) {
              if (pdfResults[pt].text) pdfTexts.push(pdfResults[pt].text);
            }
            var gemResult = askGemini(buildGeminiPrompt(subject, from, body, pdfTexts, knownForGemini));
            geminiCalls++;

            if (gemResult) {
              var parsed = parseGeminiJson(gemResult);
              if (parsed) {
                // Gemini "fatura değil" ama VIP veya audit keyword varsa → şüpheli
                if (parsed.isInvoice === false) {
                  var hasAuditWord = false;
                  var allPdfText = pdfTexts.join(" ").toLowerCase();
                  for (var aw = 0; aw < CONFIG.AUDIT_KEYWORDS.length; aw++) {
                    if (allPdfText.indexOf(CONFIG.AUDIT_KEYWORDS[aw]) >= 0 || lowerAll.indexOf(CONFIG.AUDIT_KEYWORDS[aw]) >= 0) {
                      hasAuditWord = true; break;
                    }
                  }
                  if (hasAuditWord || isVip) {
                    suspicious.push({
                      date: msg.getDate(), from: from, subject: subject,
                      reason: isVip ? "VIP mail — fatura olmayabilir" : "Finansal terim içeriyor",
                      link: "https://mail.google.com/mail/u/0/#inbox/" + threads[ti].getId()
                    });
                  }
                  newProcessedIds.push(msgId);
                  continue;
                }
                // Gemini sonuçlarını override et
                if (parsed.vendor) vendorGuess = parsed.vendor;
                if (parsed.invoiceNo) invoiceNoGuess = parsed.invoiceNo;
                if (parsed.amount) amountGuess = String(parsed.amount);
                if (parsed.currency) currencyGuess = parsed.currency;
                if (parsed.notes) geminiNotes = parsed.notes;
                if (parsed.confidence === "high") score += 3;
                else if (parsed.confidence === "medium") score += 1;
              }
            }
          } catch(ge) {
            // Gemini hata — regex sonuçlarıyla devam
          }
        }

        // ── KAYDET ──
        invoices.push({
          date: msg.getDate(),
          from: from,
          subject: subject,
          vendor: vendorGuess,
          amount: amountGuess,
          currency: currencyGuess,
          invoiceNo: normalizeInvoiceNo(invoiceNoGuess),
          score: score,
          isVip: isVip,
          hasPdf: pdfResults.length > 0,
          pdfIsInvoice: pdfResults.some(function(p) { return p.isInvoice; }),
          geminiNotes: geminiNotes,
          link: "https://mail.google.com/mail/u/0/#inbox/" + threads[ti].getId()
        });

        newProcessedIds.push(msgId);
      }
    }
  }

  // İşlenmiş ID'leri kaydet
  saveProcessedIds(newProcessedIds);

  invoices.sort(function(a, b) { return b.score - a.score; });
  return { invoices: invoices, suspicious: suspicious, unreadablePdfs: unreadablePdfs };
}

// ══════════════════════════════════════════════════════════════
// İŞLENMİŞ ID YÖNETİMİ (JSON array, 5000 limit)
// ══════════════════════════════════════════════════════════════
function loadProcessedIds() {
  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty("PROCESSED_IDS");
  if (!stored) return {};
  try {
    var arr = JSON.parse(stored);
    var map = {};
    for (var i = 0; i < arr.length; i++) map[arr[i]] = true;
    return map;
  } catch(e) { return {}; }
}

function saveProcessedIds(newIds) {
  if (newIds.length === 0) return;
  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty("PROCESSED_IDS");
  var existing = [];
  try { existing = stored ? JSON.parse(stored) : []; } catch(e) { existing = []; }
  for (var i = 0; i < newIds.length; i++) existing.push(newIds[i]);
  if (existing.length > 5000) existing = existing.slice(existing.length - 5000);
  props.setProperty("PROCESSED_IDS", JSON.stringify(existing));
}

// ══════════════════════════════════════════════════════════════
// GEMINI AI
// ══════════════════════════════════════════════════════════════
function askGemini(prompt) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    Logger.log("⚠️ GEMINI_API_KEY tanımlı değil");
    return null;
  }
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + apiKey;
  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.1, maxOutputTokens: 400 }
  };
  var options = {
    method: "post", contentType: "application/json",
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  // Retry: 0s, 10s, 20s
  var delays = [0, 10000, 20000];
  for (var attempt = 0; attempt < delays.length; attempt++) {
    if (attempt > 0) Utilities.sleep(delays[attempt]);
    try {
      var resp = UrlFetchApp.fetch(url, options);
      var code = resp.getResponseCode();
      if (code === 200) {
        var json = JSON.parse(resp.getContentText());
        return json.candidates[0].content.parts[0].text;
      }
      if (code === 429) continue; // Rate limit — retry
      return null;
    } catch(e) {
      if (attempt === delays.length - 1) return null;
    }
  }
  return null;
}

function buildGeminiPrompt(subject, from, body, pdfTexts, knownInvoices) {
  var prompt = "Sen bir fatura analiz asistanısın. Bu maili analiz et.\n\n";
  prompt += "Gönderen: " + from + "\nKonu: " + subject + "\n";
  prompt += "Body (ilk 1000 kar): " + (body || "(boş)").substring(0, 1000) + "\n\n";

  if (pdfTexts.length > 0) {
    for (var i = 0; i < pdfTexts.length; i++) {
      prompt += "PDF " + (i+1) + " (ilk 2000 kar):\n" + pdfTexts[i].substring(0, 2000) + "\n\n";
    }
  }

  if (knownInvoices.length > 0) {
    prompt += "ÖDEME BEKLEYEN FATURALAR (Excel):\n";
    for (var k = 0; k < knownInvoices.length; k++) {
      prompt += knownInvoices[k].invoiceNo + " | " + knownInvoices[k].vendor + " | " + knownInvoices[k].status + "\n";
    }
    prompt += "\n";
  }

  prompt += 'SADECE JSON döndür, başka bir şey yazma. Backtick kullanma.\n';
  prompt += '{"isInvoice":true,"vendor":"...","invoiceNo":"...","amount":"...","currency":"EUR","confidence":"high","notes":"..."}\n';
  prompt += "isInvoice: Ödeme gerektiren fatura mı? Expense report, davet, newsletter ise false.\n";
  prompt += "vendor: Faturayı kesen firma. amount: KDV dahil toplam. confidence: high/medium/low.";
  return prompt;
}

function parseGeminiJson(text) {
  if (!text) return null;
  try {
    var cleaned = text.replace(/```json\s*/g, "").replace(/```\s*/g, "").trim();
    return JSON.parse(cleaned);
  } catch(e) { return null; }
}

function getOverdueInvoices(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_KNOWN);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][3] || "").toLowerCase();
    if (status === "paid" || status === "revoked") continue;
    result.push({
      vendor: String(data[i][0] || ""),
      invoiceNo: String(data[i][1] || ""),
      status: String(data[i][3] || "")
    });
  }
  return result;
}

// ══════════════════════════════════════════════════════════════
// PDF ANALİZİ (3 yöntem + finally cleanup)
// ══════════════════════════════════════════════════════════════
function analyzePdf(attachment) {
  var fileName = attachment.getName();
  var tempFileIds = [];
  var result = { isInvoice: false, failed: true, text: "", error: "Bilinmeyen", summary: fileName + ": 🔴 HATA" };

  function cleanup() {
    for (var i = 0; i < tempFileIds.length; i++) {
      try { DriveApp.getFileById(tempFileIds[i]).setTrashed(true); } catch(e) {}
    }
  }

  try {
    var blob = attachment.copyBlob();
    if (blob.getBytes().length > 10 * 1024 * 1024) {
      result = { isInvoice: false, failed: true, text: "", error: "Dosya >10MB", summary: fileName + ": ⚠️ Büyük" };
      return result;
    }

    var text = "";

    // Yöntem 1: Drive.Files.insert + OCR
    try {
      var file = Drive.Files.insert(
        { title: "TEMP_OCR_" + Date.now(), mimeType: "application/pdf" },
        blob, { ocr: true, convert: true }
      );
      tempFileIds.push(file.id);
      text = DocumentApp.openById(file.id).getBody().getText();
    } catch(e1) {
      // Yöntem 2: DriveApp + insert
      try {
        var tempPdf = DriveApp.createFile(blob.setName("TEMP_OCR_" + Date.now() + ".pdf"));
        tempFileIds.push(tempPdf.getId());
        var docFile = Drive.Files.insert(
          { title: "TEMP_OCR_DOC_" + Date.now(), mimeType: "application/pdf" },
          tempPdf.getBlob(), { ocr: true, convert: true }
        );
        tempFileIds.push(docFile.id);
        text = DocumentApp.openById(docFile.id).getBody().getText();
      } catch(e2) {
        // Yöntem 3: Binary extraction
        try {
          var bytes = blob.getBytes();
          var raw = "";
          var limit = Math.min(bytes.length, 50000);
          for (var i = 0; i < limit; i++) {
            var c = bytes[i];
            if (c >= 32 && c <= 126) raw += String.fromCharCode(c);
          }
          var matches = raw.match(/\(([^)]{3,})\)/g);
          if (matches) text = matches.map(function(m) { return m.slice(1,-1); }).join(" ");
        } catch(e3) {}
      }
    }

    if (!text || text.trim().length < 20) {
      result = { isInvoice: false, failed: true, text: "", error: "Okunamadı", summary: fileName + ": 🔴 OKUNAMADI" };
      return result;
    }

    // Fatura belirtileri
    var indicators = ["invoice","fatura","total","subtotal","vat","kdv","amount due","iban","due date","toplam","bank account"];
    var matchCount = 0;
    var lower = text.toLowerCase();
    for (var i = 0; i < indicators.length; i++) {
      if (lower.indexOf(indicators[i]) >= 0) matchCount++;
    }

    result = {
      isInvoice: matchCount >= 3,
      failed: false,
      text: text.substring(0, 3000),
      summary: fileName + ": " + (matchCount >= 3 ? "✅" : "❓") + " (" + matchCount + ")",
      matchCount: matchCount
    };
    return result;

  } catch(e) {
    result = { isInvoice: false, failed: true, text: "", error: e.message, summary: fileName + ": 🔴 HATA" };
    return result;
  } finally {
    cleanup();
  }
}

// ══════════════════════════════════════════════════════════════
// REGEX ÇIKARIM FONKSİYONLARI
// ══════════════════════════════════════════════════════════════
function calculateScore(subject, body, from, pdfResults) {
  var score = 0;
  var ls = subject.toLowerCase();
  var lb = body.toLowerCase();
  var lf = from.toLowerCase();

  for (var i = 0; i < CONFIG.INVOICE_KEYWORDS.length; i++) {
    if (ls.indexOf(CONFIG.INVOICE_KEYWORDS[i]) >= 0) { score += 3; break; }
  }
  for (var i = 0; i < CONFIG.INVOICE_KEYWORDS.length; i++) {
    if (lb.indexOf(CONFIG.INVOICE_KEYWORDS[i]) >= 0) { score += 2; break; }
  }
  for (var i = 0; i < CONFIG.VENDOR_DOMAINS.length; i++) {
    if (lf.indexOf(CONFIG.VENDOR_DOMAINS[i]) >= 0) { score += 2; break; }
  }
  if (pdfResults.length > 0) score += 1;
  for (var i = 0; i < pdfResults.length; i++) {
    if (pdfResults[i].isInvoice) { score += 3; break; }
  }
  return score;
}

function guessVendor(from) {
  var match = from.match(/"?([^"<]+)"?\s*</);
  if (match) {
    var name = match[1].trim();
    for (var i = 0; i < ALL_KNOWN_VENDORS.length; i++) {
      if (vendorMatch(name, ALL_KNOWN_VENDORS[i])) return ALL_KNOWN_VENDORS[i];
    }
    return name;
  }
  var emailMatch = from.match(/@([^>]+)/);
  if (emailMatch) return emailMatch[1].split(".")[0];
  return "Bilinmeyen";
}

function extractAmount(body, pdfResults) {
  var allText = body + " " + pdfResults.map(function(p) { return p.text || ""; }).join(" ");
  var lower = allText.toLowerCase();
  var patterns = [
    /amount\s*due[:\s]*[\€\$\£]?\s*([\d.,]+)/i,
    /genel\s*toplam[:\s]*[\€\$\£]?\s*([\d.,]+)/i,
    /balance\s*due[:\s]*[\€\$\£]?\s*([\d.,]+)/i,
    /(?:^|\s)total[:\s]*[\€\$\£]?\s*([\d.,]+)/im,
    /([\d.,]+)\s*(EUR|USD|TRY|AED|GBP|CHF)/i
  ];

  for (var p = 0; p < patterns.length; p++) {
    var m = patterns[p].exec(allText);
    if (m) {
      // Subtotal/discount kontrolü
      if (m.index > 0) {
        var pre = lower.substring(Math.max(0, m.index - 5), m.index);
        if (pre.indexOf("sub") >= 0 || pre.indexOf("dis") >= 0) continue;
      }
      var amt = m[1];
      var lastComma = amt.lastIndexOf(",");
      var lastDot = amt.lastIndexOf(".");
      if (lastComma > lastDot) amt = amt.replace(/\./g, "").replace(",", ".");
      else if (lastDot > lastComma) amt = amt.replace(/,/g, "");
      else if (lastComma >= 0 && lastDot < 0) amt = amt.replace(",", ".");
      return amt;
    }
  }
  return "";
}

function extractCurrency(body, pdfResults) {
  var allText = body + " " + pdfResults.map(function(p) { return p.text || ""; }).join(" ");
  var m = allText.match(/(EUR|USD|TRY|AED|GBP|CHF|€|\$|£)/i);
  if (m) {
    if (m[1] === "€") return "EUR";
    if (m[1] === "$") return "USD";
    if (m[1] === "£") return "GBP";
    return m[1].toUpperCase();
  }
  return "";
}

function normalizeInvoiceNo(raw) {
  if (!raw) return "";
  return String(raw).trim().toUpperCase().replace(/[\s\-\/\.]+/g, "");
}

function extractInvoiceNumber(subject, body, pdfResults) {
  // Subject öncelikli
  var subPatterns = [
    /\b(TC\d{3,})\b/i, /\b(\d{7,})\b/, /\b([A-Z]{2}\d{4,})\b/,
    /invoice[:\s\-]*([A-Z0-9\-\/]{4,})/i, /fatura[:\s\-]*([A-Z0-9\-\/]{4,})/i
  ];
  var badWords = ["description","date","total","amount","invoice","fatura","service","payment","from","subject"];

  for (var i = 0; i < subPatterns.length; i++) {
    var m = subject.match(subPatterns[i]);
    if (m) {
      var candidate = m[1];
      var isBad = false;
      for (var b = 0; b < badWords.length; b++) {
        if (candidate.toLowerCase() === badWords[b]) { isBad = true; break; }
      }
      if (!isBad && candidate.length >= 3) return candidate;
    }
  }

  // Body
  var bodyPatterns = [
    /invoice\s*(?:#|no\.?|number)[:\s]*([A-Z0-9\-\/]{3,})/i,
    /fatura\s*(?:#|no\.?|numar)[:\s]*([A-Z0-9\-\/]{3,})/i,
    /\b(TC\d{3,})\b/i,
    /reference[:\s]*([A-Z0-9\-\/]{4,})/i
  ];
  var allText = body + " " + pdfResults.map(function(p) { return p.text || ""; }).join(" ");
  for (var i = 0; i < bodyPatterns.length; i++) {
    var m = allText.match(bodyPatterns[i]);
    if (m && m[1].length >= 3 && m[1].length <= 30) return m[1];
  }
  return "";
}

// ══════════════════════════════════════════════════════════════
// VENDOR EŞLEŞTİRME
// ══════════════════════════════════════════════════════════════
function vendorMatch(a, b) {
  if (!a || !b) return false;
  a = String(a).toLowerCase().trim();
  b = String(b).toLowerCase().trim();
  if (a === b) return true;
  if (a.length >= 7 && b.indexOf(a) >= 0) return true;
  if (b.length >= 7 && a.indexOf(b) >= 0) return true;

  var skip = ["the","a","an","de","van","von","al","llc","ltd","inc","gmbh","bv","nv"];
  function getKeys(s) {
    var words = s.split(/[\s.,\-&]+/);
    var keys = [];
    for (var i = 0; i < words.length; i++) {
      if (skip.indexOf(words[i]) < 0 && words[i].length >= 4) keys.push(words[i]);
    }
    return keys;
  }

  var ka = getKeys(a), kb = getKeys(b);
  if (ka.length === 0 || kb.length === 0) return false;
  if (ka[0].length >= 5 && ka[0] === kb[0]) return true;

  var mc = 0;
  for (var i = 0; i < ka.length; i++) {
    for (var j = 0; j < kb.length; j++) {
      if (ka[i] === kb[j] && ka[i].length >= 4) mc++;
    }
  }
  return mc >= 2 || (mc === 1 && ka[0].length >= 6 && kb[0].length >= 6);
}

// ══════════════════════════════════════════════════════════════
// SHEET İŞLEMLERİ (Batch setValues)
// ══════════════════════════════════════════════════════════════
function getOrCreateSpreadsheet() {
  var ss;
  try { ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); } catch(e) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var needed = [CONFIG.SHEET_FOUND, CONFIG.SHEET_CROSSCHECK, CONFIG.SHEET_SUSPICIOUS,
                CONFIG.SHEET_KNOWN, CONFIG.SHEET_VENDORS, CONFIG.SHEET_DASHBOARD, CONFIG.SHEET_LOG];
  for (var i = 0; i < needed.length; i++) {
    if (!ss.getSheetByName(needed[i])) {
      var sh = ss.insertSheet(needed[i]);
      // Header'lar
      var headers = {
        "📧 Faturalar": ["Tarih","Gönderen","Konu","Vendor","Fatura No","Tutar","Para Birimi","Skor","VIP?","PDF?","AI Notlar","Link","İşlendi?"],
        "🔍 Cross-Check": ["Fatura No (Gmail)","Vendor","Tarih","Excel'de Var?","Durum"],
        "⚠️ Şüpheli": ["Tarih","Gönderen","Konu","Sebep","Link","İncelendi?"],
        "📋 Bilinen Faturalar": ["Vendor","Fatura No","Tutar","Status"],
        "🏷️ Eksik Vendorlar": ["Ay","Beklenen Vendor","Durum"],
        "📊 Dashboard": ["Metrik","Değer"],
        "📝 Log": ["Zaman","Mesaj"]
      };
      if (headers[needed[i]]) {
        sh.appendRow(headers[needed[i]]);
        sh.getRange(1, 1, 1, headers[needed[i]].length).setFontWeight("bold").setBackground("#1a237e").setFontColor("white");
        sh.setFrozenRows(1);
      }
    }
  }
  return ss;
}

function writeFoundInvoices(ss, invoices) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_FOUND);
  if (!sheet) return;

  // Duplikasyon kontrolü: permalink + fatura no
  var existing = sheet.getDataRange().getValues();
  var existingLinks = {};
  var existingNos = {};
  for (var i = 1; i < existing.length; i++) {
    if (existing[i][11]) existingLinks[existing[i][11]] = true;
    if (existing[i][4]) existingNos[normalizeInvoiceNo(existing[i][4])] = true;
  }

  var rows = [];
  for (var i = 0; i < invoices.length; i++) {
    var inv = invoices[i];
    if (existingLinks[inv.link]) continue;
    if (inv.invoiceNo && existingNos[normalizeInvoiceNo(inv.invoiceNo)]) continue;

    rows.push([
      inv.date, inv.from, inv.subject, inv.vendor,
      inv.invoiceNo, inv.amount, inv.currency, inv.score,
      inv.isVip ? "👑 VIP" : "", inv.hasPdf ? "✅" : "❌",
      inv.geminiNotes, inv.link, ""
    ]);
  }

  if (rows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, 13).setValues(rows);
    // VIP satırları vurgula
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][8] === "👑 VIP") {
        sheet.getRange(startRow + i, 1, 1, 13).setBackground("#FFF3E0");
      }
    }
  }
}

function writeSuspicious(ss, suspicious) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_SUSPICIOUS);
  if (!sheet || suspicious.length === 0) return;

  var existing = sheet.getDataRange().getValues();
  var existingLinks = {};
  for (var i = 1; i < existing.length; i++) {
    if (existing[i][4]) existingLinks[existing[i][4]] = true;
  }

  var rows = [];
  for (var i = 0; i < suspicious.length; i++) {
    var s = suspicious[i];
    if (existingLinks[s.link]) continue;
    rows.push([s.date, s.from, s.subject, s.reason, s.link, ""]);
  }

  if (rows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, 6).setValues(rows);
    sheet.getRange(startRow, 1, rows.length, 6).setBackground("#FFF8E1");
  }
}

function crossCheck(ss, invoices) {
  var checkSheet = ss.getSheetByName(CONFIG.SHEET_CROSSCHECK);
  var knownSheet = ss.getSheetByName(CONFIG.SHEET_KNOWN);
  if (!checkSheet || !knownSheet) return;

  var knownData = knownSheet.getDataRange().getValues();
  var knownNos = {};
  for (var i = 1; i < knownData.length; i++) {
    if (knownData[i][1]) knownNos[normalizeInvoiceNo(knownData[i][1])] = true;
  }

  // Mevcut cross-check kayıtlarını koru
  var existingCC = checkSheet.getDataRange().getValues();
  var existingCCNos = {};
  for (var i = 1; i < existingCC.length; i++) {
    if (existingCC[i][0]) existingCCNos[normalizeInvoiceNo(existingCC[i][0])] = true;
  }

  var rows = [];
  for (var i = 0; i < invoices.length; i++) {
    var inv = invoices[i];
    if (!inv.invoiceNo) continue;
    var norm = normalizeInvoiceNo(inv.invoiceNo);
    if (existingCCNos[norm]) continue;

    var inExcel = knownNos[norm] ? true : false;
    rows.push([
      inv.invoiceNo, inv.vendor, inv.date,
      inExcel ? "✅ Evet" : "❌ HAYIR",
      inExcel ? "Tamam" : "⚠️ Excel'de yok"
    ]);
  }

  if (rows.length > 0) {
    var startRow = checkSheet.getLastRow() + 1;
    checkSheet.getRange(startRow, 1, rows.length, 5).setValues(rows);
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][3] === "❌ HAYIR") {
        checkSheet.getRange(startRow + i, 1, 1, 5).setBackground("#FFC7CE");
      }
    }
  }
}

function checkMissingVendors(ss, invoices) {
  var vendorSheet = ss.getSheetByName(CONFIG.SHEET_VENDORS);
  if (!vendorSheet) return [];

  var currentMonth = new Date().getMonth() + 1;
  var expected = MONTHLY_VENDOR_PATTERN[currentMonth] || [];
  var foundNames = invoices.map(function(inv) { return inv.vendor; });

  // Bilinen faturalardan da kontrol
  var knownSheet = ss.getSheetByName(CONFIG.SHEET_KNOWN);
  if (knownSheet && knownSheet.getLastRow() > 1) {
    var knownData = knownSheet.getDataRange().getValues();
    for (var i = 1; i < knownData.length; i++) {
      if (knownData[i][0]) foundNames.push(String(knownData[i][0]));
    }
  }

  var lastRow = vendorSheet.getLastRow();
  if (lastRow > 1) vendorSheet.getRange(2, 1, lastRow - 1, 3).clearContent();

  var missing = [];
  var rows = [];
  var monthNames = ["","Ocak","Şubat","Mart","Nisan","Mayıs","Haziran",
                     "Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"];

  for (var v = 0; v < expected.length; v++) {
    var found = false;
    for (var f = 0; f < foundNames.length; f++) {
      if (vendorMatch(expected[v], foundNames[f])) { found = true; break; }
    }
    rows.push([monthNames[currentMonth], expected[v], found ? "✅ Geldi" : "❌ Eksik"]);
    if (!found) missing.push(expected[v]);
  }

  if (rows.length > 0) {
    vendorSheet.getRange(2, 1, rows.length, 3).setValues(rows);
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][2] === "❌ Eksik") {
        vendorSheet.getRange(2 + i, 1, 1, 3).setBackground("#FFC7CE");
      }
    }
  }
  return missing;
}

// ══════════════════════════════════════════════════════════════
// DASHBOARD
// ══════════════════════════════════════════════════════════════
function updateDashboard(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_FOUND);
  var dash = ss.getSheetByName(CONFIG.SHEET_DASHBOARD);
  if (!sheet || !dash || sheet.getLastRow() <= 1) return;

  dash.clear();
  dash.appendRow(["Metrik", "Değer"]);
  dash.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#1a237e").setFontColor("white");

  var data = sheet.getDataRange().getValues();
  var count = 0, byCurrency = {}, byMonth = {};

  for (var i = 1; i < data.length; i++) {
    var amt = parseFloat(data[i][5]) || 0;
    var cur = data[i][6] || "TRY";
    var dateVal = data[i][0];
    if (dateVal instanceof Date) {
      var mk = dateVal.getFullYear() + "-" + String(dateVal.getMonth() + 1).padStart(2, "0");
      byMonth[mk] = (byMonth[mk] || 0) + amt;
    }
    byCurrency[cur] = (byCurrency[cur] || 0) + amt;
    count++;
  }

  var dashRows = [["Toplam Fatura Sayısı", count]];
  var currencies = Object.keys(byCurrency);
  for (var i = 0; i < currencies.length; i++) {
    dashRows.push(["Toplam " + currencies[i], byCurrency[currencies[i]].toFixed(2)]);
  }
  dashRows.push(["", ""]);
  dashRows.push(["Ay", "Toplam"]);
  var monthKeys = Object.keys(byMonth).sort();
  for (var i = 0; i < monthKeys.length; i++) {
    dashRows.push([monthKeys[i], byMonth[monthKeys[i]].toFixed(2)]);
  }

  if (dashRows.length > 0) {
    dash.getRange(2, 1, dashRows.length, 2).setValues(dashRows);
  }
}

// ══════════════════════════════════════════════════════════════
// DRIVE EXCEL SYNC (değişmişse)
// ══════════════════════════════════════════════════════════════
function syncFromDriveExcel(ss) {
  var files = DriveApp.searchFiles('title contains "Pending" and title contains "Authorization" and trashed = false');
  if (!files.hasNext()) {
    files = DriveApp.searchFiles('title contains "Invoice_Pending" and trashed = false');
  }
  if (!files.hasNext()) return;

  var latestFile = null, latestDate = new Date(0);
  while (files.hasNext()) {
    var file = files.next();
    if (file.getLastUpdated() > latestDate) { latestDate = file.getLastUpdated(); latestFile = file; }
  }

  // Değişmemişse atla
  var props = PropertiesService.getScriptProperties();
  var lastSync = props.getProperty("LAST_SYNC_DATE");
  if (lastSync === latestDate.toISOString()) return;

  log(ss, "📥 Excel sync: " + latestFile.getName());

  var tempFile = Drive.Files.insert(
    { title: "TEMP_SYNC_" + Date.now(), mimeType: "application/vnd.google-apps.spreadsheet" },
    latestFile.getBlob(), { convert: true }
  );

  try {
    var tempSs = SpreadsheetApp.openById(tempFile.id);
    var sourceSheet = null;
    var sheets = tempSs.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (name.indexOf("FUND") >= 0 && name.indexOf("Invoice") >= 0) { sourceSheet = sheets[i]; break; }
    }
    if (!sourceSheet) sourceSheet = sheets[0];

    var data = sourceSheet.getDataRange().getValues();

    // Header bul
    var headerRow = -1;
    for (var r = 0; r < Math.min(data.length, 10); r++) {
      for (var c = 0; c < data[r].length; c++) {
        if (String(data[r][c]).indexOf("Service Provider") >= 0) { headerRow = r; break; }
      }
      if (headerRow >= 0) break;
    }
    if (headerRow < 0) { log(ss, "❌ Header bulunamadı"); return; }

    // Sütun mapping (öncelikli)
    var colMap = { vendor: 1, date: 2, invoiceNo: 5, total: 8, status: 9 };
    var headers = data[headerRow];
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c]).toLowerCase().trim();
      if (h.indexOf("service provider") >= 0) colMap.vendor = c;
      if (h === "invoice date" || h === "invoice dates") colMap.date = c;
      else if (h === "date" && !colMap.dateSet) { colMap.date = c; colMap.dateSet = true; }
      if (h.indexOf("invoice") >= 0 && h.indexOf("#") >= 0) colMap.invoiceNo = c;
      if (h === "total") colMap.total = c;
      if (h === "status") colMap.status = c;
    }

    // Import
    var knownSheet = ss.getSheetByName(CONFIG.SHEET_KNOWN);
    var lastKnownRow = knownSheet.getLastRow();
    if (lastKnownRow > 1) knownSheet.getRange(2, 1, lastKnownRow - 1, 4).clearContent();

    var importRows = [];
    var seen = {};
    for (var r = headerRow + 1; r < data.length; r++) {
      var row = data[r];
      var vendor = String(row[colMap.vendor] || "").trim();
      var invNo = normalizeInvoiceNo(row[colMap.invoiceNo]);
      if (!invNo || !vendor) continue;
      if (vendor.indexOf("Service Provider") >= 0 || vendor === "Entity") continue;
      if (invNo.indexOf("TOTAL") >= 0 || invNo.indexOf("TOBEPAID") >= 0) continue;
      if (seen[invNo]) continue;
      seen[invNo] = true;

      importRows.push([vendor, invNo, row[colMap.total] || "", String(row[colMap.status] || "")]);
    }

    if (importRows.length > 0) {
      knownSheet.getRange(2, 1, importRows.length, 4).setValues(importRows);
    }

    log(ss, "✅ " + importRows.length + " fatura import edildi");
    props.setProperty("LAST_SYNC_DATE", latestDate.toISOString());

  } finally {
    try { DriveApp.getFileById(tempFile.id).setTrashed(true); } catch(e) {}
  }
}

// ══════════════════════════════════════════════════════════════
// CFO RAPORU (HTML Mail)
// ══════════════════════════════════════════════════════════════
function sendCfoReport(ss, data) {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
  var newCount = data.invoices.length;
  var suspCount = data.suspicious.length;
  var pdfFailCount = data.unreadablePdfs.length;
  var missCount = data.missingVendors.length;

  // Subject
  var subject = "🧾 212 Finans Raporu — " + today;
  var alerts = [];
  if (newCount > 0) alerts.push(newCount + " yeni fatura");
  if (pdfFailCount > 0) alerts.push(pdfFailCount + " okunamayan PDF");
  if (suspCount > 0) alerts.push(suspCount + " şüpheli");
  if (missCount > 0) alerts.push(missCount + " eksik vendor");
  if (alerts.length > 0) subject += " ⚠️ " + alerts.join(", ");

  // Toplam tutarlar
  var totals = {};
  for (var i = 0; i < data.invoices.length; i++) {
    var inv = data.invoices[i];
    var amt = parseFloat(inv.amount) || 0;
    var cur = inv.currency || "TRY";
    if (amt > 0) totals[cur] = (totals[cur] || 0) + amt;
  }

  var html = '<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;">';

  // Header
  html += '<div style="background:#1a237e;color:white;padding:20px;border-radius:8px 8px 0 0;">';
  html += '<h2 style="margin:0;">🧾 212 Invoice Tracker — CFO Raporu</h2>';
  html += '<p style="margin:5px 0 0;opacity:0.8;">' + today + '</p></div>';

  // Stats
  html += '<div style="display:flex;background:#f5f5f5;padding:15px;gap:10px;">';
  html += statBox(newCount, "Yeni Fatura", newCount > 0 ? "#4caf50" : "#9e9e9e");
  html += statBox(pdfFailCount, "Okunamayan PDF", pdfFailCount > 0 ? "#ff9800" : "#9e9e9e");
  html += statBox(suspCount, "Şüpheli", suspCount > 0 ? "#ff9800" : "#9e9e9e");
  html += statBox(missCount, "Eksik Vendor", missCount > 0 ? "#f44336" : "#9e9e9e");
  html += '</div>';

  // Toplam tutarlar
  var curKeys = Object.keys(totals);
  if (curKeys.length > 0) {
    html += '<div style="display:flex;padding:10px 15px;gap:10px;">';
    for (var c = 0; c < curKeys.length; c++) {
      html += '<div style="background:white;border-top:3px solid #1a73e8;padding:12px;border-radius:6px;text-align:center;flex:1;">';
      html += '<div style="font-size:10px;color:#666;">' + esc(curKeys[c]) + ' TOPLAM</div>';
      html += '<div style="font-size:18px;font-weight:bold;">' + totals[curKeys[c]].toLocaleString() + '</div></div>';
    }
    html += '</div>';
  }

  // Yeni faturalar
  if (newCount > 0) {
    html += '<div style="padding:15px;">';
    html += '<h3 style="color:#1a237e;border-bottom:1px solid #eee;padding-bottom:8px;">✅ Bulunan Faturalar (' + newCount + ')</h3>';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    for (var i = 0; i < data.invoices.length; i++) {
      var inv = data.invoices[i];
      var bg = inv.isVip ? "background:#FFF3E0;" : "";
      html += '<tr style="' + bg + '">';
      html += '<td style="padding:8px;border-bottom:1px solid #eee;">' + (inv.isVip ? "👑 " : "") + '<b>' + esc(inv.vendor) + '</b>';
      if (inv.invoiceNo) html += '<br><small style="color:#666;">#' + esc(inv.invoiceNo) + '</small>';
      html += '</td>';
      html += '<td style="padding:8px;border-bottom:1px solid #eee;text-align:right;"><b>' + esc(inv.amount) + ' ' + esc(inv.currency) + '</b></td>';
      html += '<td style="padding:8px;border-bottom:1px solid #eee;text-align:center;"><a href="' + esc(inv.link) + '">Aç</a></td>';
      html += '</tr>';
    }
    html += '</table></div>';
  }

  // Şüpheli
  if (suspCount > 0) {
    html += '<div style="background:#FFF8E1;padding:15px;margin:10px 15px;border-radius:8px;">';
    html += '<h3 style="margin:0 0 10px;color:#ff9800;">⚠️ İnceleme Gereken Mailler</h3>';
    html += '<ul style="font-size:12px;">';
    for (var i = 0; i < data.suspicious.length; i++) {
      var s = data.suspicious[i];
      html += '<li>' + esc(s.from) + ' — ' + esc(s.reason) + ' — <a href="' + esc(s.link) + '">Kontrol Et</a></li>';
    }
    html += '</ul></div>';
  }

  // Okunamayan PDF'ler
  if (pdfFailCount > 0) {
    html += '<div style="background:#FFEBEE;padding:15px;margin:10px 15px;border-radius:8px;">';
    html += '<h3 style="margin:0 0 10px;color:#d32f2f;">🔴 Okunamayan PDF\'ler</h3>';
    html += '<ul style="font-size:12px;">';
    for (var i = 0; i < data.unreadablePdfs.length; i++) {
      var p = data.unreadablePdfs[i];
      html += '<li>' + esc(p.fileName) + ' — <a href="' + esc(p.link) + '">Maili Aç</a></li>';
    }
    html += '</ul></div>';
  }

  // Eksik vendor'lar
  if (missCount > 0) {
    html += '<div style="background:#FFEBEE;padding:15px;margin:10px 15px;border-radius:8px;">';
    html += '<h3 style="margin:0 0 10px;color:#d32f2f;">❌ Eksik Vendor\'lar</h3>';
    html += '<ul style="font-size:12px;">';
    for (var i = 0; i < data.missingVendors.length; i++) {
      html += '<li>' + esc(data.missingVendors[i]) + '</li>';
    }
    html += '</ul></div>';
  }

  // Footer
  html += '<div style="padding:15px;text-align:center;">';
  html += '<a href="https://docs.google.com/spreadsheets/d/' + CONFIG.SPREADSHEET_ID + '/edit" ';
  html += 'style="background:#1a237e;color:white;padding:10px 25px;text-decoration:none;border-radius:5px;">';
  html += '📊 Tracker Sheet\'i Aç</a></div></div>';

  try {
    GmailApp.sendEmail(CONFIG.NOTIFICATION_EMAIL, subject, "", {
      htmlBody: html, name: "212 Invoice Tracker"
    });
  } catch(e) {
    Logger.log("Mail gönderilemedi: " + e.message);
  }
}

function statBox(num, label, color) {
  return '<div style="flex:1;background:white;border-radius:8px;padding:15px;text-align:center;border-top:3px solid ' + color + ';">' +
    '<div style="font-size:28px;font-weight:bold;color:' + color + ';">' + num + '</div>' +
    '<div style="font-size:12px;color:#666;">' + label + '</div></div>';
}

function esc(s) {
  return String(s || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;");
}

// ══════════════════════════════════════════════════════════════
// YARDIMCI
// ══════════════════════════════════════════════════════════════
function log(ss, msg) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_LOG);
  if (sheet) sheet.appendRow([new Date(), msg]);
  Logger.log(msg);
}

// ══════════════════════════════════════════════════════════════
// TEMP DOSYA TEMİZLİĞİ
// ══════════════════════════════════════════════════════════════
function cleanupTempFiles() {
  var queries = ['title contains "TEMP_OCR_"', 'title contains "TEMP_SYNC_"'];
  var cutoff = new Date();
  cutoff.setHours(cutoff.getHours() - 24);
  var count = 0;
  for (var q = 0; q < queries.length; q++) {
    var files = DriveApp.searchFiles(queries[q] + " and trashed = false");
    while (files.hasNext()) {
      var file = files.next();
      if (file.getDateCreated() < cutoff) { file.setTrashed(true); count++; }
    }
  }
  if (count > 0) Logger.log("🗑️ " + count + " temp dosya temizlendi");
}

// ══════════════════════════════════════════════════════════════
// KURULUM & MENÜ
// ══════════════════════════════════════════════════════════════
function setupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);

  ScriptApp.newTrigger("scanInvoices").timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger("weeklyScan").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  ScriptApp.newTrigger("cleanupTempFiles").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();

  Logger.log("✅ Trigger'lar kuruldu: Günlük 08:00, Haftalık Pazartesi 09:00, Temizlik Pazar 03:00");
}

function weeklyScan() {
  var originalDays = CONFIG.SCAN_DAYS;
  CONFIG.SCAN_DAYS = 30;
  scanInvoices();
  CONFIG.SCAN_DAYS = originalDays;
}

function resetCache() {
  PropertiesService.getScriptProperties().deleteProperty("PROCESSED_IDS");
  PropertiesService.getScriptProperties().deleteProperty("LAST_SYNC_DATE");
  Logger.log("Hafıza sıfırlandı — sonraki çalışma sıfırdan tarayacak");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧾 212 Finans")
    .addItem("🔍 Şimdi Tara (AI)", "scanInvoices")
    .addItem("📊 Dashboard Güncelle", "updateDashboard")
    .addSeparator()
    .addItem("🔄 Hafıza Sıfırla", "resetCache")
    .addItem("⚙️ Trigger Kur", "setupTriggers")
    .addToUi();
}