/**
 * PERSONAL EXPENSE TRACKER - Final (captures Info: BOOKMYSHOW and all earlier cases)
 * - Info: <merchant> reliably extracted (e.g. Info: BOOKMYSHOW.)
 * - UPI VPA, UPI-ref, Reference ID extraction
 * - Duplicate rule: (ReferenceID OR UPI) + exact Date+Time required to skip
 * - MessageId idempotency, no email body saved
 */

/* CONFIGURATION */
const CONFIG = {
  MAIN_SHEET: 'All Transactions',
  SUMMARY_SHEET: 'Monthly Summary',
  CATEGORY_SHEET: 'Category Analysis',
  EMAIL_ACCOUNTS: [
    'youremail@gmail.com'
  ],
  DAYS_TO_SCAN: 1,
  MIN_AMOUNT: 1,
  DEBUG: true,
  SIMILARITY_THRESHOLD: 0.85,
  LEVENSHTEIN_MAX: 4
};

/* BANK PATTERNS */
const BANK_PATTERNS = {
  hdfc: { senders: ['alerts@hdfcbank.net', 'creditcards@hdfcbank.com'],
    patterns: {
      amount: [/Rs\.?\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i, /INR\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i],
      merchant: [/(?:at|paid to)\s+([A-Z0-9\s\-\.\/]+?)(?:\s+on|\s+dt|\.)/i, /Transaction at\s+([^,\n\r]+)/i, /merchant\s*[:\-]\s*([^,\n\r]+)/i],
      card: [/Card ending (\d{4})/i, /xx(\d{4})/i],
      type: { debit: ['debited','debit','purchase','withdrawal','spent'], credit: ['credited','credit','refund','received','deposit'] }
    }
  },
  icici: { senders: ['credit_cards@icicibank.com','no-reply@icicibank.com','cards@icicibank.com'],
    patterns: {
      amount: [/Rs\.?\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i, /INR\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i, /₹\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i, /amount\s*:?\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i],
      merchant: [/at\s+([A-Z0-9\s\-\.\/]+?)(?:\s+on|\s+dt|\.)/i, /towards\s+([^,\n\r]+)/i, /(?:info|description|remarks)[:\s]+([A-Za-z0-9\s\.\-&]+)/i],
      card: [/card\s*ending\s*(\d{4})/i, /xx(\d{4})/i],
      type: { debit: ['debited','debit','purchase','withdrawal','spent','used for a transaction','has been used for a transaction'], credit: ['credited','credit','refund','received','deposit'] }
    }
  },
  sbi: { senders: ['onlinesbicard@sbicard.com','sbicard@sbi.co.in'],
    patterns: {
      amount: [/Rs\.?\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i],
      merchant: [/at\s+([^,\n\r]+)/i, /merchant\s*:?\s*([^,\n\r]+)/i],
      card: [/ending\s*(\d{4})/i],
      type: { debit: ['debited','debit','purchase','withdrawal','spent'], credit: ['credited','credit','refund','received','deposit'] }
    }
  },
  axis: { senders: ['alerts@axisbank.com','cards@axisbank.com'],
    patterns: {
      amount: [/Rs\.?\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i, /INR\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i, /₹\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i],
      merchant: [/at\s+([^,\n\r]+)/i, /paid to\s+([^,\n\r]+)/i],
      card: [/ending\s*(\d{4})/i],
      type: { debit: ['debited','debit','purchase','withdrawal','spent'], credit: ['credited','credit','refund','received','deposit'] }
    }
  },
  iob: { senders: ['alerts@iob.in','sms@iob.in','noreply@iob.in'],
    patterns: {
      amount: [/Rs\.?\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i, /INR\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i, /₹\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i],
      merchant: [/at\s+([A-Z0-9\s\-\.\/]+?)(?:\s+on|\s+dt|\.)/i, /Transaction at\s+([^,\n\r]+)/i, /merchant\s*[:\-]\s*([^,\n\r]+)/i, /paid to\s+([^,\n\r]+)/i],
      card: [/card ending (\d{4})/i, /ending\s*(\d{4})/i, /xx(\d{4})/i],
      type: { debit: ['debited','debit','purchase','withdrawal','spent'], credit: ['credited','credit','refund','received','deposit'] }
    }
  }
};

/* MAIN */
function scanAllExpenseEmails() {
  try {
    log('Start scan...');
    const sheet = initializeSpreadsheet();
    const queries = buildSearchQueries();
    let processedCount = 0, newTransactions = 0;
    queries.forEach(q => {
      try {
        log('Searching', q);
        const threads = GmailApp.search(q);
        threads.forEach(thread => {
          thread.getMessages().forEach(message => {
            if (isProcessableMessage(message)) {
              const res = processExpenseEmail(message, sheet);
              if (res.success) { processedCount++; if (res.isNew) newTransactions++; }
              else log('Skipped:', res.reason);
            }
          });
        });
      } catch (e) { console.error('Query error', e); }
    });
    updateSummarySheets();
    log(`Done. Processed:${processedCount}, New:${newTransactions}`);
    if (newTransactions > 0) sendNotificationEmail(newTransactions);
  } catch (err) {
    console.error('scanAllExpenseEmails error', err);
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), 'Expense Tracker Error', `Error: ${err.toString()}`);
  }
}

/* Initialize spreadsheet */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.MAIN_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.MAIN_SHEET);
    const headers = ['Date','Time','Amount','Transaction Type','Merchant/Description','Category','Bank/Service','Card Number','UPI ID','Reference ID','Balance','Email Subject','MessageId','Processed At'];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 120);
  }
  return sheet;
}

/* Build queries */
function buildSearchQueries() {
  const dateFilter = `newer_than:${CONFIG.DAYS_TO_SCAN}d`;
  const queries = [];
  Object.values(BANK_PATTERNS).forEach(b => b.senders.forEach(s => queries.push(`from:(${s}) ${dateFilter}`)));
  queries.push(`subject:(transaction alert) ${dateFilter}`, `subject:(payment) ${dateFilter}`, `subject:(debited) ${dateFilter}`, `subject:(credited) ${dateFilter}`, `(UPI OR IMPS OR NEFT OR RTGS) ${dateFilter}`);
  return [...new Set(queries)];
}

/* Message filter */
function isProcessableMessage(message) {
  const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - CONFIG.DAYS_TO_SCAN);
  return message.getDate() > cutoff && !message.isInTrash();
}

/* Process single message */
function processExpenseEmail(message, sheet) {
  try {
    const subject = message.getSubject() || '';
    const body = message.getPlainBody() || '';
    const htmlBody = message.getBody() || '';
    const dateReceived = message.getDate();
    const sender = message.getFrom() || '';
    const messageId = (typeof message.getId === 'function') ? message.getId() : null;
    const thread = message.getThread();

    let processedLabel = GmailApp.getUserLabelByName('ExpenseTrackerProcessed'); if (!processedLabel) processedLabel = GmailApp.createLabel('ExpenseTrackerProcessed');

    if (messageId && isMessageIdInSheet(sheet, messageId)) {
      try { thread.addLabel(processedLabel); } catch(e){log('label add failed',e);}
      return { success: false, reason: 'Message already processed (MessageId)' };
    }

    const serviceInfo = identifyService(sender, subject, body, htmlBody);
    if (!serviceInfo) { try { thread.addLabel(processedLabel); } catch(e){}; return { success:false, reason:'Unknown service' }; }

    const tx = extractTransactionData(body, htmlBody, subject, serviceInfo);
    if (!tx.amount || tx.amount < CONFIG.MIN_AMOUNT) { try { thread.addLabel(processedLabel); } catch(e){}; return { success:false, reason:'No valid amount' }; }

    // parse date/time; if no time in parsed, use email receive time for time portion
    const parsed = parseDateVarious(subject + ' ' + body + ' ' + htmlBody);
    let usedDate = parsed ? parsed.date : dateReceived;
    if (parsed && !parsed.hasTime) { usedDate.setHours(dateReceived.getHours(), dateReceived.getMinutes(), dateReceived.getSeconds(), 0); }

    const usedDateStr = Utilities.formatDate(usedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const usedTimeStr = Utilities.formatDate(usedDate, Session.getScriptTimeZone(), 'HH:mm:ss');

    // duplicate by (ref OR upi) + exact date+time
    if (tx.referenceId || tx.upiId) {
      const found = findRowByUpiOrRefAndDateTime(sheet, tx.upiId, tx.referenceId, usedDateStr, usedTimeStr);
      if (found) { try { thread.addLabel(processedLabel); } catch(e){}; return { success:false, reason:'Duplicate by Reference/UPI + Date+Time' }; }
    }

    // fallback fuzzy duplicate
    if (isDuplicate(sheet, usedDate, tx.amount, tx.merchant)) { try { thread.addLabel(processedLabel); } catch(e){}; return { success:false, reason:'Duplicate transaction' }; }

    const row = [
      usedDateStr,
      usedTimeStr,
      tx.amount,
      tx.type,
      tx.merchant,
      categorizeExpense(tx.merchant, tx.amount),
      serviceInfo.name,
      tx.cardNumber || '',
      tx.upiId || '',
      tx.referenceId || '',
      tx.balance || '',
      subject,
      messageId || '',
      new Date()
    ];
    sheet.appendRow(row);
    try { thread.addLabel(processedLabel); } catch(e){ log('label add failed', e); }

    return { success:true, isNew:true };
  } catch (err) {
    console.error('processExpenseEmail error', err);
    return { success:false, reason:err.toString() };
  }
}

/* Find by UPI or Reference + exact date/time */
function findRowByUpiOrRefAndDateTime(sheet, upiId, referenceId, dateStr, timeStr) {
  try {
    if (!upiId && !referenceId) return null;
    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return null;
    const header = data[0].map(h => String(h||'').trim());
    const upiIdx = header.indexOf('UPI ID'), refIdx = header.indexOf('Reference ID'), dateIdx = header.indexOf('Date'), timeIdx = header.indexOf('Time');
    if ((upiIdx === -1 && refIdx === -1) || dateIdx === -1 || timeIdx === -1) return null;
    const normUpi = normalizeUpi(upiId||''), normRef = (referenceId||'').replace(/\D/g,'').trim();
    for (let i=1;i<data.length;i++) {
      const rowUpi = String((upiIdx!==-1?data[i][upiIdx]:'')||'').trim();
      const rowRef = String((refIdx!==-1?data[i][refIdx]:'')||'').trim();
      const rowDate = String(data[i][dateIdx]||'').trim();
      const rowTime = String(data[i][timeIdx]||'').trim();
      if (rowDate!==dateStr || rowTime!==timeStr) continue;
      if (normRef && rowRef && rowRef.replace(/\D/g,'').trim()===normRef) return {row:i+1,rowData:data[i]};
      if (normUpi && rowUpi && normalizeUpi(rowUpi)===normUpi) return {row:i+1,rowData:data[i]};
    }
    return null;
  } catch(e) { log('findRowByUpiOrRefAndDateTime error', e); return null; }
}

/* normalize upi (avoid support addresses) */
function normalizeUpi(upi) {
  if (!upi) return '';
  let s = String(upi).trim();
  s = s.replace(/^(UPI[:\-_s]*)/i, '');
  s = s.replace(/[^\w@.]/g,'').toLowerCase();
  if (/customer\.care|noreply|alerts|creditcards|service|support|sms@|no-reply|customersupport|info@/i.test(s)) return '';
  const sendersFlat = Object.values(BANK_PATTERNS).flatMap(b => b.senders.map(x => x.toLowerCase()));
  if (sendersFlat.includes(s)) return '';
  return s;
}

/* identify service */
function identifyService(sender, subject, body, htmlBody) {
  let email = sender;
  const m = sender.match(/<([^>]+)>/);
  if (m) email = m[1].toLowerCase(); else email = (sender||'').toLowerCase();
  const subj = (subject||'').toLowerCase(), btext = (body||'').toLowerCase() + ' ' + (htmlBody||'').toLowerCase();
  for (const [k,cfg] of Object.entries(BANK_PATTERNS)) if (cfg.senders.some(sp => email.includes(sp.toLowerCase()))) return {name:k.toUpperCase(),config:cfg,type:'bank'};
  for (const [k,cfg] of Object.entries(BANK_PATTERNS)) if (subj.includes(k) || btext.includes(k)) return {name:k.toUpperCase(),config:cfg,type:'bank'};
  if (subj.includes('upi') || btext.includes('upi') || subj.includes('googlepay') || btext.includes('googlepay') || subj.includes('phonepe') || btext.includes('paytm')) return {name:'UPI', config:{}, type:'upi'};
  return null;
}

/* extract transaction data (tuned) */
function extractTransactionData(body, htmlBody, subject, serviceInfo) {
  const config = serviceInfo.config || {};
  const fullText = `${subject} ${body} ${htmlBody}`;
  const fullTextLower = fullText.toLowerCase();

  // amount
  let amount = null;
  if (config.patterns && config.patterns.amount) {
    for (const p of config.patterns.amount) { const m = fullText.match(p); if (m && m[1]) { amount = toNumber(m[1]); break; } }
  }
  if (!amount) {
    const fall = [/(?:inr|rs|₹)\s*(\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?)/i, /(?:amount|txn|transaction)\s*[:\-]?\s*(\d+(?:,\d+)*(?:\.\d{2})?)/i];
    for (const p of fall) { const m = fullText.match(p); if (m && m[1]) { amount = toNumber(m[1]); break; } }
  }

  // merchant (patterns)
  let merchant = 'Unknown Merchant';
  if (config.patterns && config.patterns.merchant) {
    for (const p of config.patterns.merchant) {
      const m1 = fullText.match(p), m2 = body.match(p), m3 = subject.match(p), m4 = htmlBody.match(p);
      const match = m1 || m2 || m3 || m4;
      if (match && match[1]) { merchant = sanitizeMerchant(match[1]); break; }
    }
  }

  // HDFC VPA pattern e.g. "to VPA vpeee@okaxis VIGNESH on 07-09-25"
  let referenceId = null, upiId = null;
  const vpaPattern = /VPA\s+([a-zA-Z0-9.\-_]+@[a-zA-Z0-9.\-]+)\s+([A-Za-z0-9\s\.\-&'()]+?)\s+on\s+(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/i;
  const vpaMatch = fullText.match(vpaPattern);
  if (vpaMatch) {
    upiId = vpaMatch[1].trim();
    merchant = sanitizeMerchant(vpaMatch[2].trim());
  }

  // reference pattern: "reference number is 561627842670"
  const refPattern1 = /reference\s*(?:number|id)\s*(?:is|:)?\s*[:\s\-]*?(\d{6,20})/i;
  const refMatch1 = fullText.match(refPattern1);
  if (refMatch1 && refMatch1[1]) referenceId = refMatch1[1].trim();

  // Info: UPI-<digits>-Merchant  OR Info: MERCHANT.
  const infoUpiPattern = /Info[:\s]+UPI[-_\s]?(\d{6,15})(?:[-_\s]+([A-Za-z0-9\s\.\-&'()]+))?/i;
  const infoUpiMatch = fullText.match(infoUpiPattern);
  if (infoUpiMatch && infoUpiMatch[1]) {
    referenceId = referenceId || infoUpiMatch[1].trim();
    if (infoUpiMatch[2]) merchant = sanitizeMerchant(infoUpiMatch[2].trim());
  } else {
    // Info: <merchant> fallback (robust capture including single word + trailing dot)
    const infoMerchant = fullText.match(/Info[:\s]+([A-Za-z0-9\s\.\-&'()]{1,80}?)(?:[.\n]|$)/i);
    if (infoMerchant && infoMerchant[1]) {
      const cand = sanitizeMerchant(infoMerchant[1]);
      if (!merchant || /unknown/i.test(merchant) || merchant.length < 3) merchant = cand;
      else merchant = cand; // prefer Info: merchant because that's authoritative for many card notifications
    }
  }

  // UPI VPA like vpeee@okaxis or vpa-like shorter forms (avoid support addresses)
  const upiMailMatch = fullText.match(/([a-zA-Z0-9.\-_]{2,}@[a-zA-Z0-9.\-]{2,}\.[a-zA-Z]{2,})/i);
  if (upiMailMatch && upiMailMatch[1]) {
    const candidate = upiMailMatch[1].trim();
    if (!isLikelySupportOrSenderAddress(candidate)) upiId = upiId || candidate;
  }
  if (!upiId) {
    const upiAlt = fullText.match(/([a-zA-Z0-9.\-_]{2,}@[a-zA-Z]{2,})/i);
    if (upiAlt && upiAlt[1] && !isLikelySupportOrSenderAddress(upiAlt[1])) upiId = upiAlt[1].trim();
  }

  merchant = sanitizeMerchant(merchant);

  // type detection
  let transactionType = 'Debit';
  const explicitDebit = ['has been used for a transaction','used for a transaction','was used for a transaction','transaction of','debited','debit','paid','sent','has been debited'];
  const explicitCredit = ['credited','refund','deposit','has been credited','was credited','you have received','received'];
  for (const p of explicitDebit) if (fullTextLower.indexOf(p)!==-1) { transactionType='Debit'; break; }
  for (const p of explicitCredit) if (fullTextLower.indexOf(p)!==-1) { transactionType='Credit'; break; }

  // card number
  let cardNumber = null;
  if (config.patterns && config.patterns.card) {
    for (const pat of config.patterns.card) { const mm = fullText.match(pat); if (mm && mm[1]) { cardNumber = '****' + mm[1]; break; } }
  } else {
    const mxx = fullText.match(/xx(\d{4})/i); if (mxx && mxx[1]) cardNumber = '****' + mxx[1];
  }

  // balance
  let balance = null;
  const balancePattern = /balance[:\s]*(?:rs\.?\s*)?(\d+(?:,\d+)*(?:\.\d{2})?)/i;
  const balanceMatch = fullText.match(balancePattern);
  if (balanceMatch) balance = toNumber(balanceMatch[1]);

  return { amount, merchant, type: transactionType, cardNumber, upiId, referenceId, balance };
}

/* Helpers */
function isLikelySupportOrSenderAddress(addr) {
  if (!addr) return false;
  const s = String(addr).toLowerCase();
  if (/customer\.care|noreply|alerts|creditcards|service|support|sms@|no-reply|customersupport|info@|helpdesk/i.test(s)) return true;
  const sendersFlat = Object.values(BANK_PATTERNS).flatMap(b => b.senders.map(x => x.toLowerCase()));
  if (sendersFlat.includes(s)) return true;
  return false;
}
function toNumber(s) { if (!s) return null; try { return parseFloat(String(s).replace(/[^0-9.\-]/g,'').replace(/,/g,'')) || null; } catch(e){return null;} }
function sanitizeMerchant(raw) {
  if (!raw) return 'Unknown Merchant';
  let m = String(raw).trim();
  m = m.replace(/UPI[-_\s]*\d+/ig,'');
  m = m.replace(/\btxn[:\-]?\d+\b/ig,'');
  m = m.replace(/\bref[:\-]?\d+\b/ig,'');
  m = m.replace(/\b(?:info|remarks|description)[:\-]?\s*/ig,'');
  m = m.replace(/[-_\/\|]+/g,' ');
  m = m.replace(/\s{2,}/g,' ');
  m = m.replace(/^[^A-Za-z0-9]+|[^A-Za-z0-9]+$/g,'');
  m = m.replace(/[.,\s]+$/g,'');
  return m || 'Unknown Merchant';
}
function categorizeExpense(merchant, amount) {
  const cats = { 'Food & Dining':['swiggy','zomato','dominos','pizza','restaurant','cafe','hotel','kfc','mcdonalds','subway','dunkin','starbucks','food','dining'], 'Transportation':['uber','ola','metro','irctc','bus','train','taxi','fuel','petrol','diesel','fastag','toll','parking','auto'], 'Shopping':['amazon','flipkart','myntra','ajio','nykaa','mall','store','shop','market'], 'Groceries':['bigbasket','grofers','blinkit','instamart','grocery','supermarket'], 'Utilities':['electricity','water','gas','internet','broadband','wifi','phone','bill','recharge'], 'Entertainment':['netflix','prime','hotstar','spotify','movie','cinema','game','book'], 'Healthcare':['hospital','pharmacy','medical','doctor','clinic'], 'Investment':['mutual fund','sip','fd','shares','zerodha','groww','upstox'], 'Education':['school','college','course','tuition'] };
  const ml = (merchant||'').toLowerCase();
  for (const [k,w] of Object.entries(cats)) if (w.some(x=>ml.includes(x))) return k;
  if (amount>10000) return 'High Value';
  if (amount<50) return 'Small Transactions';
  return 'Others';
}

/* Duplicate fuzzy */
function isDuplicate(sheet, date, amount, merchant) {
  try {
    const data = sheet.getDataRange().getValues();
    if (!data || data.length<2) return false;
    const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const mnorm = (merchant||'').toLowerCase().trim();
    for (let i=1;i<data.length;i++){
      const row = data[i]; const rowDate=row[0]; const rowAmount = parseFloat(row[2])||0; const rowMerchant=(row[4]||'').toString().toLowerCase();
      if (rowDate===dateStr && Math.abs(rowAmount-amount)<0.01) {
        if (rowMerchant===mnorm) return true;
        if (rowMerchant && (rowMerchant.indexOf(mnorm)!==-1||mnorm.indexOf(rowMerchant)!==-1)) return true;
        const lev = levenshteinDistance(shorten(rowMerchant), shorten(mnorm));
        const maxLen = Math.max(shorten(rowMerchant).length, shorten(mnorm).length, 1);
        const sim = 1 - (lev/maxLen);
        if ((maxLen <= CONFIG.LEVENSHTEIN_MAX && lev <= CONFIG.LEVENSHTEIN_MAX) || sim >= CONFIG.SIMILARITY_THRESHOLD) return true;
      }
    }
    return false;
  } catch(e) { log('isDuplicate error',e); return false; }
}
function shorten(s){ if(!s) return ''; return s.replace(/\d+/g,'').replace(/\b(upi|txn|ref|info|neft|imps|rtgs)\b/ig,'').replace(/[^a-z]/ig,'').substr(0,40); }
function levenshteinDistance(a,b){ if(!a||!b) return (a?a.length:0)+(b?b.length:0); const m=a.length,n=b.length; const d=Array.from({length:m+1},()=>new Array(n+1).fill(0)); for(let i=0;i<=m;i++)d[i][0]=i; for(let j=0;j<=n;j++)d[0][j]=j; for(let i=1;i<=m;i++){ for(let j=1;j<=n;j++){ const cost = a[i-1]===b[j-1]?0:1; d[i][j]=Math.min(d[i-1][j]+1,d[i][j-1]+1,d[i-1][j-1]+cost); } } return d[m][n]; }

/* MessageId check */
function isMessageIdInSheet(sheet, messageId) {
  if (!messageId) return false;
  try {
    const data = sheet.getDataRange().getValues();
    if (!data || data.length<2) return false;
    const idx = data[0].indexOf('MessageId');
    if (idx===-1) return false;
    for (let i=1;i<data.length;i++) if (String(data[i][idx]).trim()===String(messageId).trim()) return true;
    return false;
  } catch(e){ log('isMessageIdInSheet error', e); return false; }
}

/* Summary functions (same as earlier) */
function updateSummarySheets(){ const ss=SpreadsheetApp.getActiveSpreadsheet(); createMonthlySummarySheet(ss); createCategoryAnalysisSheet(ss); }
function createMonthlySummarySheet(ss){ let summary=ss.getSheetByName(CONFIG.SUMMARY_SHEET); if(!summary) summary=ss.insertSheet(CONFIG.SUMMARY_SHEET); else summary.clear(); summary.getRange('A1:E1').setValues([['Month','Total Debit','Total Credit','Net Spending','Transaction Count']]); summary.getRange('A1:E1').setFontWeight('bold'); const currentMonth=new Date(); for(let i=0;i<6;i++){ const month=new Date(currentMonth.getFullYear(), currentMonth.getMonth()-i,1); const row=i+2; summary.getRange(`A${row}`).setValue(Utilities.formatDate(month, Session.getScriptTimeZone(), 'yyyy-MM')); summary.getRange(`B${row}`).setFormula(`=SUMIFS('${CONFIG.MAIN_SHEET}'!C:C,'${CONFIG.MAIN_SHEET}'!A:A,">="&DATE(${month.getFullYear()},${month.getMonth()+1},1),'${CONFIG.MAIN_SHEET}'!A:A,"<"&DATE(${month.getFullYear()},${month.getMonth()+2},1),'${CONFIG.MAIN_SHEET}'!D:D,"Debit")`); summary.getRange(`C${row}`).setFormula(`=SUMIFS('${CONFIG.MAIN_SHEET}'!C:C,'${CONFIG.MAIN_SHEET}'!A:A,">="&DATE(${month.getFullYear()},${month.getMonth()+1},1),'${CONFIG.MAIN_SHEET}'!A:A,"<"&DATE(${month.getFullYear()},${month.getMonth()+2},1),'${CONFIG.MAIN_SHEET}'!D:D,"Credit")`); summary.getRange(`D${row}`).setFormula(`=B${row}-C${row}`); summary.getRange(`E${row}`).setFormula(`=COUNTIFS('${CONFIG.MAIN_SHEET}'!A:A,">="&DATE(${month.getFullYear()},${month.getMonth()+1},1),'${CONFIG.MAIN_SHEET}'!A:A,"<"&DATE(${month.getFullYear()},${month.getMonth()+2},1))`); } summary.getRange('B:D').setNumberFormat('₹#,##0.00'); }
function createCategoryAnalysisSheet(ss){ let cat=ss.getSheetByName(CONFIG.CATEGORY_SHEET); if(!cat) cat=ss.insertSheet(CONFIG.CATEGORY_SHEET); else cat.clear(); cat.getRange('A1:C1').setValues([['Category','Total Spent','Transaction Count']]); cat.getRange('A1:C1').setFontWeight('bold'); const categories=['Food & Dining','Transportation','Shopping','Groceries','Utilities','Entertainment','Healthcare','Investment','Education','Others']; categories.forEach((category, idx)=>{ const row=idx+2; cat.getRange(`A${row}`).setValue(category); cat.getRange(`B${row}`).setFormula(`=SUMIFS('${CONFIG.MAIN_SHEET}'!C:C,'${CONFIG.MAIN_SHEET}'!F:F,"${category}",'${CONFIG.MAIN_SHEET}'!D:D,"Debit")`); cat.getRange(`C${row}`).setFormula(`=COUNTIFS('${CONFIG.MAIN_SHEET}'!F:F,"${category}",'${CONFIG.MAIN_SHEET}'!D:D,"Debit")`); }); cat.getRange('B:B').setNumberFormat('₹#,##0.00'); }

/* Notifications & utilities */
function sendNotificationEmail(newTransactionCount){ try{ const email=Session.getActiveUser().getEmail(); const subject=`Expense Tracker: ${newTransactionCount} new transactions processed`; const body=`Your expense tracker has processed ${newTransactionCount} new transactions.\n\nView your expense sheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}\n\nThis is an automated message from your Personal Expense Tracker.`; GmailApp.sendEmail(email, subject, body); }catch(e){console.error('Failed to send notification email',e);} }
function setupAutomatedTrigger(){ ScriptApp.getProjectTriggers().forEach(trigger=>{ if(trigger.getHandlerFunction()==='scanAllExpenseEmails') ScriptApp.deleteTrigger(trigger); }); ScriptApp.newTrigger('scanAllExpenseEmails').timeBased().everyHours(6).create(); log('Trigger set'); }
function testSystem(){ log('Testing (2 days)'); const o=CONFIG.DAYS_TO_SCAN; CONFIG.DAYS_TO_SCAN=2; try{ scanAllExpenseEmails(); log('Test done'); }catch(e){ console.error('Test failed',e); } CONFIG.DAYS_TO_SCAN=o; }
function clearAllData(){ const resp=Browser.msgBox('Clear All Data','This will delete all transaction data. Are you sure?', Browser.Buttons.YES_NO); if(resp==='yes'){ const ss=SpreadsheetApp.getActiveSpreadsheet(); [CONFIG.MAIN_SHEET, CONFIG.SUMMARY_SHEET, CONFIG.CATEGORY_SHEET].forEach(n=>{ const s=ss.getSheetByName(n); if(s) ss.deleteSheet(s); }); log('Cleared'); } }
function getSystemStats(){ const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MAIN_SHEET); if(!sheet){ log('No data'); return; } const data=sheet.getDataRange().getValues(); const total=Math.max(0,data.length-1); if(total===0){ log('No tx'); return; } const totalDebit=data.slice(1).filter(r=>r[3]==='Debit').reduce((s,r)=>s+(parseFloat(r[2])||0),0); const totalCredit=data.slice(1).filter(r=>r[3]==='Credit').reduce((s,r)=>s+(parseFloat(r[2])||0),0); log(`SYSTEM: Total:${total} Debit:₹${totalDebit.toFixed(2)} Credit:₹${totalCredit.toFixed(2)} Net:₹${(totalDebit-totalCredit).toFixed(2)} URL:${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`); }

/* Date parsing helpers */
function parseDateVarious(text){
  if(!text) return null;
  const long = text.match(/\b([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})\s+(?:at\s*)?(\d{1,2}:\d{2}:\d{2})\b/i);
  if(long){ const months={jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11}; const monKey=long[1].toLowerCase().substr(0,3); const mm=months[monKey]; if(mm===undefined) return null; const day=parseInt(long[2],10); const year=parseInt(long[3],10); const t=long[4].split(':'); return {date:new Date(year,mm,day,parseInt(t[0],10),parseInt(t[1],10),parseInt(t[2],10)), hasTime:true}; }
  const d1 = text.match(/\b(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})\b/);
  if(d1){ let day=parseInt(d1[1],10); let month=parseInt(d1[2],10)-1; let yr=parseInt(d1[3],10); if(yr<100) yr+=2000; return {date:new Date(yr,month,day,0,0,0), hasTime:false}; }
  return null;
}

/* Logger */
function log(){ if(!CONFIG.DEBUG) return; const args=Array.prototype.slice.call(arguments); console.log.apply(console,args); }
