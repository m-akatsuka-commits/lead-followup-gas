/**
 * 定数定義
 */
const SS = SpreadsheetApp.getActiveSpreadsheet();
const MAIN_SHEET_NAME = "2026.01以降"; 
const SETTING_SHEET_NAME = "設定";
const CC_ADDRESS = "knowleful-contact@clinks.jp";

/**
 * 設定シートから担当者情報を取得
 */
function getStaffInfoFromSheet(staffKey) {
  const settingSheet = SS.getSheetByName(SETTING_SHEET_NAME);
  const values = settingSheet.getRange("B15:B36").getValues();
  const staffKeys = ["西山", "伊藤", "箱崎", "村上", "藤野", "赤塚", "久野", "江島", "佐藤", "吉住", "時田"];
  const idx = staffKeys.indexOf(staffKey);
  
  if (idx === -1) return { fullName: staffKey, email: "" };
  
  return {
    fullName: values[idx * 2][0],
    email: values[idx * 2 + 1][0]
  };
}

/**
 * Webアプリ表示
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('ナレフル追客管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 署名を生成
 */
function generateSignatureFromTemplate(staffKey, settingSheet) {
  let template = settingSheet.getRange(8, 2).getValue();
  if (!template) return "";
  const info = getStaffInfoFromSheet(staffKey);
  template = template.replace(/<J列>/g, info.fullName);
  template = template.replace(/Mail:[^\s\n]*/, "Mail:" + info.email);
  return template;
}

/**
 * コンテンツ生成
 */
function generateMailContent(rowNum, step, selectedStaffKey) {
  const sheet = SS.getSheetByName(MAIN_SHEET_NAME);
  const settingSheet = SS.getSheetByName(SETTING_SHEET_NAME);
  const data = sheet.getRange(rowNum, 1, 1, 28).getValues()[0];
  
  const staffKey = selectedStaffKey || String(data[9]).trim(); 
  const staffInfo = getStaffInfoFromSheet(staffKey);
  const textSignature = generateSignatureFromTemplate(staffKey, settingSheet);
  
  const typeK = String(data[10]);
  const segmentD = String(data[3]);

  const linksMap = {
    "最上位": "https://app.spirinc.com/t/0NqgU99uz_dNR3kl_tZ_g/as/z6Q_30liqHsQ9DYhfpdys/confirm",
    "上位": "https://app.spirinc.com/t/0NqgU99uz_dNR3kl_tZ_g/as/aCJST23TlXVMollOxedQE/confirm",
    "一般": "https://app.spirinc.com/t/0NqgU99uz_dNR3kl_tZ_g/as/8mTyiZzbutMCO7rAiZnRe/confirm",
    "リスキリング研修": "https://app.spirinc.com/t/0NqgU99uz_dNR3kl_tZ_g/as/3h-P6_jpxDbyLNx93wPIv/confirm"
  };
  let link = linksMap["一般"];
  if (typeK.includes("リスキリング研修")) link = linksMap["リスキリング研修"];
  else if (segmentD.includes("最上位")) link = linksMap["最上位"];
  else if (segmentD.includes("上位")) link = linksMap["上位"];

  let baseIdx;
  const isReskilling = typeK.includes("リスキリング研修");
  if (isReskilling) {
    baseIdx = (step === 2) ? 9 : (step === 3 ? 11 : 13);
  } else {
    baseIdx = (step === 2) ? 2 : (step === 3 ? 4 : 6);
  }

  const subjectTemplate = settingSheet.getRange(baseIdx, 2).getValue();
  const bodyTemplate = settingSheet.getRange(baseIdx + 1, 2).getValue();

  const resolvedSubject = subjectTemplate.replace(/<J列>|<営業氏名>/g, staffKey);
  const resolvedBody = bodyTemplate
    .replace(/<G列>/g, data[6])
    .replace(/<F列>/g, data[5])
    .replace(/<営業氏名>/g, staffKey)
    .replace(/<J列>/g, staffInfo.fullName)
    .replace(/<日程調整リンク>/g, link);

  return {
    to: data[7],
    from: Session.getActiveUser().getEmail(),
    subject: resolvedSubject,
    body: resolvedBody,
    htmlBody: resolvedBody.replace(/\n/g, '<br>') + '<br><br>' + textSignature.replace(/\n/g, '<br>'),
    signature: textSignature,
    staffName: staffInfo.fullName,
    staffEmail: staffInfo.email,
    staffSurname: staffKey,
    company: data[6],
    contact: data[5],
    segment: segmentD,
    type: typeK,
    step: step,
    isReskilling: isReskilling
  };
}

/**
 * 送信実行
 */
function processMail(rowNum, step, selectedStaffKey) {
  const content = generateMailContent(rowNum, step, selectedStaffKey);
  
  GmailApp.sendEmail(content.to, content.subject, content.body, {
    cc: CC_ADDRESS,
    htmlBody: content.htmlBody,
    name: "CLINKS株式会社 " + content.staffName,
    replyTo: content.staffEmail
  });

  const sheet = SS.getSheetByName(MAIN_SHEET_NAME);
  const todayStr = Utilities.formatDate(new Date(), "JST", "MM/dd");
  const col = (step === 2) ? 13 : (step === 3 ? 14 : 15);
  sheet.getRange(rowNum, col).setValue(todayStr);
  
  // 返却ラベルも「追客2〜4」に合わせる
  const labelNum = step; 
  return `送信完了: ${content.company}様 宛 (追客${labelNum})`;
}

/**
 * UI用：行データ取得 (完了済みは除外)
 */
function getRowData(startRow, endRow) {
  const sheet = SS.getSheetByName(MAIN_SHEET_NAME);
  const numRows = endRow - startRow + 1;
  const data = sheet.getRange(startRow, 1, numRows, 28).getValues();
  return data.map((row, index) => {
    const hasData = row[6] || row[7];
    const statusInfo = getFollowUpStatus(row);
    return {
      rowNum: startRow + index,
      itemNum: row[0],
      isWon: row[1] === "商談獲得",
      isTerminated: row[27] === true,
      segment: row[3],
      company: row[6],
      type: row[10],
      status: !hasData ? "空行" : statusInfo.label,
      nextStep: (hasData) ? statusInfo.step : 0
    };
  }).filter(r => r.status !== "空行" && r.nextStep > 0);
}

/**
 * ステータス判定（ラベルを追客2, 3, 4に変更）
 */
function getFollowUpStatus(row) {
  if (row[1] === "商談獲得") return { label: "商談獲得済", step: 0 };
  if (row[27] === true) return { label: "追客終了フラグ", step: 0 }; 
  if (!row[11]) return { label: "1通目未送付", step: 0 };
  if (!row[12]) return { label: "追客2対象", step: 2 }; 
  if (!row[13]) return { label: "追客3対象", step: 3 }; 
  if (!row[14]) return { label: "追客4対象", step: 4 }; 
  return { label: "完了", step: 0 };
}

function getPreview(rowNum, step, selectedStaffKey) {
  return generateMailContent(rowNum, step, selectedStaffKey);
}
