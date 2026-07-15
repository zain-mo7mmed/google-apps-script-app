// ==========================================
// إرسال التقرير اليومي وتوليد ملفات الإكسل
// ==========================================

function generateAndSendDailyReport() {
  // 1. إعدادات الإيميلات (أدخل الإيميلات هنا، مفصولة بفارزة)
  const recipients = ["zmohammed@alfarah.com.iq", "mmohamed@alfarah.com.iq", "mhekmat@alfarah.com.iq"];
  const companyLogoUrl = "https://i.ibb.co/Z18h62nh/Al-Farah-Logo-Transparent-Background-01.png"; // رابط شعار شركة الفرح

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemsSheet = ss.getSheetByName("Items");
  const dailyStockSheet = ss.getSheetByName("Daily Stock");

  if (!itemsSheet) throw new Error("ورقة Items غير موجودة.");
  if (!dailyStockSheet) throw new Error("ورقة Daily Stock غير موجودة.");

  // الحصول على تاريخ أمس
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  const dateString = Utilities.formatDate(yesterday, "GMT+3", "yyyy-MM-dd");
  const fullDateAr = dateString + " " + yesterday.toLocaleDateString("ar-IQ", { weekday: "long" });

  // 2. جلب كافة المواد من Items مع دعم أعمدة المخزون الجديدة
  const itemsData = itemsSheet.getDataRange().getDisplayValues();
  const itemHeaders = itemsData[0] || [];
  const allItems = itemsData.slice(1); // إزالة الرأس
  const stockColumns = getStockColumnIndexes_(itemHeaders);

  // 3. جلب بيانات تقرير النقوصات (الجرد اليومي) لتاريخ أمس
  const dailyData = dailyStockSheet.getDataRange().getDisplayValues();
  const todayAudits = dailyData.filter(row => row[4] === dateString);

  if (todayAudits.length === 0) {
    Logger.log("لا يوجد تقرير نقوصات مسجل لهذا اليوم لإرساله.");
    return;
  }

  // 4. تعريف التقارير المطلوبة وربط كل تقرير بمحافظاته وعمود المخزون الخاص به
  const reportConfigs = [
    {
      key: "kurdistan",
      title: "تقرير كردستان",
      fileTitle: "تقرير كردستان",
      stockIndex: stockColumns.erbil,
      provinces: ["أربيل", "اربيل", "السليمانية", "دهوك"]
    },
    {
      key: "baghdad",
      title: "تقرير بغداد",
      fileTitle: "تقرير بغداد",
      stockIndex: stockColumns.baghdad,
      provinces: ["بغداد"]
    },
    {
      key: "middleEuphrates",
      title: "تقرير الفرات الأوسط",
      fileTitle: "تقرير الفرات الأوسط",
      stockIndex: stockColumns.najaf,
      provinces: ["النجف", "نجف", "الحلة", "حلة", "بابل", "كربلاء"]
    }
  ];

  const groupedReports = {};
  reportConfigs.forEach(config => {
    groupedReports[config.key] = {
      config,
      marketGroups: {}
    };
  });

  todayAudits.forEach(row => {
    const province = String(row[1] || "").trim();
    const reportConfig = reportConfigs.find(config => provinceMatches_(province, config.provinces));
    if (!reportConfig) return;

    const marketName = String(row[3] || "").trim();
    const reportGroup = groupedReports[reportConfig.key].marketGroups;
    if (!reportGroup[marketName]) {
      reportGroup[marketName] = {
        promoter: row[0],
        province: row[1],
        marketCode: row[2],
        presentItemCodes: []
      };
    }
    reportGroup[marketName].presentItemCodes.push(row[5]); // كود الصنف
  });

  // 5. إدارة المجلد والملفات في جوجل درايف
  let folder;
  const folders = DriveApp.getFoldersByName("تقارير النقوصات");
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder("تقارير النقوصات");
  }

  const reportFiles = reportConfigs.map(config => {
    const marketGroups = groupedReports[config.key].marketGroups;
    return createRegionalDailyReportFile_(config, marketGroups, allItems, dateString, folder);
  });

  // 6. بناء محتوى الإيميل HTML بدون جدول مضمن
  const reportLinksHtml = reportFiles.map(report => `
    <div style="border: 1px solid #e2e8f0; border-radius: 10px; padding: 14px 16px; margin: 10px 0; background-color: #f8fafc;">
      <div style="font-weight: 800; color: #1f2937; margin-bottom: 4px;">${report.title}</div>
      <div style="font-size: 13px; color: #64748b; margin-bottom: 12px;">عدد الماركتات: ${report.marketCount}</div>
      <a href="${report.url}" style="background-color: #16a34a; color: white; padding: 10px 22px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 14px; display: inline-block;">فتح الملف</a>
    </div>
  `).join("");

  const emailHtml = `
    <div dir="rtl" style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #1f2937; max-width: 650px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
      
      <!-- رأس الإيميل -->
      <div style="background-color: #0f766e; padding: 25px 20px; text-align: center; color: white;">
        <h1 style="margin: 0; font-size: 26px; font-weight: bold; letter-spacing: 0.5px;">تقرير النقوصات</h1>
        <div style="display: inline-block; background: rgba(255,255,255,0.15); padding: 6px 18px; border-radius: 20px; font-size: 13px; margin-top: 12px; letter-spacing: 0.5px;">${fullDateAr}</div>
      </div>
      
      <!-- جسم الإيميل -->
      <div style="padding: 30px 25px; line-height: 1.7; background-color: #ffffff;">
        <p style="margin-top: 0; font-weight: bold; font-size: 17px; color: #111827;">السادة المحترمون،</p>
        <p style="margin-bottom: 20px; color: #4b5563;">تحية طيبة..</p>
        <p style="color: #374151;">إليكم ملفات تقرير النقوصات اليومي ليوم أمس، وقد تم تقسيمها حسب المنطقة واعتماد عمود المخزون المناسب لكل تقرير:</p>
        
        <div style="margin: 25px 0;">${reportLinksHtml}</div>

        <p style="color: #4b5563; margin-bottom: 5px;">مع فائق الود والتقدير،</p>
      </div>

      <!-- التوقيع الاحترافي -->
      <div style="background-color: #f8fafc; padding: 25px; border-top: 1px solid #e2e8f0;">
        <table style="width: 100%; border-collapse: collapse;">
          <tr>
            <td style="vertical-align: top; padding-right: 15px;">
              <p style="margin: 0 0 5px 0; font-weight: 900; color: #0f766e; font-size: 18px; letter-spacing: 0.5px;">زين محمد</p>
              <p style="margin: 0 0 15px 0; font-size: 13px; color: #64748b; font-weight: 600;">مطور نظم برمجية</p>
              
              <table style="font-size: 13px; color: #475569; border-collapse: separate; border-spacing: 0 6px;">
                <tr>
                  <td style="width: 20px; color: #0f766e;">📞</td>
                  <td><a href="tel:+9647767798799" style="color: #475569; text-decoration: none;">+964 776 779 8799</a></td>
                </tr>
                <tr>
                  <td style="width: 20px; color: #0f766e;">✉️</td>
                  <td><a href="mailto:zmohammed@alfarah.com.iq" style="color: #475569; text-decoration: none;">zmohammed@alfarah.com.iq</a></td>
                </tr>
                <tr>
                  <td style="width: 20px; color: #0f766e;">🌐</td>
                  <td><a href="http://www.alfarah.com.iq" style="color: #0369a1; text-decoration: none; font-weight: 600;">www.alfarah.com.iq</a></td>
                </tr>
              </table>
            </td>
            <td style="vertical-align: middle; text-align: left; width: 120px;">
              <img src="${companyLogoUrl}" width="100" style="display: block; margin-left: auto;" alt="Al-Farah Logo">
            </td>
          </tr>
        </table>
      </div>
      
    </div>
  `;

  // 7. إرسال الإيميل
  MailApp.sendEmail({
    to: recipients.join(","),
    subject: "تقرير النقوصات اليومي - " + dateString,
    htmlBody: emailHtml
  });

  cleanupOldReportFiles_(folder, 1);
}

function createRegionalDailyReportFile_(config, marketGroups, allItems, dateString, folder) {
  const fileName = "[" + dateString + "] - " + config.fileTitle;
  const newReportFile = SpreadsheetApp.create(fileName);
  const fileId = newReportFile.getId();

  const fileInDrive = DriveApp.getFileById(fileId);
  folder.addFile(fileInDrive);
  DriveApp.getRootFolder().removeFile(fileInDrive);

  const marketNames = Object.keys(marketGroups);

  if (marketNames.length === 0) {
    const emptySheet = newReportFile.getSheets()[0];
    emptySheet.setName("لا توجد تقارير");
    emptySheet.getRange("A1:E1").merge().setValue(config.title + " - " + dateString)
      .setBackground("#f1f5f9").setFontWeight("bold").setHorizontalAlignment("center");
    emptySheet.getRange("A3:E3").merge().setValue("لا توجد تقارير نقوصات ضمن نطاق هذا التقرير لهذا التاريخ.")
      .setFontColor("#64748b").setHorizontalAlignment("center");
    emptySheet.setColumnWidths(1, 5, 160);
    return {
      title: config.title,
      url: newReportFile.getUrl(),
      marketCount: 0
    };
  }

  marketNames.forEach(marketName => {
    const group = marketGroups[marketName];
    const sheet = newReportFile.insertSheet(sanitizeSheetName_(marketName));

    sheet.getRange("A1:E1").merge().setValue(config.title + ": " + marketName)
      .setBackground("#f1f5f9").setFontWeight("bold").setHorizontalAlignment("center");

    sheet.appendRow(["إسم الماركت", marketName, "", "التاريخ", dateString]);
    sheet.appendRow(["إسم المروّج", group.promoter, "", "المحافظة", group.province]);
    sheet.appendRow([""]);

    const tableHeader = ["البراند", "كود الصنف", "إسم الصنف", "المخزون", "التواجد"];
    sheet.appendRow(tableHeader);
    const headerRange = sheet.getRange(sheet.getLastRow(), 1, 1, 5);
    headerRange.setBackground("#0f766e").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");

    const tableData = [];
    allItems.forEach(item => {
      const isPresent = group.presentItemCodes.includes(item[0]);
      const stockValue = item[config.stockIndex] || "";

      tableData.push([
        item[3], // Brand
        item[0], // Code
        item[2], // Name
        stockValue,
        isPresent ? "✔️" : ""
      ]);
    });

    if (tableData.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, tableData.length, 5).setValues(tableData);
    }

    const fullTableRange = sheet.getRange(5, 1, tableData.length + 1, 5);
    fullTableRange.setBorder(true, true, true, true, true, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
    sheet.setColumnWidths(1, 5, 150);
    sheet.setColumnWidth(3, 350);
    if (tableData.length > 0) {
      sheet.getRange(6, 5, tableData.length, 1).setFontColor("green").setFontWeight("bold").setHorizontalAlignment("center");
    }
  });

  const defaultSheet = newReportFile.getSheetByName("Sheet1");
  if (defaultSheet) newReportFile.deleteSheet(defaultSheet);

  return {
    title: config.title,
    url: newReportFile.getUrl(),
    marketCount: marketNames.length
  };
}

function getStockColumnIndexes_(headers) {
  return {
    baghdad: findHeaderIndex_(headers, ["Baghdad Stock", "Stock"], 4),
    erbil: findHeaderIndex_(headers, ["Erbil Stock"], 5),
    najaf: findHeaderIndex_(headers, ["Najaf Stock", "NajafStock"], 6)
  };
}

function findHeaderIndex_(headers, names, fallbackIndex) {
  const normalizedNames = names.map(normalizeHeader_);
  for (let i = 0; i < headers.length; i++) {
    if (normalizedNames.includes(normalizeHeader_(headers[i]))) return i;
  }
  return fallbackIndex;
}

function normalizeHeader_(value) {
  return String(value || "").toLowerCase().replace(/\s+/g, "");
}

function provinceMatches_(province, acceptedProvinces) {
  const normalizedProvince = normalizeArabicText_(province);
  return acceptedProvinces.some(name => normalizedProvince === normalizeArabicText_(name));
}

function normalizeArabicText_(value) {
  return String(value || "").trim().replace(/[أإآ]/g, "ا").replace(/\s+/g, "");
}

function sanitizeSheetName_(name) {
  const safeName = String(name || "Sheet").replace(/[\\\/\?\*\[\]\:]/g, "-").trim();
  return safeName.substring(0, 99) || "Sheet";
}

function cleanupOldReportFiles_(folder, monthsToKeep) {
  const cutoffDate = new Date();
  cutoffDate.setMonth(cutoffDate.getMonth() - monthsToKeep);

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() < cutoffDate) {
      file.setTrashed(true);
    }
  }
}
