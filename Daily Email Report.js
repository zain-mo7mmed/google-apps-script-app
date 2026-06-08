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
  



  //const today = new Date();
  //const dateString = Utilities.formatDate(today, "GMT+3", "yyyy-MM-dd");
  //const fullDateAr = Utilities.formatDate(today, "GMT+3", "yyyy-MM-dd") + " " + today.toLocaleDateString("ar-IQ", { weekday: "long" });
  
  
  // الحصول على تاريخ اليوم
  const yesterday = new Date();
  
  // استخدام دالة setDate لطرح يوم واحد مباشرة
  yesterday.setDate(yesterday.getDate() - 1);
  
  const dateString = Utilities.formatDate(yesterday, "GMT+3", "yyyy-MM-dd");
  
  // تنسيق التاريخ مع اليوم من الأسبوع باللغة العربية
  const fullDateAr = dateString + " " + yesterday.toLocaleDateString("ar-IQ", { weekday: "long" });





  // 2. جلب كافة المواد من Items
  const itemsData = itemsSheet.getDataRange().getDisplayValues();
  const allItems = itemsData.slice(1); // إزالة الرأس

  // 3. جلب بيانات تقرير النقوصات (الجرد اليومي) لتاريخ اليوم
  const dailyData = dailyStockSheet.getDataRange().getDisplayValues();
  const todayAudits = dailyData.filter(row => row[4] === dateString);

  if (todayAudits.length === 0) {
    Logger.log("لا يوجد تقرير نقوصات مسجل لهذا اليوم لإرساله.");
    return;
  }

  // 4. تجميع البيانات حسب الماركت
  const marketGroups = {};
  todayAudits.forEach(row => {
    const marketName = row[3];
    if (!marketGroups[marketName]) {
      marketGroups[marketName] = {
        promoter: row[0],
        province: row[1],
        marketCode: row[2],
        presentItemCodes: []
      };
    }
    marketGroups[marketName].presentItemCodes.push(row[5]); // كود الصنف
  });

  // 5. إدارة المجلد والملف في جوجل درايف
  let folder;
  const folders = DriveApp.getFoldersByName("تقارير النقوصات");
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder("تقارير النقوصات");
  }

  const fileName = "تقرير نقوصات [" + dateString + "]";
  const newReportFile = SpreadsheetApp.create(fileName);
  const fileId = newReportFile.getId();
  
  // نقل الملف للمجلد الصحيح
  const fileInDrive = DriveApp.getFileById(fileId);
  folder.addFile(fileInDrive);
  DriveApp.getRootFolder().removeFile(fileInDrive);

  // 6. إنشاء الأوراق داخل الملف وتعبئتها
  const summaryList = [];

  for (let marketName in marketGroups) {
    const group = marketGroups[marketName];
    const sheet = newReportFile.insertSheet(marketName);
    
    // الترويسة (الصف الأول)
    sheet.getRange("A1:E1").merge().setValue("تقرير نقوصات: " + marketName)
         .setBackground("#f1f5f9").setFontWeight("bold").setHorizontalAlignment("center");
    
    sheet.appendRow(["إسم الماركت", marketName, "", "التاريخ", dateString]);
    sheet.appendRow(["إسم المروّج", group.promoter, "", "", ""]);
    sheet.appendRow([""]); // سطر فارغ للجمالية
    
    // رؤوس الأعمدة
    const tableHeader = ["البراند", "كود الصنف", "إسم الصنف", "المخزون", "التواجد"];
    sheet.appendRow(tableHeader);
    const headerRange = sheet.getRange(sheet.getLastRow(), 1, 1, 5);
    headerRange.setBackground("#0f766e").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");

    // تعبئة جدول المواد بالكامل
    const tableData = [];
    let presentCount = 0;
    let missingCount = 0; // عداد الأصناف المفقودة

    allItems.forEach(item => {
      const isPresent = group.presentItemCodes.includes(item[0]);
      
      if (isPresent) {
        presentCount++;
      } else {
        // التحقق مما إذا كان هناك مخزون متوفر لهذا الصنف (العمود 4 في جدول الآيتمز)
        const stockValue = parseFloat(item[4]);
        if (!isNaN(stockValue) && stockValue > 0) {
          missingCount++; // زيادة العداد إذا كان غير متواجد بالماركت وله مخزون
        }
      }
      
      tableData.push([
        item[3], // Brand
        item[0], // Code
        item[2], // Name
        item[4], // Stock
        isPresent ? "✔️" : "" // التواجد (صح خضراء أو فارغ)
      ]);
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, tableData.length, 5).setValues(tableData);
    
    // تنسيق الجدول
    const fullTableRange = sheet.getRange(5, 1, tableData.length + 1, 5);
    fullTableRange.setBorder(true, true, true, true, true, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
    sheet.setColumnWidths(1, 5, 150);
    sheet.setColumnWidth(3, 350); // توسيع عمود الاسم
    sheet.getRange(6, 5, tableData.length, 1).setFontColor("green").setFontWeight("bold").setHorizontalAlignment("center");

    summaryList.push({ 
      name: marketName, 
      count: presentCount, 
      missing: missingCount, // إضافة العدد المفقود لملخص الإيميل
      province: group.province,
      promoter: group.promoter 
    });
  }

  newReportFile.deleteSheet(newReportFile.getSheetByName("Sheet1")); // حذف الورقة الافتراضية
  
  // 7. ترتيب القائمة للإيميل (من الأكثر تواجداً للأقل)
  summaryList.sort((a, b) => b.count - a.count);

  // 8. بناء محتوى الإيميل HTML
  const reportUrl = newReportFile.getUrl();
  const emailHtml = `
    <div dir="rtl" style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #1f2937; max-width: 760px; margin: 0 auto; background-color: #f8fafc; border: 1px solid #e5e7eb; border-radius: 18px; overflow: hidden;">
      
      <!-- رأس الإيميل -->
      <div style="background-color: #ffffff; padding: 24px 28px 20px; border-top: 5px solid #0f766e; border-bottom: 1px solid #e5e7eb;">
        <table style="width: 100%; border-collapse: collapse;">
          <tr>
            <td style="vertical-align: middle;">
              <h1 style="margin: 0 0 8px; font-size: 24px; line-height: 1.3; font-weight: 800; color: #0f172a;">تقرير النقوصات اليومي</h1>
              <p style="margin: 0; color: #64748b; font-size: 13px; font-weight: 600;">${fullDateAr}</p>
            </td>
            <td style="vertical-align: middle; text-align: left; width: 96px;">
              <img src="${companyLogoUrl}" width="82" style="display: block; margin-right: auto;" alt="Al-Farah Logo">
            </td>
          </tr>
        </table>
      </div>
      
      <!-- جسم الإيميل -->
      <div style="padding: 28px; line-height: 1.75; background-color: #ffffff;">
        <p style="margin: 0 0 10px; font-weight: 800; font-size: 16px; color: #111827;">السادة المحترمون،</p>
        <p style="margin: 0 0 22px; color: #475569; font-size: 14px;">تحية طيبة، مرفق أدناه ملخص تقرير النقوصات اليومي حسب المحافظة والماركت والمروّج.</p>
        
        <!-- جدول الملخص -->
        <table style="width: 100%; border-collapse: collapse; margin: 24px 0; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; background-color: #ffffff;">
          <thead>
            <tr style="background-color: #f8fafc;">
              <th style="padding: 12px 10px; text-align: right; border-bottom: 1px solid #e2e8f0; color: #475569; font-size: 12px; width: 14%;">المحافظة</th>
              <th style="padding: 12px 10px; text-align: right; border-bottom: 1px solid #e2e8f0; color: #475569; font-size: 12px; width: 28%;">إسم الماركت</th>
              <th style="padding: 12px 10px; text-align: center; border-bottom: 1px solid #e2e8f0; color: #0f766e; font-size: 12px; width: 16%;">المتوفر</th>
              <th style="padding: 12px 10px; text-align: center; border-bottom: 1px solid #e2e8f0; color: #dc2626; font-size: 12px; width: 16%;">المفقود</th>
              <th style="padding: 12px 10px; text-align: right; border-bottom: 1px solid #e2e8f0; color: #475569; font-size: 12px; width: 26%;">إسم المروّج</th>
            </tr>
          </thead>
          <tbody>
            ${summaryList.map((m, index) => `
              <tr style="background-color: ${index % 2 === 0 ? '#ffffff' : '#fbfdff'};">
                <td style="padding: 12px 10px; border-bottom: 1px solid #edf2f7; color: #0f766e; font-size: 13px; font-weight: 700;">${m.province || "-"}</td>
                <td style="padding: 12px 10px; border-bottom: 1px solid #edf2f7; color: #1f2937; font-size: 13px; font-weight: 700;">${m.name}</td>
                <td style="padding: 12px 10px; text-align: center; border-bottom: 1px solid #edf2f7; font-size: 13px; font-weight: 800; color: #0f766e;">${m.count}</td>
                <td style="padding: 12px 10px; text-align: center; border-bottom: 1px solid #edf2f7; font-size: 13px; font-weight: 800; color: #dc2626;">${m.missing}</td>
                <td style="padding: 12px 10px; border-bottom: 1px solid #edf2f7; color: #64748b; font-size: 13px; font-weight: 600;">${m.promoter}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <p style="color: #475569; font-size: 14px; margin: 0;">يمكنكم الاطلاع على التقرير الكامل من خلال الرابط التالي:</p>
        
        <!-- زر فتح الملف -->
        <div style="text-align: center; margin: 28px 0 8px;">
          <a href="${reportUrl}" style="background-color: #0f766e; color: white; padding: 13px 28px; text-decoration: none; border-radius: 10px; font-weight: 800; font-size: 14px; display: inline-block;">
            عرض تقرير النقوصات
          </a>
        </div>

        <p style="color: #64748b; margin: 24px 0 0; font-size: 13px;">مع فائق الود والتقدير،</p>
      </div>

      <!-- التوقيع الاحترافي -->
      <div style="background-color: #f8fafc; padding: 22px 28px; border-top: 1px solid #e2e8f0;">
        <table style="width: 100%; border-collapse: collapse;">
          <tr>
            <td style="vertical-align: top;">
              <p style="margin: 0 0 4px; font-weight: 800; color: #0f172a; font-size: 16px;">زين محمد</p>
              <p style="margin: 0 0 14px; font-size: 12px; color: #64748b; font-weight: 600;">مطور نظم برمجية</p>
              
              <table style="font-size: 12px; color: #475569; border-collapse: separate; border-spacing: 0 6px;">
                <tr>
                  <td style="width: 52px; color: #94a3b8; font-weight: 700;">هاتف</td>
                  <td><a href="tel:+9647767798799" style="color: #475569; text-decoration: none;">+964 776 779 8799</a></td>
                </tr>
                <tr>
                  <td style="width: 52px; color: #94a3b8; font-weight: 700;">بريد</td>
                  <td><a href="mailto:zmohammed@alfarah.com.iq" style="color: #475569; text-decoration: none;">zmohammed@alfarah.com.iq</a></td>
                </tr>
                <tr>
                  <td style="width: 52px; color: #94a3b8; font-weight: 700;">موقع</td>
                  <td><a href="http://www.alfarah.com.iq" style="color: #0369a1; text-decoration: none; font-weight: 600;">www.alfarah.com.iq</a></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </div>
      
    </div>
  `;

  // 9. إرسال الإيميل
  MailApp.sendEmail({
    to: recipients.join(","),
    subject: "تقرير النقوصات اليومي - " + dateString,
    htmlBody: emailHtml
  });
}
