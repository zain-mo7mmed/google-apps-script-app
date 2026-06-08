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
        <p style="color: #374151;">إليكم التقرير اليومي لنقوصات الماركتات ليوم أمس، فقد تم رصد الأصناف التابعة لشركة الفرح العالمية المتواجدة في الأسواق، وكان نصيب كل ماركت من التوزيع كالآتي:</p>
        
        <!-- جدول الملخص -->
        <table style="width: 100%; border-collapse: separate; border-spacing: 0; margin: 25px 0; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
          <thead>
            <tr style="background-color: #f8fafc;">
              <th style="padding: 12px 10px; text-align: right; border-bottom: 2px solid #e2e8f0; color: #475569; font-size: 13px; width: 15%;">المحافظة</th>
              <th style="padding: 12px 10px; text-align: right; border-bottom: 2px solid #e2e8f0; color: #475569; font-size: 13px; width: 30%;">إسم الماركت</th>
              <th style="padding: 12px 10px; text-align: center; border-bottom: 2px solid #e2e8f0; color: #0f766e; font-size: 13px; width: 15%;">الأصناف المتوفرة</th>
              <th style="padding: 12px 10px; text-align: center; border-bottom: 2px solid #e2e8f0; color: #ef4444; font-size: 13px; width: 15%;">الأصناف المفقودة</th>
              <th style="padding: 12px 10px; text-align: right; border-bottom: 2px solid #e2e8f0; color: #475569; font-size: 13px; width: 25%;">إسم المروّج</th>
            </tr>
          </thead>
          <tbody>
            ${summaryList.map((m, index) => `
              <tr style="background-color: ${index % 2 === 0 ? '#ffffff' : '#f8fafc'};">
                <td style="padding: 12px 10px; border-bottom: 1px solid #e2e8f0; color: #0f766e; font-weight: 600;">${m.province || "-"}</td>
                <td style="padding: 12px 10px; border-bottom: 1px solid #e2e8f0; color: #1f2937; font-weight: 600;">${m.name}</td>
                <td style="padding: 12px 10px; text-align: center; border-bottom: 1px solid #e2e8f0; font-weight: bold; color: #0f766e;">${m.count}</td>
                <td style="padding: 12px 10px; text-align: center; border-bottom: 1px solid #e2e8f0; font-weight: bold; color: #ef4444;">${m.missing}</td>
                <td style="padding: 12px 10px; border-bottom: 1px solid #e2e8f0; color: #64748b;">${m.promoter}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <p style="color: #374151;">يمكنكم الإطلاع على التقرير الكامل عبر الرابط أدناه:</p>
        
        <!-- زر فتح الملف -->
        <div style="text-align: center; margin: 35px 0;">
          <a href="${reportUrl}" style="background-color: #16a34a; color: white; padding: 14px 35px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 15px; display: inline-flex; align-items: center; gap: 10px; box-shadow: 0 4px 6px rgba(22, 163, 74, 0.25);">
            <span style="font-size: 18px;">📊</span> عرض تقرير النقوصات
          </a>
        </div>

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

  // 9. إرسال الإيميل
  MailApp.sendEmail({
    to: recipients.join(","),
    subject: "تقرير النقوصات اليومي - " + dateString,
    htmlBody: emailHtml
  });
}
