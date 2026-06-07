// ==========================================
// الإعدادات الأولية للاتصال بجوجل شيت
// ==========================================
 
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('نظام إدارة المندوبين')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// دوال جلب البيانات (Read)
// ==========================================

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getInitialData() {
  try {
    const ss = getSpreadsheet();
    if(!ss) throw new Error("لا يمكن الوصول للملف، تأكد من أن السكربت مرتبط بجدول البيانات.");
    
    const employees = [];
    const empSheet = ss.getSheetByName('Employees');
    if (!empSheet) throw new Error("ورقة (Employees) غير موجودة في الجدول.");
    
    const empData = empSheet.getDataRange().getDisplayValues();
    for(let i=1; i<empData.length; i++) {
      if(empData[i][0] && String(empData[i][0]).trim() !== '') {
        employees.push({
          name: String(empData[i][0]).trim(), province: String(empData[i][1]).trim(),
          role: String(empData[i][2]).trim(), side: String(empData[i][3]).trim(), password: String(empData[i][4]).trim()
        });
      }
    }

    const markets = [];
    const marketSheet = ss.getSheetByName('Markets');
    if (!marketSheet) throw new Error("ورقة (Markets) غير موجودة في الجدول.");
    
    const marketData = marketSheet.getDataRange().getDisplayValues();
    for(let i=1; i<marketData.length; i++) {
      if(marketData[i][0] && String(marketData[i][0]).trim() !== '') {
        markets.push({
          code: String(marketData[i][0]).trim(), name: String(marketData[i][1]).trim(),
          province: String(marketData[i][2]).trim(),
          side: String(marketData[i][3]).trim(),
          longitude: String(marketData[i][5]).trim(), latitude: String(marketData[i][4]).trim()
        });
      }
    }

    const items = [];
    const itemsSheet = ss.getSheetByName('Items');
    if (!itemsSheet) throw new Error("ورقة (Items) غير موجودة في الجدول.");
    
    const itemsData = itemsSheet.getDataRange().getDisplayValues();
    for(let i=1; i<itemsData.length; i++) {
      if(itemsData[i][0] && String(itemsData[i][0]).trim() !== '') {
        items.push({
          code: String(itemsData[i][0]).trim(), image: String(itemsData[i][1]).trim(),
          name: String(itemsData[i][2]).trim(), brand: String(itemsData[i][3]).trim()
        });
      }
    }

    return { employees: employees, markets: markets, items: items };
  } catch (error) {
    throw new Error(error.message); 
  }
}

function loginUser(username, password) {
  try {
    const data = getInitialData();
    const user = data.employees.find(e => e.name === username && String(e.password) === String(password));
    if(user) return { success: true, user: user };
    return { success: false, message: 'كلمة المرور غير صحيحة' };
  } catch(err) {
    throw new Error(err.message);
  }
}

function getEmployeeDashboardData(employeeName) {
  try {
    const ss = getSpreadsheet();
    const visitsSheet = ss.getSheetByName('Visits');
    if(!visitsSheet) return { visits: [] };

    const visitsData = visitsSheet.getDataRange().getDisplayValues();
    let myVisits = [];
    for(let i = visitsData.length - 1; i > 0; i--) {
      if(visitsData[i][0] === employeeName) {
        myVisits.push({
          marketName: String(visitsData[i][3]),
          date: String(visitsData[i][4]),
          reason: String(visitsData[i][7] || 'غير محدد'),
          notes: String(visitsData[i][8] || '')
        });
        if(myVisits.length >= 7) break;
      }
    }
    return { visits: myVisits };
  } catch(err) {
    throw new Error("خطأ أثناء جلب زيارات الموظف: " + err.message);
  }
}

// ==========================================
// دوال إدخال البيانات (Write)
// ==========================================

function saveVisitRecord(data) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Visits');
  sheet.appendRow([
    data.employeeName, data.province, data.marketCode, data.marketName,
    data.date, data.lng, data.lat, data.reason, data.notes
  ]);
  return true;
}

function saveStockData(type, commonData, itemsList) {
  const ss = getSpreadsheet();
  const sheetName = type === 'daily' ? 'Daily Stock' : 'Monthly Stock';
  const sheet = ss.getSheetByName(sheetName);
  
  itemsList.forEach(item => {
    let row = [
      commonData.employeeName, commonData.province, commonData.marketCode,
      commonData.marketName, commonData.date, item.code, item.name, item.brand
    ];
    if(type === 'monthly') {
      row.push(item.quantity);
      row.push(item.expiryDate); 
    }
    sheet.appendRow(row);
  });
  return true;
}

function saveCompetitorReport(data) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('Competitors'); 
  
  if (!sheet) {
    sheet = ss.insertSheet('Competitors');
    sheet.appendRow([
      'Employee Name', 'Market Code', 'Market Name', 'Date', 'Product Name', 
      'Price', 'Weight', 'Country', 'Quantity', 'Image Sample'
    ]);
  }

  sheet.appendRow([
    data.employeeName, data.marketCode, data.marketName, data.date, data.productName, 
    data.price, data.weight, data.origin, data.cartonQty, data.imageBase64
  ]);
  return true;
}

// ==========================================
// دوال الإدارة (Admin)
// ==========================================

function getAdminDashboardData() {
  try {
    const ss = getSpreadsheet();
    
    // جلب الزيارات
    const visitsSheet = ss.getSheetByName('Visits');
    let latestVisits = [];
    if(visitsSheet) {
      const vData = visitsSheet.getDataRange().getDisplayValues();
      for(let i = vData.length - 1; i > Math.max(0, vData.length - 100); i--) {
        if(!vData[i][0]) continue;
        latestVisits.push({
          employeeName: String(vData[i][0]), 
          province: String(vData[i][1]), // تم إضافته ليُستخدم في الفلاتر
          marketName: String(vData[i][3]),
          date: String(vData[i][4]), 
          reason: String(vData[i][7] || 'غير محدد'), 
          notes: String(vData[i][8] || '')
        });
      }
    }

    // جلب الجرد اليومي
    const dailySheet = ss.getSheetByName('Daily Stock');
    let latestDaily = [];
    let addedDailyMarkets = new Set();
    if(dailySheet) {
      const dData = dailySheet.getDataRange().getDisplayValues();
      for(let i = dData.length - 1; i > 0; i--) {
        if(!dData[i][0]) continue;
        let marketDateKey = dData[i][0] + '_' + dData[i][3] + '_' + dData[i][4];
        if(!addedDailyMarkets.has(marketDateKey)) {
          latestDaily.push({
            employeeName: String(dData[i][0]), province: String(dData[i][1]), marketCode: String(dData[i][2]), marketName: String(dData[i][3]),
            date: String(dData[i][4]),
            items: getMarketStockItems(dData, String(dData[i][0]), String(dData[i][3]), String(dData[i][4]), false)
          });
          addedDailyMarkets.add(marketDateKey);
        }
      }
    }

    // جلب الجرد الشهري
    const monthlySheet = ss.getSheetByName('Monthly Stock');
    let latestMonthly = [];
    let addedMonthlyMarkets = new Set();
    if(monthlySheet) {
      const mData = monthlySheet.getDataRange().getDisplayValues();
      for(let i = mData.length - 1; i > 0; i--) {
        if(!mData[i][0]) continue;
        let marketDateKey = mData[i][0] + '_' + mData[i][3] + '_' + mData[i][4];
        if(!addedMonthlyMarkets.has(marketDateKey)) {
          latestMonthly.push({
            employeeName: String(mData[i][0]), province: String(mData[i][1]), marketCode: String(mData[i][2]), marketName: String(mData[i][3]),
            date: String(mData[i][4]),
            items: getMarketStockItems(mData, String(mData[i][0]), String(mData[i][3]), String(mData[i][4]), true)
          });
          addedMonthlyMarkets.add(marketDateKey);
        }
      }
    }

    return { visits: latestVisits, daily: latestDaily, monthly: latestMonthly };
  } catch(err) {
    throw new Error("خطأ في جلب بيانات الإدارة: " + err.message);
  }
}

function getMarketStockItems(dataArr, empName, marketName, dateValue, isMonthly) {
  let items = [];
  const dateStr = String(dateValue);
  for(let i = 1; i < dataArr.length; i++) {
    if(String(dataArr[i][0]) === String(empName) && String(dataArr[i][3]) === String(marketName) && String(dataArr[i][4]) === dateStr) {
      let itemObj = { rowIndex: i + 1, itemCode: String(dataArr[i][5]), itemName: String(dataArr[i][6]), brand: String(dataArr[i][7]) };
      if(isMonthly) {
        itemObj.quantity = String(dataArr[i][8]);
        itemObj.expiryDate = dataArr[i][9] ? String(dataArr[i][9]) : '';
      }
      items.push(itemObj);
    }
  }
  return items;
}

function deleteRecord(sheetName, rowIndex) {
  getSpreadsheet().getSheetByName(sheetName).deleteRow(rowIndex);
  return true;
}

function updateStockItem(rowIndex, newQuantity, newExpiry) {
  const sheet = getSpreadsheet().getSheetByName('Monthly Stock');
  sheet.getRange(rowIndex, 9).setValue(newQuantity); 
  if(newExpiry) {
    sheet.getRange(rowIndex, 10).setValue(newExpiry);
  }
  return true;
}

function adminAddSingleStockItem(data) {
  const sheetName = data.type === 'daily' ? 'Daily Stock' : 'Monthly Stock';
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  let row = [ data.employeeName, data.province, data.marketCode, data.marketName, data.date, data.itemCode, data.itemName, data.brand ];
  if(data.type === 'monthly') {
    row.push(data.quantity);
    row.push(data.expiryDate || '');
  }
  sheet.appendRow(row);
  return true;
}
