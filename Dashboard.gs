/*************************************************
 * S Video ERP
 * Module : Dashboard
 * File   : Dashboard.gs
 *************************************************/

function getDashboardInitData() {
  return {
    studios: getStudios() 
  };
}

function getStudioFinancials(studioId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Booking Master");
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  
  let jobCount = 0;
  let totalAmount = 0;
  let totalAdvance = 0;

  // Start at row 1 to skip headers
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentStudioId = String(row[3]).trim(); // Column D is Studio ID
    
    if (currentStudioId === String(studioId).trim()) {
      jobCount++;
      
      // Clean commas just in case manual edits were typed as text
      const advanceClean = String(row[18] || "").replace(/,/g, '');
      const amountClean = String(row[19] || "").replace(/,/g, '');
      
      totalAdvance += Number(advanceClean) || 0; 
      totalAmount += Number(amountClean) || 0;  
    }
  }

  // BULLETPROOF MATH: Calculate Balance dynamically instead of trusting the sheet column
  const totalBalance = totalAmount - totalAdvance;

  return {
    jobCount: jobCount,
    totalAmount: totalAmount,
    totalAdvance: totalAdvance,
    totalBalance: totalBalance
  };
}