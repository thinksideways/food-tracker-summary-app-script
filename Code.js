function generateDailySummary(e) {
  if (e && e.source.getActiveSheet().getName() !== "Daily Meals") {
    return;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetId = scriptProperties.getProperty('SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    Logger.log("Error: SPREADSHEET_ID script property is missing.");
    return;
  }
  
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const mealsSheet = ss.getSheetByName("Daily Meals");
  const summarySheet = ss.getSheetByName("Daily Summary");
  
  if (!mealsSheet || !summarySheet) {
    SpreadsheetApp.getUi().alert("Error: Make sure your sheets are named exactly 'Daily Meals' and 'Daily Summary'.");
    return;
  }
  
  // 1. Get all raw meal data (assuming Row 1 has headers: Date | Meal Description | Calories | Protein | Carbs | Fat | Fiber)
  const mealsData = mealsSheet.getDataRange().getValues();
  if (mealsData.length <= 1) return; // No data to process
  
  const summaryMap = {};
  
  // 2. Aggregate data in memory
  for (let i = 1; i < mealsData.length; i++) {
    const row = mealsData[i];
    let dateVal = row[0];
    
    if (!dateVal) continue; // Skip empty rows
    
    // Format date as a string standard key (e.g., YYYY-MM-DD or MM/DD/YYYY based on cell display)
    if (dateVal instanceof Date) {
      dateVal = Utilities.formatDate(dateVal, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
    } else {
      dateVal = dateVal.toString();
    }
    
    const description = row[1] || "";
    const calories = Number(row[2]) || 0;
    const protein = Number(row[3]) || 0;
    const carbs = Number(row[4]) || 0;
    const fat = Number(row[5]) || 0;
    const fiber = Number(row[6]) || 0;
    
    // If it's a new date, initialize the object
    if (!summaryMap[dateVal]) {
      summaryMap[dateVal] = {
        meals: [],
        calories: 0,
        protein: 0,
        carbs: 0,
        fat: 0,
        fiber: 0
      };
    }
    
    // Bundle individual meal details into the daily array
    summaryMap[dateVal].meals.push({
      item: description,
      cals: calories,
      p: protein,
      c: carbs,
      f: fat,
      fib: fiber
    });
    
    // Sum the totals
    summaryMap[dateVal].calories += calories;
    summaryMap[dateVal].protein += protein;
    summaryMap[dateVal].carbs += carbs;
    summaryMap[dateVal].fat += fat;
    summaryMap[dateVal].fiber += fiber;
  }
  
  // 3. Prepare the rows to write back to the Summary Sheet
  const outputRows = [];
  const sortedDates = Object.keys(summaryMap).sort((a, b) => new Date(a) - new Date(b));
  
  for (const date of sortedDates) {
    const data = summaryMap[date];
    outputRows.push([
      date,
      JSON.stringify(data.meals), // Valid JSON string bundle
      data.calories,
      data.protein,
      data.carbs,
      data.fat,
      data.fiber
    ]);
  }
  
  // 4. Safe clear and batch write
  if (summarySheet.getLastRow() > 1) {
    summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, summarySheet.getLastColumn()).clearContent();
  }
  
  if (outputRows.length > 0) {
    summarySheet.getRange(2, 1, outputRows.length, 7).setValues(outputRows);
  }
}