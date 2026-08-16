function updateNext30Days() {
  var form = FormApp.getActiveForm();
  
  // Find the "التاريخ" question
  var items = form.getItems(FormApp.ItemType.LIST);
  var dateItem;
  
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === "التاريخ") {
      dateItem = items[i].asListItem();
      break;
    }
  }
  
  if (!dateItem) {
    Logger.log("Question 'التاريخ' not found.");
    return;
  }

  var dateChoices = [];
  
  // 1. Add the first consistent option
  dateChoices.push("اقرب وقت");

  var timeZone = Session.getScriptTimeZone();
  
  // Array for Arabic day names
  var arabicDays = ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"];
  
  // 2. Loop through the next 30 days and drop Fridays and Saturdays
  for (var i = 0; i < 30; i++) {
    var targetDate = new Date();
    targetDate.setDate(targetDate.getDate() + i); // Increment day by day
    
    var dayOfWeek = targetDate.getDay(); // 0 = Sunday, 1 = Monday, ..., 5 = Friday, 6 = Saturday
    
    // Skip Fridays (5) and Saturdays (6)
    if (dayOfWeek === 5 || dayOfWeek === 6) {
      continue;
    }
    
    // Format date as standard "YYYY-MM-DD"
    var formattedDate = Utilities.formatDate(targetDate, timeZone, "yyyy-MM-dd");
    var dayNameArabic = arabicDays[dayOfWeek];
    
    // Combine date on the left and Arabic day name on the right (e.g., 2026-08-16 - الأحد)
    var combinedOption = formattedDate + " - " + dayNameArabic;
    dateChoices.push(combinedOption);
  }
  
  // Update the dropdown choice options
  dateItem.setChoiceValues(dateChoices);
}