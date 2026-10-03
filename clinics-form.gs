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

  // Always work in Egypt time so "today" and the weekday are the Egyptian ones.
  var timeZone = "Africa/Cairo";
  
  // Array for Arabic day names
  var arabicDays = ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"];
  
  // 2. Loop through the next 30 days and drop Fridays and Saturdays
  for (var i = 0; i < 30; i++) {
    var targetDate = new Date();
    targetDate.setDate(targetDate.getDate() + i); // Increment day by day
    
    // ISO weekday in Egypt: 1 = Monday ... 5 = Friday, 6 = Saturday, 7 = Sunday
    var isoDay = Number(Utilities.formatDate(targetDate, timeZone, "u"));
    var dayOfWeek = isoDay % 7; // 0 = Sunday ... matches arabicDays
    
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


// ==========================================================
// RUN THIS ONCE (manually) to skip the date question for the
// clinics "اسنان" and "عيادة عامة".
//
// Google Forms can't grey out a dropdown, so this uses section
// branching: the date question gets its own section, and picking
// "اسنان" or "عيادة عامة" in the clinic question submits the form
// right away. The form has exactly two sections: section 1 (clinic, email,
// image, ... asked of everyone) and section 2 (the date question only).
// Every other clinic goes through the date section as before.
// The clinic question must be a Multiple choice or Dropdown, and it
// must come BEFORE the date question (the script moves the date
// question after it if it isn't). Safe to run again: it reuses the
// sections it already created.
// ==========================================================
function setupClinicDateBranching() {
  var form = FormApp.getActiveForm();
  var NO_DATE_CLINICS = ["اسنان", "أسنان", "عيادة عامة"];
  var DATE_SECTION_TITLE = "اختيار التاريخ";
  var AFTER_DATE_SECTION_TITLE = "باقي البيانات";

  function norm(text) {
    return String(text).replace(/[أإآ]/g, "ا").replace(/\s+/g, " ").trim();
  }
  var noDateSet = NO_DATE_CLINICS.map(norm);
  // Match by keyword too, so choices like "عيادة الأسنان" or "العيادة العامة" are caught.
  function skipsDate(value) {
    var v = norm(value);
    // Exact names, plus anything containing "اسنان" (عيادة الأسنان) or exactly the
    // general clinic ("عيادة عامة" / "العيادة العامة"). A bare "عامة" substring would
    // also catch e.g. "جراحة عامة", which must still ask for a date.
    return noDateSet.indexOf(v) !== -1 ||
      v.indexOf("اسنان") !== -1 ||
      v === "العيادة العامة" || v === "عيادة العامة";
  }

  var clinicItem = null, dateItem = null;
  form.getItems().forEach(function (item) {
    var type = item.getType();
    var title = item.getTitle();
    if (!clinicItem && title.indexOf("العيادة") !== -1 &&
        (type === FormApp.ItemType.MULTIPLE_CHOICE || type === FormApp.ItemType.LIST)) {
      clinicItem = item;
    }
    if (!dateItem && title === "التاريخ" && type === FormApp.ItemType.LIST) {
      dateItem = item;
    }
  });
  if (!clinicItem || !dateItem) {
    throw new Error("Could not find the clinic question and/or the 'التاريخ' question.");
  }

  // The date question has to come after the clinic question.
  if (dateItem.getIndex() < clinicItem.getIndex()) {
    form.moveItem(dateItem.getIndex(), clinicItem.getIndex());
  }

  // Move item at its current position so it sits directly BEFORE the anchor.
  function placeBefore(item, anchor) {
    var from = item.getIndex();
    var to = anchor.getIndex();
    if (from < to) to = to - 1;   // account for the removal shifting the anchor up
    if (from !== to) form.moveItem(from, to);
  }

  // Move item so it sits directly AFTER the anchor.
  function placeAfter(item, anchor) {
    var from = item.getIndex();
    var to = anchor.getIndex();
    if (from > to) to = to + 1;
    if (from !== to) form.moveItem(from, to);
  }

  // Find or create the section that starts right before the date question.
  var datePage = null, afterPage = null;
  form.getItems(FormApp.ItemType.PAGE_BREAK).forEach(function (pb) {
    if (pb.getTitle() === DATE_SECTION_TITLE) datePage = pb.asPageBreakItem();
    if (pb.getTitle() === AFTER_DATE_SECTION_TITLE) afterPage = pb.asPageBreakItem();
  });

  if (!datePage) {
    datePage = form.addPageBreakItem().setTitle(DATE_SECTION_TITLE);
  }
  placeBefore(datePage, dateItem); // sits directly before the date question

  // Email + image must be asked of EVERYONE (including اسنان / عيادة عامة,
  // who skip the date section), so move them into section 1, right before
  // the date section starts. Section 2 is then the date question only.
  form.getItems().filter(function (it) {
    var type = it.getType();
    var title = String(it.getTitle()).toLowerCase();
    var isImage = type === FormApp.ItemType.FILE_UPLOAD ||
      title.indexOf("صورة") !== -1 || title.indexOf("image") !== -1 || title.indexOf("photo") !== -1;
    var isEmail = (type === FormApp.ItemType.TEXT) &&
      (title.indexOf("بريد") !== -1 || title.indexOf("email") !== -1 || title.indexOf("ايميل") !== -1 || title.indexOf("إيميل") !== -1);
    return (isImage || isEmail) && it.getIndex() > datePage.getIndex();
  }).forEach(function (it) {
    placeBefore(it, datePage);
    Logger.log("Moved to section 1: " + it.getTitle());
  });

  // Only TWO sections: section 1 (clinic, email, image, ...) and section 2
  // (the date question).
  // If an older run created the third "باقي البيانات" section, remove it;
  // its questions simply stay in section 2.
  if (afterPage) {
    form.deleteItem(afterPage);
    afterPage = null;
  }

  // Route the clinic answers.
  var choices = (clinicItem.getType() === FormApp.ItemType.LIST
    ? clinicItem.asListItem() : clinicItem.asMultipleChoiceItem()).getChoices();
  var newChoices = choices.map(function (choice) {
    var value = choice.getValue();
    var skipDate = skipsDate(value);
    Logger.log((skipDate ? "SKIPS date  : " : "asks for date: ") + value);
    var target = skipDate
      ? FormApp.PageNavigationType.SUBMIT   // skip the date section
      : datePage;
    return (clinicItem.getType() === FormApp.ItemType.LIST
      ? clinicItem.asListItem() : clinicItem.asMultipleChoiceItem()).createChoice(value, target);
  });
  (clinicItem.getType() === FormApp.ItemType.LIST
    ? clinicItem.asListItem() : clinicItem.asMultipleChoiceItem()).setChoices(newChoices);

  Logger.log("Done: اسنان / عيادة عامة now skip the date question.");
}
