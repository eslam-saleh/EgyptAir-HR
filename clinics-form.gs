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
  
  // 2. Loop through the next 31 days and drop Fridays and Saturdays
  for (var i = 0; i < 31; i++) {
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
//
// ALIGNMENT: choice labels with a note, e.g. "مسالك (ايام الخميس)", are
// rebuilt as  name + " ـــــ " + "(note)"  (a stretch line "ـ", tatweel),
// with the number of tatweels chosen so every "(" lines up. Google Forms
// collapses every kind of space, but it keeps tatweel. Labels without a
// "(note)" (e.g. "صدر") are left alone. The note and the tatweels are
// ignored when deciding whether a clinic skips the date; the dashboard
// script trims them from the value saved in the sheet.
//
// To align by hand instead: set AUTO_ALIGN = false and add/delete "ـ"
// in the option text in the form. The script then keeps your labels as
// they are and only sets up the routing.
// ==========================================================
function setupClinicDateBranching() {
  var form = FormApp.getActiveForm();
  var NO_DATE_CLINICS = ["اسنان", "أسنان", "عيادة عامة"];
  var DATE_SECTION_TITLE = "اختيار التاريخ";
  var AFTER_DATE_SECTION_TITLE = "باقي البيانات";

  // ── Alignment settings (tune if the "(" don't line up on your screen) ──
  var AUTO_ALIGN = true;      // true = rebuild the tatweel padding on every run (overwrites hand edits)
                              // false = keep the labels exactly as they are in the form
  var TATWEEL = "\u0640";     // "ـ"
  var PAD_FACTOR = 1;         // tatweels added per character of difference in name length
  var MIN_PAD = 3;            // tatweels on the longest name

  function norm(text) {
    return String(text).replace(/[أإآ]/g, "ا").replace(/\s+/g, " ").trim();
  }
  var noDateSet = NO_DATE_CLINICS.map(norm);

  // Clinic name only: no "(note)", no padding characters (also cleans old
  // labels padded with other kinds of spaces by earlier versions).
  function cleanName(value) {
    return String(value)
      .split(/[(（]/)[0]
      .replace(/[\u200B\u2000-\u200A\u2003\u00A0\u3000\u0640]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  // Takes the current labels, returns labels with every "(" aligned.
  // Labels without a "(note)" are returned untouched.
  function alignLabels(values) {
    var parts = values.map(function (v) {
      var s = String(v);
      var i = s.search(/[(（]/);
      return i === -1
        ? { name: s, note: "" }
        : { name: cleanName(s), note: s.slice(i).trim() };
    });
    var maxLen = 0;
    parts.forEach(function (p) {
      if (p.note && p.name.length > maxLen) maxLen = p.name.length;
    });
    return parts.map(function (p) {
      if (!p.note) return p.name;
      var count = Math.max(1, Math.round((maxLen - p.name.length) * PAD_FACTOR) + MIN_PAD);
      return p.name + " " + new Array(count + 1).join(TATWEEL) + " " + p.note;
    });
  }

  // Match by keyword too, so choices like "عيادة الأسنان" or "العيادة العامة" are caught.
  // Anything from the first "(" onward is ignored ("عيادة عامة (ايام ...)").
  function skipsDate(value) {
    var v = norm(cleanName(value));
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

  // Route the clinic answers, with every "(" aligned.
  var isList = clinicItem.getType() === FormApp.ItemType.LIST;
  var clinicQ = isList ? clinicItem.asListItem() : clinicItem.asMultipleChoiceItem();

  var currentLabels = clinicQ.getChoices().map(function (c) { return c.getValue(); });
  var labels = AUTO_ALIGN ? alignLabels(currentLabels) : currentLabels;

  var newChoices = labels.map(function (label) {
    var skipDate = skipsDate(label);
    Logger.log((skipDate ? "SKIPS date  : " : "asks for date: ") + label);
    var target = skipDate
      ? FormApp.PageNavigationType.SUBMIT   // skip the date section
      : datePage;
    return clinicQ.createChoice(label, target);
  });
  clinicQ.setChoices(newChoices);

  Logger.log("Done: اسنان / عيادة عامة skip the date question" +
    (AUTO_ALIGN ? ", and the \"(\" are aligned." : "."));
}


// ==========================================================
// DIAGNOSTIC: shows what is stored in each clinic choice right now.
// "640" = tatweel (kept by Forms), "20" = normal space.
// ==========================================================
function dumpClinicLabels() {
  var form = FormApp.getActiveForm();
  form.getItems().forEach(function (it) {
    var t = it.getType();
    if (it.getTitle().indexOf("العيادة") === -1) return;
    if (t !== FormApp.ItemType.MULTIPLE_CHOICE && t !== FormApp.ItemType.LIST) return;
    var q = t === FormApp.ItemType.LIST ? it.asListItem() : it.asMultipleChoiceItem();
    q.getChoices().forEach(function (c) {
      var v = c.getValue();
      var head = v.split(/[(（]/)[0];
      Logger.log(v + "   | before ( : " + head.slice(-12).split("").map(function (ch) {
        return ch.charCodeAt(0).toString(16);
      }).join(" "));
    });
  });
}
