// Google Apps Script to scrape Coursera course data and populate Google Sheets

function scrapeCourseraData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COURSERA');
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  // Find the Link column index
  const headers = data[0];
  const linkColIndex = headers.indexOf('Link');
  
  if (linkColIndex === -1) {
    SpreadsheetApp.getUi().alert('Link column not found!');
    return;
  }
  
  for (let i = 1; i < data.length; i++) {
    const courseLink = data[i][linkColIndex];
    if (courseLink && !data[i][headers.indexOf('Course Title')]) { // Process only if Course Title is empty
      const courseInfo = getCourseraCourseInfo(courseLink);
      
      if (courseInfo) {
        sheet.getRange(i + 1, headers.indexOf('Course Title') + 1).setValue(courseInfo.title);
        sheet.getRange(i + 1, headers.indexOf('Institution') + 1).setValue(courseInfo.institution);
        sheet.getRange(i + 1, headers.indexOf('Recommended Experience') + 1).setValue(courseInfo.experience);
        sheet.getRange(i + 1, headers.indexOf('Approx. Hours') + 1).setValue(courseInfo.hours);
        sheet.getRange(i + 1, headers.indexOf('No. of Modules') + 1).setValue(courseInfo.modules);
        sheet.getRange(i + 1, headers.indexOf('Enrollees') + 1).setValue(courseInfo.enrollees);
        sheet.getRange(i + 1, headers.indexOf('Course Rating') + 1).setValue(courseInfo.rating);
        sheet.getRange(i + 1, headers.indexOf('No. of Reviews') + 1).setValue(courseInfo.reviews);
      }
    }
  }

  setDropdowns();
  setCourseType();
  addNewRow();
}

function getCourseraCourseInfo(url) {
  try {
    const response = UrlFetchApp.fetch(url);
    const content = response.getContentText();

    // Extracting course details with multiple regex attempts for hours and rating
    const titleMatch = content.match(/<h1[^>]*data-e2e="hero-title"[^>]*>(.*?)<\/h1>/);
    const institutionMatch = content.match(/<div class="css-1ujzbfc">.*?<img[^>]*alt="(.*?)"[^>]*><\/div>/);
    const experienceMatch = content.match(/<div class="css-fk6qfz">(Beginner level|Intermediate level|Advanced level)<\/div>/);

    // Multiple attempts for Approx. Hours to capture different formats
    const hoursMatch1 = content.match(/<div class="css-fk6qfz">(\d+(\.\d+)?)\s*hours?<\/div>/); // Decimal or whole hours (e.g., "1.5 hours", "1 hour")
    const hoursMatch2 = content.match(/<div class="css-fk6qfz">Approx\.\s*(\d+(\.\d+)?)\s*hours?<\/div>/); // "Approx. hours" format
    const hoursMatch3 = content.match(/<div class="css-fw9ih3"><div>Approx\.\s*(\d+(\.\d+)?)\s*hours?<\/div>/); // "Flexible schedule" format (simplified)
    const hoursMatch4 = content.match(/<div class="css-fk6qfz">(\d+)\s*hour?s?\s*<\/div>/); // "X hour(s)" format
    const hoursMatch5 = content.match(/<div class="css-fk6qfz">(\d+)\s*hours?\s*to\s*complete<\/div>/); // "X hours to complete" format
    const minutesMatch = content.match(/<div class="css-fk6qfz">(\d+)\s*minutes?<\/div>/); // Minutes format
    const hoursAndMinutesMatch = content.match(/<div class="css-fk6qfz">(\d+)\s*hours?\s*(\d+)\s*mins?<\/div>/); // "X hours Y mins" format

    // Prioritize the matches
    const hoursMatch = hoursMatch1 || hoursMatch2 || hoursMatch3 || hoursMatch4 || hoursMatch5 || minutesMatch || hoursAndMinutesMatch;

    // Convert minutes to hours if minutes are found
    let hours = 'N/A';
    if (hoursAndMinutesMatch) {
      const hoursPart = parseInt(hoursAndMinutesMatch[1]);
      const minutesPart = parseInt(hoursAndMinutesMatch[2]);
      hours = (hoursPart + minutesPart / 60).toFixed(2);  // Combine hours and convert minutes to decimal
    } else if (minutesMatch) {
      const minutes = parseInt(minutesMatch[1]);
      hours = (minutes / 60).toFixed(2);  // Convert minutes to hours (to 2 decimal places)
    } else if (hoursMatch) {
      hours = hoursMatch[1];
    }

    // Multiple attempts for Course Rating to capture different formats (aria-label or content)
    const ratingMatch1 = content.match(/aria-label="(\d+\.\d+)\s*stars"/); // Rating in aria-label
    const ratingMatch2 = content.match(/aria-label="\d+\.\d+" role="group">(\d+\.\d+)<\/div>/); // Rating in div content
    const ratingMatch3 = content.match(/role="group">(\d+\.\d+)<\/div>/); // Another rating div pattern
    const ratingMatch = ratingMatch1 || ratingMatch2 || ratingMatch3;  // Prioritize matches

    // const modulesMatch = content.match(/<a[^>]*href="#modules"[^>]*>(\d+ module[s]?|Guided Project)<\/a>/);
    const modulesMatch = content.match(/<a[^>]*href="#modules"[^>]*>(\d+)\s*modules?|Guided Project<\/a>/);
    const enrolleesMatch = content.match(/<span><strong><span>([\d,]+)<\/span><\/strong> already enrolled<\/span>/);
    const reviewsMatch = content.match(/<p class="css-vac8rf">\((\d+,?\d*) reviews\)<\/p>/);

    return {
      title: titleMatch ? titleMatch[1] : 'N/A',
      institution: institutionMatch ? institutionMatch[1] : 'N/A',
      experience: experienceMatch ? experienceMatch[1] : 'N/A',
      hours: hours,
      modules: modulesMatch ? modulesMatch[1] : 0,
      enrollees: enrolleesMatch ? enrolleesMatch[1].replace(/,/g, '') : 'N/A',
      rating: ratingMatch ? ratingMatch[1] : 'N/A', // Use the first match (ratingMatch1, ratingMatch2, or ratingMatch3)
      reviews: reviewsMatch ? reviewsMatch[1].replace(/,/g, '') : 'N/A'
    };

  } catch (e) {
    Logger.log('Error fetching data: ' + e);
    return null;
  }
}

function setDropdowns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COURSERA');

  // Get the headers row (assumed to be row 1)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Get the current active row
  const currentRow = sheet.getLastRow();

  // Dynamically calculate the column index for "State" and "Added to LinkedIn"
  const stateColumn = headers.indexOf('State') + 1;
  const linkedinColumn = headers.indexOf('Added to LinkedIn') + 1;

  // Set dropdown for "State"
  const stateRange = sheet.getRange(currentRow, stateColumn);
  const stateValues = ["STARTED", "NOT STARTED"];
  const stateValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(stateValues)
    .setAllowInvalid(false)
    .setHelpText('Select either STARTED or NOT STARTED')
    .build();
  stateRange.setDataValidation(stateValidation);
  stateRange.setValue("NOT STARTED");  // Set default value
  stateRange.setHorizontalAlignment('center');

  // Set dropdown for "Added to LinkedIn"
  const linkedinRange = sheet.getRange(currentRow, linkedinColumn);
  const linkedinValues = ["YES", "NO"];
  const linkedinValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(linkedinValues)
    .setAllowInvalid(false)
    .setHelpText('Select YES or NO')
    .build();
  linkedinRange.setDataValidation(linkedinValidation);
  linkedinRange.setValue("NO");  // Set default value
  linkedinRange.setHorizontalAlignment('center');

  // Apply conditional formatting for "State"
  const stateColors = {
    "STARTED": "#bfe1f6",      // Light blue
    "NOT STARTED": "#e6e6e6"   // Grey
  };
  applyConditionalFormatting(sheet, stateRange, stateColors);

  // Apply conditional formatting for "Added to LinkedIn"
  const linkedinColors = {
    "YES": "#bfe1f6",  // Light blue
    "NO": "#e6e6e6"    // Grey
  };
  applyConditionalFormatting(sheet, linkedinRange, linkedinColors);
}

// Helper function to apply conditional formatting for specific values
function applyConditionalFormatting(sheet, range, colorMap) {
  const existingRules = sheet.getConditionalFormatRules();
  const newRules = [];

  // Add conditional formatting for each value in the color map
  Object.keys(colorMap).forEach(value => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(value)
      .setBackground(colorMap[value])  // Set background color
      .setFontColor("#000000")         // Ensure text remains visible
      .setRanges([range])
      .build();
    newRules.push(rule);
  });

  // Preserve existing rules and add new ones for the current row
  sheet.setConditionalFormatRules(existingRules.concat(newRules));
}

function setCourseType() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COURSERA');

  // Get the headers row (assumed to be row 1)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Get the current active row
  const currentRow = sheet.getLastRow();

  // Dynamically find the column index for "Course Type" and the column for H (assumed "Hours" or similar)
  const courseTypeColumn = headers.indexOf('Course Type') + 1;
  const hoursColumn = headers.indexOf('No. of Modules') + 1;

  // Get the corresponding column letter for the "Hours" column
  const hoursColumnLetter = String.fromCharCode(64 + hoursColumn);  // 64 + 8 = 72 -> 'H'

  // Construct the formula dynamically
  const formula = `=IF(${hoursColumnLetter}${currentRow}<>0,"MODULAR","PROJECT/GUIDED PROJECT")`;

  // Apply the formula to the "Course Type" cell
  sheet.getRange(currentRow, courseTypeColumn).setFormula(formula);
}

function addNewRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COURSERA");
  var lastRow = sheet.getLastRow(); // Get the last row with data
  sheet.insertRowAfter(lastRow); // Insert a new row after the last populated row
}
