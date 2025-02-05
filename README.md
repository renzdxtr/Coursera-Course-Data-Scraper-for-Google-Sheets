# Coursera Course Data Scraper for Google Sheets

## Introduction
This Google Apps Script automates the extraction of Coursera course data and populates it into a Google Sheet. It simplifies the process of gathering course details like title, institution, duration, modules, ratings, and more by scraping data directly from provided Coursera course links.

## Problem Statement
Manually gathering information for multiple Coursera courses is time-consuming and error-prone. This script automates the process, ensuring consistent and accurate data collection.

## Who Will Benefit?
- **Students** looking to track their progress on various Coursera courses.
- **Educators** who want to compile course data for recommendations.
- **Professionals** managing personal learning paths or team training.
- **Data Analysts** who need structured course data for insights.

## Features
- Automatically scrapes course details from Coursera.
- Dynamically populates Google Sheets with data.
- Adds dropdowns for course status and LinkedIn updates.
- Auto-detects course type (Modular or Project).
- Applies conditional formatting based on course status.

## How to Use this

1. Make a copy of this Google Spreadsheet:  
   https://docs.google.com/spreadsheets/d/1spMQqnkGoi_0j0z_234sbq8PRReCZ1pAvka1lhzkGrE/edit?usp=sharing
   
2. I have provided a sample course link. Click the "Scrape Coursera Data" button. You should see a "Script running" popup.

**Note:**  
If you are prompted with "Authorisation required. A script attached to this document needs your permission to run," click "OK". Then, log into your Google Account. Click "ADVANCED" below and then "Go to Coursera Scraper (unsafe)".

This is required because:

- **See, edit, create and delete all your Google Sheets spreadsheets**: This allows the script to access and modify your Google Sheets document to populate course data.
- **Connect to an external service**: The script needs permission to access and scrape data from Coursera.

Make sure you trust the **Coursera Scraper** tool.  
On the window, click "Allow" and the script will run. The data should be scraped and available in your spreadsheet.

## Script Overview

### 1. scrapeCourseraData()
This function scrapes course data from the URLs provided in the "Link" column of the 'COURSERA' sheet and populates the relevant details.

```javascript
function scrapeCourseraData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COURSERA');
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  const headers = data[0];
  const linkColIndex = headers.indexOf('Link');

  if (linkColIndex === -1) {
    SpreadsheetApp.getUi().alert('Link column not found!');
    return;
  }

  for (let i = 1; i < data.length; i++) {
    const courseLink = data[i][linkColIndex];
    if (courseLink && !data[i][headers.indexOf('Course Title')]) {
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
```

### 2. getCourseraCourseInfo(url)
This function fetches course details from the provided Coursera course URL.

```javascript
function getCourseraCourseInfo(url) {
  try {
    const response = UrlFetchApp.fetch(url);
    const content = response.getContentText();

    // Extracting course details with regex
    const titleMatch = content.match(/<h1[^>]*data-e2e="hero-title"[^>]*>(.*?)<\/h1>/);
    const institutionMatch = content.match(/<div class="css-1ujzbfc">.*?<img[^>]*alt="(.*?)"[^>]*><\/div>/);
    const experienceMatch = content.match(/<div class="css-fk6qfz">(Beginner level|Intermediate level|Advanced level)<\/div>/);

    const hoursMatch = content.match(/<div class="css-fk6qfz">(\d+(\.\d+)?)\s*hours?<\/div>/);
    const minutesMatch = content.match(/<div class="css-fk6qfz">(\d+)\s*minutes?<\/div>/);

    let hours = 'N/A';
    if (minutesMatch) {
      const minutes = parseInt(minutesMatch[1]);
      hours = (minutes / 60).toFixed(2);
    } else if (hoursMatch) {
      hours = hoursMatch[1];
    }

    const ratingMatch = content.match(/aria-label="(\d+\.\d+)\s*stars"/);
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
      rating: ratingMatch ? ratingMatch[1] : 'N/A',
      reviews: reviewsMatch ? reviewsMatch[1].replace(/,/g, '') : 'N/A'
    };
  } catch (e) {
    Logger.log('Error fetching data: ' + e);
    return null;
  }
}
```

### 3. setDropdowns()
Adds dropdown menus for "State" and "Added to LinkedIn" columns and applies conditional formatting.

```javascript
function setDropdowns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COURSERA');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const currentRow = sheet.getLastRow();

  const stateColumn = headers.indexOf('State') + 1;
  const linkedinColumn = headers.indexOf('Added to LinkedIn') + 1;

  const stateRange = sheet.getRange(currentRow, stateColumn);
  const stateValues = ["STARTED", "NOT STARTED"];
  const stateValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(stateValues)
    .setAllowInvalid(false)
    .setHelpText('Select either STARTED or NOT STARTED')
    .build();
  stateRange.setDataValidation(stateValidation);
  stateRange.setValue("NOT STARTED");
  stateRange.setHorizontalAlignment('center');

  const linkedinRange = sheet.getRange(currentRow, linkedinColumn);
  const linkedinValues = ["YES", "NO"];
  const linkedinValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(linkedinValues)
    .setAllowInvalid(false)
    .setHelpText('Select YES or NO')
    .build();
  linkedinRange.setDataValidation(linkedinValidation);
  linkedinRange.setValue("NO");
  linkedinRange.setHorizontalAlignment('center');

  applyConditionalFormatting(sheet, stateRange, {"STARTED": "#bfe1f6", "NOT STARTED": "#e6e6e6"});
  applyConditionalFormatting(sheet, linkedinRange, {"YES": "#bfe1f6", "NO": "#e6e6e6"});
}

```

### 4. applyConditionalFormatting(sheet, range, colorMap)
Applies color-based conditional formatting for dropdown selections.

```javascript
function applyConditionalFormatting(sheet, range, colorMap) {
  const existingRules = sheet.getConditionalFormatRules();
  const newRules = [];

  Object.keys(colorMap).forEach(value => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(value)
      .setBackground(colorMap[value])
      .setFontColor("#000000")
      .setRanges([range])
      .build();
    newRules.push(rule);
  });

  sheet.setConditionalFormatRules(existingRules.concat(newRules));
}
```

### 5. setCourseType()
Determines whether the course is "MODULAR" or a "PROJECT/GUIDED PROJECT" based on the number of modules.

```javascript
function setCourseType() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COURSERA');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const currentRow = sheet.getLastRow();

  const courseTypeColumn = headers.indexOf('Course Type') + 1;
  const modulesCountColumn = headers.indexOf('No. of Modules') + 1;
  const modulesCountColumnLetter = String.fromCharCode(64 + modulesCountColumn);

  const formula = `=IF(${modulesCountColumnLetter}${currentRow}<>0,"MODULAR","PROJECT/GUIDED PROJECT")`;
  sheet.getRange(currentRow, courseTypeColumn).setFormula(formula);
}
```

### 6. Add New Row Function
To add a new row below the last row in the Google Sheets document, you can use the following function:

```javascript
function addNewRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COURSERA");
  var lastRow = sheet.getLastRow();
  sheet.insertRowAfter(lastRow);
  
  var range = sheet.getRange(lastRow + 1, 1, 1, sheet.getLastColumn());
  range.clearDataValidations();
}
```


## Conclusion
This script provides an efficient solution for collecting Coursera course data directly into Google Sheets, saving time and reducing manual effort. With dropdown menus and conditional formatting, it's easy to manage and track course progress.

For best results, ensure the 'COURSERA' sheet contains headers for 'Link', 'Course Title', 'Institution', 'Recommended Experience', 'Approx. Hours', 'No. of Modules', 'Enrollees', 'Course Rating', 'No. of Reviews', 'State', 'Added to LinkedIn', and 'Course Type'.
