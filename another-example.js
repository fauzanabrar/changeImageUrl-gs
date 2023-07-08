const CONFIG = {
  SHEET_NAME: "Sheet1",
};

const COLUMN = {
  ADDED_TIME: {
    NAME: "Added Time",
    INDEX_COLUMN: 1,
  },
  IP_ADDRESS: {
    NAME: "IP Address",
    INDEX_COLUMN: 2,
  },
  NO: {
    NAME: "No",
    INDEX_COLUMN: 3,
  },
  PENGIRIM: {
    NAME: "Pengirim",
    INDEX_COLUMN: 4,
  },
  DEPARTEMEN: {
    NAME: "Departemen",
    INDEX_COLUMN: 5,
  },
  PENERIMA: {
    NAME: "Penerima",
    INDEX_COLUMN: 6,
  },
  DOKUMEN: {
    NAME: "Dokumen",
    INDEX_COLUMN: 7,
  },
  EKSPEDISI: {
    NAME: "Ekspedisi",
    INDEX_COLUMN: 8,
  },
  IMAGE_UPLOAD: {
    NAME: "Image Upload",
    INDEX_COLUMN: 9,
  },
  SIGNATURE: {
    NAME: "Signature",
    INDEX_COLUMN: 10,
  },
  IMAGE_URL: {
    NAME: "Image URL",
    INDEX_COLUMN: 11,
  },
  SIGNATURE_URL: {
    NAME: "Signature URL",
    INDEX_COLUMN: 12,
  },
};

// Initialize the sheet
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

function onSheetChange(e) {
  var changedSheet = e.source.getActiveSheet();

  // Check if the change occurred in the desired sheet (optional)
  if (changedSheet.getName() !== CONFIG.SHEET_NAME) return;
  formatSheet();
}

function formatSheet() {
  splitAndAddRowsBelow();
  changeimagetourl();
}


// This function splits rows based on comma-separated values in column F and adds new rows with the split values
function splitAndAddRowsBelow() {
  let run = true;

  while (run) {
    const refreshedSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = refreshedSheet.getDataRange();
    const data = dataRange.getValues();
    const numRows = dataRange.getNumRows();

    const newData = [];
    for (let i = 0; i < numRows; i++) {
      const row = data[i];
      const fValue = row[COLUMN.PENERIMA.INDEX_COLUMN - 1]; // Assuming column F is the fifth column (index 4)

      // Check if column F value contains commas
      if (fValue?.toString().includes(",")) {
        const fValues = fValue.split(",");
        const numFValues = fValues.length;

        // Duplicate row for each split value in column F
        for (let j = 0; j < numFValues; j++) {
          const newRow = row.slice(0); // Duplicate the entire row
          newRow[COLUMN.PENERIMA.INDEX_COLUMN - 1] = fValues[j].trim(); // Replace column F with split value
          newData.push(newRow);

          if (j === 0) {
            refreshedSheet
              .getRange(i + 1, 1, 1, newRow.length)
              .setValues([newRow]);
          } else {
            refreshedSheet.insertRowsAfter(i + 1, 1);
            i++;
            refreshedSheet
              .getRange(i + 1, 1, 1, newRow.length)
              .setValues([newRow]);
          }
        }
      }
    }

    if (newData.length === 0) {
      run = false;
    }
  }
}

// This function changes image URLs in columns I and J to display the images in Google Sheets
function changeimagetourl() {
  const dataRange = sheet.getRange(2, COLUMN.IMAGE_UPLOAD.INDEX_COLUMN, sheet.getLastRow() - 1, 2);

  // Retrieves the data range for columns I (column 9) and J (column 10)
  const data = dataRange.getValues();

  for (let i = 0; i < data.length; i++) {
    const valueI = data[i][0];
    const valueJ = data[i][1];

    addImageUrlFromExistedCellObject(
      valueI,
      sheet.getRange(i + 2, COLUMN.IMAGE_UPLOAD.INDEX_COLUMN),
      sheet.getRange(i + 2, COLUMN.IMAGE_URL.INDEX_COLUMN)
    );
    addImageUrlFromExistedCellObject(
      valueJ,
      sheet.getRange(i + 2, COLUMN.SIGNATURE.INDEX_COLUMN),
      sheet.getRange(i + 2, COLUMN.SIGNATURE_URL.INDEX_COLUMN)
    );

    // Checks if the value in column I is a string and contains 'drive.google.com'
    changeDriveUrlToImageUrl(
      valueI,
      sheet.getRange(i + 2, COLUMN.IMAGE_UPLOAD.INDEX_COLUMN),
      sheet.getRange(i + 2, COLUMN.IMAGE_URL.INDEX_COLUMN)
    );
    
    // Checks if the value in column J is a string and contains 'drive.google.com'
    changeDriveUrlToImageUrl(
      valueJ,
      sheet.getRange(i + 2, COLUMN.SIGNATURE.INDEX_COLUMN),
      sheet.getRange(i + 2, COLUMN.SIGNATURE_URL.INDEX_COLUMN)
    );
  }
}

function addImageUrlFromExistedCellObject(cellValue, currentRange, addedRange) {
  if (typeof cellValue === "object") {
    try {
      const formula = currentRange.getFormula();
      const imageUrl = formula.match(/\"(.+?)\"/)[1];
      
      const width = 400;
      const imageId = imageUrl.split('id=')[1];
      const srcfile = DriveApp.getFileById(imageId);
      const newImageUrl = Drive.Files.get(srcfile.getId()).thumbnailLink.replace(/\=s.+/, "=s" + width);
      
      insertImageIntoCell(newImageUrl, currentRange);
      insertImageUrlIntoCell(imageUrl, addedRange);
    } catch (e) {
      
      if(e instanceof TypeError) return;
      
      Logger.info("Error : " + e);
    }
  }
}

function insertImageUrlIntoCell(imageUrl, cellRange) {
  cellRange.setValue(imageUrl);
}

function changeDriveUrlToImageUrl(
  cellValue,
  cellRange,
  imageUrlRange,
  useFormula = false
) {
  if (typeof cellValue === "string" && cellValue.includes("drive.google.com")) {
    const adaptImageURL = cellValue.split("/")[5];
    const imageUrl = "https://drive.google.com/uc?id=" + adaptImageURL;

    if (useFormula) {
      cellRange.setFormula(`=IMAGE("${imageUrl}",2)`);
    } else {
      try{
        insertImageIntoCell(imageUrl, cellRange);
      }catch(e){
        const width = 400;
        const imageId = imageUrl.split('id=')[1];
        const srcfile = DriveApp.getFileById(imageId);
        const newImageUrl = Drive.Files.get(srcfile.getId()).thumbnailLink.replace(/\=s.+/, "=s" + width);

        insertImageIntoCell(newImageUrl, cellRange);
      }
    }

    if (imageUrlRange) {
      insertImageUrlIntoCell(imageUrl, imageUrlRange);
    }
  }
}

function insertImageIntoCell(imageUrl, range) {
  const image = SpreadsheetApp.newCellImage()
    .setSourceUrl(imageUrl)
    .build()
    .toBuilder();
  range.setValue(image);
}

function autoResize() {
  resizeColumn(sheet);
  resizeRow(sheet);

  centerTextVerticallyInRange();
}

function resizeColumn(sheet) {
  sheet.setColumnWidth(1, 150);
  sheet.autoResizeColumn(2);
  sheet.setColumnWidth(3, 50);
  sheet.autoResizeColumn(4);
  sheet.autoResizeColumn(5);
  sheet.autoResizeColumn(6);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 100);
}

function resizeRow(sheet) {
  for (let i = 2; i <= sheet.getLastRow(); i++) {
    sheet.setRowHeight(i, 90);
  }
}

function centerTextVerticallyInRange() {
  // Set the vertical alignment to center for the selected range
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, COLUMN.length);
  range.setVerticalAlignment("middle");
}

