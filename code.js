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
  PENERIMA: {
    NAME: "Penerima",
    INDEX_COLUMN: 4,
  },
  PENGIRIM: {
    NAME: "Pengirim",
    INDEX_COLUMN: 5,
  },
  DOKUMEN: {
    NAME: "Dokumen",
    INDEX_COLUMN: 6,
  },
  IMAGE_UPLOAD: {
    NAME: "Image Upload",
    INDEX_COLUMN: 7,
  },
  SIGNATURE: {
    NAME: "Signature",
    INDEX_COLUMN: 8,
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


// This function splits rows based on comma-separated values in column E and adds new rows with the split values
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
      const eValue = row[COLUMN.PENGIRIM.INDEX_COLUMN - 1]; // Assuming column E is the fifth column (index 4)

      // Check if column E value contains commas
      if (eValue?.toString().includes(",")) {
        const eValues = eValue.split(",");
        const numEValues = eValues.length;

        // Duplicate row for each split value in column E
        for (let j = 0; j < numEValues; j++) {
          const newRow = row.slice(0); // Duplicate the entire row
          newRow[COLUMN.PENGIRIM.INDEX_COLUMN - 1] = eValues[j].trim(); // Replace column E with split value
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

// This function changes image URLs in columns G and H to display the images in Google Sheets
function changeimagetourl() {
  const dataRange = sheet.getRange(2, 7, sheet.getLastRow() - 1, 2);

  // Retrieves the data range for columns G (column 7) and H (column 8)
  const data = dataRange.getValues();

  for (let i = 0; i < data.length; i++) {
    const valueG = data[i][0];
    const valueH = data[i][1];

    addImageUrlFromExistedCellObject(
      valueG,
      sheet.getRange(i + 2, 7),
      sheet.getRange(i + 2, 9)
    );
    addImageUrlFromExistedCellObject(
      valueH,
      sheet.getRange(i + 2, 8),
      sheet.getRange(i + 2, 10)
    );

    // Checks if the value in column G is a string and contains 'drive.google.com'
    changeDriveUrlToImageUrl(
      valueG,
      sheet.getRange(i + 2, 7),
      sheet.getRange(i + 2, 9)
    );
    changeDriveUrlToImageUrl(
      valueH,
      sheet.getRange(i + 2, 8),
      sheet.getRange(i + 2, 10)
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

