// Fisher-Yates Shuffle Function
function shuffleArray() {
 var sheet = SpreadsheetApp.getActiveSheet();
 var lastRow = sheet.getLastRow();
 var numberOfRandomizations = sheet.getRange("A2").getValue();

 // Get the range of data in column B
 var rangeToRandomize = sheet.getRange("B3:B" + lastRow);
 var valuesToRandomize = rangeToRandomize.getValues();

 // Filter out any blank cells from the range
 var filteredValues = valuesToRandomize.filter(function (row) {
   return row[0] != "";
 });

 // Apply the Fisher-Yates shuffle algorithm for a number of times
 for (var i = 0; i < numberOfRandomizations; i++) {
   for (var j = filteredValues.length - 1; j > 0; j--) {
     var k = Math.floor(Math.random() * (j + 1));
     var temp = filteredValues[j];
     filteredValues[j] = [filteredValues[k][0]];
     filteredValues[k] = [temp[0]];
   }
 }

 // Clear the previous values and set new randomized values
 rangeToRandomize.clear();
 rangeToRandomize = sheet.getRange("B3:B" + (filteredValues.length + 2)); // +2 to account for the offset starting at B3
 rangeToRandomize.setValues(filteredValues);

 // Preserve font style and size
 rangeToRandomize.setFontSize(14); // Set the font size to 14
 rangeToRandomize.setFontFamily("Barlow"); // Set the font to Barlow

 return filteredValues;
}


// This function creates random pairs from a list of students.
function createRandomPartners() {
 var sheet = SpreadsheetApp.getActiveSheet();
 var rangeToClear = sheet.getRange("D3:F500");
 rangeToClear.clear();

 var range = sheet.getRange("B3:B500");
 var values = range.getValues();

 // Filter out any blank cells from the range
 var filteredValues = values.filter(function (row) {
   return row[0] != "";
 });

 // Shuffle the filtered values to create random pairs
 var randomValues = shuffleArray(filteredValues);
 var numStudents = randomValues.length;
 var pairs = [];

 // Create pairs of students
 for (var i = 0; i < numStudents - 1; i += 2) {
   pairs.push([randomValues[i][0], randomValues[i + 1][0]]);
 }

 // Add a single student to an existing pair if the total number is odd
 if (numStudents % 2 !== 0) {
   var soloStudent = randomValues[numStudents - 1];
   var randomPairIndex = Math.floor(Math.random() * pairs.length);
   pairs[randomPairIndex].push(soloStudent[0]);
 }

 // Ensure all sub-arrays in `flatPairs` have a length of 3
 var flatPairs = pairs.map(function (pair) {
   while (pair.length < 3) {
     pair.push("");
   }
   return pair;
 });

 var startCell = sheet.getRange("D3"); // Get the starting cell of the range
 var numRows = flatPairs.length; // Get the number of rows in the pairs array
 var targetRange = sheet.getRange(startCell.getRow(), startCell.getColumn(), numRows, 3); // Set the range using the starting cell, number of rows, and a fixed number of columns

 targetRange.setValues(flatPairs); // Set the values in the target range
 targetRange.setFontSize(14); // Set the font size to 14
 targetRange.setFontFamily("Barlow"); // Set the font to Barlow
}




// This function creates random trios from a list of students.
function createRandomTrios() {
 var sheet = SpreadsheetApp.getActiveSheet();
 var rangeToClear = sheet.getRange("H3:K500");
 rangeToClear.clear();

 var range = sheet.getRange("B3:B500");
 var values = range.getValues();

 // Filter out any blank cells from the range
 var filteredValues = values.filter(function (row) {
   return row[0] != "";
 });

 // Shuffle the filtered values to create random trios
 var randomValues = shuffleArray(filteredValues);
 var numStudents = randomValues.length;
 var trios = [];

 // Create trios of students
 for (var i = 0; i < numStudents - 2; i += 3) {
   trios.push([randomValues[i][0], randomValues[i + 1][0], randomValues[i + 2][0]]);
 }

 // Handle the remaining one or two students
 if (numStudents % 3 === 1) {
   // One extra student. Add them to an existing random trio to make a group of four.
   var randomTrioIndex = Math.floor(Math.random() * trios.length);
   trios[randomTrioIndex].push(randomValues[numStudents - 1][0]);
 } else if (numStudents % 3 === 2) {
   // Two extra students. They form a duo.
   trios.push([randomValues[numStudents - 2][0], randomValues[numStudents - 1][0]]);
 }

 // Ensure all sub-arrays in `trios` have a length of 4
 var flatTrios = trios.map(function (trio) {
   while (trio.length < 4) {
     trio.push("");
   }
   return trio;
 });

 var startCell = sheet.getRange("H3"); // Get the starting cell of the range
 var numRows = flatTrios.length; // Get the number of rows in the trios array
 var targetRange = sheet.getRange(startCell.getRow(), startCell.getColumn(), numRows, 4); // Set the range using the starting cell, number of rows, and a fixed number of columns

 targetRange.setValues(flatTrios); // Set the values in the target range
 targetRange.setFontSize(14); // Set the font size to 14
 targetRange.setFontFamily("Barlow"); // Set the font to Barlow
}



// This function creates random quartets from a list of students.
function createRandomQuartets() {
 var sheet = SpreadsheetApp.getActiveSheet();
 var rangeToClear = sheet.getRange("M3:Q500");
 rangeToClear.clear();

 var range = sheet.getRange("B3:B500");
 var values = range.getValues();

 // Filter out any blank cells from the range
 var filteredValues = values.filter(function (row) {
   return row[0] != "";
 });

 // Shuffle the filtered values to create random quartets
 var randomValues = shuffleArray(filteredValues);
 var numStudents = randomValues.length;
 var quartets = [];

 // Create quartets of students
 for (var i = 0; i < numStudents - numStudents % 4; i += 4) {
   quartets.push([randomValues[i][0], randomValues[i + 1][0], randomValues[i + 2][0], randomValues[i + 3][0], ""]);
 }

 // Handle the remaining one, two or three students
 if (numStudents % 4 === 1) {
   var randomQuartetIndex = Math.floor(Math.random() * quartets.length);
   quartets[randomQuartetIndex][4] = randomValues[numStudents - 1][0];
 } else if (numStudents % 4 === 2) {
   var usedIndexes = new Set();
   for (var j = 0; j < 2; j++) {
     var randomQuartetIndex;
     do {
       randomQuartetIndex = Math.floor(Math.random() * quartets.length);
     } while (usedIndexes.has(randomQuartetIndex));
     quartets[randomQuartetIndex][4] = randomValues[numStudents - 2 + j][0];
     usedIndexes.add(randomQuartetIndex);
   }
 } else if (numStudents % 4 === 3) {
   quartets.push([randomValues[numStudents - 3][0], randomValues[numStudents - 2][0], randomValues[numStudents - 1][0], "", ""]);
 }

 var startCell = sheet.getRange("M3"); // Get the starting cell of the range
 var numRows = quartets.length; // Get the number of rows in the quartets array
 var targetRange = sheet.getRange(startCell.getRow(), startCell.getColumn(), numRows, 5); // Set the range using the starting cell, number of rows, and a fixed number of columns

 targetRange.setValues(quartets); // Set the values in the target range
 targetRange.setFontSize(14); // Set the font size to 14
 targetRange.setFontFamily("Barlow"); // Set the font to Barlow
}

// Code by Avi Megiddo & ChatGPT