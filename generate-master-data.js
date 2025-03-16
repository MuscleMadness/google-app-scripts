function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Workout Planner')
    .addItem('Export Planner', 'showJSONInDialog')
    .addItem('Export MasterData', 'exportMasterData')
    .addToUi();
}

function exportWeeklyPlanData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const dataValues = dataRange.getValues();
  
  // Read the values starting from E11
  const coachName = sheet.getRange("F11").getValue();
  const gymName = sheet.getRange("F12").getValue();
  
  const coach = {
    name: coachName,
    gymName: gymName
  };
  const createdDate = new Date().toISOString().split('T')[0];

  const jsonOutput = {
    days: [],
    goal: dataValues[1][6],
    fitnessLevel: dataValues[1][7],
    daysPerWeek: dataValues[1][8],
    equipment: Array.from(new Set(dataValues.slice(1).flatMap(row => row[9].split(';')))),
    focusAreas: Array.from(new Set(dataValues.slice(1).flatMap(row => row[10].split(';')))),
    duration: dataValues[1][11],
    preferences: Array.from(new Set(dataValues.slice(1).flatMap(row => row[12].split(';')))),
    coach : coach,
    createdDate : createdDate
  };
  
  const maxRow = 7;
  for (let i = 1; i < 6; i++) {
    const row = dataValues[i];
    const dayObject = {
      day: row[0],
      muscleGroups: row[1].split(';'),
      exercises: [],
      exerciseIds: row[2].split(', '),
      sets: row[3],
      reps: row[4],
      rest: row[5]
    };
    jsonOutput.days.push(dayObject);
  }
  
  Logger.log(JSON.stringify(jsonOutput, null, 2));
  return jsonOutput;
}

function showJSONInDialog() {
  const jsonOutput = exportWeeklyPlanData();
  const htmlOutput = HtmlService.createHtmlOutputFromFile('PlannerDialogue')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'JSON Output');
  
  // Pass the JSON data to the HTML file
  const script = `<script>displayJSON(${JSON.stringify(jsonOutput)});</script>`;
  htmlOutput.append(script);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'JSON Output');
}


function exportMasterData() {
  // Get the active spreadsheet and the first sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];
  
  // Get all the data in the sheet
  var data = sheet.getDataRange().getValues();
  
  // Get the headers
  var headers = data[0];
  
  // Initialize an array to hold the JSON objects
  var jsonArray = [];
  
  // Loop through the rows of data, starting from the second row
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var jsonObject = {};
    
    // Loop through the columns
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      var value = row[j];

      // Ignore the 'instructions' column
      if (header === 'instructions') {
        // continue;
      }

      // Check if the column should be an array
      if (header === 'primaryMuscles' || header === 'secondaryMuscles' || header === 'images' || header === 'videos') {
        // Split the value into an array
        jsonObject[header] = value ? value.split(',') : [];
      } else {
        jsonObject[header] = value;
      }
    }
    
    // Add the JSON object to the array
    jsonArray.push(jsonObject);
  }
  
  // Convert the array to a JSON string
  var jsonString = JSON.stringify(jsonArray, null, 2);
  
  // Get the folder of the active spreadsheet
  var fileId = spreadsheet.getId();
  var file = DriveApp.getFileById(fileId);
  var folder = file.getParents().next();
  
  // Create a file in the same folder
  var jsonFile = folder.createFile('Workouts.json', jsonString, MimeType.PLAIN_TEXT);
  
  // Get the download URL
  var downloadUrl = jsonFile.getDownloadUrl();
  
  // Log the download URL
  Logger.log('Download URL: ' + downloadUrl);
  
  // Optionally, you can return the download URL
  return downloadUrl;
}

function applyDataValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const plannerSheet = ss.getSheetByName('planner');
  const dataSheet = ss.getSheetByName('data');
  
  // Get all data from the data sheet
  const dataRange = dataSheet.getDataRange();
  const dataValues = dataRange.getValues();
  
  const muscleGroupMap = {};
  for (let i = 1; i < dataValues.length; i++) {
    const muscleGroup = dataValues[i][13]; // Assuming muscleGroups is in the 13th column (index 12)
    const exerciseId = dataValues[i][0]; // Assuming id is in the 1st column (index 0)
    const level = dataValues[i][4]; // Assuming level is in the 5th column (index 4)
    const popularity = dataValues[i][3]; // Assuming popularity is in the 4th column (index 3)
    if (i == 0) {
      console.log(muscleGroup);
      console.log(exerciseId);
      console.log(level);
      console.log(popularity);
    }
    
    // Apply additional conditions: level should be 'beginner' and popularity should be greater than 2
    if (level === 'beginner' && popularity >= 4) {
      if (!muscleGroupMap[muscleGroup]) {
        muscleGroupMap[muscleGroup] = [];
      }
      muscleGroupMap[muscleGroup].push(exerciseId);
    }
  }

  // Get all data from the workouts sheet
  const workoutsRange = plannerSheet.getDataRange();
  const workoutsValues = workoutsRange.getValues();
  
  // Apply data validation to column C based on the muscle group in column B
  for (let i = 1; i < workoutsValues.length; i++) {
    const muscleGroup = workoutsValues[i][1]; // Assuming muscleGroups is in the 2nd column (index 1)
    const cell = plannerSheet.getRange(i + 1, 3); // Column C (index 2)
    
    if (muscleGroupMap[muscleGroup]) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(muscleGroupMap[muscleGroup])
        .setAllowInvalid(false)
        .build();
      cell.setDataValidation(rule);
    } else {
      cell.clearDataValidations();
    }
  }
}
