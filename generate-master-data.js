function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Workout Planner')
    .addItem('Export Planner', 'showJSONInDialog')
    .addItem('Export MasterData', 'exportMasterData')
    .addItem('Export Equipments', 'convertEquipmentsToJson')
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
    coach: coach,
    createdDate: createdDate
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
function convertEquipmentsToJson() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("equipments");
  if (!sheet) {
    Logger.log("Sheet 'equipments' not found!");
    return;
  }

  var data = sheet.getDataRange().getValues();
  var jsonDataEn = [];
  var jsonDataTa = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var exercises = row[3] ? row[3].split(", ") : [];
    var enVideo = row[4] ? [row[4]] : [];
    var taVideo = row[5] ? [row[5]] : [];
    
    jsonDataEn.push({
      id: row[0],          // Equipment ID
      name: row[1],        // Equipment Name
      thumbnail: row[2],   // Thumbnail URL
      exercises: exercises,
      videos: enVideo
    });

    jsonDataTa.push({
      id: row[0],
      name: row[1],
      thumbnail: row[2],
      exercises: exercises,
      videos: taVideo
    });
  }

  // Convert JSON to string
  var jsonStringEn = JSON.stringify(jsonDataEn, null, 2);
  var jsonStringTa = JSON.stringify(jsonDataTa, null, 2);

  // Save JSON files
  saveJsonToFolder("equipments_en.json", jsonStringEn);
  saveJsonToFolder("equipments_ta.json", jsonStringTa);

  Logger.log("JSON files saved in the same folder as the Google Sheet.");
}

function saveJsonToFolder(filename, jsonString) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetFile = DriveApp.getFileById(ss.getId());
  var parentFolder = sheetFile.getParents().next(); // Get parent folder

  var existingFiles = parentFolder.getFilesByName(filename);

  // Delete existing file if it already exists
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  // Create new JSON file in the sheet’s folder
  parentFolder.createFile(filename, jsonString, MimeType.PLAIN_TEXT);
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
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var jsonArrayEn = [];
  var jsonArrayTa = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var jsonObjectEn = {};
    var jsonObjectTa = {};

    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      var value = row[j];

      if (header === 'instructions(en)') {
        // continue; // Skip instructions column
      }

      const headerKey = header.replace(/\s*\((en|ta)\)\s*/g, ''); // Remove language code

      if (['primaryMuscles', 'secondaryMuscles', 'images'].includes(header)) {
        jsonObjectEn[header] = value ? value.split(',') : [];
        jsonObjectTa[header] = value ? value.split(',') : [];
      } else if (header.endsWith('(en)')) {
        if (typeof value === 'string' && value.startsWith('[')) {
          value = JSON.parse(value);
          jsonObjectEn[headerKey] = value ? value : [];
        } else {
          jsonObjectEn[headerKey] = value ? value.split(',') : [];
        }
      } else if (header.endsWith('(ta)')) {
        if (typeof value === 'string' && value.startsWith('[')) {
          try {
            value = JSON.parse(value);
          } catch (e) {
            value = value ? value.split(',') : [];
          }
          jsonObjectTa[headerKey] = value ? value : [];
        } else {
          jsonObjectTa[headerKey] = value ? value.split(',') : [];
        }
      } else {
        jsonObjectEn[header] = value;
        jsonObjectTa[header] = value;
      }

    }

    jsonArrayEn.push(jsonObjectEn);
    jsonArrayTa.push(jsonObjectTa);
  }

  var jsonStringEn = JSON.stringify(jsonArrayEn, null, 2);
  var jsonStringTa = JSON.stringify(jsonArrayTa, null, 2);

  var fileId = spreadsheet.getId();
  var file = DriveApp.getFileById(fileId);
  var folder = file.getParents().next();

  var jsonFileEn = folder.createFile('Workouts_en.json', jsonStringEn, MimeType.PLAIN_TEXT);
  var jsonFileTa = folder.createFile('Workouts_ta.json', jsonStringTa, MimeType.PLAIN_TEXT);

  Logger.log('Download URL (EN): ' + jsonFileEn.getDownloadUrl());
  Logger.log('Download URL (TA): ' + jsonFileTa.getDownloadUrl());

  return {
    en: jsonFileEn.getDownloadUrl(),
    ta: jsonFileTa.getDownloadUrl()
  };
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
    const muscleGroup = dataValues[i][15]; // Assuming muscleGroups is in the 13th column (index 12)
    const exerciseId = dataValues[i][0]; // Assuming id is in the 1st column (index 0)
    const level = dataValues[i][4]; // Assuming level is in the 5th column (index 4)
    const popularity = dataValues[i][3]; // Assuming popularity is in the 4th column (index 3)
    if (i == 1) {
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
  //console.log(workoutsValues)

  // Apply data validation to column C based on the muscle group in column B
  for (let i = 1; i < workoutsValues.length; i++) {
    const muscleGroup = workoutsValues[i][1]; // Assuming muscleGroups is in the 2nd column (index 1)
    const cell = plannerSheet.getRange(i + 1, 3); // Column C (index 2)
    console.log(muscleGroup)
    if (muscleGroupMap[muscleGroup]) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(muscleGroupMap[muscleGroup])
        .setAllowInvalid(false)
        .build();
      cell.setDataValidation(rule);
    } else {
     // cell.clearDataValidations();
    }
  }
}
