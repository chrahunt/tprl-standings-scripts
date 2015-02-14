// Add Menu Option
function onOpen(e) {
  var ss = SpreadsheetApp.getActive();
  var newMenu = [
    {name: 'Open Sidebar', functionName: 'showSidebar'},
    {name: 'Recalculate...', functionName: 'calculateRaces'}
  ];
  ss.addMenu('TagPro Racing', newMenu);
}

// Display sidebar.
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)    
      .setTitle('Stats');
  SpreadsheetApp.getUi().showSidebar(ui);
}

/*
 * Include a file from the same Script project from within a template.
 * 
 * @param  {string}  filename  The filename of the file to include.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .getContent();
}

/**
 * Take in a CSV string output by the Lap Timer userscript and return either an array of arrays
 * representing rows (with header row included), or an array of objects, if the `objects` param
 * is set to true. If an array of objects is returned, they have the following properties:
 *  * mongoId: The player's mongoId
 *  * name: The player's name
 *  * laps: An array with the lap times for each lap. If a player did not complete a lap, an empty string
 *      will be present in the array at the index corresponding to the lap.
 * This function does no error checking, and may break if the userscript output changes. This function
 * does properly handle cases where players have a comma in their name.
 *
 * @param {string} csvString  The string read in from the userscript file.
 * @param {boolean} objects  Whether or not to return an array of objects.
 */
function parseCSV(csvString, objects) {
  if (typeof objects == "undefined") objects = false;
  var dataRowRegex = "([0-9a-f]*),(.*)";
  var lapRegex = ",(\\d*)";
  var rows = csvString.split("\n");
  var header = rows.splice(0, 1)[0];
  // Get number of laps;
  var laps = header.match(/ (\d+)"$/)[1];
  laps = +laps;
  for (var i = 0; i < laps; i++) {
    dataRowRegex += lapRegex;
  }
  dataRowRegex += "$";
  
  var re = new RegExp(dataRowRegex);
  var data = [];
  if (objects) {
    rows.forEach(function(row) {
      var rowData = {};
      var match = row.match(re);
      rowData.mongoId = match[1];
      rowData.name = match[2];
      rowData.laps = [];
      for (var i = 0; i < laps; i++) {
        rowData.laps.push(match[i + 3]);
      }
      data.push(rowData);
    });
  } else {
    var headers = ["mongoId", "name"];
    for (var i = 0; i < laps; i++) {
      headers.push("Lap " + (i + 1));
    }
    data.push(headers);
    rows.forEach(function(row) {
      var rowData = [];
      var match = row.match(re);
      rowData.push(match[1]); // mongoId
      rowData.push(match[2]); // player name
      
      for (var i = 0; i < laps; i++) {
        rowData.push(match[i + 3]);
      }
      data.push(rowData);
    });
  }
  return data;
}

/* Sheet Creation and Formatting Functions */
/**
 * Populates the sheet with name `name` in the current spreadsheet with the given `data`, creating the sheet if necessary, at index `index`.
 * @param {string} name The name of the sheet.
 * @param {array} data A 2d array with the information to be populated into the sheet.
 * @param {integer} index [optional] The index at which to create the sheet if it does not already exist. Defaults to 0.
 * @returns {Sheet} The created/retrieved sheet.
 */
function createAndPopulateSheet(name, data, index) {
  if (typeof index == "undefined") index = 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create summary sheet.
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name, index);
  } else {
    sheet.clear();
  }
  
  // Ensure correct number of rows and columns in sheet.
  var numRows = sheet.getMaxRows();
  var numCols = sheet.getMaxColumns();
  
  if (numRows < data.length) {
    sheet.insertRows(1, data.length - numRows);
  } else if (numRows > data.length) {
    sheet.deleteRows(1, numRows - data.length);
  }
  
  if (numCols < data[0].length) {
    sheet.insertColumns(1, data[0].length - numCols);
  } else if (numCols > data[0].length) {
    sheet.deleteColumns(1, numCols - data[0].length);
  }
  
  // Write values to summary sheet.
  var targetRange = sheet.getRange(1, 1, data.length, data[0].length);
  targetRange.setValues(data);
  return sheet;
}

/**
 * Properly auto-resizes the columns of the provided sheet.
 * @param {Sheet} sheet The sheet to autoResize the columns of.
 * @param {integer} buffer [optional] The number of pixels to add or subtract from the default autoResize width. Defaults to 0.
 */
function autoResize(sheet, buffer) {
  if (typeof buffer == "undefined") buffer = 0;
  // Adjust for headers that contain \n by setting headers to longest length line in header before auto resize.
  var headerRange = sheet.getRange("1:1");
  var headers = headerRange.getValues()[0];
  var oldHeaders = headers.slice();
  for (var i in headers) {
    headers[i] = oldHeaders[i].split("\n").reduce(function(prev, current) {
      return prev.length > current.length ? prev : current;
    });
  }
  headerRange.setValues([headers]);
  
  for (var i = 0; i < sheet.getMaxColumns(); i++) {
    sheet.autoResizeColumn(i + 1);
    sheet.setColumnWidth(i + 1, sheet.getColumnWidth(i + 1) + buffer);
  }
  
  headerRange.setValues([oldHeaders]);
}

/**
 * Escape text with leading equal sign for insertion into sheet.
 */
function escapeText(str) {
  if (str.length > 0 && str[0] == "=") {
    return "'" + str;
  } else {
    return str;
  }
}

/**
 * Takes in a file output from Lap Timer userscript and adds a sheet with the information to the spreadsheet and an
 * entry into the "Rounds" sheet. Currently only supports addition of a single file.
 * @param {Array} files An array of objects with properties:
 *  * name: The name of the file.
 *  * content: The content of the file.
 */
function addRounds(files) {
  // For now we're only handling one file.
  var file = files[0];
  var name = file.name;
  var content = file.content;
  // Parse the CSV into an array of rows.
  var eventData = parseCSV(content);
  
  // Extract the time and the name/author of the map played.
  var fileNameRe = /tagproracing-(\d+)-(.+)\.csv$/;
  var nameMatch = name.match(fileNameRe);
  var time = nameMatch[1];
  var mapName = nameMatch[2];
  
  time = +time;
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  
  // Get or create Rounds sheet.
  var roundSheet = ss.getSheetByName("Rounds");
  if (!roundSheet) {
    var roundValues = [["Sheet Name", "Number"]];
    var lastRow = 0; // Number of the last round.
  } else {
    // Retrieve existing data.
    var roundValues = roundSheet.getRange(1, 1, roundSheet.getLastRow(), 2).getValues();
    var lastRow = roundSheet.getLastRow() - 1;
  }
  
  var sheetName = "R" + (lastRow + 1) + ": " + mapName;
  
  // Escape player names.
  eventData.forEach(function(row, i) {
    if (i !== 0) {
      row[1] = escapeText(row[1]);
    }
  });
  
  // Create sheet for event data.
  createAndPopulateSheet(sheetName, eventData);  
  
  // Update list of rounds.
  roundValues.push([sheetName, lastRow + 1]);
  roundSheet = createAndPopulateSheet("Rounds", roundValues);
  
  // Move rounds sheet to front.
  roundSheet.activate();
  ss.moveActiveSheet(0);
}

/**
 * This function is responsible for generating the summary sheet. Prereqs: A sheet named "Rounds" needs to exist
 * in the spreadsheet, and it should have 2 columns, "Name", and "Number. "Name" should have the name of the sheet
 * that will contain the information (userscript export) for a given round, and number should have the round number
 * (e.g. 1 for round 1, 2 for round 2, where "round" corresponds to the playing of a single map in the weekly tournament).
 * 
 * The generated summary sheet will have columns for player name, player mongoId, the player's total score, and each of the rounds.
 */
function calculateRaces() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roundSheet = ss.getSheetByName("Rounds");
  // Get list of rounds, with names corresponding to the sheets with round information.
  var roundData = roundSheet.getRange("A2:B").getValues();
  roundData = roundData.filter(function(data) { return data[0] !== ""; });
  var rounds = roundData.map(function(data) {
    return {name: data[0], number: data[1]};
  });
  
  /**
   * Holds mongoId => player object values. Player objects may have properties:
   * * id: the mongoId of the player
   * * name: the name of the player
   * * scores: an object with the scores the player recieved in each round
   * * total: the total score the player recieved.
   */
  var playersById = {};
  
  // Get the scores for each player/round.
  rounds.forEach(function(round) {
    var sheet = ss.getSheetByName(round.name);
    
    // TODO: Actual error handling.
    if (!sheet) {
      Browser.msgBox("Error retrieving sheet with name: " + round.name + "\nPlease check your spelling and try again.");
      return;
    }
    
    var vals = sheet.getSheetValues(2, 1, -1, -1);
    
    getRaceScores(round, vals, playersById);
  });
  
  // Total up scores for each player.
  for (var id in playersById) {
    var player = playersById[id];
    player.total = getTotalScore(player.scores);
  }
  
  // Get and sort array of players for display.
  var players = [];
  for (var playerId in playersById) {
    players.push(playersById[playerId]);
  }
  
  players.sort(function(playerA, playerB) { return playerB.total - playerA.total; });

  // Output.
  var output = [];
  
  rounds.sort(function(roundA, roundB) { return roundA.number - roundB.number; });
  var headers = ["Name", "mongoId", "Total"].concat(rounds.map(function(round) {
    return round.name;
  }));
  output.push(headers);
  
  // Write player information rows to output array.
  players.forEach(function(player) {
    // Escape player name text.
    var name = escapeText(player.name);
    var playerRow = [name, player.id, player.total];
    
    rounds.forEach(function(round) {
      if (player.scores.hasOwnProperty(round.name)) {
        playerRow.push(player.scores[round.name]);
      } else {
        playerRow.push("");
      }
    });
    output.push(playerRow);
  });

  var summarySheet = createAndPopulateSheet("Summary", output);
  
  // Formatting.
  // Resize columns.
  autoResize(summarySheet, 5);
  summarySheet.showColumns(1, summarySheet.getLastColumn());
  // Hide mongoId column.
  var mongoIds = summarySheet.getRange("B:B");
  Logger.log(mongoIds.getValues()); // DEBUG
  summarySheet.hideColumn(mongoIds);
  // Freeze header and first column.
  if (summarySheet.getFrozenColumns() !== 1)
    summarySheet.setFrozenColumns(1);
  if (summarySheet.getFrozenRows() !== 1)
    summarySheet.setFrozenRows(1);
  
  // Move to front.
  summarySheet.activate();
  ss.moveActiveSheet(0);
}

/**
 * Get the score for a single race, given the results and an object which will be populated
 * with players by player id. Currently calculates score based on position, where achieving
 * `i`-th place out of `n` racers that finished achieves a score of `n - i + 1` for that round.
 * @param {object} round The information for the round.
 * @param {array} results The data rows for the results of the race, as output by the Lap Timer userscript.
 * @param {object} players The object to be populated with player information from this round.
 */
function getRaceScores(round, results, players) {
  var roundPlayers = [];
  results.forEach(function(player) {
    var lastLapTime = player[player.length - 1];
    var mongoId = player[0];
    var playerName = player[1];
    if (mongoId == "" || lastLapTime == "") {
      return;
    } else {
      roundPlayers.push({ id: mongoId, name: playerName, time: lastLapTime });
    }
  });
  roundPlayers.sort(function(playerA, playerB) { return playerA.time - playerB.time; });
  roundPlayers.forEach(function(player, position) {
    // Create player if one does not exist.
    if (!players.hasOwnProperty(player.id)) {
      players[player.id] = {
        id: player.id,
        name: player.name,
        scores: {}
      };
    }
    
    // Set score for this round.
    players[player.id].scores[round.name] = roundPlayers.length - position;
  });
}

/**
 * Get the total score that a player earned based on their individual scores.
 * @param {object} scores An object with eventId => score values.
 * @returns {number} The total score earned by the player.
 */
function getTotalScore(scores) {
  var total = 0;
  for (var score in scores) {
    total += scores[score];
  }
  return total;
}
