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
  
  var playersById = {};
  
  // Get the scores for each player/round.
  rounds.forEach(function(round) {
    var sheet = ss.getSheetByName(round.name);
    
    if (!sheet) {
      Browser.msgBox("Error retrieving sheet with name: " + round.name + "\nPlease check your spelling and try again.");
      //throw new Error("Sheet not found: " + round.name);
    }
    
    var vals = sheet.getSheetValues(2, 1, -1, -1);
    var roundPlayers = [];
    vals.forEach(function(player) {
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
      if (!playersById.hasOwnProperty(player.id)) {
        playersById[player.id] = {
          id: player.id,
          name: player.name,
          scores: {}
        };
      }
      
      // Set score for this round.
      playersById[player.id].scores[round.name] = roundPlayers.length - position;
    });
    
    //Logger.log(vals);
  });
  
  // Total up scores for each player.
  for (var id in playersById) {
    var player = playersById[id];
    player.total = 0;
    for (var score in player.scores) {
      player.total += player.scores[score];
    }
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
    return "Round " + round.number + ": " + round.name;
  }));
  output.push(headers);
  
  // Write player information rows to output array.
  players.forEach(function(player) {
    var playerRow = [player.name, player.id, player.total];
    
    rounds.forEach(function(round) {
      if (player.scores.hasOwnProperty(round.name)) {
        playerRow.push(player.scores[round.name]);
      } else {
        playerRow.push("");
      }
    });
    output.push(playerRow);
  });
  
  // Get or create summary sheet.
  var summarySheet = ss.getSheetByName("Summary");
  if (!summarySheet) {
    summarySheet = ss.insertSheet("Summary", 0);
  } else {
    summarySheet.clear();
  }
  
  // Ensure correct number of rows and columns in sheet.
  var numRows = summarySheet.getMaxRows();
  var numCols = summarySheet.getMaxColumns();
  
  if (numRows < output.length) {
    summarySheet.insertRows(1, output.length - numRows);
  } else if (numRows > output.length) {
    summarySheet.deleteRows(1, numRows - output.length);
  }
  
  if (numCols < output[0].length) {
    summarySheet.insertColumns(1, output[0].length - numCols);
  } else if (numCols > output[0].length) {
    summarySheet.deleteColumns(1, numCols - output[0].length);
  }
  
  // Write values to summary sheet.
  var targetRange = summarySheet.getRange(1, 1, output.length, output[0].length);
  targetRange.setValues(output);
  
  // Formatting.
  // Resize columns.
  for (var i = 1; i <= output[0].length; i++) {
    summarySheet.autoResizeColumn(i);
  }
  
}

// Add custom menu option.
function onOpen() {
  var ss = SpreadsheetApp.getActive();
  var newMenu = [
    {name: 'Calculate weekly scores...', functionName: 'calculateRaces'}
  ];
  ss.addMenu('TagPro Racing', newMenu);
}
