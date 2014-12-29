/**
 * A player holds all the information about a single player, identified by their mongoId.
 * @param {varies} id The unique id of the player.
 */
var Player = function(id)  {
  this.name = "";
  this.played = 0;
  this.id = id; // Corresponds to mongoId.
  // Scores are organized by eventId. See getScore for structure of objects accessible by event id.
  this.scores = {};
  this.totalScore = 0;
  this.events = []; // Array of event ids.
  this.eventRanks = {};
  this.cumulativeRanks = {};
  this.streak = {};
};

/**
 * Set name for the player.
 * @param {string} name
 */
Player.prototype.setName = function(name) {
  this.name = name;
}

/**
 * Add a score for a given event id. Event scores must be added in order
 * of occurrence. Event ids must be unique.
 * @param {varies} eventId The unique identifier for the event.
 * @param {integer,float} score The score for the event.
 */
Player.prototype.addScore = function(eventId, score) {
  this.totalScore += score;
  this.played++;
  this.scores[eventId] = {
    event: score,
    cumulative: this.totalScore
  }
}

/**
 * Set cumulative score as of an event not attended.
 */
Player.prototype.addCumulativeScore = function(eventId) {
  if (!this.scores.hasOwnProperty(eventId)) {
    this.scores[eventId] = {};
  }
  this.scores[eventId].cumulative = this.totalScore;
}

/**
 * Set rank achieved in specific event and updates streak.
 * @param {varies} eventId
 * @param {integer} rank
 * @param {integer} total The total participants in the event.
 */
Player.prototype.addEventRank = function(eventId, rank, total) {
  this.eventRanks[eventId] = rank;
  var win = (rank <= total / 2);
  if (win) {
    if (!this.streak.hasOwnProperty("wins")) {
      this.streak.wins = 0;
      delete this.streak.losses;
    }
    this.streak.wins++;
  } else {
    if (!this.streak.hasOwnProperty("losses")) {
      this.streak.losses = 0;
      delete this.streak.wins;
    }
    this.streak.losses++;
  }
}

/**
 * Set cumulative rank as of a specific event.
 */
Player.prototype.addCumulativeRank = function(eventId, rank) {
  this.cumulativeRanks[eventId] = rank;
}

/**
 * Returns the score achieved for a given event, or nothing if event
 * was not attended.
 * @param {varies} eventId The identifier for the event.
 * @returns {object} The score for the event, or false if event was not attended. Object has properties `event`,
 * for the event-specific score, and `cumulative` for the cumulative score up to that point.
 */
Player.prototype.getScore = function(eventId) {
  if (this.scores.hasOwnProperty(eventId)) {
    return this.scores[eventId];
  }
}

/**
 * Get score for specific event.
 */
Player.prototype.getEventScore = function(eventId) {
  if (this.scores.hasOwnProperty(eventId)) {
    return this.scores[eventId].event;
  }
}

/**
 * Get cumulative score as of specific event.
 * @param {varies} eventId
 * @param {boolean} total Return total score if eventId not found.
 *
 */
Player.prototype.getCumulativeScore = function(eventId) {
  if (this.scores.hasOwnProperty(eventId)) {
    return this.scores[eventId].cumulative;
  }
}

Player.prototype.getEventRank = function(eventId) {
  return this.eventRanks[eventId];
}

Player.prototype.getCumulativeRank = function(eventId) {
  return this.cumulativeRanks[eventId];
}

/**
 * Returns the total score for all games played so far.
 * @returns {integer,float} The total score attained so far.
 */
Player.prototype.getTotal = function() {
  return this.totalScore;
}

/**
 * Returns a string identifying the win/lose streak.
 * @returns {string} A string of the form nW or nL for wins and losses, respectively, where n is the number
 * of wins/losses.
 */
Player.prototype.getStreak = function() {
  if (this.streak.wins) {
    return this.streak.wins + "W";
  } else {
    return this.streak.losses + "L";
  }
}

/**
 * Class function that returns a function for sorting, descending, based on score in specific game.
 */
Player.eventScoreSort = function(eventId) {
  return function(playerA, playerB) {
    return playerB.getScore(eventId).event - playerA.getScore(eventId).event;
  };
}

/**
 * Class function that returns a function for sorting, descending, based on score in specific game.
 */
Player.cumulativeScoreSort = function(eventId) {
  return function(playerA, playerB) {
    return playerB.getCumulativeScore(eventId) - playerA.getCumulativeScore(eventId);
  };
}

/**
 * Class function that returns a function for sorting, descending, based on total score.
 */
Player.totalScoreSort = function() {
  return function(playerA, playerB) {
    return playerB.getTotal() - playerA.getTotal();
  };
}

Player.cumulativeRankSort = function(eventId) {
  return function(playerA, playerB) {
    return playerA.getCumulativeRank(eventId) - playerB.getCumulativeRank(eventId);
  }
}

/**
 * An Event holds the summary of results from a set of races held as part of a single event.
 * @param {varies} id The unique identifier for the game.
 */
var Event = function(id, date) {
  this.id = id;
  this.date = date;
  this.sorted = true;
  this.players = [];
  this.playersGameSort = [];
  this.playersCumulativeSort = [];
}

Event.prototype.addPlayer = function(player) {
  this.players.push(player);
  this.sorted = false;
}

/**
 * Returns formatted date of the event.
 * @returns {string}
 */
Event.prototype.getFormattedDate = function() {
  return (this.date.getMonth() + 1) + "/" + this.date.getDate() + "/" + this.date.getYear();
}

/**
 * Returns the players sorted by their score in the event.
 */
Event.prototype.getPlayersByEvent = function() {
  this._sortPlayers();
  return this.playersGameSort;
}

Event.prototype.getPlayersByCumulative = function() {
  this._sortPlayers();
  return this.playersCumulativeSort;
}

/**
 * Calculates player ranks and assigns them to each of the players.
 */
Event.prototype.calculateRanks = function() {
  this._sortPlayers();
  this.playersGameSort.forEach(function(player, rank, players) {
    player.addEventRank(this.id, rank + 1, players.length);
  }, this);
}

Event.prototype._sortPlayers = function() {
  if (!this.sorted) {
    this.playersGameSort = this.players.slice().sort(Player.eventScoreSort(this.id));
    this.playersCumulativeSort = this.players.slice().sort(Player.cumulativeScoreSort(this.id));
    this.sorted;
  }
}

/**
 * Class function that returns comparison function that orders ascending by date.
 */
Event.dateSort = function() {
  return function(eventA, eventB) {
    return eventA.date - eventB.date;
  }
}

/**
 * A DataSheet is an interface to a Sheet and can be used to separate the population and formatting of multiple
 * sheets in a single script.
 * @param {string} name The name of the sheet that will be overwritten/created.
 * @param {array} headers The headers of the data sheet.
 * @param {function} format [optional] Callback function to format the sheet after it is created; passed the created sheet as a parameter.
 */
var DataSheet = function(name, headers, format) {
  this.name = name;
  this.headers = headers;
  if (typeof format == "function") {
    this.format = format.bind(this);
  }
}

/**
 * Create sheet and populate with provided data.
 * @param {array} data The 2d array of data to put into the sheet.
 */
DataSheet.prototype.create = function(data) {
  var output = [this.headers].concat(data);
  this.sheet = DataSheet.createAndPopulateSheet(this.name, output);
  if (this.format) {
    this.format(this.sheet);
  } else {
    DataSheet._autoResize(this.sheet);
  }
}

/**
 * Properly auto-resizes the columns of the provided sheet.
 * @param {Sheet} sheet The sheet to autoResize the columns of.
 * @param {integer} buffer [optional] The number of pixels to add or subtract from the default autoResize width. Defaults to 0.
 */
DataSheet._autoResize = function(sheet, buffer) {
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
 * Populates the sheet with name `name` in the current spreadsheet with the given `data`, creating the sheet if necessary, at index `index`.
 * @param {string} name The name of the sheet.
 * @param {array} data A 2d array with the information to be populated into the sheet.
 * @param {integer} index [optional] The index at which to create the sheet if it does not already exist. Defaults to 0.
 * @returns {Sheet} The created/retrieved sheet.
 */
DataSheet.createAndPopulateSheet = function(name, data, index) {
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

// The main function that executes when the menu option is selected.
function computeWeeklyRanks() {
  // Get list of files with weekly info.
  // For each of them, get the information from the summary sheet (just the name, mongoid, points)
  // Populate the information into a weekly points sheet that acts as a summary of the individual weeks (can also be used directly to get the week-to-week top 10 standings).
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var eventSummaries = getEventSummaries();
  
  // Create event objects.
  var eventSummariesById = {};
  var eventsById = {};
  var events = [];
  
  eventSummaries.forEach(function(eventSummary) {
    var event = new Event(eventSummary.id, eventSummary.date);
    eventsById[event.id] = event;
    eventSummariesById[event.id] = eventSummary;
    events.push(event);
  });
  
  // Sort events oldest to most recent.
  events.sort(Event.dateSort());
  
  
  // Create player objects.
  var playersById = {};
  var players = [];
  
  events.forEach(function(event) {
    var eventSummary = eventSummariesById[event.id];
    var playerData = eventSummary.playerData;
    playerData.forEach(function(playerDatum) {
      if (!playersById.hasOwnProperty(playerDatum.id)) {
        var player = new Player(playerDatum.id);
        playersById[playerDatum.id] = player;
        players.push(player);
      } else {
        var player = playersById[playerDatum.id];
      }
            
      // Set player name to that used in the most recent event participated in.
      player.setName(playerDatum.name);
      
      player.addScore(event.id, playerDatum.score);
      
      event.addPlayer(player);
    });
    
    // Set cumulative scores for players that didn't attend event but have attended a past event.
    players.forEach(function(player) {
      player.addCumulativeScore(event.id);
    });
    
    // Calculate rankings for players in event.
    event.calculateRanks();
    
    // Calculate cumulative rankings for players.
    players.sort(Player.cumulativeScoreSort(event.id));
    players.forEach(function(player, rank) {
      player.addCumulativeRank(event.id, rank + 1);
    });
  });
  
  
  // Creating sheets.
  // Event summaries, scores acquired in each event.
  var eventSummaryHeaders = ["Name", "mongoId"].concat(events.map(function(event) {
    return "Event " + event.id + "\n(" + event.getFormattedDate() + ")";
  }));
  
  var eventSummaryFormat = function(sheet) {
    var idCol = sheet.getRange("B:B");
    var headers = sheet.getRange("1:1");
    var players = sheet.getRange("A2:A");
    
    // Show columns before hiding columns.
    sheet.showColumns(1, sheet.getLastColumn());
    sheet.hideColumn(idCol);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
    sheet.setRowHeight(headers.getRow(), 50)
    headers.setFontWeight("bold");
    headers.setHorizontalAlignment("center");
    players.setHorizontalAlignment("center");
    players.setFontWeight("bold");
    DataSheet._autoResize(sheet, 5);
    headers.setBackground("#ffe599");
    sheet.getRange("A1").setBackground("#f9cb9c"); // Player name column header.
    
  }
  
  var eventSummarySheet = new DataSheet("Event Summaries", eventSummaryHeaders, eventSummaryFormat);

  var eventSummaryData = players.map(function(player) {
    return [player.name, player.id].concat(events.map(function(event) {
      var score = player.getEventScore(event.id);
      if (score) {
        return score;
      } else {
        return "";
      }
    }));
  });
  
  eventSummarySheet.create(eventSummaryData);
  
  
  // Event summaries with cumulative points listed.
  var cumulativeEventSummarySheet = new DataSheet("Event Summaries (cumulative)", eventSummaryHeaders, eventSummaryFormat);
  
  var cumulativeEventSummaryData = players.map(function(player) {
    return [player.name, player.id].concat(events.map(function(event) {
      var score = player.getCumulativeScore(event.id)
      if (score) {
        return score;
      } else {
        return "";
      }
    }));
  });
  
  cumulativeEventSummarySheet.create(cumulativeEventSummaryData);

  
  // Top ten list week to week.
  var eventTopTenFormat = function(sheet) {
    var headers = sheet.getRange("1:1");
    DataSheet._autoResize(sheet, 5);
    headers.setBackground("#ffe599");
    sheet.setRowHeight(headers.getRow(), 50)
    headers.setFontWeight("bold");
    headers.setHorizontalAlignment("center");
  }
  
  var eventTopTenSheet = new DataSheet("Weekly Rankings (Top Ten)", eventSummaryHeaders.slice(2), eventTopTenFormat);
  var eventTopTenData = events.map(function(event) {
    var players = event.getPlayersByEvent().slice(0, 10).map(function(player) {
      return player.name + " (" + player.getEventScore(event.id) + ")";
    });
    // Handle cases where there were less than 10 players.
    if (players.length < 10) {
      for (var i = players.length; i < 10; i++) {
        players.push("");
      }
    }
    return players;
  });
  
  eventTopTenSheet.create(transpose(eventTopTenData));
  
  
  // Main standings sheet  
  var standingsHeaders = [
    "Rank",
    "Previous\nRank",
    "▲▼",
    "Player",
    "mongoId",
    "Points",
    "Streak\nW=Top 1/2\nLow=Bottom 1/2",
    "Best Rank",
    "Played",
    "PPG",
    "Ranks Below\nBest Rank"
  ];
  
  // Formatting function for standings sheet.
  var standingsSheetFormat = function(sheet) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stdHeader = sheet.getRange("1:1");
    var stdRanks = sheet.getRange("A2:B");
    var stdNames = sheet.getRange("D:D");
    var stdDiffs = sheet.getRange("C:C");
    var stdStreak = sheet.getRange("G2:G");
    var stdPoints = sheet.getRange("F:F");
    var stdId = sheet.getRange("E:E");
    var stdPpg = sheet.getRange("J2:J");
    var stdBestRanks = sheet.getRange("K2:K");
    
    
    // Standings formatting.
    var formatSheet = ss.getSheetByName("Formatting");
    var conditionalFormattingRange = formatSheet.getRange("A2");
    var bestRankFormattingRange = formatSheet.getRange("C2");
    var streakFormattingRange = formatSheet.getRange("B2");
    
    conditionalFormattingRange.copyFormatToRange(sheet, stdDiffs.getColumn(), stdDiffs.getLastColumn(), stdDiffs.getRow(), stdDiffs.getLastRow());
    bestRankFormattingRange.copyFormatToRange(sheet, stdBestRanks.getColumn(), stdBestRanks.getLastColumn(), stdBestRanks.getRow(), stdBestRanks.getLastRow());
    streakFormattingRange.copyFormatToRange(sheet, stdStreak.getColumn(), stdStreak.getLastColumn(), stdStreak.getRow(), stdStreak.getLastRow());
    
    sheet.getRange("I1").setNote("Number of games played.");
    sheet.getRange("J1").setNote("Points per game.");
    sheet.showColumns(1, sheet.getLastColumn());
    sheet.hideColumns(stdId.getColumn());
    sheet.setFrozenRows(stdHeader.getRow());
    sheet.setRowHeight(stdHeader.getRow(), 50);
    DataSheet._autoResize(sheet, 30);
    
    sheet.getDataRange().setHorizontalAlignment("center");
    
    stdRanks.setFontWeight("bold");
    stdDiffs.setFontWeight("bold");
    stdStreak.setFontWeight("bold");
    stdHeader.setFontWeight("bold");
    stdNames.setFontWeight("bold");
    
    stdHeader.setBackground("#ffe599");
    sheet.getRange("D1").setBackground("#f9cb9c"); // Player name column header.
  }
  
  var standingsSheet = new DataSheet("Standings", standingsHeaders, standingsSheetFormat);

  // Get players for each event by cumulative points.
  var eventPlayers = events.map(function(event) {
    var relevantPlayers = players.filter(function(player) {
      return player.getCumulativeRank(event.id);
    });
    relevantPlayers.sort(Player.cumulativeRankSort(event.id));
    return relevantPlayers;
  });
  
  // Get most recent rankings for ordering.
  var standingPlayers = eventPlayers.slice(-1)[0];
  //Logger.log(standingPlayers);
  var currentEvent = events.slice(-1)[0];
  if (events.length > 1) {
    var lastEvent = events.slice(-2, events.length - 1)[0];
  } else {
    var lastEvent = false;
  }
  
  standingPlayers.sort(Player.totalScoreSort());
  
  var standingsData = standingPlayers.map(function(player) {
    var currentRank = player.getCumulativeRank(currentEvent.id);
    // Get best rank for this player.
    var bestRank = currentRank;
    for (var id in eventsById) {
      var cumulativeRank = player.getCumulativeRank(id);
      if (cumulativeRank && cumulativeRank < bestRank) {
        bestRank = cumulativeRank;
      }
    }
    var bestRankDiff = currentRank - bestRank;
    
    // Get previous rank if it exists.
    if (lastEvent) {
      
      var previousRank = player.getCumulativeRank(lastEvent.id);
      if (previousRank) {
        var previousRankDiff = previousRank - currentRank;
      }
    }
    
    var pointsPerEvent = player.getCumulativeScore(currentEvent.id) / player.played;
    return [
      currentRank,
      lastEvent && previousRank ? previousRank : "",
      lastEvent && previousRank ? previousRankDiff : "",
      player.name,
      player.id,
      player.getCumulativeScore(currentEvent.id),
      player.getStreak(),
      bestRank,
      player.played,
      Math.floor(pointsPerEvent),
      bestRankDiff
    ];
  });
  
  standingsSheet.create(standingsData);
  
  // Champion Status spreadsheet.
  var champStatsHeaders = [
    "Player",
    "Champion\nSince",
    "Champion\nStatus?",
    "Current Reign\nas Champion",
    "Longest Reign\nas Champion",
    "Total Weeks as\nChampion"
  ]
  
  var champStatsFormat = function(sheet) {
    DataSheet._autoResize(sheet, 30);
    var headers = sheet.getRange("1:1");
    var playerHeader = sheet.getRange("A1");
    var players = sheet.getRange("A2:A");
    var champStatus = sheet.getRange("C2:C");
    headers.setBackground("#ffe599");
    playerHeader.setBackground("#f9cb9c");
    headers.setFontWeight("bold");
    players.setFontWeight("bold");
    headers.setHorizontalAlignment("center");
    players.setHorizontalAlignment("center");
  }
  
  var champStatsSheet = new DataSheet("Champ Stats", champStatsHeaders, champStatsFormat);
  
  // Winning players in each event.
  var championPlayerData = {};
  events.forEach(function(event, index, events) {
    var player = event.getPlayersByCumulative()[0];
    if (!championPlayerData.hasOwnProperty(player.id)) {
      championPlayerData[player.id] = {
        player: player,
        date: event.date,
        longestReign: 0,
        total: 0,
        currentReign: 0,
        current: false
      };
    }
    var champion = championPlayerData[player.id];
    champion.total++;
    
    if (index != 0) {
      if (champion.current) {
        champion.currentReign++;
      } else {
        // Reset current reign and resent current status on all other champions.
        champion.currentReign = 1;
        for (var id in championPlayerData) {
          championPlayerData[id].current = false;
        }
      }
    } else {
      champion.currentReign = 1;
    }
    champion.current = true;
    
    // Update longest reign.
    if (champion.currentReign > champion.longestReign) {
      champion.longestReign = champion.currentReign;
    }
  });
  
  var championPlayers = [];
  for (var id in championPlayerData) {
    championPlayers.push(championPlayerData[id]);
  }
  
  championPlayers.sort(function(playerA, playerB) {
    return playerA.date - playerB.date;
  });
  
  var champStatsData = championPlayers.map(function(champion) {
    var formattedDate = (champion.date.getMonth() + 1) + "/" + champion.date.getDate() + "/" + champion.date.getYear();
    return [
      champion.player.name,
      champion.date,
      champion.current ? "Current" : "Past",
      champion.currentReign,
      champion.longestReign,
      champion.total
    ];
  });
  
  champStatsSheet.create(champStatsData);
}

/**
 * Transpose a provided array.
 * @param {array} arr The array to transpose.
 * @returns {array} The transposed array.
 */
function transpose(arr) {
  return arr[0].map(function(col, i) {
    return arr.map(function(row) {
      return row[i];
    });
  });
}

/**
 * Retrieves the summary information from the event sheets. Event URLs should be specified in the Events sheet.
 * Returns array of objects with information on each of the event sheets. Objects have the following properties:
 * - `number` {integer} the number of the event in the season.
 * - `url` {string} the url of the spreadsheet for the event.
 * - `date` {Date} the date of the event.
 * - `playerData` {array} an array of objects containing the summary of player standings for the event. Objects have:
 *   + `name` {string} the display name of the player.
 *   + `id` {string} the id of the player, corresponds to the mongoId.
 *   + `score` {integer} the earned points for the player for this event.
 */
function getEventSummaries() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eventSheet = ss.getSheetByName("Events");
  
  var eventValues = eventSheet.getDataRange().getValues();
  var eventSheets = eventValues.slice(1).map(function(row) {
    return { id: row[0], url: row[1], date: row[3] };
  });
  
  // Get summary data from event sheets.
  eventSheets.forEach(function(eventSheet) {
    var otherSpreadsheet = SpreadsheetApp.openByUrl(eventSheet.url);
    var summarySheet = otherSpreadsheet.getSheetByName("Summary");
    eventSheet.playerData = summarySheet.getRange("A2:C").getValues().map(function(data) {
      return {name: data[0], id: data[1], score: data[2]};
    }).filter(function(player) {
      return player.name !== "";
    });
  });
  
  return eventSheets;
}

// Add custom menu option.
function onOpen() {
  var ss = SpreadsheetApp.getActive();
  var newMenu = [
    {name: 'Recalculate Standings...', functionName: 'computeWeeklyRanks'}
  ];
  ss.addMenu('TagPro Racing', newMenu);
}
