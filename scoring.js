// WEEKLY FF SCORING - Updated 11.25.2025

// Creates a trigger to automatically run the `sleeperScoring` function every X minutes
function sleeperLiveScoringOn() {
  let frequency = 5; // minutes
  // Run the function at onset to fetch scores initially, then the trigger should run every X minutes
  
  sleeperScoringLogging();
  ScriptApp.newTrigger(`sleeperScoringLogging`)
      .timeBased()
      .everyMinutes(frequency) // must be 1, 5, 10, 15, 30
      .create();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getRangeByName(`LIVE`).setValue(`LIVE SCORING`)
    .setFontColor(`#FFD900`)
    .setFontSize(12)
    .setFontWeight(`bold`);
  ss.toast(`Live Scoring ON`,`🟢 LIVE SCORING`);
}

// Deletes the trigger based on any trigger associated that trigger the script `sleeperScoring`
function sleeperLiveScoringOff() {
  let triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if ( triggers[i].getHandlerFunction() == `sleeperScoringLogging` ) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getRangeByName(`LIVE`).setValue('');
  ss.toast(`Live Scoring OFF`,`🔴 LIVE SCORING`);
}

// SHOW/HIDE SCORING
function scoringShow(silent) {
  toggleScoringCells(true,silent);
}

function scoringHide(silent) {
  toggleScoringCells(false,silent);
}

function toggleScoringCells(visible,silent) {
  visible = visible == undefined ? false : visible;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const sheet = ss.getSheetByName('ROSTERS');
    const maxCols = sheet.getMaxColumns();
    const maxRows = sheet.getMaxRows();
    
    // Hide/show columns (every 3rd column starting at 4)
    for (let col = 4; col <= maxCols; col += 3) {
      const range = sheet.getRange(1, col);
      visible ? sheet.unhideColumn(range) : sheet.hideColumn(range);
    }
    
    // Hide/show last two rows
    const lastTwoRows = sheet.getRange(maxRows - 1, 1, 2, 1);
    visible ? sheet.unhideRow(lastTwoRows) : sheet.hideRow(lastTwoRows);
    if (!silent) ss.toast(visible ? `Scoring cells revealed.` : `Scoring cells hidden`,`${visible ? '👀' : '😶‍🌫️'} SCORING ${visible ? 'SHOWN' : 'HIDDEN'}`);
    return visible;
  } catch (err) {
    Logger.log(`Error ${visible ? 'revealing' : 'hiding'} scoring cells on the "ROSTERS" sheet, try unhiding manually: ${err.stack}`);
    if (visible && !silent) ss.toast(`Unable to reveal scoring columns and final rows on the "ROSTERS" sheet, try unhiding manually.`,`⚠️ SCORING REVEAL ERROR`);
  }
}

// SLEEPER SCORES
// pull down players' scoring for a week in Sleeper (returns OBJECT with format {player_id}:{points scored})
function sleeperScoring(ppr,year,week) {
  try {
    if (!ppr || !week || !year) {
      const docProps = PropertiesService.getDocumentProperties();
      const matchups = JSON.parse(docProps.getProperty('matchups'));
      week = week || matchups.week;
      year = year || matchups.year;
      
      const configuration = JSON.parse(docProps.getProperty('configuration'));
      ppr = ppr || configuration.ppr || 0.5;
    }
    const scoring = leagueScoring(ppr);
    
    const obj = JSON.parse(UrlFetchApp.fetch(`https://api.sleeper.app/stats/nfl/${year}/${week}?season_type=regular&position[]=DEF&position[]=FLEX&position[]=QB&position[]=RB&position[]=TE&position[]=WR&position[]=K&order_by=player_id`));
    const players = obj.length;

    let ids = [];
    let pts = {};
    let id, key, score;
    for (let a = 0; a < players; a++) {
      score = 0;
      id = obj[a].player_id;
      for (const key in scoring) {
        if ( isNaN(parseFloat(obj[a].stats[key])*parseFloat(scoring[key])) == false && parseFloat(obj[a].stats[key]) != null ) {
          score = score + parseFloat(obj[a].stats[key])*parseFloat(scoring[key]);
        }
      }
      pts[id] = Math.round(score*100)/100;
    }
    Logger.log(`🔢 Scoring Fetched for Week ${week}`);
    return pts;
  } catch (err) {
    Logger.log(`⚠️ Error fetching sleeper scoring - ${err.stack}`);
  } 
}

// SLEEPER SCORING LOGGING
// write Sleeper scores to sheet
function sleeperScoringLogging() {
  const docProps = PropertiesService.getDocumentProperties();
  const matchupsString = docProps.getProperty('matchups');
  let week = null, year = null, ppr = null;
  if (matchupsString) {
    const matchups = JSON.parse(matchupsString);
    try {
      week = matchups.week;
      year = matchups.year;
    } catch (err) {
      Logger.log(`⚠️ Error fetching the matchups data, please ensure matchups have been configured for at least 1 week of the season: ${err.stack}`);
    }
  }
  const configurationString = docProps.getProperty('configuration');
  if (configurationString) {
    const configuration = JSON.parse(configurationString);
    try {
      ppr = configuration.ppr;
    } catch (err) {
      Logger.log(`⚠️Error getting the PPR setting, please check configuration and try again: ${err.stack}`);
    }
  }

  let obj = sleeperScoring(ppr,year,week)
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sourceID = ss.getRangeByName(`SLPR_PLAYER_ID`).getValues().flat();
  let scoreRange = ss.getRangeByName(`SLPR_SCORE`);
  let arr = [];
  let data = [];
  for (let a = 0; a < sourceID.length; a++) {
    arr = [];
    if (obj[sourceID[a]] != null) {
      arr = [obj[sourceID[a]]];
    } else {
      arr = [''];
    }
    data.push(arr);
  }
  scoreRange.setHorizontalAlignment(`right`);
  // scoreRange.setValues(data);

  if (contestComplete()) {
    ss.toast(`Trigger for live scoring disabled, matchups are all complete. Checking for winner.`,`🏆 WINNER CHECK`)
    const sheet = ss.getSheetByName('ROSTERS');
    const namesRange = ss.getRangeByName(`ROSTER_NAMES`);
    const pointsRange = ss.getRangeByName(`ROSTER_POINTS`);
    let names = namesRange.getValues().flat();
    const regex = /^(?!Score$|\s*$).+/;
    let namesFiltered = names.filter(value => regex.test(value));
    let points = pointsRange.getValues().flat();
    const isNumeric = value => typeof value === `number` && !isNaN(value);
    let pointsFiltered = points.filter(isNumeric);
    // Create an array of indices to maintain the original order
    const indices = Array.from({ length: pointsFiltered.length }, (_, index) => index);
    indices.sort((a, b) => pointsFiltered[b] - pointsFiltered[a]);
    namesFiltered = indices.map(index => namesFiltered[index]);
    pointsFiltered = indices.map(index => pointsFiltered[index]);
    let multiple = (count = pointsFiltered.filter((value) => value === pointsFiltered[0]).length > 1) ? true : false;
    const titleRegEx = /^(?!a|A|the|The|THE).*/;
    
    recapPanel();

    sleeperLiveScoringOff();
    
    toolbar();
    try {
      for (let a = 0; a < names.length; a++) {
        if (points[a] == pointsFiltered[0]) {
          let row = namesRange.getRow();
          let col = namesRange.getColumn();
          multiple ? sheet.getRange(row,col+a).setValue(`🥇 CO-CHAMP: ${names[a]}`) : sheet.getRange(row,col+a).setValue(`🏆 CHAMPION: ${names[a]}`);
        }
      }
    } catch (err) {
      Logger.log(`⚠️ Error setting champion cell on "ROSTERS" page. ${err.stack}`)
    }
  } else {
    scoringShow();
    ss.toast(`Updated scores for all players in week ${week} successfully`,`✏️ SCORES UPDATED`);
  }
}

function contestComplete() {
  let teams = []; year = null, week = null;
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const matchups = JSON.parse(docProps.getProperty('matchups'));
    const configuration = JSON.parse(docProps.getProperty('configuration'));
    week = week || matchups.week;
    year = year || matchups.year;
    teams = matchups[week].teams;
    const awayTeams = teams.filter((_, i) => i % 2 === 0);
    const homeTeams = teams.filter((_, i) => i % 2 === 1);
    const url = `https://api.sleeper.com/schedule/nfl/regular/${year}`;
    const json = JSON.parse(UrlFetchApp.fetch(url));
    let completedMatchups, remaining;
    try {
      completedMatchups = json.filter(matchup => (matchup.week == week && awayTeams.indexOf(matchup.away) >= 0 && homeTeams.indexOf(matchup.home) >= 0));
      Logger.log(`🧹 Filtered games in the contest slate from API: ${completedMatchups.map(item => item.away + '@' + item.home).join(', ')}`);
    } catch (err) {
      Logger.log(`⚠️Check for completion issue: the teams selected not being found for the provided week when checking for game completion. ${err.stack}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(`There's an issue with the selected teams not being found for the provided week when checking for game completion`,`⚠️ COMPLETION CHECK ERROR`);
    }
    remaining = homeTeams.length - completedMatchups.filter(matchup => matchup.status === 'complete').length;
    const done = completedMatchups.every(matchup => matchup.status === 'complete');
    if (done) {
      configuration.complete = true;
      docProps.setProperty('configuration',JSON.stringify(configuration));
      Logger.log(`✅ Detected completion of ${homeTeams.length > 1 ? 'all ' + homeTeams.length + ' matchups' : 'the one matchup'} for the contest!`);
    } else {
      Logger.log(`🕒 Week ${week} has ${remaining == 1 ? '1' : remaining} matchup${remaining > 1 ? 's' : ''} remaining.`);
    }
    return done;
  } catch (err) {
    Logger.log(`⚠️ Issue in evaluating the completeness of the scoring for the matchups in question: ${err.stack}`);
  }
}

// 2025 - Created by Ben Powers
// ben.powers.creative@gmail.com
