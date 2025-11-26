// WEEKLY FF INJURIES - Updated 11.25.2025

// INJURIES
// Function to pull injury comments, practice status, and designations for all players during the active week from Pro Football Reference
function injuries(){
  const url = 'https://www.pro-football-reference.com/players/injuries.htm';
  const proFootballTeams = {
    "BUF":"BUF","CAR":"CAR","CHI":"CHI","CIN":"CIN","CLE":"CLE","ARI":"ARI","DAL":"DAL","DEN":"DEN","DET":"DET","GNB":"GB","HOU":"HOU","JAX":"JAX","KAN":"KC","MIA":"MIA",
    "MIN":"MIN","NYG":"NYG","NYJ":"NYJ","TEN":"TEN","PHI":"PHI","PIT":"PIT","LVR":"LV", "LAR":"LAR","BAL":"BAL","LAC":"LAC","SEA":"SEA","SFO":"SF","TAM":"TB","WAS":"WAS"
  };
  const practices = {
    "Did Not Participate In Practice":"DNP",
    "Limited Participation In Practice":"LIMITED PRACTICE",
    "Full Participation In Practice":"FULL PRACTICE"
  }
  const positions = ['QB','RB','FB','WR','TE','K'];
  const resValues = new RegExp(/(((?<=\"\>)|(?<=\"\ \>))[^<"]{1,}|((?<=status\"\>)|(?<=status\"\ \>))[^<"]{0,})/,'g');
  const resRows = new RegExp(/((?<=\<tr\ \>)|(?<=\<tr\>))(.+?)(?=<\/tr>)/,'g');
  const table = (UrlFetchApp.fetch(url).getContentText().split('<table '))[1].split('</table>')[0];
  const week = table.match(/((?<!\<caption\ \>)|(?<!\<caption\>))[A-Za-z0-9\ ]+(?=<\/caption>)/)[0].match(/[0-9]+/)[0];
  let data = {"week":week};
  let players = 0;
  let undesignated = 0;
  let rows = table.match(resRows);
  for (let a = 0; a < rows.length; a++) {
    let id, rowName, rowTeam, rowPos;
    let player = {};
    let row = rows[a].match(resValues);
    if (positions.indexOf(row[2]) >= 0) {
      if (row[4].slice(0,3) == "NIR" || row[4].slice(0,4) == "Rest") {
        // Avoid creating entry if the player had a rest day
        // Logger.log('Not Injury Related - ' + row[0]); // Prints out player name who was a non-participant for non-injury reasons
      } else {
        try {
          rowName = row[0];
          id = nameFinder(rowName);
          if (id == null) {
            Logger.log(`❔ No ID match found for ${rowName}`);
          } else {
            rowTeam = proFootballTeams[row[1]];
            row[2] == 'FB' ? rowPos == 'RB' : rowPos = row[2];
            player = {}; // Creates object for player ID after matching with custom global function

            // Assign object values
            player.name = rowName;
            player.team = rowTeam;
            player.pos = rowPos;
            player.status = row[3];
            player.injury = row[4];
            player.practice = practices[row[5]];
            player.status == '' ? undesignated++ : null;
            players++;
            // Add player to object
            data[id] = player;
          }
        } catch (err) {
          Logger.log('⚠️ Error with data in regard to ' + row[0] + '| ' + err.stack);
        }
      }
    }
  }
  data.undesignated = undesignated;
  data.players = players;
  return data;
}

// INJURY FILTER
// Removes all players who are not within the specified games for the weekly competition
function filterInjuries() {
  const injuryReport = injuries();
  const matchupsString = PropertiesService.getDocumentProperties().getProperty('matchups');
  let matchups = {}, teams = null;
  try {
    if (matchupsString) {
      matchups = JSON.parse(matchupsString);
      teams = matchups[matchups.week].teams;
    }
    Object.keys(injuryReport).forEach(player => {
      if(injuryReport[player] instanceof Object) {
        if( teams.indexOf(injuryReport[player].team) == -1) {
          injuryReport[player].status == '' ? injuryReport.undesignated-- : null;
          injuryReport.players--;
          delete injuryReport[player];
        }
      }
    });
  } catch (err) {
    Logger.log(`⚠️ Error fetching teams for inclusion, returning all results`);
  }
  // Logger.log(JSON.stringify(injuryReport));
  return injuryReport;
}

// UPDATE INJURIES
// Function to gather and update all injury designations on the draft board
function injuryCheck() {

  const injuryReport = filterInjuries();
  const injuryWeek = injuryReport.week;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const configWeek = fetchWeek();
  let percent = ((injuryReport.players-injuryReport.undesignated)/injuryReport.players * 100).toFixed(1);
  let prompt;
  if ( configWeek != injuryWeek ) {
    prompt = ui.alert(`❗ NO INJURY REPORT`,`You've set the event to take place for week ${configWeek}, but the injury report available is for week ${injuryWeek}.\r\n\r\nWould you like to continue importing the injury report?`, ui.ButtonSet.YES_NO);
  } else if (percent == 0) {
    prompt = ui.alert(`🤕 INJURY REPORT`,`Injury report available for week ${injuryWeek} but no players have injury designations confirmed yet.\r\n\r\nWould you still like to update injuries now?`, ui.ButtonSet.OK_CANCEL);
  } else if (percent < 80) {
    prompt = ui.alert(`🤕 INJURY REPORT`,`Injury report available for week ${injuryWeek} but only ${percent}% have designations.\r\n\r\nWould you still like to update injuries now?`, ui.ButtonSet.OK_CANCEL);
  } else {
    prompt = ui.alert(`🤕 INJURY REPORT`,`Injury report available for week ${injuryWeek}.\r\n\r\nUpdate injuries now?`, ui.ButtonSet.OK_CANCEL);
  }
  if (prompt == 'YES' || prompt == 'OK') {
  
    deleteTriggers();

    let rangeIds = ss.getRangeByName('SLPR_PLAYER_ID');
    let ids = rangeIds.getValues().flat();
    let rangeInjury = ss.getRangeByName('SLPR_INJURY');
    
    let healthIds = [];
    let healthNotes = [];

    for (let a = 0; a < ids.length; a++) {
      let id = ids[a];
      if (injuryReport[id] instanceof Object) {
        let injured = injuryReport[id];
        let string = '';
        if (injured.practice == '') {
          string = 'NO INFO';
        } else {
          string = injured.practice;
        }
        
        if (injured.injury != '') {
          string = string.concat(': ' + injured.injury);
        }
        
        if (injured.status != '') {
          string = string + ' (' + injured.status.toUpperCase() + ')';
        }
        healthIds.push([id]);
        healthNotes.push([string]);
        
      } else {
        healthIds.push([id]);
        healthNotes.push(['']);
      }
    }

    rangeInjury.setValues(healthNotes);

    let noteIds = ss.getRangeByName('DRAFT_ID').getValues().flat();
    let noteRange = ss.getRangeByName('DRAFT_HEALTH');
    let noteComments = [];
    noteRange.clearNote();

    for (let a = 0; a < noteIds.length; a++) {
      let id = noteIds[a];
      if (injuryReport[id] instanceof Object) {
        let injured = injuryReport[id];
        let string = '';
        if (injured.practice == '') {
          string = 'NO INFO';
        } else {
          string = injured.practice;
        }
        
        if (injured.injury != '') {
          string = string.concat(': ' + injured.injury);
        }
        
        if (injured.status != '') {
          string = string + ' (' + injured.status.toUpperCase() + ')';
        }
        noteComments.push([string]);
      } else {
        noteComments.push(['']);
      }
    }
    noteRange.setNotes(noteComments);

    createTriggers();

  }
}

// 2025 - Created by Ben Powers
// ben.powers.creative@gmail.com


