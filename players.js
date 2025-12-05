// WEEKLY FF PLAYERS - Updated 12.05.2025

// Function to run all player pool setup scripts (auto = 1 escapes prompt)
function playersRefresh(auto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let prompt;
  if ( auto != 1 ) {
    const ui = SpreadsheetApp.getUi();
    prompt = ui.alert(`🔄 PLAYERS REFRESH`,`Click okay to update the following:\n\n🔹 Sleeper players list\n🔹 Sleeper projections\n🔹 ESPN projections\n🔹 Fantasy Pros projections\n🔹 Player status\n🔹 Player notes`, ui.ButtonSet.OK_CANCEL)
  } else {
    prompt = 'OK';
  }
  
  if ( prompt == 'OK' ) {
    players();
    ss.toast(`All player data updated`,`🏃 PLAYERS UPDATED`);
    draftList();
    ss.toast(`Draft list created for eligible players`,`📋 DRAFT LIST UPDATED`);
    draftLobbyClean(ss);
    ss.toast(`Updated and prepped the draft lobby`,`🚪 DRAFT LOBBY READY`);
  } else {
    ss.toast(`Canceled update of players`,`❌ CANCELED`);
  }
}

function players(selectedWeek){
  let data = [];
  try {        
    // Gets spreadsheet and sheet (creates if not existing)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PLAYERS') || ss.insertSheet('PLAYERS');
    
    let year = null;
    
    const docProps = PropertiesService.getDocumentProperties();
    const matchupsString = docProps.getProperty('matchups');
    let matchups = {}, teams = null;
    if (matchupsString) {
      matchups = JSON.parse(matchupsString);
      if (!selectedWeek) {
        try {
          selectedWeek = matchups.week;
        } catch (err) {
          Logger.log(`⚠️ Error fetching the matchups data, please ensure matchups have been configured for at least 1 week of the season`);
          throw new Error(`⚠️ MATCHUP FETCH ERROR`,`Ensure you've configured at least one matchup to use for the week prior to deploying the draft setup!`, ui.ButtonSet.OK);
        }
      }
      year = matchups.year || '2025';
      teams = matchups[selectedWeek].teams;
      
    }
    ss.toast(`Fetching draftable players based on your selections from these teams: ${matchups[selectedWeek].teams.flat()}`,`✳️ FETCHING PLAYERS`)
    Logger.log(`✳️ Teams to Include: ${matchups[selectedWeek].teams.join(', ')}`);
    const matchupMap = new Map(teams.flatMap((team, i) => 
      i % 2 === 0 
        ? [[team, `@${teams[i + 1]}`], [teams[i + 1], team]]
        : []
    ));
    selectedWeek = parseInt(selectedWeek); 
    const rosterText = docProps.getProperty('roster') || [];
    let roster;
    if (!rosterText) {
      throw new Error(`⚠️ No roster variable created yet!`);
    } else {
      roster = JSON.parse(rosterText);
    }

    // Returns an array of viable positions, excluding FLX/SPFLX (e.g. ['QB','RB','WR','TE','K','DEF'])
    let positions = roster.filter(pos => pos.count > 0 && pos.pos !== 'FLX' && pos.pos !== 'SPFLX').map(pos => { return pos.pos });
    
    // Fetch JSON object from Sleeper's API
    let players = fetchFilteredSleeperPlayers(positions, teams)

    // Initial variables -- look at notes to customize
    let ppr = 0.5;
    try {
      ppr = parseInt(JSON.parse(docProps.getProperty('configuration')).ppr);
    } catch (err) {
      Logger.log(`⚠️ Issue fetching existing PPR format object, please ensure configuration has occurred`)
    }
     
    const format = ppr == 0.5 ? 'HALF' : ppr == 1.0 ? 'PPR' : ppr == 0 ? 'STANDARD' : 'HALF';
    const week = parseInt(JSON.parse(UrlFetchApp.fetch(SCOREBOARD)).week.number);

    selectedWeek = selectedWeek || week;
    const available = selectedWeek == week ? true : false;

    Logger.log(`➕ Positions to include: ${positions}`);

    // Fetch images from Fantasy Pros site with function below
    // let images = fantasyProsImages(positions);

    // Fetch previous scoring from Sleeper stats
    let score, previous, previousTwo;
    score = sleeperScoring(ppr,year,week);
    week > 1 ? previous = sleeperScoring(ppr,year,week-1) : null;
    week > 2 ? previousTwo = sleeperScoring(ppr,year,week-2) : null;

    if (positions.indexOf('RB') > -1) {
      positions.splice(positions.indexOf('RB')+1,0,'FB');
    }

    let projectionSources = [];
    let slpr = {};
    try{
      Logger.log('🟪 Fetching projections from Sleeper...');
      slpr = slprProjectionFetch(ppr,selectedWeek,year);
      Logger.log('✅ Sleeper Projections Done');
      projectionSources.push(`🟪 Sleeper Projections`);
    }
    catch (err){
      Logger.log(`⚠️ Sleeper Projections Failed ${err.stack}`)
    }
    let espn = {};
    let fp = {};
    let injury = {};
    try {
      Logger.log('🟥 Fetching projections and outlooks from ESPN...');
      const pprId = format == 'HALF' ? 2 : format == 'FULL' ? 3 : format == 'STANDARD' ? 1 : 2; // default to half
      espn = espnFetch(year, week, pprId, 500);
      Logger.log('✅ ESPN Projections Done');
      projectionSources.push(`🟥 ESPN Projections & Outlooks`);
    }
    catch (err){
      Logger.log('⚠️ ESPN Projections Failed: ' + err.message);
    }   
    if (available) { 
      try {
        Logger.log('🟦 Fetching projections from Fantasy Pros...');
        fp = fpProjectionFetch(ppr,players);
        Logger.log('✅ FantasyPros Projections Done');
        projectionSources.push(`🟦 FantasyPros Projections`)
      }
      catch (err){
        Logger.log('⚠️ FantasyPros Projections Failed: ' + err.message);
      }
      try {
        Logger.log('🟩 Fetching practice reports, injuries, and official injury designations...');
        injury = filterInjuries();
        Logger.log('✅ Injury and Practice Reports Done');
        projectionSources.push(`🟩 Pro Football Reference Injuries`)
      }
      catch (err){
        Logger.log('⚠️ Injury Fetching Failed: ' + err.message);
      }
    } else {
      Logger.log('⚠️ Current week and selected week are not the same, no FP projections, ESPN outlooks, or injury designations fetched.');
    }
    ss.toast(`Fetched from these sources:${projectionSources.join(', ')}`,`📊 FETCHED DATA:`,30)

    // Modify these as needed
    const dataPoints = {
      'player_id':{'width':50,'hide':false,'named_range':true},
      'full_name':{'width':170,'hide':false,'named_range':true},
      // 'last_name':{'width':100,'hide':false,'named_range':true},
      // 'first_name':{'width':100,'hide':false,'named_range':true},
      'team':{'width':50,'hide':false,'named_range':true},
      'opp':{'width':50,'hide':false,'named_range':true}, // Manually created to show opponent, either with an "@" or only abbreviation depending on location
      // 'height':{'width':50,'hide':false,'named_range':true},
      // 'weight':{'width':50,'hide':false,'named_range':true},
      // 'age':{'width':50,'hide':false,'named_range':true},
      // 'birth_date':{'width':50,'hide':false,'named_range':true},
      // 'years_exp':{'width':50,'hide':false,'named_range':true},
      // 'position':{'width':50,'hide':true,'named_range':false},
      'fantasy_positions':{'width':50,'hide':false,'named_range':true},
      // 'depth_chart_position':{'width':50,'hide':true,'named_range':false},
      'depth_chart_order':{'width':50,'hide':false,'named_range':true},
      // 'number':{'width':50,'hide':true,'named_range':false},
      // 'college':{'width':50,'hide':true,'named_range':false},
      // 'status':{'width':50,'hide':true,'named_range':false},
      // 'active':{'width':50,'hide':true,'named_range':false},
      'espn_id':{'width':50,'hide':false,'named_range':true},
      // 'yahoo_id':{'width':50,'hide':false,'named_range':true},
      // 'rotowire_id':{'width':50,'hide':true,'named_range':false},
      // 'rotoworld_id':{'width':50,'hide':true,'named_range':false},
      // 'fantasy_data_id':{'width':50,'hide':false,'named_range':false},
      // 'gsis_id':{'width':50,'hide':true,'named_range':false},
      // 'sportradar_id':{'width':50,'hide':true,'named_range':false},
      // 'stats_id':{'width':50,'hide':true,'named_range':false},
      // 'news_updated':{'width':50,'hide':true,'named_range':false}
      'proj':{'width':50,'hide':false,'named_range':true},
      'proj_espn':{'width':50,'hide':false,'named_range':true},
      'proj_fp':{'width':50,'hide':false,'named_range':true},
      // 'proj_ff':{'width':50,'hide':false,'named_range':true},
      'previous':{'width':50,'hide':false,'named_range':true},
      'score':{'width':50,'hide':false,'named_range':true},
      'injury_status':{'width':50,'hide':false,'named_range':true},
      // 'injury_start_date':{'width':50,'hide':true,'named_range':false},
      // 'injury_body_part':{'width':50,'hide':true,'named_range':false},
      // 'injury_notes':{'width':50,'hide':true,'named_range':false},
      'injury':{'width':300,'hide':false,'named_range':true},
      'image':{'width':500,'hide':true,'named_range':true},
      'outlook':{'width':1000,'hide':true,'named_range':true}
    };
    
    // Defense ESPN IDs (Sleeper API lacks these)
    const espnIds = {
      'ARI':-16022,'ATL':-16001,'BAL':-16033,'BUF':-16002,'CAR':-16029,'CHI':-16003,'CIN':-16004,'CLE':-16005,'DAL':-16006,'DEN':-16007,'DET':-16008,'GB':-16009,
      'HOU':-16034,'IND':-16011,'JAX':-16030,'KC':-16012,'LV':-16013,'LAC':-16024,'LAR':-16014,'MIA':-16015,'MIN':-16016,'NE':-16017,'NO':-16018,'NYG':-16019,
      'NYJ':-16020,'PHI':-16021,'PIT':-16023,'SF':-16025,'SEA':-16026,'TB':-16027,'TEN':-16010,'WAS':-16028
    };
    // Injury status shorthand for easier representation in cells
    const injuries = {
      'Questionable':'Q',
      'Doubtful':'D',
      'Out':'O',
      'IR':'IR',
      'PUP':'PUP',
      'COV':'COV',
      'NA':'NA',
      'Sus':'SUS',
      'DNR':'DNR'
    };
    
    // Creates an array of the header values to use
    let headers = [];
    for (let a = 0; a < Object.keys(dataPoints).length; a++){
      headers.push(Object.keys(dataPoints)[a]);
    }
    
    // Sets the header values to the first row of the array 'keys' to be written to the sheet
    data.push(headers);
    // Loops through all 'key' entries (players) in the JSON object that was fetched
    let playerRow = [];
    Object.keys(players).forEach(key => {
      playerRow = [];
      try {
        const id = key;
        const player = players[id];
        let name = player.first_name + ' ' + player.last_name;
        for ( let col = 0; col < Object.keys(dataPoints).length; col++ ) {
          const dataPoint = Object.keys(dataPoints)[col];
          switch (dataPoint) {
            case 'full_name':
              // Creates the full name entry alongside the first/last entries in the JSON data
              playerRow.push(name);
              break;
            case 'espn_id':
              if ( player.position == 'DEF' ) {
                playerRow.push(espnIds[id] || '');
              } else {
                playerRow.push(espnId(id) || '');
              }
              break;
            case 'injury_status':
              if ( player[dataPoint] == null && injury[id] == null) {
                // Pushes a 'G' for 'good' to any player without an injury tag
                playerRow.push('G');
              } else if (injury[id] != null) {
                if (injury[id].status == '') {
                  // If the player is listed on the injury report at all, it will default to giving that player a 'Q'
                  playerRow.push('Q');
                } else {
                  playerRow.push(injury[id].status.charAt(0));
                }
              } else {
                // If player has injury designation, assigns the shorthand to that player
                playerRow.push(injuries[player[dataPoint]]);
              }
              break;
            case 'proj':
              playerRow.push(slpr[id]);
              break;
            case 'proj_espn' :
              try {
                playerRow.push(espn[id].points);
              }
              catch (err) {
                Logger.log(`❗ No ESPN Projection for ${name}`);
                try {
                  let avg = [];
                  slpr[id] != null ? avg.push(slpr[id]) : null;
                  fp[id] != null ? avg.push(fp[id]) : null;
                  // projFF[id] != null ? avg.push(projFF[id]) : null;
                  let sum = 0;
                  for (let p = 0; p < avg.length; p++) {
                    sum = parseFloat(sum) + parseFloat(avg[p]);
                  }
                  playerRow.push((sum/avg.length).toFixed(2));
                }
                catch (err) {
                  Logger.log(`❗ Missing either Sleeper or FantasyPros projection for ${name}`);
                  try {
                    playerRow.push(slpr[id]);
                  }
                  catch (err) {
                    Logger.log(`❗ Missing Sleeper projection for ${name}`)
                    try {
                      playerRow.push(fp[id])
                    }
                    catch (err) {
                      Logger.log(`❗ No player projections available for ${name}`);
                      playerRow.push('');
                    }
                  }
                }
              } 
              break;
            case 'proj_fp':
              if (available) {
                playerRow.push(fp[id]);
              } else {
                playerRow.push('');
              }
              break;
            case 'previous':
              if(previous[id] == null || previous[id] == undefined) {
                if(previousTwo[id] == null == previousTwo[id] == undefined) {
                  playerRow.push('NA');
                } else {
                  playerRow.push(previousTwo[id]);
                }
              } else {
                playerRow.push(previous[id]);
              }
              break;
            case 'score':
              if(score[id] == null) {
                playerRow.push('');
              } else {
                playerRow.push(score[id])
              }
              break;
            case 'image':
              if (player.position == 'DEF') {
                playerRow.push(`https://sleepercdn.com/images/team_logos/nfl/${id.toLowerCase()}.png`);
              } else {
                playerRow.push(`https://sleepercdn.com/content/nfl/players/thumb/${id}.jpg`);
              }
              break;
              // could use checking, this is a placeholder image: 'https://images.fantasypros.com/images/players/nfl/missing/headshot/210x210.webp'
            case 'injury':
              if(injury[id] == null) {
                playerRow.push('');
              } else {
                let injured = injury[id];
                let string = '';
                if (injured.practice == '') {
                  string = 'NO INFO';
                } else {
                  string = injured.practice;
                }
                
                if (injured.injury != '') {
                  string = string.concat(`: ${injured.injury}`);
                }
                
                if (injured.status != '') {
                  string += ` (${injured.status.toUpperCase()})`;
                }
                
                playerRow.push(string);
              }
              break;
            case 'depth_chart_order':
              if (player.position == 'DEF') {
                playerRow.push(1);
              } else {
                playerRow.push(player[dataPoint]);
              }
              break;
            case 'outlook':
              try {
                playerRow.push(espn[key].outlook);
              }
              catch (err) {
                playerRow.push('');
              }
              break;
            case 'opp': 
              try {
                playerRow.push(matchupMap.get(player.team));
              }
              catch (err) {
                playerRow.push('');
              }
              break;
            default:
              playerRow.push(player[dataPoint]);
          }
        }
        // so long as the array mapped values, it pushes the array into the array ('data') of arrays
        data.push(playerRow);
        // resets the 'playerRow' variable to start over
        playerRow = [];
      } catch (err) {
        Logger.log(`${name}: ${err.stack}`)
        ss.toast(`Error bringing in data`,`⚠️ ERROR`);
      }
    });

    // Clear the sheet for new data
    sheet.clear();
    // Gets range for setting data and headers
    let playerTable = sheet.getRange(1,1,data.length,data[0].length);
    // Sets data in place
    playerTable.setValues(data);
    // Sorts based on 
    sheet.getRange(2,1,data.length-1,data[0].length).sort([{column: headers.indexOf('proj')+1, ascending: false}]);
    
    // Creates named ranges for doing VLOOKUP functions in Google Sheets; only for keys in 'headers' object tagged with 'true' for 'named_range', also checks for column widths and hidden status
    for (let col = 0; col < Object.keys(dataPoints).length; col++ ) {
      const dataPoint = Object.keys(dataPoints)[col];
      if (dataPoints[dataPoint].named_range) {
        ss.setNamedRange(`SLPR_${headers[col].toUpperCase()}`,sheet.getRange(2,col+1,data.length-1,1));
      }
      sheet.setColumnWidth(col+1,dataPoints[dataPoint].width);
      if (dataPoints[dataPoint].hide){
        sheet.hideColumns(col+1,1);
      } else {
        sheet.unhideColumn(sheet.getRange(1,col+1,sheet.getMaxRows(),1));
      }
    }
    
    // Notification text creation
    let positionsString = '';
    if (positions.indexOf('FB') >= 0) {
      positions.splice(positions.indexOf('FB'),1);
    }
    for (let a = 0; a < positions.length; a++) {
      if (positions[a+2] == undefined) {
        positionsString = positionsString.concat(` and ${positions[a]}`);
      } else {
        positionsString = positionsString.concat(`${positions[a]}, `);
      }
    }
    ss.toast(`All Sleeper player information placed successfully for ${positionsString}`,`👥 PLAYERS IMPORTED`);

    // Update for correct rows
    let rows = data.length;
    adjustRows(sheet,rows);

    // Update for correct columns
    let columns = data[0].length;
    adjustColumns(sheet,columns);

    let alignments = sheet.getRange(1,1,data.length,data[0].length);
    alignments.setHorizontalAlignment('left');

    // Locks data on sheet
    sheet.protect(); 
  }
  catch (err) {
    Logger.log(`DATA STATE: ${data}`);
    Logger.log(`⚠️ Issue with players fetching | ${err.stack}`);
    throw new Error(`⚠️ Issue fetching players: ${err.message}`);
  }
}

//-------------------------------------------------------------
// OPTIMIZED SLEEPER PLAYER FETCH WITH FILTERING
function fetchFilteredSleeperPlayers(positions, teams) {
  positions = positions || ['QB', 'RB', 'WR', 'TE', 'K', 'DEF'];
  teams = teams || null; // null = all teams
  
  const url = 'https://api.sleeper.app/v1/players/nfl';
  let response, allPlayers;
  
  try {
    response = UrlFetchApp.fetch(url);
    allPlayers = JSON.parse(response.getContentText());
  } catch (err) {
    Logger.log('⚠️ Error fetching Sleeper data: ' + err);
    return {};
  }
  
  let filteredPlayers = {};
  let skipped = { inactive: 0, wrongTeam: 0, wrongPosition: 0 };
  
  // Filter during iteration (most efficient)
  for (const playerId in allPlayers) {
    const player = allPlayers[playerId];
    
    // Skip inactive players
    if (player.active === false || player.status === 'Inactive') {
      skipped.inactive++;
      continue;
    }
    // player['status'] == 'Active' || player.position == 'DEF'
    // Skip wrong positions
    if (!positions.includes(player.position)) {
      skipped.wrongPosition++;
      continue;
    }
    
    // Skip wrong teams (if teams filter provided)
    if (teams && !teams.includes(player.team)) {
      skipped.wrongTeam++;
      continue;
    }
    
    // Player passed all filters
    filteredPlayers[playerId] = player;
  }
  
  Logger.log(`✅ Filtered ${Object.keys(filteredPlayers).length} active players from Sleeper API endpoint`);
  Logger.log(`❌ Skipped: ${skipped.inactive} inactive, ${skipped.wrongTeam} wrong team, ${skipped.wrongPosition} wrong position`);
  
  return filteredPlayers;
}

//-------------------------------------------------------------
// SLEEPER League-Specific Projections
function slprProjectionFetch(ppr,week,year) {
  const scoring = leagueScoring(ppr);
  if (!week || !year) {
    const docProps = PropertiesService.getDocumentProperties();
    const matchupsString = docProps.getProperty('matchups');
    if (matchupsString) {
      const matchups = JSON.parse(matchupsString);
      week = matchups.week || 1;
      year = matchups.year || 2025;
    } else {
      week = week || 1;
      year = year || 2025;
    }
  }

  const obj = JSON.parse(UrlFetchApp.fetch(`https://api.sleeper.app/projections/nfl/${year}/${week}?season_type=regular&position[]=DEF&position[]=FLEX&position[]=QB&position[]=RB&position[]=TE&position[]=WR&position[]=K&order_by=player_id`));
  const players = obj.length;
  let a;
  let score;
  let slprProjections = {};
  for (a = 0; a < players; a++) {
    score = 0;
    let id = obj[a].player_id;
    for (let keys in scoring) {
      if ( isNaN(parseFloat(obj[a].stats[keys])*parseFloat(scoring[keys])) == false && parseFloat(obj[a].stats[keys]) != null ) {
        score = score + parseFloat(obj[a].stats[keys])*parseFloat(scoring[keys]);
      }
    }
    slprProjections[id] = score.toFixed(1); // Round to the 10ths
  }
  // Logger.log(slprProjections);
  return slprProjections;
}

function testESPN() {
  Logger.log(JSON.stringify(espnFetch(2025,13,3,)))
}

//-------------------------------------------------------------
// FUNCTION TO FETCH OBJECT OF ESPN PROJECTIONS
function espnFetch(year, week, pprId, limit) {
  let half = true;
  year = year || 2025;
  week = week || 12;
  if (pprId == 2) {
    pprId = 1;
  } else if (pprId) {
    half = false;
  } else {
    pprId = 3;
  }
  limit = limit || 1000;

  const injuryStatuses = {
    "ACTIVE": "G",
    "SUSPENSION": "S",
    "QUESTIONABLE": "Q",
    "OUT": "O",
    "INJURY_RESERVE": "IR",
    "DOUBTFUL":"D",
    "null": "NA",
  }
  // ESPN PROJECTION FETCHING
  let espnProjections = {};
  let data = {};  
  const projectionId = '11' + year + week;
  const url = `https://lm-api-reads.fantasy.espn.com/apis/v3/games/ffl/seasons/${year.toFixed(0)}/segments/0/leaguedefaults/${pprId.toFixed()}?view=kona_player_info`;
  
  const headers = {
    'X-Fantasy-Filter': JSON.stringify({
      "players": {
        "limit": limit,
        "sortPercOwned": {
          "sortPriority": 4,
          "sortAsc": false
        }
      }
    })
  };
  
  const options = {
    'method': 'get',
    'headers': headers,
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    data = JSON.parse(response.getContentText()).players;
  } catch (err) {
    Logger.log('⚠️ Error fetching data: ' + err.stack);
    return [];
  }
  
  let bye = [], error = [], injuryOptions = [];
  for (let a = 0; a < data.length; a++) {
    pos = espnPositionId[data[a].player.defaultPositionId];
    if (pos == "QB" || pos == "RB" || pos == "WR" || pos == "TE" || pos == "DEF" || pos == "K"){
      team = espnProTeamId[data[a].player.proTeamId];
      idEspn = data[a].player.id;
      name = data[a].player.fullName.replace(/D\/ST/g,'defense');    
      if (pos === 'DEF'){
        id = team;
      } else {
        id = nameFinder(name);
      }
      if (id != null) {
        if (data[a].player.stats.some(item => item.id === projectionId)) {
          const proj = (data[a].player.stats).find(x => x.id === ('11'+year+week));
          if (proj) {
            const rec = parseFloat(proj.stats["53"]) || 0; // Identifier for receptions
            let points;
            if (half) { // Rounded to the nearest, ignored if full ppr or non, since API returns those two
              points = parseInt(10*(proj.appliedTotal+(rec*0.5)))/10; // Rounding to 10ths
            } else {
              points = parseInt(10*proj.appliedTotal)/10; // Rounding to 10ths
            }
            
            // Get injury if available
            let injury = null;
            if (data[a].player.injuryStatus && data[a].player.injuryStatus !== 'ACTIVE') {
              injury = injuryStatuses[data[a].player.injuryStatus];
            }

            // Get outlook if available
            let outlook = null;
            if (data[a].player.outlooks && data[a].player.outlooks.outlooksByWeek[week]) {
              outlook = data[a].player.outlooks.outlooksByWeek[week];
            }
            
            // Build the object with points and injury
            espnProjections[id] = {
              points: points
            };
            
            espnProjections[id].injured = data[a].player.injured;

            // Only add injury if it exists
            if (injury) {
              espnProjections[id].injury = injury;
            }

            // Only add outlook if it exists
            if (outlook) {
              espnProjections[id].outlook = outlook;
            }
          } else {
            error.push(name);
          }
        } else {
          bye.push(name);
        }
      }
    }
  }
  if (bye.length > 0) {
    Logger.log('❔ ESPN - Likely on bye: ' + bye);
  }
  if (error.length > 0) {
    Logger.log('⚠️ ESPN - Error getting projection: ' + error);
  }
  // Logger.log(JSON.stringify(espnProjections));
  return espnProjections;
}

//-------------------------------------------------------------
// FUNCTION TO FETCH OBJECT OF FP PROJECTIONS
function fpProjectionFetch(ppr,playersObj) {
  format = ppr == 1.0 ? 'PPR' : ppr == 0.5 ? 'HALF' : ppr == 0.0 ? null : 'HALF';
  playersObj = playersObj || JSON.parse(UrlFetchApp.fetch('https://api.sleeper.app/v1/players/nfl'));
  let playersNames = [];
  let playersIds = [];
  for (let key in playersObj){
    if ( playersObj[key].fantasy_positions == 'DEF'){
      playersNames.push(`${playersObj[key].first_name} ${playersObj[key].last_name}`);
      playersIds.push(playersObj[key].player_id);
    } else if ( playersObj[key].fantasy_positions == 'QB' ||  playersObj[key].fantasy_positions == 'RB' ||  playersObj[key].fantasy_positions == 'WR' ||  playersObj[key].fantasy_positions == 'TE' ||  playersObj[key].fantasy_positions == 'K') {
      playersNames.push(playersObj[key].first_name + ' ' + playersObj[key].last_name);
      playersIds.push(playersObj[key].player_id);
    }    
  }
  let obj = {};
  let url, table, count, output = [], values = [], missed = [], name, id, points, len;
  let baseUrl = 'https://www.fantasypros.com/nfl/projections';
  let positionList = ['qb', 'rb', 'wr', 'te', 'dst', 'k'];
  for (let b = 0 ; b < positionList.length ; b++) {
    //https://www.fantasypros.com/nfl/projections/qb.php
    url = (baseUrl+'/'+positionList[b]+'.php');
    if (format && ['rb','wr','te'].indexOf(positionList[b]) >= 0) url += '?scoring=' + format; // Only needed if RB/WR/TE and half or full PPR
    table = UrlFetchApp.fetch(url).getContentText();
    table = table.substring(table.indexOf('<table cellpadding="0" cellspacing="0" border="0" id="data"') - 1).split('</table>')[0].split('<tbody>')[1].split('</tbody>')[0].split('<tr class=').slice(1);
    
    count = table.length;
    for (let c = 0 ; c < count ; c++) {
      arr = [];
      id = table[c].match(/fp\-id\-[0-9]+/g);
      id = id[0].substring(6);
      values = table[c].split('<td class');
      len = values.length;
      values = values[len-1].split('"');
      len = values.length;
      points = parseFloat(values[len-1].match(/[0-9\.]+/g)[0]).toFixed(1); // Round to 10ths
      name = table[c].split('fp-player-name="')[1].split('</a>')[0];
      name = name.split('">')[1];
      
      let skip = false;
      nameFound = nameFinder(name);
      if ( nameFound == null ) {
          missed.push(name);
          skip = true;
      }
      if (positionList[b] != 'dst'){
        nameFound = parseInt(nameFound);
      }
      if ( skip == false) {
        obj[nameFound] = points;
      }
    }
  }
    
  if ( output.length > 0 ) {
    Logger.log(`❗ These players missed: ${output}`);
  } else {
    Logger.log(`✅ All FP Players matched`);
  }
  return obj;
}

// 2025 - Created by Ben Powers
// ben.powers.creative@gmail.com
