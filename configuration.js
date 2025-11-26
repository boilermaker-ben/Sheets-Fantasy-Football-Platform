// WEEKLY FF DRAFT - Updated 11.25.2025

// Toolbar function that dynamically builds the spreadsheet menu based on the current state of the draft setup process.
function toolbar() {
  const ui = SpreadsheetApp.getUi();
  const docProps = PropertiesService.getDocumentProperties();
  const props = docProps.getProperties(); // Get all properties at once for efficiency
  // --- Check the state of the setup process ---
  // Safely parse properties, providing null/empty objects as defaults
  const members = props.members ? JSON.parse(props.members) : null;
  const matchups = props.matchups ? JSON.parse(props.matchups) : null;
  const roster = props.roster ? JSON.parse(props.roster) : null;
  const configuration = props.configuration ? JSON.parse(props.configuration) : null;
  // These simple flags are easier to manage than digging into a nested object
  let isDraftReady = false;
  let isDraftComplete = false;
  let isDrafting = false;
  let isComplete = false;
  if (configuration) {
    isDraftReady = configuration.draftReady === true;
    isDraftComplete = configuration.draftComplete === true;
    isDrafting = configuration.drafting === true;
    isComplete = configuration.complete === true;
  }
  // --- Build the Menu ---
  const menu = ui.createMenu('🏈 Football Tools');
  
  if (isComplete) {
    menu.addItem('🏆 Draft Recap','recapPanel');
  } else if (isDraftComplete) {
    // --- STAGE 3: SCORING ---
    // The draft is over, only show scoring-related items.
    scoringShow();
    menu.addSubMenu(ui.createMenu('🔢 Scoring')
      .addItem('📡 Get Scores', 'sleeperScoringLogging')
      .addItem('🟢 Live Scoring ON', 'sleeperLiveScoringOn')
      .addItem('🔴 Live Scoring OFF', 'sleeperLiveScoringOff'));
  } else if (isDraftReady) {
    // --- STAGE 2: DRAFTING ---
    // Setup is done, show the drafting tools.
    menu.addItem(`🔄 Refresh Players`,`playersRefresh`);
    scoringHide(); // Call your existing UI function
    if (isDrafting) {
      // The draft has been started, show trigger management
      menu.addSubMenu(ui.createMenu('📝 Drafting')
        .addItem('✔️ Enable Triggers', 'triggersDrafting')
        .addItem('❌ Disable Triggers', 'deleteTriggers')
        .addItem('🛑 Stop Drafting', 'stopDrafting'));
    } else {
      menu.addSubMenu(ui.createMenu('📝 Drafting')
        .addItem('▶️ Start Draft', 'startDrafting'));
    }    
  }

  // --- STAGE 1: SETUP ---
  // The setup is not yet complete. Reveal items sequentially.
  scoringHide();
  
  if (!isDrafting) {
    // 1. Always show Member Manager as the first step
    if (members && members.length > 0) {
      menu.addItem('👥 Member Manager ✔️', 'manageMembers');
    } else {
      menu.addItem('👥 Member Manager', 'manageMembers');
    }

    const hasGames = matchups && Object.values(matchups).some(weekData => weekData.length > 0);

    // 2. If members exist, show Matchup Selection
    if (members && members.length > 0) {
      if (hasGames) {
        menu.addItem('📆 Matchup Selection ✔️', 'manageMatchups');
      } else if (members && members.length > 0) {
        menu.addItem('📆 Matchup Selection', 'manageMatchups');
      }
    }
    
    // 3. If members AND matchups exist, show Roster Configuration
    // (We check if any games have been selected in any week)
    if (members && members.length > 0 && hasGames) {
      if (roster) {
        menu.addItem('🧮 Roster Configuration ✔️', 'manageRoster');
      } else {
        menu.addItem('🧮 Roster Configuration', 'manageRoster');
      }
    }

    // 4. If all previous steps are done, show the final setup confirmation
    if (members && members.length > 0 && hasGames && roster && configuration && !isDraftComplete && !isDraftReady && !isComplete) {
      menu.addItem(`🔽 Confirm and Deploy`,'draftSetup');
    } else {
      menu.addItem('⚙️ Configuration Panel', 'draftSetup');
    }
  }
  if (isDraftComplete) {
    menu.addItem('♻️ Reset Draft', 'resetDraft');
  }
  // --- Add common items and build the final menu ---
  menu.addSeparator()
      .addItem('❔ Help & Support', 'showSupportDialog')
      .addToUi();
}

//------------------------------------------------------------------------
// FETCH CURRENT WEEK
function fetchWeek() {
  let selectedWeek = null;
  const matchupsString = PropertiesService.getDocumentProperties().getProperty('matchups');
  let matchups = {};
  if (matchupsString) {
    matchups = JSON.parse(matchupsString);
    if (!selectedWeek) {
      try {
        selectedWeek = matchups.week;
        Logger.log(`✅ Found selected week of ${selectedWeek}`);
      } catch (err) {
        Logger.log(`⚠️ Error fetching the week, please run matchup selection to ensure you've got a week specified`);
      }
    }
  }
  return selectedWeek;
}

// FETCH CURRENTLY SELECTED YEAR
function fetchYear() {
  let selectedYear = null;
  const matchupsString = PropertiesService.getDocumentProperties().getProperty('matchups');
  let matchups = {};
  if (matchupsString) {
    matchups = JSON.parse(matchupsString);
    if (!selectedYear) {
      try {
        selectedYear = matchups.year;
        Logger.log(`✅ Found selected year of ${selectedYear}`);
      } catch (err) {
        Logger.log(`⚠️ Error fetching the year, please run matchup selection to ensure you've got a year specified`);
      }
    }
  }
  return selectedYear;
}

// FETCH FORMAT
function fetchFormat() {
  const configuration = PropertiesService.getDocumentProperties().getProperty('configuration');
  let format = 'HALF';
  if (configuration) {
    const pprString = JSON.parse(configuration).ppr;
    if (pprString) {
      const ppr = parseInt(pprString);
      format = ppr === 1 ? 'FULL' : ppr === 0 ? 'STANDARD' : format;
    }
  }
  return format;
}

// FETCH FORMAT
function fetchDraftComplete() {
  const configuration = PropertiesService.getDocumentProperties().getProperty('configuration');
  let draftComplete = false;
  if (configuration) {
    draftComplete = JSON.parse(configuration).draftComplete;
  }
  return draftComplete;
}

// This is the set of scripts used to configure the daily pool and prep it before kickoff.
/**
 * Creates and displays the HTML modal dialog for member management.
 */
function manageMembers() {
  const html = HtmlService.createHtmlOutputFromFile('manageMembers')
      .setWidth(600) // A bit wider to accommodate new columns
      .setHeight(700); // Initial height, the script will adjust it
  SpreadsheetApp.getUi().showModalDialog(html, 'Member Manager');
}

/**
 * Retrieves the current member list for the panel.
 * For now, this returns dummy data. Later, you can have it read
 * from PropertiesService or a spreadsheet.
 *
 * @returns {Array<Object>} An array of member objects.
 */
function fetchMembers() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const members = docProps.getProperty('members') ? JSON.parse(docProps.getProperty('members')) : [];
    const settings = docProps.getProperty('configuration') ? JSON.parse(docProps.getProperty('configuration')) : {};
    if (members && settings) {
      return { members, settings };
    } else if (members) {
      return { 
        members: members,
        setting: {}
      }
    } else {
      return {
        members: [],
        settings: settings
      };
    }
  } catch (err) {
    Logger.log('Error fetching member data: ' + err.stack);
    return {
      members: [],
      settings: {}
    };
  }
}

/**
 * Receives the final member list from the panel and saves it.
 * @param {Array<Object>} memberData The array of member objects to save.
 */
function saveMembers(submissionData) {
  try {
    Logger.log(`Received data: ${JSON.stringify(submissionData)}`);
    const docProps = PropertiesService.getDocumentProperties();
    const { members, settings } = submissionData;
    
    // Save members and settings to their respective properties
    docProps.setProperty('members', JSON.stringify(members));
    docProps.setProperty('configuration', JSON.stringify(settings));
    
    console.log(`Saved ${members.length} members with configuration settings:`, settings);
    
    return { success: true };

  } catch (err) {
    Logger.log('⚠️ Error saving member data: ' + err.message);
    throw new Error(`⚠️ ERROR: ${err.stack}`);
  }
}

function saveMembersSuccess() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const ui = SpreadsheetApp.getUi();
    if (!docProps.getProperty('members')) {
      ui.alert(`⚠️ ISSUE WITH MEMBERS`,`There was a problem recording your members.
      
      Please try running "👥 Member Manager"
      again from the "🏈 Football Tools" menu`,ui.ButtonSet.OK);
      Logger.log(`⚠️ Issue with members when recording.`)
      toolbar();
    } else if (!docProps.getProperty('matchups')) {
      ui.alert(`2️⃣ GAME SELECTION IS NEXT`,`Now run the "📆 Matchup Selection" function
      from the "🏈 Football Tools" menu`,ui.ButtonSet.OK);
      Logger.log(`👥 Members recorded successfully, prompted for matchups.`)
      toolbar();
    } else {
      Logger.log(`🔄 Updated members after previously initializing.`)
    }
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error with 'saveMemberSuccess: ${err.stack}`,`⚠️ ERROR`)
    Logger.log(`⚠️ Error with 'saveMemberSuccess: ${err.stack}`);
  }
}



function manageMatchups() {
  const html = HtmlService.createHtmlOutputFromFile('manageMatchups')
      .setWidth(600)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Matchup Selection');
}

/**
 * Fetches schedule data from the ESPN API. If no week is provided, it
 * automatically determines the current week from the API. It also retrieves
 * any previously saved game selections for that week.
 *
 * @param {number} week The optional NFL week to fetch.
 * @returns {object} An object containing the schedule, selected game IDs, and the week number used.
 */
function fetchScheduleData(week) {
  try {
    let weekToFetch = week;
    // 0. If no week provided and variable set, use it
    if (!weekToFetch) {
      try {
        const docProps = PropertiesService.getDocumentProperties();
        // Get the existing selections object, or create a new one
        const matchups = JSON.parse(docProps.getProperty('matchups') || '{}');
        weekToFetch = matchups.week;
      }
      catch (err) {
        Logger.log(`No existing document property for 'matchups' to refer to... moving on`);
      }
    }

    // 1. If no week is provided, find the current week from the base API endpoint
    
    const currentJson = JSON.parse(UrlFetchApp.fetch(SCOREBOARD, { 'muteHttpExceptions': true }).getContentText());
    const currentWeek = currentJson.week.number;
    let year = currentJson.season.year;

    if (!weekToFetch) { 
      weekToFetch = currentWeek;
    }

    // 2. Fetch the schedule for the determined week
    let schedule = {};
    const apiUrl = `${SCOREBOARD}?week=${weekToFetch}`;
    const json = JSON.parse(UrlFetchApp.fetch(apiUrl, { 'muteHttpExceptions': true }).getContentText());
    if (json.events.length === 0) {
      Logger.log(`⭕ Unpopulated post-season week selected, reverting to the current week`);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Unpopulated post-season week selected, reverting to the current week`,`⭕ INVALID WEEK SELECTED`);
      weekToFetch = currentWeek;
      schedule = currentJson.events.map(event => {
        const homeTeam = event.competitions[0].competitors.find(c => c.homeAway === 'home');
        const awayTeam = event.competitions[0].competitors.find(c => c.homeAway === 'away');
        return {
          id: event.id,
          date: event.date,
          homeTeamAbbr: homeTeam.team.abbreviation,
          homeTeamLogo: homeTeam.team.logo,
          awayTeamAbbr: awayTeam.team.abbreviation,
          awayTeamLogo: awayTeam.team.logo
        };
      });
    } else {
      year = json.season.year;
      schedule = json.events.map(event => {
        const homeTeam = event.competitions[0].competitors.find(c => c.homeAway === 'home');
        const awayTeam = event.competitions[0].competitors.find(c => c.homeAway === 'away');
        return {
          id: event.id,
          date: event.date,
          homeTeamAbbr: homeTeam.team.abbreviation,
          homeTeamLogo: homeTeam.team.logo,
          awayTeamAbbr: awayTeam.team.abbreviation,
          awayTeamLogo: awayTeam.team.logo
        };
      });
    }
    
    // 3. Fetch previously saved selections
    const docProps = PropertiesService.getDocumentProperties();
    const allSelections = JSON.parse(docProps.getProperty('matchups')) || {};
    let selectedIds = [];
    if (allSelections) {
      if (allSelections[weekToFetch]) {
        selectedIds = allSelections[weekToFetch].ids;
      }
    }

    // 4. Return all data, including the week number that was used
    return { schedule, selectedIds, year, week: weekToFetch, minWeek: currentWeek };

  } catch (err) {
    Logger.log(`⚠️ Failed to fetch schedule for ${week}. Error: ${err.toString()}`);
    throw new Error(`⚠️ Error with schedule pulling function | ${err.stack}`);
  }
}


/**
 * Saves the user's selected game IDs for a specific week.
 * @param {object} data An object containing the week and an array of gameIds.
 */
function saveMatchups(data) {
  try {
    const { week, year, gameIds, teams } = data;
    const docProps = PropertiesService.getDocumentProperties();
    Logger.log(data);
    // Get the existing selections object, or create a new one
    const allSelections = JSON.parse(docProps.getProperty('matchups') || '{}');
    
    // Update the selections for the specific week
    allSelections[week] = allSelections[week] || {};
    allSelections[week].ids = gameIds;
    allSelections[week].teams = teams;
    allSelections.week = week;
    allSelections.year = year;

    // Save the updated object back to properties
    docProps.setProperty('matchups', JSON.stringify(allSelections));

    console.log(`Saved ${gameIds.length} games for week ${week}.`);
    return { success: true };

  } catch (err) {
    Logger.log('Error saving game selections: ' + err.toString());
    throw new Error('An error occurred while trying to save your selections.');
  }
}

function saveMatchupsSuccess() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const ui = SpreadsheetApp.getUi();
    if (!docProps.getProperty('matchups')) {
      ui.alert(`⚠️ ISSUE WITH MATCHUPS`,`There was a problem recording your matchup selections.
      
      Please try running "📆 Matchup Selection" again
      from the "🏈 Football Tools" menu`,ui.ButtonSet.OK);
      Logger.log(`⚠️ Issue with matchups when recording.`);
      toolbar();
    } else if (!docProps.getProperty('roster')) {
      ui.alert(`3️⃣ ROSTER CREATION IS NEXT`,`Now run the "🧮 Roster Configuration" function
      from the "🏈 Football Tools" menu`,ui.ButtonSet.OK);
      Logger.log(`📆 Matchups recorded successfully, prompted for matchups.`);
      toolbar();
    } else {
      Logger.log(`🔄 Updated matchups after previously initializing.`);
    }
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error with 'saveMatchupsSuccess: ${err.stack}`,`⚠️ ERROR`);
    Logger.log(`⚠️ Error with 'saveMatchupsSuccess: ${err.stack}`);
  }
}



function manageRoster() {
  const html = HtmlService.createHtmlOutputFromFile('manageRoster')
      .setWidth(700)
      .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Roster Configuration');
}

/**
 * Fetches the initial configuration data for the roster setup tool.
 * It loads the saved roster state and calculates the number of participants
 * from the saved member list.
 * @returns {object} An object containing the saved roster state, participants, and games selected.
 */
function fetchRoster() {
  let roster = [];
  let members = 6; // assumed
  let membersStored = false;
  let week = null;
  let matchups = 3;
  let matchupsStored = false;
  let ppr = '0.5';
  

  try {
    const docProps = PropertiesService.getDocumentProperties();

    try {
      roster = JSON.parse(docProps.getProperty('roster'));
    } catch (err) {
      Logger.log(`No 'roster' data stored'`);
    }

    try {
      const savedMembers = JSON.parse(docProps.getProperty('members'));
      members = savedMembers.length;
      membersStored = true;
      Logger.log(`Found members and populating with ${members}`);
    } catch (err) {
      Logger.log(`No 'members' data recorded, assuming ${members}`);
    }

    try {
      const savedMatchups = JSON.parse(docProps.getProperty('matchups'));
      week = savedMatchups.week;
      if (week) {
        matchups = savedMatchups[week].length;
        matchupsStored = true;
      }
      Logger.log(`Found matchups and populating with ${matchups}`);
    } catch (err) {
      Logger.log(`No 'matchups' data recorded, assuming ${matchups}`);
    }

    try {
      const configuration = JSON.parse(docProps.getProperty('configuration'));
      ppr = configuration.ppr;
      Logger.log(`Found entry for ppr: ${ppr}`);
    } catch (err) {
      Logger.log(`No entry for PPR found`);
    }

    return {
      roster: roster, // This will be the full { roster, participants, gamesSelected } object from the last save
      members: members,
      membersStored: membersStored,
      matchups: matchups,
      matchupsStored: matchupsStored,
      week: week,
      ppr: ppr
    };

  } catch (err) {
    Logger.log('Error fetching roster configuration inputs: ' + err.stack);
    // Return a safe default if anything goes wrong
    return {
      roster: [],
      members: members,
      membersStored: membersStored,
      matchups: matchups,
      matchupsStored: matchupsStored,
      week: week,
      ppr: '0.5'
    };
  }
}

/**
 * Saves the final roster configuration from the panel to Document Properties.
 * @param {object} state The entire state object from the client { roster, participants, gamesSelected }.
 */
function saveRoster(state) {
  try {
    // We save the entire state object as a single JSON string.
    const docProps = PropertiesService.getDocumentProperties()
    docProps.setProperty('roster', JSON.stringify(state.roster));
    let configuration = {};
    if (docProps.getProperty('configuration')) {
      configuration = JSON.parse(docProps.getProperty('configuration'));
    }
    configuration.ppr = state.ppr;

    docProps.setProperty('configuration', JSON.stringify(configuration));

    console.log('Roster configuration saved successfully.');
    
    return { success: true }; // Send a success confirmation back to the client

  } catch (err) {
    console.error('Error in saveRoster: ' + err.toString());
    throw new Error('Failed to save the roster configuration.');
  }
}

function saveRosterSuccess() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const ui = SpreadsheetApp.getUi();
    if (!docProps.getProperty('roster')) {
      ui.alert(`⚠️ ISSUE WITH ROSTER`,`There was a problem recording your roster selections.
      
      Please try running "🧮 Roster Configuration" again
      from the "🏈 Football Tools" menu`,ui.ButtonSet.OK);
      Logger.log(`⚠️ Issue with matchups when recording.`);
      toolbar();
    } else if (!JSON.parse(docProps.getProperty('configuration')).draftReady) {
      ui.alert(`4️⃣ DEPLOY DRAFT IS NEXT`,`Now run the "🔽 Confirm and Deploy" function
      from the "🏈 Football Tools" menu`,ui.ButtonSet.OK);
      Logger.log(`🧮 Roster recorded successfully, prompted for draft deployment.`);
      toolbar();
    } else {
      Logger.log(`🔄 Updated roster after previously initializing.`);
    }
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error with 'saveMatchupsSuccess: ${err.stack}`,`⚠️ ERROR`);
    Logger.log(`⚠️ Error with 'saveMatchupsSuccess: ${err.stack}`);
  }
}

/**
 * Launches the final confirmation panel.
 */
function draftSetup() {
  const html = HtmlService.createHtmlOutputFromFile('reviewConfiguration')
      .setWidth(800)
      .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '🏈 Sheets Fantasy Football Configuration');
}

/**
 * Fetches all saved data from Document Properties and the required ESPN
 * data to populate the final confirmation panel.
 */
function fetchFinalConfirmationData() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    
    // 1. Load all properties with safe defaults
    const members = JSON.parse(docProps.getProperty('members') || '[]');
    const configuration = JSON.parse(docProps.getProperty('configuration') || '{}');
    const roster = JSON.parse(docProps.getProperty('roster') || '{}');
    const matchupsObj = JSON.parse(docProps.getProperty('matchups') || '{}');

    // 2. Determine the current week to show games for
    const currentWeek = matchupsObj.week;
    let selectedGameIds = [];
    if (currentWeek) selectedGameIds = matchupsObj[currentWeek].ids;

    // 3. Fetch details for only the selected games from ESPN
    let matchups = [];
    if (selectedGameIds.length > 0) {
      const apiUrl = `${SCOREBOARD}?week=${currentWeek}`;
      const response = UrlFetchApp.fetch(apiUrl, { 'muteHttpExceptions': true });
      const json = JSON.parse(response.getContentText());
      
      const allGamesForWeek = json.events.map(event => {
        const homeTeam = event.competitions[0].competitors.find(c => c.homeAway === 'home');
        const awayTeam = event.competitions[0].competitors.find(c => c.homeAway === 'away');
        return { id: event.id, homeTeamAbbr: homeTeam.team.abbreviation, homeTeamLogo: homeTeam.team.logo, awayTeamAbbr: awayTeam.team.abbreviation, awayTeamLogo: awayTeam.team.logo };
      });
      
      // Filter the full list to get only the games we care about
      matchups = allGamesForWeek.filter(game => selectedGameIds.includes(game.id));
    }
    
    // 4. Return the complete package of data
    return {
      members,
      configuration,
      roster,
      matchups,
      currentWeek
    };
  } catch (err) {
    Logger.log("⚠️ Error in fetchFinalConfirmationData: " + err.stack);
    throw new Error("Could not load all setup data. Please configure each section first.");
  }
}

/**
 * Placeholder function for what happens after the user confirms the final setup.
 * In a real application, this would trigger the main script to build the draft sheet.
 */
function processFinalSetup() {
  try {
    Logger.log("▶️ Final setup submitted. Main script execution starting...");
    const docProps = PropertiesService.getDocumentProperties();
    const configuration = JSON.parse(docProps.getProperty('configuration'));
    configuration.draftReady = true;
    docProps.setProperty('configuration',JSON.stringify(configuration));
    draftSetupExecute(docProps);
    return { success: true };
  } catch (err) {
    Logger.log(`Process Final Setup issue: ${err.stack}`);
  }
}

//------------------------------------------------------------------------
// SUPPORT POPUP FOR HELP - Loads HTML "supportPrompt.html" file
function showSupportDialog() {
  let html = HtmlService.createHtmlOutputFromFile(`supportPrompt.html`)
      .setWidth(500)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, ` `);
}

// Returns version to help and support popup to quickly review version by users
function getSupportPromptInfo() {
  Logger.log(`Returning support prompt info, version: ${VERSION}`);
  return {
    version: VERSION
  };
}

function reviewDocumentProperties() {
  const docProps = PropertiesService.getDocumentProperties().getProperties();
  Logger.log('DOCUMENT PROPERTIES:');
  Object.keys(docProps).forEach(key => {
    Logger.log(`${key}: ${docProps[key]}`);
  });
}
