const VERSION = '1.2';
/**
 * 🏈 SHEETS FANTASY FOOTBALL PLATFORM
 * Single-day or single-week drafting and scoring for NFL games.
 * Last Updated: 11/25/2025
 * 
 * Created by Ben Powers
 * ben.powers.creative@gmail.com
 * 
 * ------------------------------------------------------------------------
 * 📝 DESCRIPTION:
 * This tool transforms a Google Sheet into a fully-featured platform for running
 * one-day or one-week fantasy football contests. It's perfect for making
 * 🦃 Thanksgiving, 🎄 Christmas, or any 🏆 NFL game day more exciting with friends.
 * The tool guides you through setup, runs the draft, and provides a final recap.
 * 
 * ------------------------------------------------------------------------
 * 🚀 INSTRUCTIONS:
 * After authorizing the script, a new menu called "🏈 Football Tools" will
 * appear in your spreadsheet's menu bar. First you'll run the "▶️ Click Here to Setup"
 * Authorization will guide you through a few screens, one is tricky where you'll have 
 * to click the "Advanced" option in the lower left to proceed through the authorization.
 * 
 * Once you've authorized the script, the toolbar should reload (refresh if not)
 * You should be prompted to start working through the setup, which is designed to be
 * completed in order, with checkmarks (✔️) appearing as you complete each. Details below.
 *  
 * ------------------------------------------------------------------------
 * ⚙️ USAGE - SETUP MENU:
 * 
 * 1. 👥 Members & Draft Order
 *    - Add/remove members and assign them fun, auto-generated team names.
 *    - Set the draft order by dragging-and-dropping (Manual) or have the
 *      script randomize it for you upon submission.
 *    - Configure advanced draft settings like "3rd Round Reversal".
 * 
 * 2. 📆 Game Selection
 *    - Automatically loads the current NFL week's schedule.
 *    - Select which NFL matchups you want to include in your contest's player pool.
 * 
 * 3. 🧮 Roster Configuration
 *    - Define the structure of each team's roster (e.g., 1 QB, 2 RBs, etc.).
 *    - Set the league's scoring format (PPR: 1.0, 0.5, or 0.0).
 *    - The tool will automatically calculate player availability and warn you
 *      if you need to add "Extra Copies" of players to support your league size.
 * 
 * 4. 🚀 Finalize & Deploy
 *    - Opens a final confirmation panel that summarizes all your settings.
 *    - From here, you can launch back into any of the setup panels to make last-minute
 *      changes or click "Submit Setup" to lock in your configuration and prepare for the draft.
 * 
 * ------------------------------------------------------------------------
 * 
 * I'm thrilled you're here and checking this out. If you're feeling generous and have the
 * means to support my work, you can support my wife, five kiddos (sixth on the way!), and me:
 * https://www.buymeacoffee.com/benpowers
 * OR
 * https://venmo.com/benpowerscreative
 * 
 * Thanks for checking out the script!
 * 
 **/

function onOpen() {
  // Use userProperties to track if THIS user has seen setup for THIS sheet
  const sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const setupKey = 'footballWeeklySpreadsheetInitialized_v' + VERSION + '_' + sheetId.slice(1,10);
  Logger.log(setupKey);
  const isSetup = PropertiesService.getUserProperties().getProperty(setupKey);
  Logger.log(isSetup);

  if (isSetup) {
    toolbar();
  } else {
    SpreadsheetApp.getUi().createMenu('🏈 Football Tools')
      .addItem('▶️ Click Here to Setup', 'runSetup')
      .addToUi();
  }
}

function runSetup() {
  const ui = SpreadsheetApp.getUi();

  try {
    SpreadsheetApp.getActiveSpreadsheet().getName();
    
    // Mark as setup for this specific sheet
    const userProperties = PropertiesService.getUserProperties();
    const sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const setupKey = 'footballWeeklySpreadsheetInitialized_v' + VERSION + '_' + sheetId.slice(1,10);
    userProperties.setProperty(setupKey, 'true');
    
    // Optionally, also mark it in documentProperties (for other features)
    const docProperties = PropertiesService.getDocumentProperties();
    docProperties.setProperty('setupComplete', 'true');
    
    runFirst();
    
  } catch (err) {
    Logger.log(err.stack);
    ui.alert('⚠️ Authorization Required', 'Please complete authorization and try again.', ui.ButtonSet.OK);
  }
}

function runFirst() {
  // Deletes any existing triggers
  deleteTriggers();

  // Creates a trigger to automatically load the toolbars upon opening
  toolbar();

  const ui = SpreadsheetApp.getUi();
  let setupSteps = ui.alert(`🎉 CONGRATULATIONS`,`You've successfully enabled the scripts to run!
  

  There are FOUR steps to begin your draft:
  
  1️⃣ MEMBERS enter names, team names, and determine draft order
  
  2️⃣ GAMES selection for which NFL matchups to include
  
  3️⃣ ROSTERS configuration of which positional spots you'll draft
  
  4️⃣ DEPLOY the setup to generate the draft lobby
  

  Click "OK" to bring up the configuration overview panel.
  
  🤝 Have fun!`, ui.ButtonSet.OK);
  if (setupSteps == 'OK') {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Launching the setup screen now...`,`🔄 SETUP LOADING`);
    draftSetup();
  }
}

