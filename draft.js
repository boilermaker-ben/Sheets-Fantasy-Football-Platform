// WEEKLY FF DRAFT - Updated 11.25.2025

// DRAFT SETUP
// Function to remake "PICKS" (table), "DRAFT" (display of snake draft), and "ROSTERS" sheets for new configuration data
function draftSetupExecute(docProps) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  docProps = docProps || PropertiesService.getDocumentProperties();
  const configuration = JSON.parse(docProps.getProperty('configuration'));
  const trr = configuration.thirdRoundReversal;
  const pprText = configuration.ppr;
  const roster = JSON.parse(docProps.getProperty('roster'));
  const members = JSON.parse(docProps.getProperty('members'));

  if (configuration.draftReady) {
    deleteTriggers();
    const positionsRanges = ['QB','RB','WR','TE','FLX','SPFLX','DEF','K'];
    const positionArr = positionsRanges.flatMap(pos => roster.find(item => item.pos === pos)?.count);
    const positionDraftList = roster.flatMap(item => Array(item.count).fill(item.pos));
    let positions = positionDraftList.length;
    
    let draftComplete = fetchDraftComplete();
    let active = false;
    try {
      if (!draftComplete) {
        const existing = existingDraft();
        active = existing[0] > 0;
        draftComplete = (existing[1] == 0 && existing[0] > 0);
      }
    } catch (err) {
      Logger.log(`❕ Encountered an issue parsing previous draft data or it was unavailable. Moving on...`)
    }
    let backup, failed, promptText = '';
    if (draftComplete) {
      backup = ui.alert(`🔎 EXISTING DATA FOUND`,'It looks like your draft board has been populated by a previous draft.\r\n\r\nDo you want to make a copy before proceeding?',ui.ButtonSet.YES_NO_CANCEL)
    }
    if (backup == 'YES') {
      try {
        let sheetName;
        try {
          let namedRanges = ss.getNamedRanges();
          for(let a = 0; a < namedRanges.length; a++){
            if (namedRanges[a].getRange().getSheet().getName() == 'ROSTERS') {
              namedRanges[a].remove();
            }
          }
        }
        catch (err) {
          Logger.log(`"ROSTERS" backup failed to remove existing named ranges`);
        }
        let sheetNames = ss.getSheets().map(sheet => sheet.getName());
        let rosterSheets = sheetNames.filter(sheetName => /^ROSTERS_\d{2}$/.test(sheetName));
        if (rosterSheets.length === 0) {
          Logger.log(`No matching sheets found.`);
          sheetName = `ROSTERS_01`;
        } else {
          let highestIndex = Math.max(...rosterSheets.map(sheetName => parseInt(sheetName.match(/\d{2}$/)[0], 10)));
          let index = highestIndex;
          highestIndex < 9 ? index = `0` + (index+1) : index++;
          sheetName = `ROSTERS_` + index;
        }
        ss.getSheetByName(`ROSTERS`).copyTo(ss).setName(sheetName);
        ss.toast(`Backed up previous draft to "${sheetName}".`,`💾 BACKUP COMPLETE`);
      }
      catch (err) {
        Logger.log(err.stack)
        failed = ui.alert(`⚠️ ERROR`,`Error encounter while trying to copy over existing "ROSTERS" sheet.\r\n\r\nWould you still like to continue?`, ui.ButtonSet.YES_NO);
      }
      if (failed == 'NO') {
        ss.toast(`Canceled setup`);
        Logger.log(`Canceled setup`)
        return null;
      }
    }
    let prompt;
    promptText = active ? `Reset draft board and start new draft?` : `Set up the draft board for a new draft?`;
    promptText += `
    
    Please be patient, this step takes a while to process`;
    
    if ( backup == 'YES' || !draftComplete ) {
      prompt = ui.alert(`📋 DRAFT BOARD`,promptText, ui.ButtonSet.OK_CANCEL );
    } else {
      prompt = 'OK';
    }
    if ( prompt == 'OK' ) {
      const pprValue = parseFloat(pprText);
      const ppr = pprValue == 0.0 ? 'NO PPR' : pprValue == 0.5 ? 'HALF PPR' : pprValue == 1.0 ? 'FULL PPR' : 'PPR ERR';
      ss.getRangeByName('PPR').setValue(ppr);

      let checkboxes = ss.getRangeByName(`DRAFT_CHECKBOXES`).getValues();
      let draftBoard = ss.getSheetByName(`DRAFT_LOBBY`);
      ss.getRangeByName(`DRAFT_CHECKBOXES`).clearContent();
      draftBoard.showRows(3,checkboxes.length);

      const drafters = [...members].sort((a, b) => a.draftSlot - b.draftSlot).map(m => m.name);
      const draftersTeamNames = [...members].sort((a, b) => a.draftSlot - b.draftSlot).map(m => m.teamName);
      
      let a, b, c, rows, cols, full, topRow, rowMultiplier, rowAdditional, range;
          
      // Updates all the player information and available pool of players
      playersRefresh(1); // 1 for auto
      ss.toast(`Updated player pool, projections, and outlooks for players`,`🎯 UPDATED PLAYERS`);
      
      try {
        injuryCheck();
        ss.toast(`Injuries notes were successfully set for available designations.`, `🤕 INJURY NOTES SET`);
      } catch (err) {
        ss.toast(`Issue denoting injuries or they were unavailable.`, `🤕 INJURY NOTES NOT SET`);
      }

      if ( prompt == 'OK' ) {

        let picker = drafters[0];
        let onDeck = drafters[1];
        let totalDrafters = drafters.length;
        positions = parseInt(positions,10);
        let sheet;
        let sheetNames = [`DRAFT`,`ROSTERS`,`PICKS`];
        const darkNavy = `#222735`;
        const lightGray = `#E5E5E5`;
        for ( a = 0; a < sheetNames.length; a++ ) {
          sheet = ss.getSheetByName(sheetNames[a]) || ss.insertSheet(sheetNames[a]);
          
          if ( sheetNames[a] == `DRAFT` ) {
            rowMultiplier = 3;
            rowAdditional = 2;
            cols = members.length + 1;
          } else if ( sheetNames[a] == `ROSTERS` ) {
            rowMultiplier = 3;
            rowAdditional = 2;
            cols = members.length*3 + 1;
          } else if ( sheetNames[a] == `PICKS` ) {
            cols = 8 + positionDraftList.length;
            rows = members.length*positions+3;
            rowMultiplier = 1;
            rowAdditional = 1;
          }
          
          // Format draft board to be the correct size
          sheet.clear();
          
          let maxCols = sheet.getMaxColumns();
          if ( cols < maxCols ) {
            sheet.deleteColumns(cols, maxCols - cols);
          } else if ( cols > maxCols ) {
            sheet.insertColumnsAfter(maxCols, cols - maxCols);
          }
          maxCols = sheet.getMaxColumns();
          
          let maxRows = sheet.getMaxRows();
          if ( sheetNames[a] != `PICKS`) {
            rows = positions * rowMultiplier + rowAdditional;
          }
          if ( rows < maxRows ) {
            sheet.deleteRows(rows, maxRows - (rows));
          } else if ( rows > maxRows ) {
            sheet.insertRowsAfter(maxRows, (rows) - maxRows);
          }
          maxRows = sheet.getMaxRows();
          if ( sheetNames[a] == `ROSTERS` ) {
            sheet.insertRowsAfter(maxRows, 4);
            maxRows = sheet.getMaxRows();
          }
          
          if ( sheetNames[a] == `DRAFT` || sheetNames[a] == `ROSTERS` ) {
            sheet.setRowHeight(1,35)
              .setRowHeight(2,25)
              .setRowHeights(3,rows-2,25)
              .setColumnWidth(1,65)
              .setColumnWidths(2,cols-1,120);
            full = sheet.getRange(1,1,maxRows,maxCols);
            
            ss.setNamedRange(`FULL`,full);
            full.breakApart()
              .setFontSize(18)
              .setBackground(`white`) // formerly `#283244`
              .setFontColor(darkNavy)
              .setHorizontalAlignment(`center`)
              .setVerticalAlignment(`middle`)
              .setBorder(true,true,true,true,false,false,darkNavy,SpreadsheetApp.BorderStyle.SOLID_THICK);
            sheet.getRange(1,2,1,maxCols-1).setBackground(`#FFD900`); // Bright Yellow
            sheet.getRange(2,2,1,maxCols-1).setBackground(`#FFF1A2`); // Desaturated Yellow
            let topCorner = sheet.getRange(1,1,2,1);
            topCorner.setBackground(darkNavy)
              .merge()
              .setHorizontalAlignment(`center`)
              .setVerticalAlignment(`middle`)
              .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
            ss.setNamedRange(`LIVE`,sheet.getRange(1,1));
            for ( b = 2; b < rows; b++ ) {
              if ( b % 3 == 0 ) {
                if ( sheetNames[a] == `DRAFT` ) {
                  if ( b == 1 ) {
                    sheet.getRange(b-1,1).setValue(`ROUND`)
                      .setFontSize(12)
                      .setFontColor(`white`)
                      .setVerticalAlignment(`bottom`);
                  }
                  sheet.getRange(b,1).setValue(b/3);
                }
                sheet.getRange(b,1,1,cols).setBorder(true,null,null,null,null,null,darkNavy,SpreadsheetApp.BorderStyle.SOLID_THICK);
                if (b/3 == 3 && trr && sheetNames[a] == `DRAFT` ) {
                  sheet.getRange(b,1,3,1).setValues([[`THIRD`],[`ROUND`],[`REVERSAL`]])
                    .setFontColor(darkNavy)
                    .setFontSize(12)
                    .setVerticalAlignment(`middle`)
                    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
                    .setHorizontalAlignment(`center`);
                  sheet.setRowHeight(b,21);      
                } else {
                  sheet.getRange(b,1,3,1).mergeVertically();
                }
              }
            }
            for ( b = 0; b <= positions-1; b++ ) {
              //First Row
              sheet.setRowHeight(b*3+3,21);
              sheet.getRange(b*3+3,2,1,cols-1).setFontSize(12)
                .setHorizontalAlignment(`left`)
                .setVerticalAlignment(`bottom`);
              
              //Second Row
              sheet.getRange(b*3+4,2,1,cols-1).setFontSize(14)
                .setHorizontalAlignment(`left`)
                .setVerticalAlignment(`bottom`)
                .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
              
              //Third Row
              sheet.getRange(b*3+5,2,1,cols-1).setFontSize(14)
                .setHorizontalAlignment(`left`)
                .setVerticalAlignment(`top`)
                .setFontWeight(`bold`)
                .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
              if ( (b+1) % 2 == 0 ) { 
                sheet.getRange(b*3+3,1,3,maxCols).setBackground(lightGray);
              }
            }
            
            //full.setFontFamily("Abel");
            //full.setFontFamily("Montserrat");
            full.setFontFamily("Teko");
            topRow = sheet.getRange(1,1,1,maxCols);
            topRow.setVerticalAlignment(`bottom`);
            sheet.getRange(2,1,1,maxCols).setFontSize(12);
            sheet.getRange(2,1,1,maxCols).setVerticalAlignment(`top`);

            // DRAFT SHEET ONLY FORMATTING            
            if ( sheetNames[a] == `DRAFT` ) {
              sheet.getRange(1,(cols+1)-members.length,1,members.length).setValues([drafters]);
              sheet.getRange(2,(cols+1)-members.length,1,members.length).setValues([draftersTeamNames]);
              for ( b = 0; b <= positionDraftList.length-1; b++ ) {
                for ( c = 0; c <= positionDraftList.length-1; c++ ) {
                  sheet.getRange(c*3+3,b+2,3,3).setBorder(false,false,false,false,false,false,darkNavy,SpreadsheetApp.BorderStyle.SOLID);
                  sheet.getRange(c*3+3,b+2,3,3).setBorder(true,true,true,true,false,false,darkNavy,SpreadsheetApp.BorderStyle.SOLID_THICK);
                  sheet.getRange(c*3+5,b+2,3,2).setBorder(null,null,true,null,null,null,darkNavy,SpreadsheetApp.BorderStyle.SOLID);
                }
              }
              ss.toast(`Successfully configured and deployed draft sheet!`,`📚 DRAFT SHEET`);
            // ROSTER SHEET ONLY FORMATTING
            } else if ( sheetNames[a] == `ROSTERS` ) {
              sheet.unhideRow(sheet.getRange(1,1,sheet.getMaxRows(),1));
              let posList = [`QB`,`RB`,`WR`,`TE`,`DEF`,`K`,`FLX`,`SPFLX`];//
              //let hexList = [`b22052`,`009288`,`4781c4`,`b37c43`,`022047`,`8e4dbf`,`a3b500`,`b50900`];
              let hexList = [`FF2A6D`,`00CEB8`,`58A7FF`,`FFAE58`,`7988A1`,`BD66FF`,`FFF858`,`E22D24`];
              let hexAltList = [`C82256`,`00A493`,`4482C6`,`CD8B45`,`5D697D`,`9650CB`,`CAC444`,`B8251E`];
              for ( b = 0; b <= positionDraftList.length-1; b++ ) {
                sheet.getRange(b*3+3,1).setValue(positionDraftList[b]);
                sheet.getRange(b*3+3,1,3,1).setBackground(`#` + hexList[posList.indexOf(positionDraftList[b])]);
              }
              sheet.getRange(maxRows-3,1).setValue(`PROJ`);
              sheet.getRange(maxRows-2,1).setValue(`RANK`);
              sheet.getRange(maxRows-1,1).setValue(`POINTS`);
              ss.setNamedRange(`ROSTER_NAMES`,sheet.getRange(1,2,1,maxCols-1));
              ss.setNamedRange(`ROSTER_POINTS`,sheet.getRange(maxRows-1,2,1,maxCols-1));
              sheet.getRange(maxRows,1).setValue(`RANK`);
              sheet.setRowHeights(maxRows-3,2,50);
              let conditionalRangePlayers = [];
              
              for ( b = 0; b < members.length; b++ ) {
                sheet.getRange(1,b*3+2).setValue(drafters[b]);
                sheet.getRange(2,b*3+2).setValue(draftersTeamNames[b]);
                sheet.getRange(2,b*3+2,1,2).merge();
                sheet.hideColumn(sheet.getRange(2,b*3+4));
                sheet.getRange(1,b*3+4).setValue(`Score`);
                sheet.getRange(1,b*3+2,1,2,).merge();
                sheet.getRange(1,b*3+4,2,1).merge();
                                
                for ( c = 0; c <= positionDraftList.length-1; c++ ) {
                  sheet.getRange(c*3+3,b*3+2,3,3).setBorder(false,false,false,false,false,false,darkNavy,SpreadsheetApp.BorderStyle.SOLID);
                  sheet.getRange(c*3+3,b*3+2,3,3).setBorder(true,true,true,true,false,false,darkNavy,SpreadsheetApp.BorderStyle.SOLID_THICK);
                  sheet.getRange(c*3+3,b*3+3,3,2).setBorder(null,null,null,null,true,null,darkNavy,SpreadsheetApp.BorderStyle.SOLID);
                  
                  let scoreRange = sheet.getRange(c*3+3,b*3+4);
                  scoreRange.setFormulaR1C1(`=iferror(vlookup(vlookup(R[1]C[-1]&" "&R[2]C[-1],{SLPR_FULL_NAME,SLPR_PLAYER_ID},2,false),{SLPR_PLAYER_ID,SLPR_SCORE},2,false))`);
                  
                  let scoringColBoxes = sheet.getRange(c*3+3,b*3+4,3,1)
                  scoringColBoxes.merge()
                    .setFontSize(18)
                    .setFontWeight(`bold`)
                    .setFontColor(darkNavy)
                    .setHorizontalAlignment(`center`)
                    .setVerticalAlignment(`middle`);

                  let headshotRange = sheet.getRange(c*3+4,b*3+2,2,1);
                  headshotRange.merge()
                    .setFormulaR1C1(`=iferror(image(vlookup(vlookup(R[0]C[1]&" "&R[1]C[1],{SLPR_FULL_NAME,SLPR_PLAYER_ID},2,false),{SLPR_PLAYER_ID,SLPR_IMAGE},2,false)))`);
                  
                  let teamRange = sheet.getRange(c*3+3,b*3+2);
                  teamRange.setHorizontalAlignment(`left`)
                    .setVerticalAlignment(`middle`)
                    .setFormulaR1C1(`=iferror(vlookup(vlookup(R[1]C[1]&" "&R[2]C[1],{SLPR_FULL_NAME,SLPR_PLAYER_ID},2,false),{SLPR_PLAYER_ID,SLPR_TEAM},2,false),)`);
                  
                  let projRange = sheet.getRange(c*3+3,b*3+3);
                  projRange.setHorizontalAlignment(`right`)
                    .setVerticalAlignment(`middle`)
                    .setFormulaR1C1(`=iferror(vlookup(vlookup(R[1]C[0]&" "&R[2]C[0],{SLPR_FULL_NAME,SLPR_PLAYER_ID},2,false),{SLPR_PLAYER_ID,SLPR_PROJ},2,false),)`);
                }
              
                sheet.getRange(maxRows-3,b*3+2).setFormulaR1C1(`=iferror(sum(R3C[1]:R[-3]C[1]))`);
                sheet.getRange(maxRows-2,b*3+2).setFormulaR1C1(`=iferror(if(R[-1]C[0]=0,,arrayformula(rank(R[-1]C[0],R[-1]C2:R[-1]C${members.length*3+1}))))`);
                sheet.getRange(maxRows-1,b*3+2).setFormulaR1C1(`=iferror(sum(R3C[2]:R[-3]C[2]))`);
                sheet.getRange(maxRows,b*3+2).setFormulaR1C1(`=iferror(if(R[-1]C[0]=0,,arrayformula(rank(R[-1]C[0],R[-1]C2:R[-1]C${members.length*3+1}))))`);
                
                conditionalRangePlayers.push(sheet.getRange(2,b*3+4,maxRows-6,1));
                
                sheet.getRange(maxRows-3,b*3+2,4,3).mergeAcross();
                
                sheet.setColumnWidth(b*3+2,50); // Headshot col width
                sheet.setColumnWidth(b*3+4,55); // Score col width
                
              }
              sheet.getRange(maxRows-3,1,4,members.length*3+1).setBorder(true,true,true,true,true,true,darkNavy,SpreadsheetApp.BorderStyle.SOLID_THICK)
                .setVerticalAlignment(`middle`)
                .setHorizontalAlignment(`center`);
              sheet.clearConditionalFormatRules();

              // const lightGray = `#b9b9b9`;
              // const darkNavy = `#00142f`;
              const slprGreen = `#45E6A7`
              const slprRed = `#F75C8D`;              

              //Formatting for projected scoring and rank
              let formatProj = SpreadsheetApp.newConditionalFormatRule()
                .setGradientMaxpointWithValue(slprGreen, SpreadsheetApp.InterpolationType.PERCENT, `100`)
                .setGradientMidpointWithValue(`white`,SpreadsheetApp.InterpolationType.PERCENT, `50`)
                .setGradientMinpointWithValue(slprRed,SpreadsheetApp.InterpolationType.PERCENT, `0`)
                .setRanges([sheet.getRange(maxRows-3,2,1,sheet.getMaxColumns())])
                .build();

              let formatProjRank = SpreadsheetApp.newConditionalFormatRule()
                .setGradientMaxpointWithValue(slprRed, SpreadsheetApp.InterpolationType.PERCENT, `100`)
                .setGradientMidpointWithValue(`white`,SpreadsheetApp.InterpolationType.PERCENT, `50`)
                .setGradientMinpointWithValue(slprGreen,SpreadsheetApp.InterpolationType.PERCENT, `0`)
                .setRanges([sheet.getRange(maxRows-2,2,1,sheet.getMaxColumns())])
                .build();
              
              //Formatting for actual scoring and rank
              let formatPoints = SpreadsheetApp.newConditionalFormatRule()
                .setGradientMaxpointWithValue(slprGreen, SpreadsheetApp.InterpolationType.PERCENT, `100`)
                .setGradientMidpointWithValue(`white`,SpreadsheetApp.InterpolationType.PERCENT, `50`)
                .setGradientMinpointWithValue(slprRed,SpreadsheetApp.InterpolationType.PERCENT, `0`)
                .setRanges([sheet.getRange(maxRows-1,2,1,sheet.getMaxColumns())])
                .build();

              let formatPointsRank = SpreadsheetApp.newConditionalFormatRule()
                .setGradientMaxpointWithValue(slprRed, SpreadsheetApp.InterpolationType.PERCENT, `100`)
                .setGradientMidpointWithValue(`white`,SpreadsheetApp.InterpolationType.PERCENT, `50`)
                .setGradientMinpointWithValue(slprGreen,SpreadsheetApp.InterpolationType.PERCENT, `0`)
                .setRanges([sheet.getRange(maxRows,2,1,sheet.getMaxColumns())])
                .build();
              
              //Formatting for points scored per player
              let formatPlayerPoints = SpreadsheetApp.newConditionalFormatRule()
                .setGradientMaxpointWithValue(slprGreen, SpreadsheetApp.InterpolationType.PERCENT, `100`)
                .setGradientMidpointWithValue(`white`,SpreadsheetApp.InterpolationType.PERCENT, `50`)
                .setGradientMinpointWithValue(slprRed,SpreadsheetApp.InterpolationType.PERCENT, `0`)
                .setRanges(conditionalRangePlayers)
                .build();

              let formatRules = sheet.getConditionalFormatRules();
              formatRules.push(formatProj);
              formatRules.push(formatProjRank);
              formatRules.push(formatPoints);
              formatRules.push(formatPointsRank);
              formatRules.push(formatPlayerPoints);
              sheet.setConditionalFormatRules(formatRules);
              
              let hideRange = sheet.getRange(maxRows-1,1,2,maxCols);
              sheet.hideRow(hideRange);
              scoringHide(true);
            }
            ss.toast(`Successfully configured and deployed roster and scoring sheets!`,`📚 SHEETS SETUP`);
          } else if ( sheetNames[a] == `PICKS` ) {
            let draftersHelper = [...drafters];
            draftersHelper.reverse();
            let arrCol = [];
            for (let a = draftersHelper.length; a > 0 ; a--) {
              arrCol = arrCol.concat(a);
            }
            let arrDraftersInitial = [];
            let arrColsInitial = [];
            sheet.setColumnWidths(1,4,30);
            sheet.setColumnWidth(5,80);
            sheet.setColumnWidths(6,4,40);
            sheet.setColumnWidths(9,positionDraftList.length,50);
            
            sheet.getRange(2,5,maxRows,maxCols-5).setHorizontalAlignment(`left`);
            full = sheet.getRange(2,1,maxRows,maxCols);

            for (let a = 2; a < (positions + 2); a++) {
              if (trr & a == 4) {
                arrDraftersInitial = arrDraftersInitial.concat(draftersHelper);
                arrColsInitial = arrColsInitial.concat(arrCol);  
              } else {
                arrDraftersInitial = arrDraftersInitial.concat(draftersHelper.reverse());
                arrColsInitial = arrColsInitial.concat(arrCol.reverse());
              }
            }
            let arr = [];
            arrDraftersInitial = arrDraftersInitial.concat(`End`);
            for (let a = 0; a < arrDraftersInitial.length; a++) {
              arr[a] = [];
              arr[a] = [a+1, Math.floor(a/totalDrafters) + 1,(a+1) - (Math.floor((a/totalDrafters)))*totalDrafters,arrColsInitial[a],arrDraftersInitial[a]];
            }

            sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).clearNote();
            let headers = [`overall`,`round`,`round_pick`,`col`,`picker`,`player`,`pos`,`pos_index`];
            headers = headers.concat(...positionDraftList);
            sheet.getRange(1,1,1,headers.length).setValues([headers]);
            const picksNamedRanges = [`OVERALL`,`ROW`,`PICK`,`COL`,`NAME`,`ID`,`POS`,`ROSTER_ROW`,`REMAINING`] // will have `PICKS_` appended;
            const picksNotes = [`Overall selection number`,
              `Draft pick row placement on DRAFT sheet, also draft round`,
              `Round pick number`,
              `Column index of drafter`,
              `Drafters name`,
              `ID of draft pick, based on Sleeper ID`,
              `Position of draft pick`,
              `Position index of draft pick on drafter's roster`,
              `Array of all remaining positions to fill after this pick`];

            picksNamedRanges.forEach(index => {
              let col = picksNamedRanges.indexOf(index)+1;
              sheet.getRange(1,col).setNote(picksNotes[picksNamedRanges.indexOf(index)]);
              if (index == `REMAINING`) {
                ss.setNamedRange(`PICKS_${index}`,sheet.getRange(2,col,maxRows,positionDraftList.length));
                ss.setNamedRange(`STARTERS`,sheet.getRange(1,col,1,positionDraftList.length))
              } else {
                ss.setNamedRange(`PICKS_${index}`,sheet.getRange(2,col,maxRows,1));
              }
            });
            
            sheet.getRange(2,1,maxRows-2,5).setValues(arr);
            ss.toast(`Configured picks based on draft order [PICKS sheet deployed successfully]`,`📃 PICK ORDER SET`);
          }
          
        }
        // Sets the active drafter to the cell on the draft board page
        ss.getRangeByName(`PICKER`).setValue(picker);
        ss.getRangeByName(`PICKER_NEXT`).setValue(onDeck);
        ss.getRangeByName(`PICKER_ROSTER`).setValue(arrayToString(arrayClean(positionDraftList),false,false));
        
        
      }
      const configuration = JSON.parse(docProps.getProperty('configuration'));
      configuration.draftReady = true;
      docProps.setProperty('configuration',JSON.stringify(configuration));
      toolbar();
      prompt = ui.alert(`👟 READY?`,`Are you ready to begin the draft?`, ui.ButtonSet.YES_NO);
      if (prompt == 'YES') {
        configuration.drafting = true;
        docProps.setProperty('configuration',JSON.stringify(configuration));
        ss.toast(`Creating triggers for active drafting...`,`🧙 TRIGGERS`);
        draftBoard.activate();
        triggersDrafting();
        ui.alert(`🚀 START DRAFTING!`,`⌛ ${drafters[0]} is on the clock!\r\n\r\nUse the Challenge Flag to UNDO the last pick.` ,ui.ButtonSet.OK);
      } else {
        ss.toast(`Configuration finished setup`,`✅ COMPLETE`);
        ui.alert(`🚀 SETUP COMPLETE`,`Use the "🏈 Fantasy Tools" menu option "▶️ Start Draft" to start.\r\n\r\nProjections and injury statuses may change before kickoff\r\nUse the "🏈 Fantasy Tools" > "🔄 Refresh Players" to get up-to-date information before drafting.`,ui.ButtonSet.OK);
        toolbar();
      }
    } else {
      deleteTriggers();
    }
  } else {
    Logger.log(`❗ Configuration did not reflect draft setup readiness. Ensure all configuration has been completed, then try again.`)
    ss.toast(`Configuration did not reflect draft setup readiness. Ensure all configuration has been completed, then try again.`,`❗ INCOMPLETE CONFIG!`);
  }
}

// FETCHES PICKS
// A more efficient way to get all `PICK` sheet data quickly than by named ranges
function picksData(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(`PICKS`);
  const table = sheet.getDataRange().getValues();
  const starterRow = table[0].splice(table[0].indexOf(`pos_index`)+1,table[0].length - table[0].indexOf(`pos_index`));
  let obj = table[0].reduce((result, key, col) => {
    result = result || {};
    if (key === `pos_index`) {
      result.remaining = table.slice(1).map((row) => row.slice(col + 1));
      result.starters = table[0].slice(1,table[0].length);
    } else {
      result[key] = result[key] || table.slice(1).map((row) => row[col]);
    }
    return result;
  }, {});
  obj.starters = starterRow;
  obj.participants = Math.max(...obj.round_pick);
  const remainingLength = obj.remaining.length;
  for (let a = obj.participants; a < remainingLength; a++) {
    if (obj.remaining[a].filter(Boolean).length === 0) {
      obj.last = a;
      break;
    }
  }
  const counts = obj.starters.reduce((acc, key) => {
    acc[key] = 0;
    return acc;
  }, {});
  for (let a = obj.last - obj.participants; a < obj.last; a++) {
    obj.remaining[a].forEach(key => {
      if (key) counts[key] += 1;
    });
  }
  const positions = [`QB`,`RB`,`WR`,`TE`,`K`,`DEF`].filter(key => obj.starters.indexOf(key) >= 0);
  const flx = [`RB`,`WR`,`TE`];
  const spflx = [`QB`,`RB`,`WR`,`TE`];
  obj.eliminated = [];
  for (let a = 0; a < positions.length; a++) {
    const pos = positions[a];
    if (counts[pos] === 0) {
      if (spflx.indexOf(pos) >= 0) {
        if (counts.hasOwnProperty(`SPFLX`)) {
          if (counts[`SPFLX`] === 0 && pos === `QB`) {
            obj.eliminated.push(pos);
          } else if (flx.indexOf(pos) >= 0) {
            if (counts.hasOwnProperty(`FLX`)) {
              if (counts[`FLX`] === 0) {
                obj.eliminated.push(pos);
              }
            }
          }
        }
      } else {
        obj.eliminated.push(pos);
      }
    }
  }
  return obj;
}

// FETCHES DRAFTER PLAYER PLACEMENT AND NEXT TWO PICKER NAMES
function nextDrafter(dataObj,pos,ss) {
  try {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  
    const sleeperId = ss.getRangeByName(`SLPR_PLAYER_ID`).getValues().flat();
    const sleeperName = ss.getRangeByName(`SLPR_FULL_NAME`).getValues().flat();
    
    let target = dataObj.player.indexOf('');
    let row = dataObj.round[target];
    let col = dataObj.col[target];
    let count = target+1;
    
    let pick = dataObj.round_pick[target];
    let picker = dataObj.picker[target];
    let nextPicker = dataObj.picker[target+1];
    let onDeck = dataObj.picker[target+2];

    // Third Round Reversal Boolean
    let third = dataObj.round.indexOf(3);
    const trr = (third > -1) ? (dataObj.picker[third] != dataObj.picker[third-1] ? true : false) : false;
    const reverse = (trr && dataObj.round[target+2] === 3) ? true : false;

    let currentRoster = [];

    let last = 0;
    for (let a = 0; a < dataObj.picker.length; a++ ) {
      if (dataObj.picker[a] == picker && dataObj.player[a] != '') {
        currentRoster = dataObj.remaining[a];
      }
    }
    if (currentRoster.length == 0) {
      currentRoster = dataObj.starters;
    }
    let nextCurrentRoster = [];
    
    if (picker === nextPicker && row === 1) {
      for (let a = 0; a < dataObj.starters.length; a++) {
        nextCurrentRoster.push(dataObj.starters[a]);
      }
      try {
        nextCurrentRoster.splice(nextCurrentRoster.indexOf(pos),1,'');
      } catch (err) {
        try {
          if (pos === `WR` || pos === `RB` || pos === `TE`) {
            nextCurrentRoster.splice(nextCurrentRoster.indexOf(`FLX`),1,'');
          } else if (pos === `QB`) {
            nextCurrentRoster.splice(nextCurrentRoster.indexOf(`SPFLX`),1,'');
          }
        } catch (err) {
          if (pos === `QB`) {
            nextCurrentRoster.splice(nextCurrentRoster.indexOf(`SPFLX`),1,'');
          }
        }
      }
    }

    for (let a = 0; a < dataObj.picker.length; a++ ) {
      if (dataObj.picker[a] == nextPicker && dataObj.player[a] != ''  && picker != nextPicker) {
        nextCurrentRoster = dataObj.remaining[a];
      } else if (dataObj.picker[a] == picker && dataObj.player[a] != '' && picker == nextPicker) {
        for (let a = 0; a < dataObj.starters.length-1; a++) {
          nextCurrentRoster.push(dataObj.starters[a]);
        }
        try {
          nextCurrentRoster.splice(dataObj.starters.indexOf(pos),1,'');
        } catch (err) {
          try {
            if (pos === `WR` || pos === `RB` || pos === `TE`) {
              nextCurrentRoster.splice(dataObj.starters.indexOf(`FLX`),1,'');
            } else if (pos === `QB`) {
              nextCurrentRoster.splice(dataObj.starters.indexOf(`SPFLX`),1,'');
            }
          } catch (err) {
            if (pos === `QB`) {
              nextCurrentRoster.splice(dataObj.starters.indexOf(`SPFLX`),1,'');
            }
          }
        }
      }
      if (dataObj.picker[a] == picker) {
        last = a;
      }
    }
    if (nextCurrentRoster.length == 0) {
      nextCurrentRoster = dataObj.starters;
    }  
    let currentPlayerIds = [];
    let currentPlayerNames = [];
    for (let a = 0; a < dataObj.player.length; a++ ) {
      if ( dataObj.picker[a] == picker ) {
        currentPlayerIds.push(dataObj.player[a]);
        currentPlayerNames.push(sleeperName[sleeperId.indexOf(dataObj.player[a])]);
      }
    }
    let object = {
      row,
      col,
      count,
      pick,
      picker,
      nextPicker,
      onDeck,
      roster:currentRoster,
      nextRoster:nextCurrentRoster,
      currentPlayerIds,
      currentPlayerNames,
      trr,
      reverse
    };
    return object;
  } catch (err) {
    Logger.log(`Issue with nextDrafter function: ${err.stack}`);
  }
}

// DYNAMIC DRAFTER
// Tool to hide/reveal all picked players and select them while moving them to the correct sheet
function dynamicDrafter(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const draftBoard = ss.getActiveSheet();
  const sheetName = ss.getSheetName();

  const range = e.range;
  const colDraft = range.getColumnIndex();
  const rowDraft = range.getRowIndex();
  if ( sheetName === `DRAFT_LOBBY` &&  colDraft === 2 ) {
    let dataObj = picksData(ss);
    if(rowDraft === 1) {
      toggleMarked(e,ss,dataObj.eliminated);
    } else if (rowDraft > 2) {
      let value = range.isChecked();
      if ( value == true ) {
        ss.toast(`Processing user selection`, `🔄 PROCESSING PICK`);
        let draftComplete = false;
        let draftHeaders = ss.getRangeByName(`DRAFT_HEADERS`).getValues().flat();
        // let id = sheet.getRange(rowDraft,draftHeaders.indexOf("ID")+1).getValues();
        let playerInfo = draftBoard.getRange(rowDraft,draftHeaders.indexOf("PLAYER")+1,1,3).getValues().flat();
        let player = playerInfo[0];
        let team = playerInfo[1];
        let pos = playerInfo[2];
        let obj = nextDrafter(dataObj,pos,ss);
        let id = nameFinder(player);
        // let team = draftBoard.getRange(rowDraft,draftHeaders.indexOf("TEAM")+1).getValue();
        // let pos = draftBoard.getRange(rowDraft,draftHeaders.indexOf("POS")+1).getValue();
        let opening = obj.roster.indexOf(pos);
        let alreadyOwned = 0;
        // const positions = [`QB`,`RB`,`WR`,`TE`,`K`,`DEF`].filter(key => obj.starters.indexOf(key) >= 0);
        const flx = [`RB`,`WR`,`TE`];
        const spflx = [`QB`,`RB`,`WR`,`TE`];
        if (opening == -1) {
          if (flx.indexOf(pos) >= 0) {
            opening = obj.roster.indexOf(`FLX`);
          }
        }
        if (opening == -1) {
          if (spflx.indexOf(pos) >= 0) {
            opening = obj.roster.indexOf(`SPFLX`);
          }
        }
        let rosterSlotUsedIndex = -1;
        let rosterString = '';
        let altRosterString = '';
        if (obj.picker == obj.nextPicker) {
          Logger.log(`▶️ CurrentRoster: ${obj.roster}`);
          Logger.log(`🔄 Pos: ${pos} (Found at ${obj.roster.indexOf(pos)})`);
          if (obj.roster.indexOf(pos) > -1) {
            rosterSlotUsedIndex = obj.roster.indexOf(pos);
          }
          if (rosterSlotUsedIndex == -1) {
            if (( pos == `RB` || pos == `WR` || pos == `TE` ) && obj.roster.indexOf(`FLX`) > -1 ) {
              rosterSlotUsedIndex = obj.roster.indexOf(`FLX`);
            }
          }
          if (rosterSlotUsedIndex == -1) {
            if (( pos == `RB` || pos == `WR` || pos == `TE` || pos == `QB` ) && obj.roster.indexOf(`SPFLX`) > -1 ) {
              rosterSlotUsedIndex = obj.roster.indexOf(`SPFLX`);
            }
          }
          let tempRoster = [];
          for (let a = 0; a < obj.roster.length; a++) {
            tempRoster.push(obj.roster[a]);
          }
          tempRoster.splice(rosterSlotUsedIndex,1);
          for ( let a = 0; a < tempRoster.length; a++ ) {
            if ( tempRoster[a] != '' ) {
              rosterString = rosterString.concat(`${tempRoster[a]}, `);
            }
          }
          rosterString = rosterString.slice(0,-2);
          altRosterString = Object.entries(
              tempRoster.filter(pos => pos && positionsEmojiMap[pos]).reduce((acc, pos) => {
                acc[pos] = (acc[pos] || 0) + 1;
                return acc;
              }, {})
            ).map(([pos, count]) => 
              numberMap[count] + pos
            ).join(",");
        } else {
          for ( let a = 0; a < obj.nextRoster.length; a++ ) {
            if ( obj.nextRoster[a] != '' ) {
              rosterString = rosterString.concat(`${obj.nextRoster[a]}, `);
            }
          }
          rosterString = rosterString.slice(0,-2);
          altRosterString = Object.entries(
              obj.nextRoster.filter(pos => pos && positionsEmojiMap[pos]).reduce((acc, pos) => {
                acc[pos] = (acc[pos] || 0) + 1;
                return acc;
              }, {})
            ).map(([pos, count]) => 
              numberMap[count] + pos
            ).join(",");
        }
        let currentPlayersString = '';
        for ( let a = 0; a < obj.currentPlayerNames.length; a++ ) {
          if ( obj.currentPlayerNames[a] != '' && obj.currentPlayerNames[a] != undefined ) {
            currentPlayersString = currentPlayersString.concat(`${obj.currentPlayerNames[a]}\r\n`);
          }
        }  
        if ( obj.currentPlayerIds.indexOf(id) > -1 ) {
          alreadyOwned = 1;
        }
        const ui = SpreadsheetApp.getUi();
        
        if ( opening == -1 && alreadyOwned == 1 ) {
          let str = ( `There isn't a roster spot left for ${player} and you already own that player.\r\n\r\nRemaining positions:\r\n\r\n${rosterString}\r\n\r\nCurrent Players:\r\n\r\n${currentPlayersString}`);
          ui.alert(`🧐 INVALID PICK`,str, ui.ButtonSet.OK);
          range.setValue(false);
        } else if ( opening == -1 ) {
          let str = ( `There isn't a roster spot left for ${player}.\r\n\r\nRemaining positions:\r\n\r\n${rosterString}`);
          ui.alert(`🧐 INVALID PICK`,str, ui.ButtonSet.OK);
          range.setValue(false);
        } else if ( alreadyOwned == 1 ) {
          let str = ( `You've already selected ${player}.\r\n\r\nCurrent Players:\r\n\r\n${currentPlayersString}`);
          ui.alert(`🧐 INVALID PICK`,str, ui.ButtonSet.OK);
          range.setValue(false);
        } else {
          obj.roster.splice(opening,1,'');
          opening++;
          let pickString = obj.row + '.' + obj.pick;
          let again = '';
          if (obj.reverse) {
            again = ` (3rd Round Reversal)`;
          } else if (obj.picker === obj.nextPicker) {
            again = ` (turn)`;
          }
          let title = (`👌 PICK ${obj.count} BY ${obj.picker.toUpperCase()}`);
          let str = `Selected: ${player}, ${pos} (${team})`;
          if (obj.nextPicker == `End`) {
            str += `\r\n\r\n✅ Draft Completed!\r\n\r\nUse the "🔢 Scoring" menu to turn on live scoring when the games are near and see who proves victorious!`;
            draftComplete = true;
          } else {
            str += `\r\n\r\n➡️ Next up: ${obj.nextPicker}${again}`;  
          }
          let notify = ui.alert(title,str, ui.ButtonSet.OK_CANCEL);
          if ( notify == 'OK' ) {
            ss.getRangeByName(`PICKER`).setValue(obj.nextPicker);
            ss.getRangeByName(`PICKER_ROSTER`).setValue(altRosterString); //rosterString);
            ss.getRangeByName(`PICKER_NEXT`).setValue(obj.onDeck);          
            
            if (draftBoard.getRange(1,2).isChecked()) {
              draftBoard.hideRow(draftBoard.getRange(rowDraft,1));
            }
            
            let playerArr = player.split(' ');
            let first = playerArr[0];
            let last = playerArr[1];
            if ( playerArr.length > 2 ) {
              for (let a = 2 ; a < playerArr.length; a++ ) {
                last = last.concat(` ${playerArr[a]}`);
              }
            }
            let picksSheet = ss.getSheetByName(`PICKS`);
            picksSheet.getRange((obj.count + 1),6,1,(3 + obj.roster.length)).setValues([[id,pos,opening].concat(...obj.roster)]);
                      
            let draftSheet = ss.getSheetByName(`DRAFT`);
            let tabs = `\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t`;
            if ( ( pickString.length + team.length ) > 6 ) {
              tabs = `\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t`;
            } else if ( ( pickString.length + team.length ) > 5 ) {
              tabs = `\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t`;
            } else if ( ( pickString.length + team.length ) > 4 ) {
              tabs = `\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t`;
            }
            draftSheet.getRange((obj.row-1)*3+3,obj.col+1,3,1).setValues([[`${pos} - ${team}${tabs}${pickString}`],[first],[last]]);
            // draftSheet.getRange((obj.row-1)*3+4,obj.col+1).setValue();
            // draftSheet.getRange((obj.row-1)*3+5,obj.col+1).setValue(last);
            let posList = [`QB`,`RB`,`WR`,`TE`,`DEF`,`K`,`FLX`,`SPFLX`];
            let hexList = [`FF2A6D`,`00CEB8`,`58A7FF`,`FFAE58`,`7988A1`,`BD66FF`,`FFF858`,`E22D24`];
            let hexAltList = [`C82256`,`00A493`,`4482C6`,`CD8B45`,`5D697D`,`9650CB`,`CAC444`,`B8251E`];
            //let hexList = [`b22052`,`009288`,`4781c4`,`b37c43`,`022047`,`8e4dbf`,`a3b500`,`b50900`]; // Old
            // let hexAltList = [`8f153f`,`01746c`,`2c5e97`,`916232`,`00142f`,`6a3096`]; // Old
            let hex = hexList[posList.indexOf(pos)];
            
            // Adjust second value to reduce/increase saturation
            let hexDesat = hexAltList[posList.indexOf(pos)];

            hexList[posList.indexOf(pos)] == '' ?  hexDesat = hexColorAdjust(hex,-15) : null;
            
            draftSheet.getRange((obj.row-1)*3+3,obj.col+1,3,1).setBackgrounds([[`#${hexDesat}`],[`#${hex}`],[`#${hex}`]]);
            
            let rostersSheet = ss.getSheetByName(`ROSTERS`);
            rostersSheet.getRange((opening-1)*3+3,(obj.col-1)*3+2).setValue(team)
            rostersSheet.getRange((opening-1)*3+4,(obj.col-1)*3+3,2,1).setValues([[first],[last]]);
            rostersSheet.getRange((opening-1)*3+3,(obj.col-1)*3+2,3,2).setBackgrounds([[`#${hexDesat}`,`#${hexDesat}`],[`#${hex}`,`#${hex}`],[`#${hex}`,`#${hex}`]]);

            ss.toast(`Draft pick recorded`,`✏️ RECORDED`);
            if (draftComplete) {
              try {
                stopDrafting(draftComplete);
              } catch (err) {
                ui.alert(`⚠️ ERROR STOPPING DRAFT`,`Issue with draft completion, please review logs for more details.`,ui.ButtonSet.OK);
                Logger.log(`⚠️ Error stopping the draft: ${err.stack}`);
              }
            }
          } else {
            range.setValue(false);
            ss.toast(`Declined to confirm pick, ${obj.picker} is still picking!`,`⌛ ON THE CLOCK`);
          }
        }
      }
    }
  }
}

// UNDO LAST PICK
// Associated with the flag to remove the most recently picked player
function undoPick() {
  deleteTriggers();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const lightGray = `#E5E5E5`;
  let picksId = ss.getRangeByName(`PICKS_ID`).getValues().flat();
  let picksName = ss.getRangeByName(`PICKS_NAME`).getValues().flat();
  let target = picksId.indexOf('') - 1;
  let id = picksId[target]; 
  
  let picker = picksName[target];
  let sleeperId, sleeperName, player, team;
  let prompt = `CANCEL`;
  if ( picksId[0] == '' ) {
    ui.alert(`✔️ NOTHING TO REVERT`,`No selections have been made, get drafting!`, ui.ButtonSet.OK);
  } else {
    sleeperId = ss.getRangeByName(`SLPR_PLAYER_ID`).getValues().flat();
    sleeperName = ss.getRangeByName(`SLPR_FULL_NAME`).getValues().flat();
    player = sleeperName[sleeperId.indexOf(id)];
    prompt = ui.alert(`🔄 UNDO`,`Revert last pick of ${player} made by ${picker}?`, ui.ButtonSet.OK_CANCEL);
  }
  if ( prompt == 'OK' ) {
    
    let picksRow = ss.getRangeByName(`PICKS_ROW`).getValues().flat();
    let picksCol = ss.getRangeByName(`PICKS_COL`).getValues().flat();
    
    let picksRosterRow = ss.getRangeByName(`PICKS_ROSTER_ROW`).getValues().flat();
    let picksStarters = ss.getRangeByName(`PICKS_REMAINING`).getValues();

    let row = picksRow[target];
    let col = picksCol[target];
    let count = target+1;
    
    let nextPicker = picksName[target+1];
    let rosterRow = picksRosterRow[target];
    let currentRoster = picksStarters[target];

    let pickerRosters = picksStarters.filter((_value, index) => (picksName[index] == picker && picksStarters[index].some(x => x.length > 0)));
    let previousRoster = pickerRosters[pickerRosters.length-2];

    // Reveal picked player
    let unhideRow = ss.getRangeByName(`DRAFT_ID`).getValues().flat().indexOf(id);
    let lobbySheet = ss.getSheetByName(`DRAFT_LOBBY`);
    lobbySheet.getRange(unhideRow+3,2).clearContent();
    lobbySheet.unhideRow(lobbySheet.getRange(unhideRow+3,1));

    // Define and clear range for `DRAFT` sheet
    let range = ss.getSheetByName(`DRAFT`).getRange((row-1)*3+3,col+1,3,1);
    range.clearContent();
    if ( row % 2 == 0 ) {
      range.setBackground(lightGray);
    } else {
      range.setBackground(`white`);
    }

    // Define and clear range for `ROSTERS` sheet
    let rowRosters = rosterRow;
    // Clear first and last name
    ss.getSheetByName(`ROSTERS`).getRange((rowRosters-1)*3+4,(col-1)*3+3,2,1).clearContent();
    // Clear team abbreviation
    ss.getSheetByName(`ROSTERS`).getRange((rowRosters-1)*3+3,(col-1)*3+2).clearContent();
    // Get Range of full box and set color
    range = ss.getSheetByName(`ROSTERS`).getRange((rowRosters-1)*3+3,(col-1)*3+2,3,2);
    if ( rowRosters % 2 == 0 ) {
      range.setBackground(lightGray);
    } else {
      range.setBackground(`white`);
    }

    // Define range that will be cleared out
    let rangePicksSheet = ss.getSheetByName(`PICKS`).getRange(count+1,6,1,currentRoster.length + 4);
    
    picksRosterRow = ss.getRangeByName(`PICKS_ROSTER_ROW`).getValues().flat();
    picksStarters = ss.getRangeByName(`PICKS_REMAINING`).getValues();

    rangePicksSheet.clearContent();
    
    ss.getRangeByName(`PICKER`).setValue(picker);
    ss.getRangeByName(`PICKER_NEXT`).setValue(nextPicker);
    ss.getRangeByName(`PICKER_ROSTER`).setValue(arrayToString(arrayClean(previousRoster),false,false)); // Cleans up the previous roster to place on DRAFT_LOBBY sheet
    ui.alert(`🔄 REVERTED`,`Last pick of ${player} removed.\r\n\r\n⌛ ${picker} is back on the clock!`,ui.ButtonSet.OK);
  } else {
    ss.toast(`Undo of last pick canceled`,`❌ UNDO CANCELED`);
  }
  triggersDrafting();
}

// MASTER FUNCTION TO TOGGLE VISIBILITY
// Two options: the fist when checked will hide all drafted players and all players of a position who cannot be drafted by anyone else
function toggleMarked(e,ss,eliminated) {
  let range = e.range;
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if( range.getRowIndex() == 1 && range.getColumnIndex() == 2 && ss.getSheetName() == `DRAFT_LOBBY` ) {
    let activeSheet = ss.getActiveSheet();
    if ( activeSheet !== null ) {
      let draftedRange = ss.getRangeByName(`DRAFT_CHECKBOXES`);
      let drafted = draftedRange.getValues().flat();
      let hidden = Array(draftedRange.getRow()-1).fill(0); // Populates with 0s for first rows
      if( range.getCell(1,1).getValue() === true ) {
        for (let a = 0 ; a < drafted.length ; a++) {
          if (drafted[a] == true) {
            hidden.push(1);
          } else {
            hidden.push(0);
          }
        }
        let elim = true;
        if (eliminated) {
          if (eliminated.length > 0) {
            hidden = hideEliminated(ss,eliminated,hidden);    
          } else {
            elim = false;
          }
        } else {
          elim = false;
        }
        hideRowsUtility(activeSheet,hidden);
        if (elim) {
          ss.toast(`${eliminated} All marked and eliminated rows hidden`,`😶‍🌫️ HIDDEN`);
        } else {
          ss.toast(`All marked rows hidden`,`😶‍🌫️ HIDDEN`);
        }
      } else {
        let firstRow = 3;
        let lastRow = drafted.length;
        activeSheet.showRows(firstRow, lastRow);
        
        ss.toast(`All marked rows revealed`, `👀 REVEALED`);
      }
    }
  }
}

// ADDS OR CREATES BINARY ARRAY OF ROWS TO HIDE BASED ON ELIMINATED POSITIONS FROM ALL ROSTERS
function hideEliminated(ss,eliminated,hidden) {
  if (eliminated.length > 0) {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    const positionsRange = ss.getRangeByName(`DRAFT_POS`);
    const positions = Array(positionsRange.getRow()-1).fill('').concat(positionsRange.getValues().flat());
    hidden = hidden || Array(positionsRange.getRow()-1).fill(0).concat(Array(positions.length).fill(0)); // Populates with 0s for first rows
    for (let a = positionsRange.getRow()-1; a < positions.length; a++) {
      if (eliminated.indexOf(positions[a]) >= 0) {
        hidden[a] = 1;
      }
    }
  }
  return hidden;
}

// HIDES ROWS IN BATCHES TO REDUCE RUNTIME
function hideRowsUtility(sheet,arr) {
  try { 
    let start = null, end = null;
    for (let a = 0; a < arr.length; a++) {
      if (arr[a] === 1) {
        if (start === null) {
          start = a + 1;
          end = a + 1;
        } else {
          end = a + 1;
        }
      } else if (start !== null) {
        // Hide the range of rows and reset
        sheet.hideRows(start, end - start + 1);
        start = null;
        end = null;
      }
    }
    if (start !== null) {
      sheet.hideRows(start, end - start + 1);
    }
    return true;
  }
  catch (err) {
    return false;
  }
}


// DRAFT LIST
// Creates a sheet that has all eligible players on it (by Sleeper ID) and has their ordered spot based on projection
function draftList() {
  deleteTriggers();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName(`DRAFT_LIST`) || ss.insertSheet(`DRAFT_LIST`);

  sheet.clear();
  
  let sourceID = ss.getRangeByName(`SLPR_PLAYER_ID`).getValues().flat();
  let sourcePOS = ss.getRangeByName(`SLPR_FANTASY_POSITIONS`).getValues().flat();
  
  const docProps = PropertiesService.getDocumentProperties();
  try {
    
    const roster = JSON.parse(docProps.getProperty('roster'));
    let positions = [`QB`,`RB`,`WR`,`TE`,`K`,`DEF`];
    
    positions = positions.filter(pos => roster.find(entry => entry.pos === pos)?.count !== 0);
    
    // Object for referring what duplicates exist
    let positionDups = roster.reduce((acc, item) => {
        acc[item.pos] = item.duplicates;
        return acc;
    }, {});

    let positionRank = Array(positions.length).fill(1);
    
    let arr = [];
    let data = [];
    let count = 1;
    for (let a = 0; a < sourcePOS.length; a++){
      if (positions.indexOf(sourcePOS[a]) >= 0) {
        for (let b = 0; b <= positionDups[sourcePOS[a]]; b++) {
          arr = [count];
          count++;
          arr.push(sourceID[a]);
          arr.push(positionRank[positions.indexOf(sourcePOS[a])]);
          data.push(arr);
        }
        positionRank[positions.indexOf(sourcePOS[a])] = positionRank[positions.indexOf(sourcePOS[a])] + 1;
      }
    }

    headers = [`RNK`,`ID`,`POS_RNK`];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    let dataRange = sheet.getRange(2,1,data.length,headers.length);
    dataRange.setValues(data);
    dataRange.setHorizontalAlignment(`right`);
    sheet.setColumnWidths(1,headers.length,70);
    adjustRows(sheet,data.length+1);
    adjustColumns(sheet,headers.length);

    ss.setNamedRange(`DRAFT_LIST_RNK`,sheet.getRange(2,headers.indexOf(`RNK`)+1,count-1,1));
    ss.setNamedRange(`DRAFT_LIST_ID`,sheet.getRange(2,headers.indexOf(`ID`)+1,count-1,1));
    ss.setNamedRange(`DRAFT_LIST_POS_RNK`,sheet.getRange(2,headers.indexOf(`POS_RNK`)+1,count-1,1));

    ss.toast(`Created list of draftable players.`,`🧾 DRAFT LIST CREATED`);
    createTriggers();
  } catch (err) {
    Logger.log(`⚠️ Error setting up draftable players | ${err.stack}`);
  }
}

// EXISTING DRAFT
// Function to check if there are any picked entries on the PICKS sheet and return an array with both picked quantity and remaining to pick quantity
function existingDraft(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const ids = ss.getRangeByName(`PICKS_ID`).getValues().flat();
    const owners = ss.getRangeByName(`PICKS_NAME`).getValues().flat();
    let picked = owners.filter((value, index) => (typeof ids[index] === `number` && owners[index] != `End`));
    let remaining = owners.filter((value, index) => (typeof ids[index] === `string` && ids[index].trim() === '' && owners[index] != `End` && owners[index].length > 0));
    return [picked.length,remaining.length];
  } catch (err) {
    Logger.log(`⚠️ PICKS sheet missing or some other error. ${err.stack}`);
    ss.toast(`Attempted to find a previous draft and was unable to located one. Moving on...`,`▶️ NO PRIOR DRAFTS`)
    return null;
  }
}

// START DRAFTING
// Creates the dynamicDrafter trigger and prompts for first draft picker as well as displays initial roster
function startDrafting(){
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const docProps = PropertiesService.getDocumentProperties();
    const configuration = JSON.parse(docProps.getProperty('configuration')) || {};
    
    const trr = configuration.thirdRoundReversal;
    const pprText = configuration.ppr;
    
    const roster = JSON.parse(docProps.getProperty('roster')).filter(item => item.count > 0);

    // const positionsRanges = ['QB','RB','WR','TE','FLX','SPFLX','DEF','K'];

    // ss.toast(roster,`ROSTER:`);
    const rosterPrompt = roster.flatMap(item => Array(item.count).fill(positionsEmojiMap[item.pos]).join("") + item.pos);
    let rosterFresh = `Starting Roster:
    
    ${rosterPrompt.join('\n')}`;
    const picker = ss.getRangeByName('PICKER').getValue();
    const startDraft = ui.alert(`🚀 LET'S GO!`,`${rosterFresh}
    
    ⌛ ${picker} is on the clock.
    
    Use the Challenge Flag to UNDO the last pick.`,ui.ButtonSet.OK_CANCEL);
    if (startDraft == 'OK') {
      configuration.drafting = true;
      docProps.setProperty('configuration',JSON.stringify(configuration));
      
      triggersDrafting();

      scoringHide();

      toolbar();
    } else {
      ss.toast(`User canceled draft start. Please try again when you're ready!`,`⏸ DRAFT NOT STARTED`)
    }

  } catch (err) {
    Logger.log(`⚠️Error starting the draft: ${err.stack}`)
    ui.alert(`⚠️ DRAFT START ERROR`,`Issue with starting the draft. Please try again and ensure everything was configured prior to running this
    
    ${err.stack}`,ui.ButtonSet.OK);
  }
}

// Refresh triggers if errors being experienced
function triggersDrafting() {
  deleteTriggers();
  createTriggers();
}

// Deletes all triggers in the current project.
function deleteTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  let scoring = false;
  for (let a = 0; a < triggers.length; a++) {
    if ( triggers[a].getHandlerFunction() == `dynamicDrafter` ) {
      ScriptApp.deleteTrigger(triggers[a]);
    }
    if ( triggers[a].getHandlerFunction() == `sleeperScoringAuto` ) {
      ScriptApp.deleteTrigger(triggers[a]);
      scoring == false ? scoring = true : null;
    }
  }
  scoring == true ? sleeperLiveScoringOn() : null;
}

// Creates an edit trigger for a spreadsheet identified by ID.
function createTriggers() {
  // createOnOpen();
  createDrafterTrigger();
}

// Creates an edit trigger for a spreadsheet identified by ID.
function createDrafterTrigger() {
  ScriptApp.newTrigger(`dynamicDrafter`)
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet().getId())
    .onEdit()
    .create();
}

function stopDrafting(draftComplete) {
  try {  
    const docProps = PropertiesService.getDocumentProperties();
    const configuration = JSON.parse(docProps.getProperty('configuration')) || {};
    configuration.draftComplete = draftComplete ? true : false;
    configuration.drafting = false;
    deleteTriggers(); 
    docProps.setProperty('configuration',JSON.stringify(configuration));
    // Reload toolbar now that config has changed
    toolbar();
    Logger.log(`🛑 Draft stopped and configuration saved: ${JSON.stringify(configuration)}`);
  } catch (err) {
    Logger.log(`⚠️ Issue stopping the draft: ${err.stack}`);
  }
}

function resetDraft() {
  const ui = SpreadsheetApp.getUi();
  let confirm = ui.alert(`♻️ RESET DRAFT?`,`Would you like to remove all data
  and prepare for another draft?`, ui.ButtonSet.YES_NO);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let draftComplete = fetchDraftComplete();
  try {
    if (!draftComplete) {
      const existing = existingDraft();
      draftComplete = (existing[1] == 0 && existing[0] > 0);
    }
  } catch (err) {
    Logger.log(`❕ Encountered an issue parsing previous draft data or it was unavailable. Moving on...`)
  }
  let backup, failed, promptText = '';
  if (draftComplete) {
    backup = ui.alert(`🔎 EXISTING DATA FOUND`,`You have a previous draft detected.
    
    Do you want to make a copy before proceeding?`,ui.ButtonSet.YES_NO_CANCEL)
  }
  if (backup == 'YES') {
    try {
      let sheetName;
      try {
        let namedRanges = ss.getNamedRanges();
        for(let a = 0; a < namedRanges.length; a++){
          if (namedRanges[a].getRange().getSheet().getName() == 'ROSTERS') {
            namedRanges[a].remove();
          }
        }
      }
      catch (err) {
        Logger.log(`"ROSTERS" backup failed to remove existing named ranges`);
      }
      let sheetNames = ss.getSheets().map(sheet => sheet.getName());
      let rosterSheets = sheetNames.filter(sheetName => /^ROSTERS_\d{2}$/.test(sheetName));
      if (rosterSheets.length === 0) {
        Logger.log(`No matching sheets found.`);
        sheetName = `ROSTERS_01`;
      } else {
        let highestIndex = Math.max(...rosterSheets.map(sheetName => parseInt(sheetName.match(/\d{2}$/)[0], 10)));
        let index = highestIndex;
        highestIndex < 9 ? index = `0` + (index+1) : index++;
        sheetName = `ROSTERS_` + index;
      }
      ss.getSheetByName(`ROSTERS`).copyTo(ss).setName(sheetName);
      ss.toast(`Backed up previous draft to "${sheetName}".`,`💾 BACKUP COMPLETE`);
    }
    catch (err) {
      Logger.log(err.stack)
      failed = ui.alert(`⚠️ ERROR`,`Error encounter while trying to copy over existing "ROSTERS" sheet.\r\n\r\nWould you still like to continue?`, ui.ButtonSet.YES_NO);
    }
    if (failed == 'NO') {
      ss.toast(`Canceled setup`);
      Logger.log(`Canceled setup`)
      return null;
    }
  }
  if (confirm == 'YES') {
    const docProps = PropertiesService.getDocumentProperties();
    let configurationString = docProps.getProperty('configuration');
    let ppr = 0.5;
    if (configurationString) {
      try {
        ppr = JSON.parse(configurationString).ppr;
      } catch (err) {
        Logger.log(`No PPR stored configuration found, defaulting to half (${err.stack})`);
      }
    }
    docProps.deleteProperty('configuration');
    docProps.deleteProperty('matchups');
    ss.toast(`Deleted matchups and draft status.`,`🧼 MATCHUP CLEAN`);

    if (docProps.getProperty('members')) {
      let memberDelete = ui.alert('💾 KEEP MEMBERS?',`Would you like to retain the 
      existing ${JSON.parse(docProps.getProperty('members')).length} members?`, ui.ButtonSet.YES_NO);
      if (memberDelete === ui.Button.NO) {
        docProps.deleteProperty('members');
        ss.toast(`Removed all previously entered members and team names.`,`🧼 MEMBERS CLEAN`);
      }
    }
    
    if (docProps.getProperty('roster')) {
      let rosterDelete = ui.alert('💾 KEEP ROSTER?',`Would you like to retain the previously
      configured team rosters?`, ui.ButtonSet.YES_NO);
      if (rosterDelete === ui.Button.NO) {
        docProps.deleteProperty('roster');
        ss.toast(`Removed roster settings.`,`🧼 ROSTER CLEAN`);
      } else {
        docProps.setProperty('configuration',JSON.stringify({'ppr':ppr}));
      }
    }
    
    // CLEAR OUT PICKS, ROSTERS, AND DRAFT DATA
    let successes = [];
    let failures = [];
    try {
      const picksSheet = ss.getSheetByName('PICKS');
      picksSheet.getRange(2,1,picksSheet.getMaxRows()-1,picksSheet.getMaxColumns()).clear({contentsOnly: true});
      successes.push('PICKS');
    } catch (err) {
      failures.push('PICKS');
      Logger.log(`Failed to remove PICKS data: ${err.stack}`);
    }
    try {
      const draftSheet = ss.getSheetByName('DRAFT');
      draftSheet.getRange(1,2,draftSheet.getMaxRows(),draftSheet.getMaxColumns()-1)
        .setBackground(null)
        .clear({contentsOnly: true});
      successes.push('DRAFT');
    } catch (err) {
      failures.push('DRAFT');
      Logger.log(`Failed to remove DRAFT data: ${err.stack}`);
    }
    try {
      const rostersSheet = ss.getSheetByName('ROSTERS');
      rostersSheet.getRange(1,2,rostersSheet.getMaxRows(),rostersSheet.getMaxColumns()-1)
        .setBackground(null)
        .clear({contentsOnly: true});        
      successes.push('ROSTERS');
    } catch (err) {
      failures.push('ROSTERS');
      Logger.log(`Failed to remove ROSTERS data: ${err.stack}`);
    }
    try {
      const playersSheet = ss.getSheetByName('PLAYERS');
      playersSheet.getRange(2,1,playersSheet.getMaxRows()-1,playersSheet.getMaxColumns()).clear({contentsOnly: true, formatOnly: true});
      successes.push('PLAYERS');
    } catch (err) {
      failures.push('PLAYERS');
      Logger.log(`Failed to remove PLAYERS data: ${err.stack}`);
    }
    if (successes.length > 0) ss.toast(`Removed sheet contents for: ${successes.join(', ')}.`,`🧼 REMOVED DATA`);
    if (failures.length > 0) ss.toast(`Failed to remove sheet contents for: ${failures.join(', ')}, this shouldn't prevent future drafts.`,`⚠️ DATA REMOVAL ERROR`);

    // CLEAR OUT DRAFT BOARD
    try {
      draftLobbyClean(ss,true);
    } catch (err) {
      Logger.log(`⚠️ Issue clearing out the draft lobby: ${err.stack}`);
      ss.toast(`The script encountered an error when attempting to clear out the draft board, re-running the setup later should still resolve it.`, `⚠️ DRAFT BOARD CLEAN FAILED`)
    }
    const another = ui.alert(`✨ NEW DRAFT?`, `Would you like to launch the configuration
    panel for running a new draft now?`,ui.ButtonSet.YES_NO);
    if (another == 'YES') draftSetup();
    toolbar();
    ss.toast(`Script completed for draft reset.`,`✅ RESET DONE`);

  } else {
    ss.toast('Canceled the clearing out of content. Try again later if desired.',`🚫 RESET CANCELED`)
  }
    
}


// DRAFT BOARD CLEAN
// Function to clean up draft board
function draftLobbyClean(ss,done) {
  deleteTriggers();
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const darkNavy = `#222735`;
  const lobbySheet = ss.getSheetByName(`DRAFT_LOBBY`);
  const rankSheet = ss.getSheetByName(`DRAFT_LIST`);
  const draftHeaders = ss.getRangeByName(`DRAFT_HEADERS`).getValues().flat();
  let players = rankSheet.getLastRow()-1;
  let firstRowFormulas = lobbySheet.getRange(3,1,1,lobbySheet.getMaxColumns()).getFormulas().flat();
  let formulaCols = firstRowFormulas.map(x => x != '');

  let draftRows = lobbySheet.getMaxRows();

  // Adding and removing rows as needed
  if (!done) {
    adjustRows(lobbySheet,ss.getRangeByName(`DRAFT_LIST_ID`).getNumRows()+2);
    draftRows = lobbySheet.getMaxRows();
  }

  // Reset all indices down column 2, as well as propogate all formulas down the specified columns in formulaCols array
  let index = Array.from({length: players}, (_, i) => [i + 1]);
  lobbySheet.getRange(3,3,draftRows-2,1).setValues(index);
  
  for (let a = 0; a < formulaCols.length; a++ ) {
    if ( formulaCols[a] == true ) {
      let draftFormula = lobbySheet.getRange(3,a+1).getFormulaR1C1();
      for ( let b = 1; b < draftRows - 2; b++ ) {
        lobbySheet.getRange(b+3,a+1).setFormulaR1C1(draftFormula);
      }
    }
  }

  let sleeperId = ss.getRangeByName(`SLPR_PLAYER_ID`).getValues().flat();
  let sleeperOutlook = ss.getRangeByName(`SLPR_OUTLOOK`).getValues().flat();
  let id, fullNotes = [];
  for (let a = 3; a <= draftRows; a++ ) {
    id = lobbySheet.getRange(a,1).getValue();
    fullNotes.push([sleeperOutlook[sleeperId.indexOf(id)] || '']);
  }
  // Set all notes at the same time
  lobbySheet.getRange(3,5,(draftRows - 2),1).setNotes(fullNotes);

  lobbySheet.getRange(3,1,draftRows - 2,lobbySheet.getMaxColumns())
    .setBorder(false,false,false,false,false,true,darkNavy,SpreadsheetApp.BorderStyle.SOLID_THICK);

  const checkboxesRange = lobbySheet.getRange(3,2,draftRows - 2,1);
  let uncheckedValues = Array(draftRows - 2).fill([false]);
  checkboxesRange.setValues(uncheckedValues);
  ss.setNamedRange(`DRAFT_ID`,lobbySheet.getRange(3,draftHeaders.indexOf(`ID`)+1,draftRows - 2,1));
  ss.setNamedRange(`DRAFT_CHECKBOXES`,checkboxesRange);
  ss.setNamedRange(`DRAFT_PLAYER`,lobbySheet.getRange(3,draftHeaders.indexOf(`PLAYER`)+1,draftRows - 2,3));
  ss.setNamedRange(`DRAFT_POS`,lobbySheet.getRange(3,draftHeaders.indexOf(`POS`)+1,draftRows - 2,1));
  const healthRange = lobbySheet.getRange(3,draftHeaders.indexOf(`HEALTH`)+1,draftRows - 2,1);
  healthRange.clearNote();
  ss.setNamedRange(`DRAFT_HEALTH`,healthRange);
  
  if (done) {
    ss.getSheetByName('DRAFT_LIST').clear();
    ss.getRangeByName('PICKER').setValue('READY?');
    ss.getRangeByName('PICKER_ROSTER').setValue('⬆️ Use "🏈 Football Tools" above');
    ss.getRangeByName('PICKER_NEXT').setValue('');
  }

  // Prep for filter values
  let ssId = ss.getId();
  let lastRow = lobbySheet.getLastRow();
  let lastColumn = lobbySheet.getLastColumn();
  const sheetId = lobbySheet.getSheetId();
  
  // Filter specifics
  const filterSettings = {
    "range": {
      "sheetId": sheetId,
      "startRowIndex": 1,
      "endRowIndex": lastRow,
      "startColumnIndex": 0,
      "endColumnIndex": lastColumn
    }
  };
  
  // Filter request
  let requests = [
    {
      "clearBasicFilter": {
        "sheetId": sheetId
      }
    },
    {
      "updateDimensionProperties": {
        "range": {
          "sheetId": sheetId,
          "dimension": "ROWS",
          "startIndex": 2,  // Start from row 3 (0-indexed)
          "endIndex": lastRow
        },
        "properties": {
          "hiddenByUser": false
        },
        "fields": "hiddenByUser"
      }
    },
    {
      "setBasicFilter": {
        "filter": filterSettings
      }
    }
  ];
  
  // Pushing the API request
  const url = "https://sheets.googleapis.com/v4/spreadsheets/" + ssId + ":batchUpdate";
  const params = {
    method:"post",
    contentType: "application/json",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    payload: JSON.stringify({"requests": requests}),
    muteHttpExceptions: true,
  };
  
  let res = UrlFetchApp.fetch(url, params).getContentText();
  ss.toast(`Draft lobby scrubbed of all data successfully.`,`🧽 DRAFT BOARD`);
}

// 2025 - Created by Ben Powers
// ben.powers.creative@gmail.com
