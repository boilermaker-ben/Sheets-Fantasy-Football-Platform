// WEEKLY FF SCORING - Updated 11.25.2025

// RECAP panel and final champion announcement tool
function recapPanel() {
  const html = HtmlService.createHtmlOutputFromFile('recapPanel')
      .setWidth(1100)
      .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Contest Completed!');
}

/**
 * Gathers and processes all data from the sheets and properties to build the recap.
 * This version is rewritten to use your specific named ranges.
 * @returns {object} A comprehensive object with all the calculated stats for the recap panel.
 */
function fetchRecapData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const playerMap = buildPlayerMap(ss);
    const picks = buildPicksArray(ss);
    const draftedPlayerIds = new Set(picks.map(p => p.id));
    const members = JSON.parse(PropertiesService.getDocumentProperties().getProperty('members') || '[]');

    const memberScores = calculateAllMemberScores(playerMap, picks, members);
    
    const sortedScores = Object.values(memberScores).sort((a, b) => b.score - a.score);
    const winners = getWinners(sortedScores);
    const podium = getPodium(sortedScores);

    const topPerformers = {};
    const bottomPerformers = {};
    // --- CHANGE 1: ADD DEF AND K ---
    const positions = ['QB', 'RB', 'WR', 'TE', 'DEF', 'K']; 
    positions.forEach(pos => {
      topPerformers[pos] = findPerformerByPosition(playerMap, picks, pos, 'top');
      bottomPerformers[pos] = findPerformerByPosition(playerMap, picks, pos, 'bottom');
    });

    const unownedOverperformer = findUnownedOverperformer(playerMap, draftedPlayerIds);
    const draftGem = findDraftValue(playerMap, picks, 'gem');
    const draftDud = findDraftValue(playerMap, picks, 'dud');

    return {
      members: members,
      winners,
      podium,
      highestScorer: sortedScores[0],
      lowestScorer: sortedScores[sortedScores.length - 1],
      topPerformers,
      bottomPerformers,
      unownedOverperformer,
      draftGem,
      draftDud
    };

  } catch (e) {
    Logger.log('Error in fetchRecapData: ' + e.stack);
    throw new Error('Could not generate recap data. Ensure all named ranges (SLPR_..., PICKS_...) are correct.');
  }
}


// --- NEW HELPER FUNCTIONS TAILORED TO YOUR NAMED RANGES ---

/**
 * Builds a Map of player data from individual SLPR_* named ranges.
 * @param {Spreadsheet} ss The active spreadsheet object.
 * @returns {Map<string, object>} A Map where keys are player IDs.
 */
function buildPlayerMap(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const playerMap = {};
  const ids = ss.getRange('SLPR_PLAYER_ID').getValues().flat();
  const names = ss.getRange('SLPR_FULL_NAME').getValues().flat();
  const positions = ss.getRange('SLPR_FANTASY_POSITIONS').getValues().flat();
  const scores = ss.getRange('SLPR_SCORE').getValues().flat();
  const projs = ss.getRange('SLPR_PROJ').getValues().flat();
  const images = ss.getRange('SLPR_IMAGE').getValues().flat();

  for (let i = 0; i < ids.length; i++) {
    const playerId = ids[i];
    if (playerId && !playerMap.hasOwnProperty(playerId)) {
      playerMap[playerId] = {
        id: playerId,
        name: names[i],
        pos: positions[i],
        // --- THIS IS THE FIX ---
        // Using Number() || 0 ensures the score is always a valid number.
        // It safely handles empty cells, text ("DNP"), or other non-numeric data.
        score: Number(scores[i]) || 0,
        proj: Number(projs[i]) || 0,
        image: images[i]
      };
    }
  }
  return playerMap;
}

/**
 * Builds an array of pick objects from PICKS_* named ranges.
 * The array index corresponds to the pick number.
 * @param {Spreadsheet} ss The active spreadsheet object.
 * @returns {Array<object>} An array of pick objects.
 */
function buildPicksArray(ss) {
  const picks = [];
  const ids = ss.getRange('PICKS_ID').getValues().flat();
  const pickers = ss.getRange('PICKS_NAME').getValues().flat();

  for (let i = 0; i < ids.length; i++) {
    if (ids[i]) {
      picks.push({
        pickNumber: i + 1,
        id: ids[i],
        picker: pickers[i]
      });
    }
  }
  return picks;
}


// --- REVISED HELPER FUNCTIONS ---

function calculateAllMemberScores(playerMap, picks, members) {
  const scores = {};
  picks.forEach(pick => {
    const ownerName = pick.picker;
    const player = playerMap[pick.id];
    if (player) {
      if (!scores[ownerName]) {
        const memberInfo = members.find(m => m.name === ownerName) || {};
        scores[ownerName] = { name: ownerName, teamName: memberInfo.teamName || `Team ${ownerName}`, score: 0 };
      }
      scores[ownerName].score += player.score;
    }
  });
  return scores;
}

function findPerformerByPosition(playerMap, picks, position, type = 'top') {
  const draftedPlayerIds = new Set(picks.map(p => p.id));
  
  // 1. Filter all players down to only those who were drafted and match the position.
  const draftedPositionPlayers = Object.values(playerMap).filter(p => 
    draftedPlayerIds.has(p.id) && p.pos.includes(position)
  );

  if (draftedPositionPlayers.length === 0) return null;
  
  // 2. Sort the filtered list.
  const sorted = draftedPositionPlayers.sort((a, b) => (type === 'top' ? b.score - a.score : a.score - b.score));
  const performer = sorted[0];
  
  // 3. Find the owner(s) from the picks array.
  const owners = picks.filter(p => p.id === performer.id).map(p => p.picker);
  
  return {
    name: performer.name,
    score: performer.score,
    image: performer.image,
    owners: owners.length > 0 ? [...new Set(owners)] : [] // Use Set to remove duplicate owners
  };
}

function findUnownedOverperformer(playerMap, draftedPlayerIds) {
  // Filter for players who were NOT drafted and scored more than 0
  const undrafted = Object.values(playerMap).filter(p => 
    !draftedPlayerIds.has(p.id) && p.score > 0
  );

  if (undrafted.length === 0) return null;

  // Sort by score to find the best one
  const sorted = undrafted.sort((a, b) => b.score - a.score);
  const topUnowned = sorted[0];

  return {
    name: topUnowned.name,
    score: topUnowned.score,
    image: topUnowned.image,
    owners: ['Undrafted'] // Special owner category
  };
}


function findDraftValue(playerMap, picks, type = 'gem') {
  let bestCandidate = null;

  picks.forEach(pick => {
    const player = playerMap[pick.id];
    if (!player || ['DEF', 'K'].some(pos => player.pos.includes(pos))) return;

    if (type === 'gem') {
      const latePickThreshold = Math.floor(picks.length * 0.85);
      const scoreThreshold = player.pos.includes('QB') ? 24 : 12;
      if (pick.pickNumber >= latePickThreshold && player.score >= scoreThreshold) {
        if (!bestCandidate || player.score > bestCandidate.score) {
          bestCandidate = { ...player, pickNumber: pick.pickNumber, owner: pick.picker };
        }
      }
    } else { // type === 'dud'
      const earlyPickThreshold = Math.ceil(picks.length * 0.10);
      const scoreThreshold = player.pos.includes('QB') ? 12 : 6;
      if (pick.pickNumber <= earlyPickThreshold && player.score <= scoreThreshold) {
        if (!bestCandidate || player.score < bestCandidate.score) {
          bestCandidate = { ...player, pickNumber: pick.pickNumber, owner: pick.picker };
        }
      }
    }
  });

  if (!bestCandidate) return null;
  return { name: bestCandidate.name, score: bestCandidate.score, image: bestCandidate.image, owners: [bestCandidate.owner], pickNumber: bestCandidate.pickNumber };
}


/**
 * Identifies the winner(s) of the competition.
 */
function getWinners(sortedScores) {
  if (sortedScores.length === 0) return [];
  const topScore = sortedScores[0].score;
  const winners = sortedScores.filter(m => m.score === topScore);
  return [{
    score: topScore,
    isTie: winners.length > 1,
    members: winners
  }];
}

/**
 * Identifies the 2nd and 3rd place finishers.
 */
function getPodium(sortedScores) {
  if (sortedScores.length < 2) return {};
  const topScore = sortedScores[0].score;
  const secondPlace = sortedScores.find(m => m.score < topScore);
  if (!secondPlace) return { second: null, third: null };
  const secondScore = secondPlace.score;
  const thirdPlace = sortedScores.find(m => m.score < secondScore);
  return { second: secondPlace, third: thirdPlace || null };
}

/**
 * Utility to convert a 2D array from a range into an array of objects.
 * Assumes the first row is headers.
 */
function getRangeDataAsObjects(range) {
    const values = range.getValues();
    const headers = values.shift();
    return values.map(row => {
        const obj = {};
        headers.forEach((header, index) => {
            obj[header] = row[index];
        });
        return obj;
    });
}
