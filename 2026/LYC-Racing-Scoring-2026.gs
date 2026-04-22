/**
 * Liberty Yacht Club Racing Scoring System - 2026
 * Automates PHRF scoring, series standings, and season standings
 *
 * 2026 Rule Change (per updated SIs):
 *   BYE races are scored as the average of the boat's other races
 *   in that series (including DNC, RET, etc. in the average).
 *   BYE still counts toward qualification.
 */

const SHEETS = {
  SCRATCH_SHEET: 'Scratch Sheet',
  RACE_ENTRY: 'Race Results Entry',
  CALCULATED: 'Calculated Results',
  SERIES: 'Series Standings',
  SEASON: 'Season Standings',
  CUMULATIVE: 'Cumulative Results'
};

const PHRF_TCF_DIVISOR = 650;
const PHRF_TCF_BASE = 550;

/**
 * Main function to recalculate all results
 * Run this after entering race data
 */
function calculateAllResults() {
  Logger.log('Starting calculation...');
  
  calculateRaceResults();
  calculateSeriesStandings();
  calculateSeasonStandings();
  calculateCumulativeResults();
  
  Logger.log('Calculation complete!');
  SpreadsheetApp.getActiveSpreadsheet().toast('Scoring updated successfully!', 'Complete', 3);
}

/**
 * Calculate race results from raw entry data
 *
 * 2026 BYE scoring uses a two-pass approach:
 *   Pass 1 - score all non-BYE races normally
 *   Pass 2 - set each BYE to the average of that boat's
 *            other scored races in the same series
 */
function calculateRaceResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrySheet = ss.getSheetByName(SHEETS.RACE_ENTRY);
  const calcSheet = ss.getSheetByName(SHEETS.CALCULATED);
  const boatSheet = ss.getSheetByName(SHEETS.SCRATCH_SHEET);
  
  // Get scratch sheet data
  const boatData = boatSheet.getDataRange().getValues();
  const boatMap = {};
  for (let i = 1; i < boatData.length; i++) {
    const boatName = String(boatData[i][0]);
    const sailNumber = String(boatData[i][1]);
    const phrf = boatData[i][7];
    boatMap[boatName] = { sailNumber, phrf };
  }
  
  // Get race entry data
  const entryData = entrySheet.getDataRange().getValues();
  const results = [];
  
  for (let i = 1; i < entryData.length; i++) {
    const row = entryData[i];
    const raceNum    = row[0];
    const series     = row[1];
    const raceType   = row[2];
    const course     = row[3];   // Course (col D)
    const boatName   = row[4];   // BoatName (col E)
    const classNum   = row[5];   // Class (col F - VLOOKUP from Scratch Sheet)
    const startTime  = row[8];   // StartDateTime (col I)
    const wind       = row[15];  // Wind (col P)
    const tide       = row[16];  // Tide (col Q)
    const finishTime = row[11];  // FinishDateTime (col L)
    const status     = row[12];  // Status (col M)
    
    if (!boatName) continue;
    
    const boat = boatMap[boatName];
    if (!boat) {
      Logger.log(`Warning: Boat "${boatName}" not found in scratch sheet`);
      continue;
    }
    
    const isPractice = String(raceType).startsWith('Practice');
    
    let elapsedSeconds = null;
    let correctedSeconds = null;
    
    if (finishTime && !status && !isPractice) {
      elapsedSeconds = calculateElapsedSeconds(startTime, finishTime);
      const tcf = PHRF_TCF_DIVISOR / (PHRF_TCF_BASE + boat.phrf);
      correctedSeconds = Math.round(elapsedSeconds * tcf);
    }
    
    results.push({
      date: startTime instanceof Date ? startTime : new Date(startTime),
      raceNum, series, raceType, classNum, course, boatName,
      sailNumber: boat.sailNumber,
      phrf: Number(boat.phrf),
      finishTime, elapsedSeconds, correctedSeconds, status, isPractice,
      place: null, points: null
    });
  }
  
  // ── PASS 1: Score all non-BYE, non-practice races ──────────────────────────
  const raceGroups = {};
  results.forEach(r => {
    if (r.isPractice) return;
    if (r.status === 'BYE') return; // skip BYEs in pass 1
    const key = `${r.raceNum}_${r.classNum}`;
    if (!raceGroups[key]) raceGroups[key] = [];
    raceGroups[key].push(r);
  });
  
  Object.values(raceGroups).forEach(group => {
    group.sort((a, b) => {
      if (a.correctedSeconds === null) return 1;
      if (b.correctedSeconds === null) return -1;
      return a.correctedSeconds - b.correctedSeconds;
    });
    
    const numStarters = group.filter(r => r.status !== 'DNC').length;
    let place = 1;
    const numFinishers = group.filter(r => !r.status).length;

    group.forEach(r => {
      if (r.status) {
        r.place = null;
        if (r.status === 'DNC' || r.status === 'DSQ' || r.status === 'DNE') {
          r.points = numStarters + 2;
        } else if (r.status === 'TLE') {
          r.points = numFinishers + 2;
        } else {
          r.points = numStarters + 1;
        }
      } else {
        r.place = place;
        r.points = place;
        place++;
      }
    });
  });

  // ── PASS 2: Score BYE races as average of boat's other races in same series ─
  //
  // For each BYE, find all other scored races for that boat in the same series
  // (same classNum + series), then average their points.
  // If the boat has no other races yet in that series, BYE scores 0
  // (will naturally update on recalculation as races are entered).

  results.forEach(r => {
    if (r.isPractice || r.status !== 'BYE') return;

    // Collect all non-BYE, non-practice points for this boat in this series
    const otherPoints = results
      .filter(other =>
        other !== r &&
        !other.isPractice &&
        other.status !== 'BYE' &&
        other.boatName === r.boatName &&
        other.classNum === r.classNum &&
        other.series === r.series &&
        other.points !== null &&
        other.points !== undefined
      )
      .map(other => other.points);

    if (otherPoints.length === 0) {
      // No other races scored yet — default to 0, will update on recalc
      r.points = 0;
      Logger.log(`BYE for ${r.boatName} in Series ${r.series} Class ${r.classNum}: no other races yet, scored 0`);
    } else {
      const sum = otherPoints.reduce((acc, p) => acc + p, 0);
      // Round to 2 decimal places; stored as string like other points
      r.points = Math.round((sum / otherPoints.length) * 100) / 100;
      Logger.log(`BYE for ${r.boatName} Series ${r.series} Class ${r.classNum}: avg of [${otherPoints.join(', ')}] = ${r.points}`);
    }
    r.place = null;
  });

  // ── Write to Calculated Results sheet ──────────────────────────────────────
  // Column order: A=Start DateTime, B=Race#, C=Series, D=RaceType, E=Class, F=Course,
  //               G=Boat Name, H=Sail#, I=PHRF, J=Finish DateTime, K=Elapsed,
  //               L=Corrected, M=Place, N=Points, O=Status
  calcSheet.clear();
  calcSheet.appendRow(['Start DateTime', 'Race#', 'Series', 'RaceType', 'Class', 'Course', 'Boat Name', 'Sail#', 'PHRF', 
                       'Finish DateTime', 'Elapsed', 'Corrected', 'Place', 'Points', 'Status']);
  
  results.forEach(r => {
    calcSheet.appendRow([
      r.date, r.raceNum, r.series, r.raceType, r.classNum, r.course, r.boatName, r.sailNumber, r.phrf,
      r.finishTime, r.elapsedSeconds, r.correctedSeconds, r.place,
      r.points !== null && r.points !== undefined ? String(r.points) : '',
      r.status
    ]);
  });

  const headerRange = calcSheet.getRange(1, 1, 1, 15);
  headerRange.setBackground('#4a86e8').setFontColor('white').setFontWeight('bold');
  if (results.length > 0) {
    calcSheet.getRange(2, 8, results.length, 1).setNumberFormat('@');  // Sail# col H as text
    calcSheet.getRange(2, 9, results.length, 1).setNumberFormat('0');  // PHRF col I as number
    calcSheet.getRange(2, 14, results.length, 1).setNumberFormat('@'); // Points col N as text
  }
}

/**
 * Calculate elapsed time in seconds
 */
function calculateElapsedSeconds(startTime, finishTime) {
  const start = startTime instanceof Date ? startTime : new Date(startTime);
  const finish = finishTime instanceof Date ? finishTime : new Date(finishTime);
  return Math.round((finish - start) / 1000);
}

/**
 * Calculate series standings
 * Calculated Results column indices (0-based):
 * 0=Start DateTime, 1=Race#, 2=Series, 3=RaceType, 4=Class, 5=Course,
 * 6=Boat Name, 7=Sail#, 8=PHRF, 9=Finish DateTime, 10=Elapsed,
 * 11=Corrected, 12=Place, 13=Points, 14=Status
 */
function calculateSeriesStandings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName(SHEETS.CALCULATED);
  const seriesSheet = ss.getSheetByName(SHEETS.SERIES);
  
  const data = calcSheet.getDataRange().getValues();
  
  const seriesGroups = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const series   = row[2];
    const raceType = row[3];
    const classNum = row[4];
    const boatName = row[6];
    const raceNum  = row[1];
    const points   = Number(row[13]);
    const status   = row[14];
    
    const isPractice = String(raceType).startsWith('Practice');
    if (isPractice) continue;
    
    const key = `${series}_${classNum}`;
    if (!seriesGroups[key]) seriesGroups[key] = {};
    if (!seriesGroups[key][boatName]) seriesGroups[key][boatName] = [];
    
    // 2026: BYE is now a scored race (average-based), include it
    seriesGroups[key][boatName].push({ raceNum, points });
  }
  
  seriesSheet.clear();
  let currentRow = 1;
  
  Object.keys(seriesGroups).sort().forEach(key => {
    const [series, classNum] = key.split('_');
    const boats = seriesGroups[key];
    
    seriesSheet.getRange(currentRow, 1).setValue(`CLASS ${classNum} - SERIES ${series} SCORES`);
    seriesSheet.getRange(currentRow, 1, 1, 10).setBackground('#ff9900').setFontWeight('bold');
    currentRow++;
    
    seriesSheet.appendRow(['Boat Name', 'Total Points', 'Throwouts', 'Net Score', 'Races']);
    seriesSheet.getRange(currentRow, 1, 1, 5).setBackground('#6d9eeb').setFontWeight('bold');
    currentRow++;
    
    const standings = [];
    Object.keys(boats).forEach(boatName => {
      const races = boats[boatName];
      const numRaces = races.length;
      const numThrowouts = Math.floor(numRaces / 7);
      
      const sortedPoints = races.map(r => r.points).sort((a, b) => b - a);
      const kept = sortedPoints.slice(numThrowouts);
      
      const total = sortedPoints.reduce((sum, p) => sum + p, 0);
      const score = kept.reduce((sum, p) => sum + p, 0);
      
      standings.push({ boatName, total, throwouts: numThrowouts, score, numRaces });
    });
    
    standings.sort((a, b) => a.score - b.score);
    
    standings.forEach(s => {
      seriesSheet.appendRow([s.boatName, s.total, s.throwouts, s.score, s.numRaces]);
      currentRow++;
    });
    
    currentRow += 2;
  });
}

/**
 * Calculate season standings with 75% Race Day participation threshold
 * A "Race Day" is a unique StartDate - boats are credited for the whole day
 * regardless of how many races were held that day.
 * Participation counts if status is anything except DNC (boat showed up).
 * 2026: BYE counts toward qualification (unchanged from 2025).
 */
function calculateSeasonStandings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName(SHEETS.CALCULATED);
  const seasonSheet = ss.getSheetByName(SHEETS.SEASON);
  
  const data = calcSheet.getDataRange().getValues();

  function toDateKey(val) {
    if (!val) return '';
    const d = val instanceof Date ? val : new Date(val);
    if (isNaN(d.getTime())) return '';
    if (d.getFullYear() < 2000) return '';
    return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
  }

  // Pass 1: collect all unique Race Days per class
  const raceDaysByClass = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const startDateTime = row[0];
    const classNum      = row[4];
    const dateKey = toDateKey(startDateTime);
    if (!dateKey) continue;
    if (!raceDaysByClass[classNum]) raceDaysByClass[classNum] = new Set();
    raceDaysByClass[classNum].add(dateKey);
  }

  Object.keys(raceDaysByClass).sort().forEach(classNum => {
    const dates = Array.from(raceDaysByClass[classNum]).sort();
    Logger.log(`Class ${classNum} (${dates.length} race days): ${dates.join(', ')}`);
  });

  // Pass 2: per-boat race day participation and points
  const classGroups = {};
  const boatRaceDays = {};
  const boatExcusedDays = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const startDateTime = row[0];
    const raceType  = row[3];
    const classNum  = row[4];
    const boatName  = row[6];
    const points    = Number(row[13]);
    const status    = row[14];
    const isPractice = String(raceType).startsWith('Practice');
    const dateKey = toDateKey(startDateTime);

    if (!boatName) continue;

    const boatKey = `${classNum}_${boatName}`;

    if (!classGroups[classNum]) classGroups[classNum] = {};
    if (!classGroups[classNum][boatName]) classGroups[classNum][boatName] = [];
    if (!boatRaceDays[boatKey]) boatRaceDays[boatKey] = new Set();
    if (!boatExcusedDays[boatKey]) boatExcusedDays[boatKey] = new Set();

    // Participation: any status except DNC means they showed up that day
    if (status !== 'DNC') {
      boatRaceDays[boatKey].add(dateKey);
    }

    // 2026: BYE counts toward qualification (add to excused days for tracking)
    if (status === 'BYE') {
      boatExcusedDays[boatKey].add(dateKey);
    }

    // 2026: BYE is now a scored race — include its points in season standings
    // Exclude only practice and DNC
    if (!isPractice && status !== 'DNC') {
      classGroups[classNum][boatName].push(points);
    }
  }

  seasonSheet.clear();
  let currentRow = 1;
  seasonSheet.getRange(1, 10, 1000, 1).setNumberFormat('@');

  Object.keys(classGroups).sort().forEach(classNum => {
    const boats = classGroups[classNum];
    const totalRaceDays = raceDaysByClass[classNum] ? raceDaysByClass[classNum].size : 0;
    const requiredDays = Math.ceil(totalRaceDays * 0.75);

    seasonSheet.getRange(currentRow, 1).setValue(`CLASS ${classNum} - SEASON STANDINGS`);
    seasonSheet.getRange(currentRow, 1, 1, 10).setBackground('#ff9900').setFontWeight('bold');
    currentRow++;

    seasonSheet.appendRow(['Boat Name', 'Races', 'Total PTS', 'Throwouts', 'Net PTS', 'AVG', 'Qualified', 'Race Days', 'Required Days', 'Participation %']);
    seasonSheet.getRange(currentRow, 1, 1, 10).setBackground('#6d9eeb').setFontWeight('bold');
    seasonSheet.getRange(1, 10, 1000, 1).setNumberFormat('@');
    currentRow++;

    const standings = [];
    Object.keys(boats).forEach(boatName => {
      const allPoints = boats[boatName];
      const numRaces = allPoints.length;
      const numThrowouts = Math.floor(numRaces / 7);

      const boatKey = `${classNum}_${boatName}`;
      const raceDaysAttended = boatRaceDays[boatKey] ? boatRaceDays[boatKey].size : 0;
      const participationPct = totalRaceDays > 0 ? (raceDaysAttended / totalRaceDays * 100) : 0;
      const isQualified = raceDaysAttended >= requiredDays;

      const sortedPoints = allPoints.sort((a, b) => b - a);
      const kept = sortedPoints.slice(numThrowouts);

      const total = sortedPoints.reduce((sum, p) => sum + p, 0);
      const netPoints = kept.reduce((sum, p) => sum + p, 0);
      const average = kept.length > 0 ? netPoints / kept.length : 0;

      standings.push({ boatName, numRaces, total, numThrowouts, netPoints, average, isQualified, raceDaysAttended, totalRaceDays, participationPct });
    });

    standings.sort((a, b) => a.average - b.average);

    standings.forEach(s => {
      seasonSheet.appendRow([
        s.boatName, s.numRaces, s.total, s.numThrowouts, s.netPoints,
        s.average.toFixed(2),
        s.isQualified ? 'YES' : 'NO',
        s.raceDaysAttended + ' of ' + s.totalRaceDays,
        requiredDays,
        s.participationPct.toFixed(1) + '%'
      ]);
      currentRow++;

      if (!s.isQualified) {
        seasonSheet.getRange(currentRow - 1, 1, 1, 10).setBackground('#f4cccc');
      }
    });

    currentRow += 2;
  });
}

/**
 * Calculate Cumulative Results in YachtScoring style
 * Groups by Series > Class, with R1, R2, R3... columns per series
 */
function calculateCumulativeResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName(SHEETS.CALCULATED);
  
  let cumSheet = ss.getSheetByName(SHEETS.CUMULATIVE);
  if (!cumSheet) {
    cumSheet = ss.insertSheet(SHEETS.CUMULATIVE);
  }
  cumSheet.clear();
  
  const data = calcSheet.getDataRange().getValues();

  const structure = {};
  const seriesRaceNums = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const raceNum  = row[1];
    const series   = row[2];
    const raceType = row[3];
    const classNum = row[4];
    const boatName = row[6];
    const points   = Number(row[13]);
    const status   = row[14];
    
    const isPractice = String(raceType).startsWith('Practice');
    if (isPractice) continue;
    if (!boatName) continue;
    
    if (!structure[series]) structure[series] = {};
    if (!structure[series][classNum]) structure[series][classNum] = {};
    if (!structure[series][classNum][boatName]) structure[series][classNum][boatName] = {};
    structure[series][classNum][boatName][raceNum] = { 
      points: typeof points === 'number' ? points : Number(points), 
      status: status || '' 
    };
    
    if (!seriesRaceNums[series]) seriesRaceNums[series] = new Set();
    seriesRaceNums[series].add(raceNum);
  }
  
  let currentRow = 1;
  
  Object.keys(structure).sort((a, b) => Number(a) - Number(b)).forEach(series => {
    const classes = structure[series];
    const raceNums = Array.from(seriesRaceNums[series]).sort((a, b) => Number(a) - Number(b));
    const numRaces = raceNums.length;
    
    const seriesHeaderRange = cumSheet.getRange(currentRow, 1, 1, 4 + numRaces);
    seriesHeaderRange.merge();
    cumSheet.getRange(currentRow, 1).setValue(`SERIES ${series}`);
    seriesHeaderRange.setBackground('#1c4587').setFontColor('white').setFontWeight('bold').setFontSize(12);
    currentRow++;
    
    Object.keys(classes).sort((a, b) => Number(a) - Number(b)).forEach(classNum => {
      const boats = classes[classNum];
      
      const classHeaderRange = cumSheet.getRange(currentRow, 1, 1, 4 + numRaces);
      classHeaderRange.merge();
      cumSheet.getRange(currentRow, 1).setValue(`Class ${classNum}`);
      classHeaderRange.setBackground('#4a86e8').setFontColor('white').setFontWeight('bold');
      currentRow++;
      
      const colHeaders = ['Boat Name', 'Sail#', 'PHRF', 'Total'];
      raceNums.forEach((rn) => colHeaders.push(`R${rn}`));
      cumSheet.appendRow(colHeaders);
      cumSheet.getRange(currentRow, 1, 1, colHeaders.length).setBackground('#6d9eeb').setFontColor('white').setFontWeight('bold');
      currentRow++;
      
      const boatStandings = [];
      Object.keys(boats).forEach(boatName => {
        const raceResults = boats[boatName];
        let total = 0;
        const raceScores = raceNums.map(rn => {
          const result = raceResults[rn];
          if (!result) return '';
          const pts = result.points;
          const stat = result.status;
          total += pts || 0;
          // 2026: BYE shows as "avg/BYE" so it's clear it's an averaged score
          if (stat) return `${pts}/${stat}`;
          return pts !== null && pts !== undefined ? String(pts) : '';
        });
        boatStandings.push({ boatName, total, raceScores });
      });
      
      boatStandings.sort((a, b) => a.total - b.total);
      
      const boatInfo = {};
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const bn = row[6];
        if (!boatInfo[bn]) {
          boatInfo[bn] = { 
            sailNumber: row[7],
            phrf: row[8]
          };
        }
      }
      
      boatStandings.forEach((s, idx) => {
        const info = boatInfo[s.boatName] || { sailNumber: '', phrf: '' };
        const rowData = [s.boatName, info.sailNumber, info.phrf, s.total, ...s.raceScores];
        cumSheet.appendRow(rowData);
        
        if (idx % 2 === 0) {
          cumSheet.getRange(currentRow, 1, 1, rowData.length).setBackground('#e8f0fe');
        }
        currentRow++;
      });
      
      currentRow++;
    });
    
    currentRow++;
  });
  
  cumSheet.getRange(1, 3, currentRow, 1).setNumberFormat('0');
  cumSheet.autoResizeColumns(1, 20);
  Logger.log('Cumulative Results calculated.');
}

/**
 * Calculate sunset time for Jersey City, NJ
 * Usage: =SUNSET(date) where date is any date value
 * Returns sunset time as a decimal fraction of day (Google Sheets time value)
 */
function SUNSET(date) {
  if (!date) return '';
  const d = date instanceof Date ? date : new Date(date);
  if (isNaN(d.getTime())) return '';

  const lat = 40.7178;
  const lng = -74.0431;

  const rad = Math.PI / 180;
  const deg = 180 / Math.PI;

  const dayOfYear = Math.floor((d - new Date(d.getFullYear(), 0, 0)) / 86400000);
  const declination = -23.45 * Math.cos(rad * (360 / 365) * (dayOfYear + 10));
  
  const cosHourAngle = -Math.tan(lat * rad) * Math.tan(declination * rad);
  if (cosHourAngle < -1) return 'No sunset';
  if (cosHourAngle > 1) return 'No sunrise';
  
  const hourAngle = Math.acos(cosHourAngle) * deg;
  const sunsetUTC = 12 + hourAngle / 15 - lng / 15;
  
  const jan = new Date(d.getFullYear(), 0, 1);
  const jul = new Date(d.getFullYear(), 6, 1);
  const stdOffset = Math.max(jan.getTimezoneOffset(), jul.getTimezoneOffset());
  const isDST = d.getTimezoneOffset() < stdOffset;
  const utcOffset = isDST ? -4 : -5;
  
  let localSunset = sunsetUTC + utcOffset;
  if (localSunset < 0) localSunset += 24;
  if (localSunset >= 24) localSunset -= 24;
  
  return localSunset / 24;
}

/**
 * Restore formulas, VLOOKUP, and dropdowns in Race Results Entry sheet
 *
 * Run this after clearing the sheet for a new season, or any time
 * formulas or dropdowns are accidentally deleted.
 *
 * Column layout:
 *   A = Race#      B = Series      C = RaceType    D = Course
 *   E = BoatName   F = Class*      G = StartDate   H = StartTime
 *   I = StartDateTime*  J = Finish Date  K = FinishTime  L = FinishDateTime*
 *   M = Status     N = Sunset Time*  O = After Sunset*  P = Wind  Q = Tide
 *   (* = formula or VLOOKUP, restored by this function)
 *
 * Dropdowns restored:
 *   A = Race# (1-17)
 *   B = Series (1, 2, 3)
 *   C = RaceType (Fleet, Practice Fleet, Pursuit)
 *   D = Course (ALPHA through QUEBEC)
 *   E = BoatName (from Scratch Sheet col A)
 *   M = Status (DNC, DNF, RET, TLE, DSQ, DNE, OCS, BYE)
 */
function restoreEntryFormulas() {
  const MAX_ROWS = 500;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RACE_ENTRY);
  const scratchSheet = ss.getSheetByName(SHEETS.SCRATCH_SHEET);
  const rule = SpreadsheetApp.newDataValidation;

  // ── FORMULAS ────────────────────────────────────────────────────────────────
  const colF = [], colI = [], colL = [], colN = [], colO = [];
  for (let row = 2; row <= MAX_ROWS + 1; row++) {
    colF.push([`=IFERROR(VLOOKUP(E${row},'Scratch Sheet'!A:J,10,FALSE),"")`]); // Class from col J
    colI.push([`=IF(AND(G${row}<>"",H${row}<>""),G${row}+H${row},"")`]);
    colL.push([`=IF(AND(J${row}<>"",K${row}<>""),J${row}+K${row},"")`]);
    colN.push([`=SUNSET(G${row})`]);
    colO.push([`=IF(L${row}="","",IF(M${row}<>"","",ROUND((MOD(L${row},1)-SUNSET(G${row}))*1440,0)))`]);
  }
  sheet.getRange(2, 6,  MAX_ROWS, 1).setFormulas(colF);  // F = Class (VLOOKUP)
  sheet.getRange(2, 9,  MAX_ROWS, 1).setFormulas(colI);  // I = StartDateTime
  sheet.getRange(2, 12, MAX_ROWS, 1).setFormulas(colL);  // L = FinishDateTime
  sheet.getRange(2, 14, MAX_ROWS, 1).setFormulas(colN);  // N = Sunset Time
  sheet.getRange(2, 15, MAX_ROWS, 1).setFormulas(colO);  // O = After Sunset

  // ── DROPDOWNS ───────────────────────────────────────────────────────────────

  // Col A: Race# (1-17)
  const raceNums = Array.from({length: 17}, (_, i) => String(i + 1));
  sheet.getRange(2, 1, MAX_ROWS, 1).setDataValidation(
    rule().requireValueInList(raceNums, true).setAllowInvalid(false).build()
  );

  // Col B: Series
  sheet.getRange(2, 2, MAX_ROWS, 1).setDataValidation(
    rule().requireValueInList(['1', '2', '3'], true).setAllowInvalid(false).build()
  );

  // Col C: RaceType
  sheet.getRange(2, 3, MAX_ROWS, 1).setDataValidation(
    rule().requireValueInList(['Fleet', 'Practice Fleet', 'Pursuit'], true).setAllowInvalid(false).build()
  );

  // Col D: Course
  const courses = [
    'ALPHA','BRAVO','CHARLIE','DELTA','ECHO',
    'FOXTROT','GOLF','HOTEL','INDIA','JULIET',
    'KILO','LIMA','MIKE','NOVEMBER','OSCAR',
    'PAPA','QUEBEC'
  ];
  sheet.getRange(2, 4, MAX_ROWS, 1).setDataValidation(
    rule().requireValueInList(courses, true).setAllowInvalid(false).build()
  );

  // Col E: BoatName (dynamic from Scratch Sheet col A, rows 2 onward)
  const lastBoatRow = scratchSheet.getLastRow();
  const boatRange = scratchSheet.getRange(2, 1, lastBoatRow - 1, 1);
  sheet.getRange(2, 5, MAX_ROWS, 1).setDataValidation(
    rule().requireValueInRange(boatRange, true).setAllowInvalid(false).build()
  );

  // Col M: Status
  const statuses = ['DNC', 'DNF', 'RET', 'TLE', 'DSQ', 'DNE', 'OCS', 'BYE'];
  sheet.getRange(2, 13, MAX_ROWS, 1).setDataValidation(
    rule().requireValueInList(statuses, true).setAllowInvalid(false).build()
  );

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Formulas and dropdowns restored in Race Results Entry',
    'Done', 4
  );
  Logger.log('restoreEntryFormulas complete.');
}

/**
 * Create menu in Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⛵ Racing Scoring')
    .addItem('Calculate All Results', 'calculateAllResults')
    .addSeparator()
    .addItem('Calculate Race Results', 'calculateRaceResults')
    .addItem('Calculate Series Standings', 'calculateSeriesStandings')
    .addItem('Calculate Season Standings', 'calculateSeasonStandings')
    .addItem('Calculate Cumulative Results', 'calculateCumulativeResults')
    .addSeparator()
    .addItem('Restore Entry Sheet Formulas', 'restoreEntryFormulas')
    .addToUi();
}

/**
 * Web app proxy for Wix display
 * Always serves data from the sheet this script is bound to —
 * no hardcoded ID needed, so copying to a new season just works.
 */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = e.parameter.sheet;
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  if (!sheet) {
    return ContentService.createTextOutput('Sheet not found: ' + sheetName)
      .setMimeType(ContentService.MimeType.TEXT);
  }
  const data = sheet.getDataRange().getDisplayValues();
  const csv = data.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(',')).join('\n');
  return ContentService.createTextOutput(csv)
    .setMimeType(ContentService.MimeType.TEXT);
}