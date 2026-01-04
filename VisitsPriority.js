/*
  Selection summary:
  - Institutions are scored by max days since last seen, avg days, sizeWeight, and blankDateBoost.
  - Selection fills in order: overdue, small quota, then score rank.
  - Selection stops when either the institution count or resident count cap is reached.
  - The Select column marks all rows for selected institutions.
*/

// Priority configuration notes:
// - sizeWeight: higher favors larger institutions; lower favors smaller ones.
// - avgWeight: higher favors broadly overdue institutions; lower emphasizes maxDays.
// - overdueThresholdDays: institutions at/above this maxDays are force-selected.
// - smallInstitutionThreshold: "small" institution size cutoff.
// - smallInstitutionQuota: fraction of targetInstitutionCount reserved for small institutions.
// - targetInstitutionCount: total institutions to select each run.
// - maxResidentsPerTrip: total residents across selected institutions.
// - blankDateThreshold: minimum blank/invalid dates before blankDateWeight applies.
// - blankDateWeight: per-blank score boost once threshold is met.
// - deprioritizedInstitutions: institutions pushed to the bottom of the report and
//   excluded from selection when any other institution is available.
const PRIORITY_CONFIG = {
  HISTORY_SHEET_NAME: "History Sheet",
  PRIORITY_SHEET_NAME: "Institution Priority",
  sizeWeight: 1,
  smallInstitutionThreshold: 15,
  smallInstitutionQuota: 0.25,
  overdueThresholdDays: 180,
  targetInstitutionCount: 10,
  maxResidentsPerTrip: 10,
  avgWeight: 0.3,
  blankDateThreshold: 3,
  blankDateWeight: 5,
  deprioritizedInstitutions: [
    "Homebound",
    "EnglishMeadows",
    "Lodge at Old Trail",
  ],
};

function buildInstitutionPriority() {
  const ss = SpreadsheetApp.getActive();
  const historySheet = ss.getSheetByName(PRIORITY_CONFIG.HISTORY_SHEET_NAME)
    || ss.getSheetByName("History");
  if (!historySheet) {
    throw new Error("History sheet not found.");
  }

  const values = historySheet.getDataRange().getValues();
  if (values.length === 0) {
    throw new Error("History sheet has no data.");
  }

  const headers = values[0];
  const institutionIndex = findHeaderIndex_(headers, ["institution"]);
  const dateIndex = findHeaderIndex_(headers, ["date last seen", "date lastseen", "last seen"]);

  if (institutionIndex === -1 || dateIndex === -1) {
    throw new Error("Required headers not found: Institution, Date Last Seen.");
  }

  const now = new Date();
  const deprioritizedNames = PRIORITY_CONFIG.deprioritizedInstitutions.map(
    normalizeInstitutionName_
  );
  const statsMap = {};

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const institution = String(row[institutionIndex] || "").trim();
    if (!institution) continue;

    const parsedDate = parseDate_(row[dateIndex]);
    const isBlankDate = !parsedDate;
    const daysSince = parsedDate ? daysSinceFromDate_(parsedDate, now) : null;
    if (!statsMap[institution]) {
      statsMap[institution] = {
        institution: institution,
        forceBottom: isDeprioritizedInstitution_(institution, deprioritizedNames),
        countPeople: 0,
        blankDateCount: 0,
        residentRows: [],
      };
    }

    const stats = statsMap[institution];
    stats.countPeople++;
    if (isBlankDate) {
      stats.blankDateCount++;
    }
    stats.residentRows.push({
      rowIndex: r + 1,
      daysSince: daysSince,
      isBlank: isBlankDate,
    });
  }

  const statsList = Object.keys(statsMap).map(key => {
    const stats = statsMap[key];
    let fallbackDaysSince = maxNonBlankDays_(stats.residentRows);

    let sumDays = 0;
    let maxDays = 0;
    const adjustedDays = [];
    const residentRows = stats.residentRows.map(resident => {
      const adjusted = resident.daysSince !== null ? resident.daysSince : fallbackDaysSince;
      adjustedDays.push(adjusted);
      sumDays += adjusted;
      if (adjusted > maxDays) maxDays = adjusted;
      return {
        rowIndex: resident.rowIndex,
        daysSince: adjusted,
      };
    });

    const avgDays = stats.countPeople ? sumDays / stats.countPeople : 0;
    const medianDays = median_(adjustedDays);
    const blankDateBoost = stats.blankDateCount >= PRIORITY_CONFIG.blankDateThreshold
      ? stats.blankDateCount * PRIORITY_CONFIG.blankDateWeight
      : 0;
    const score = maxDays
      + (avgDays * PRIORITY_CONFIG.avgWeight)
      + (Math.log(stats.countPeople + 1) * PRIORITY_CONFIG.sizeWeight)
      + blankDateBoost;

    return {
      institution: stats.institution,
      forceBottom: stats.forceBottom,
      countPeople: stats.countPeople,
      maxDaysSinceLastSeen: maxDays,
      avgDaysSinceLastSeen: avgDays,
      medianDaysSinceLastSeen: medianDays,
      blankDateCount: stats.blankDateCount,
      blankDateCountApplied: stats.blankDateCount >= PRIORITY_CONFIG.blankDateThreshold
        ? stats.blankDateCount
        : 0,
      blankDateBoost: blankDateBoost,
      score: score,
      selected: false,
      selectedReason: "",
      residentRows: residentRows,
    };
  });

  applySelection_(statsList);

  const outputHeaders = [
    "Institution",
    "countPeople",
    "maxDaysSinceLastSeen",
    "avgDaysSinceLastSeen",
    "medianDaysSinceLastSeen",
    "blankDateCount",
    "blankDateCountApplied",
    "blankDateBoost",
    "score",
    "selectedReason",
    "selected",
  ];

  statsList.sort(byScoreDescName_);
  const outputRows = statsList.map(stat => ([
    stat.institution,
    stat.countPeople,
    stat.maxDaysSinceLastSeen,
    round1_(stat.avgDaysSinceLastSeen),
    round1_(stat.medianDaysSinceLastSeen),
    stat.blankDateCount,
    stat.blankDateCountApplied,
    round1_(stat.blankDateBoost),
    round1_(stat.score),
    stat.selectedReason || "",
    stat.selected,
  ]));

  let prioritySheet = ss.getSheetByName(PRIORITY_CONFIG.PRIORITY_SHEET_NAME);
  if (!prioritySheet) {
    prioritySheet = ss.insertSheet(PRIORITY_CONFIG.PRIORITY_SHEET_NAME);
  } else {
    prioritySheet.clear();
  }

  prioritySheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
  if (outputRows.length > 0) {
    prioritySheet.getRange(2, 1, outputRows.length, outputHeaders.length).setValues(outputRows);
  }

  writeSelectionColumn_(historySheet, values, institutionIndex, statsList);
}

function buildInstitutionPriorityTrigger() {
  buildInstitutionPriority();
}

function addPastoralCareMenu_() {
  const ss = SpreadsheetApp.getActive();
  ss.addMenu("Pastoral Care", [
    { name: "Build Institution Priority", functionName: "buildInstitutionPriority" },
    { name: "Build Visit Route", functionName: "buildVisitRouteFromSheet" },
  ]);
}

function writeSelectionColumn_(historySheet, values, institutionIndex, statsList) {
  const headers = values[0];
  const selectIndex = findHeaderIndex_(headers, ["select"]);
  const selectColumn = selectIndex === -1 ? headers.length + 1 : selectIndex + 1;
  if (selectIndex === -1) {
    historySheet.getRange(1, selectColumn).setValue("Select");
  }

  const selectedSet = new Set(
    statsList.filter(stat => stat.selected).map(stat => stat.institution)
  );

  const selectValues = [];
  for (let r = 1; r < values.length; r++) {
    const institution = String(values[r][institutionIndex] || "").trim();
    const value = institution && selectedSet.has(institution) ? "YES" : "";
    selectValues.push([value]);
  }

  if (selectValues.length > 0) {
    historySheet.getRange(2, selectColumn, selectValues.length, 1).setValues(selectValues);
  }
}

function applySelection_(statsList) {
  const cfg = PRIORITY_CONFIG;
  statsList.forEach(stat => {
    stat.selected = false;
    stat.selectedReason = "";
  });

  if (statsList.length === 0) return;

  const eligible = statsList.filter(stat => !stat.forceBottom);
  const selectionPool = eligible.length > 0 ? eligible : statsList;
  const scoreRanks = buildScoreRankMap_(selectionPool);
  let remainingInstitutions = cfg.targetInstitutionCount;
  let remainingResidents = cfg.maxResidentsPerTrip;

  function trySelect(stat, reason) {
    if (stat.selected) return false;
    if (remainingInstitutions <= 0) return false;
    if (stat.countPeople > remainingResidents) return false;
    markSelected_(stat, reason);
    remainingInstitutions--;
    remainingResidents -= stat.countPeople;
    return true;
  }

  const overdue = selectionPool.filter(stat => stat.maxDaysSinceLastSeen >= cfg.overdueThresholdDays);
  overdue.sort(byMaxDaysDescName_);
  for (let i = 0; i < overdue.length; i++) {
    if (remainingInstitutions <= 0 || remainingResidents <= 0) break;
    trySelect(overdue[i], "overdue >= " + cfg.overdueThresholdDays + " days");
  }

  const smallQuotaTarget = Math.round(cfg.targetInstitutionCount * cfg.smallInstitutionQuota);
  const smallSelectedCount = selectionPool.filter(
    stat => stat.selected && stat.countPeople <= cfg.smallInstitutionThreshold
  ).length;
  const smallNeeded = Math.max(0, smallQuotaTarget - smallSelectedCount);

  if (smallNeeded > 0 && remainingInstitutions > 0 && remainingResidents > 0) {
    const smallCandidates = selectionPool.filter(
      stat => !stat.selected && stat.countPeople <= cfg.smallInstitutionThreshold
    );
    smallCandidates.sort(byMaxDaysDescName_);
    let picked = 0;
    for (let i = 0; i < smallCandidates.length; i++) {
      if (picked >= smallNeeded) break;
      if (remainingInstitutions <= 0 || remainingResidents <= 0) break;
      if (trySelect(smallCandidates[i], "small institution quota")) {
        picked++;
      }
    }
  }

  const selectedCount = selectionPool.filter(stat => stat.selected).length;
  const remaining = Math.max(0, cfg.targetInstitutionCount - selectedCount);
  if (remaining > 0 && remainingResidents > 0) {
    const candidates = selectionPool.filter(stat => !stat.selected);
    candidates.sort(byScoreDescName_);
    for (let i = 0; i < candidates.length; i++) {
      if (remainingInstitutions <= 0 || remainingResidents <= 0) break;
      const rank = scoreRanks[candidates[i].institution];
      trySelect(candidates[i], "score rank " + rank);
    }
  }
}

function markSelected_(stat, reason) {
  if (!stat.selected) {
    stat.selected = true;
    stat.selectedReason = reason;
  }
}

function normalizeHeader_(value) {
  return String(value || "").trim().toLowerCase();
}

function normalizeInstitutionName_(value) {
  return String(value || "").trim().toLowerCase();
}

function buildScoreRankMap_(statsList) {
  const sorted = statsList.slice().sort(byScoreDescName_);
  const ranks = {};
  for (let i = 0; i < sorted.length; i++) {
    ranks[sorted[i].institution] = i + 1;
  }
  return ranks;
}

function isDeprioritizedInstitution_(name, deprioritizedNames) {
  const normalized = normalizeInstitutionName_(name);
  for (let i = 0; i < deprioritizedNames.length; i++) {
    if (normalized.indexOf(deprioritizedNames[i]) !== -1) return true;
  }
  return false;
}

function findHeaderIndex_(headers, candidates) {
  const normalizedCandidates = candidates.map(c => String(c).trim().toLowerCase());
  for (let i = 0; i < headers.length; i++) {
    if (normalizedCandidates.indexOf(normalizeHeader_(headers[i])) !== -1) {
      return i;
    }
  }
  return -1;
}

function daysSinceFromDate_(date, now) {
  let diff = Math.floor((now.getTime() - date.getTime()) / 86400000);
  if (diff < 0) diff = 0;
  return diff;
}

function parseDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") return null;
    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) return parsed;
    return null;
  }
  if (typeof value === "number") {
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function maxNonBlankDays_(residentRows) {
  let maxDays = 0;
  for (let i = 0; i < residentRows.length; i++) {
    if (!residentRows[i].isBlank && residentRows[i].daysSince !== null) {
      if (residentRows[i].daysSince > maxDays) {
        maxDays = residentRows[i].daysSince;
      }
    }
  }
  return maxDays;
}

function median_(values) {
  if (!values || values.length === 0) return 0;
  const sorted = values.slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 0) {
    return (sorted[mid - 1] + sorted[mid]) / 2;
  }
  return sorted[mid];
}

function round1_(value) {
  return Math.round(value * 10) / 10;
}

function byScoreDescName_(a, b) {
  if (a.forceBottom && !b.forceBottom) return 1;
  if (!a.forceBottom && b.forceBottom) return -1;
  if (b.score !== a.score) return b.score - a.score;
  return a.institution.localeCompare(b.institution);
}

function byMaxDaysDescName_(a, b) {
  if (b.maxDaysSinceLastSeen !== a.maxDaysSinceLastSeen) {
    return b.maxDaysSinceLastSeen - a.maxDaysSinceLastSeen;
  }
  return a.institution.localeCompare(b.institution);
}
