const PRIORITY_CONFIG = {
  HISTORY_SHEET_NAME: "History Sheet",
  PRIORITY_SHEET_NAME: "Institution Priority",
  sizeWeight: 1,
  smallInstitutionThreshold: 15,
  smallInstitutionQuota: 0.25,
  overdueThresholdDays: 180,
  targetInstitutionCount: 10,
  avgWeight: 0.3,
  veryOldDays: 9999,
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
  const statsMap = {};

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const institution = String(row[institutionIndex] || "").trim();
    if (!institution) continue;

    const daysSince = daysSince_(row[dateIndex], now, PRIORITY_CONFIG.veryOldDays);
    if (!statsMap[institution]) {
      statsMap[institution] = {
        institution: institution,
        countPeople: 0,
        sumDays: 0,
        maxDaysSinceLastSeen: 0,
        daysSince: [],
      };
    }

    const stats = statsMap[institution];
    stats.countPeople++;
    stats.sumDays += daysSince;
    if (stats.countPeople === 1 || daysSince > stats.maxDaysSinceLastSeen) {
      stats.maxDaysSinceLastSeen = daysSince;
    }
    stats.daysSince.push(daysSince);
  }

  const statsList = Object.keys(statsMap).map(key => {
    const stats = statsMap[key];
    const avgDays = stats.countPeople ? stats.sumDays / stats.countPeople : 0;
    const medianDays = median_(stats.daysSince);
    const score = stats.maxDaysSinceLastSeen
      + (avgDays * PRIORITY_CONFIG.avgWeight)
      + (Math.log(stats.countPeople + 1) * PRIORITY_CONFIG.sizeWeight);

    return {
      institution: stats.institution,
      countPeople: stats.countPeople,
      maxDaysSinceLastSeen: stats.maxDaysSinceLastSeen,
      avgDaysSinceLastSeen: avgDays,
      medianDaysSinceLastSeen: medianDays,
      score: score,
      selected: false,
      selectedReason: "",
    };
  });

  applySelection_(statsList);

  const outputHeaders = [
    "Institution",
    "countPeople",
    "maxDaysSinceLastSeen",
    "avgDaysSinceLastSeen",
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
  const overdue = statsList.filter(stat => stat.maxDaysSinceLastSeen >= cfg.overdueThresholdDays);
  overdue.sort(byMaxDaysDescName_);
  overdue.forEach(stat => markSelected_(stat, "overdue >= " + cfg.overdueThresholdDays + " days"));

  const smallQuotaTarget = Math.round(cfg.targetInstitutionCount * cfg.smallInstitutionQuota);
  const smallSelectedCount = statsList.filter(
    stat => stat.selected && stat.countPeople <= cfg.smallInstitutionThreshold
  ).length;
  const smallNeeded = Math.max(0, smallQuotaTarget - smallSelectedCount);

  if (smallNeeded > 0) {
    const smallCandidates = statsList.filter(
      stat => !stat.selected && stat.countPeople <= cfg.smallInstitutionThreshold
    );
    smallCandidates.sort(byMaxDaysDescName_);
    for (let i = 0; i < Math.min(smallNeeded, smallCandidates.length); i++) {
      markSelected_(smallCandidates[i], "small institution quota");
    }
  }

  const selectedCount = statsList.filter(stat => stat.selected).length;
  const remaining = Math.max(0, cfg.targetInstitutionCount - selectedCount);
  if (remaining > 0) {
    const candidates = statsList.filter(stat => !stat.selected);
    candidates.sort(byScoreDescName_);
    for (let i = 0; i < Math.min(remaining, candidates.length); i++) {
      markSelected_(candidates[i], "top score");
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

function findHeaderIndex_(headers, candidates) {
  const normalizedCandidates = candidates.map(c => String(c).trim().toLowerCase());
  for (let i = 0; i < headers.length; i++) {
    if (normalizedCandidates.indexOf(normalizeHeader_(headers[i])) !== -1) {
      return i;
    }
  }
  return -1;
}

function daysSince_(value, now, veryOldDays) {
  let date = null;
  if (value instanceof Date && !isNaN(value.getTime())) {
    date = value;
  } else if (typeof value === "string" && value.trim() !== "") {
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) date = parsed;
  } else if (typeof value === "number") {
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) date = parsed;
  }

  if (!date) return veryOldDays;
  let diff = Math.floor((now.getTime() - date.getTime()) / 86400000);
  if (diff < 0) diff = 0;
  return diff;
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
  if (b.score !== a.score) return b.score - a.score;
  return a.institution.localeCompare(b.institution);
}

function byMaxDaysDescName_(a, b) {
  if (b.maxDaysSinceLastSeen !== a.maxDaysSinceLastSeen) {
    return b.maxDaysSinceLastSeen - a.maxDaysSinceLastSeen;
  }
  return a.institution.localeCompare(b.institution);
}
