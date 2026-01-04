const ROUTE_CONFIG = {
  INPUT_SHEET_NAME: "Visit Route Inputs",
  OUTPUT_SHEET_NAME: "Visit Route",
  ADDRESS_HEADER: "address",
  MAX_WAYPOINTS: 23,
};

function buildVisitRouteFromSheet() {
  const ss = SpreadsheetApp.getActive();
  const inputSheet = ss.getSheetByName(ROUTE_CONFIG.INPUT_SHEET_NAME);
  if (!inputSheet) {
    throw new Error("Missing input sheet: " + ROUTE_CONFIG.INPUT_SHEET_NAME);
  }

  const values = inputSheet.getDataRange().getValues();
  if (values.length < 2) {
    throw new Error("Add a header row and at least one start and stop address.");
  }

  const headers = values[0];
  const addressIndex = findRouteHeaderIndex_(headers, ROUTE_CONFIG.ADDRESS_HEADER);
  const addresses = [];
  for (let r = 1; r < values.length; r++) {
    const raw = values[r][addressIndex];
    const address = String(raw || "").trim();
    if (address) addresses.push(address);
  }

  if (addresses.length < 2) {
    throw new Error("Provide a start address and at least one stop.");
  }

  const startAddress = addresses[0];
  const stopAddresses = addresses.slice(1);
  if (stopAddresses.length > ROUTE_CONFIG.MAX_WAYPOINTS) {
    throw new Error("Too many stops. Max waypoints is " + ROUTE_CONFIG.MAX_WAYPOINTS + ".");
  }

  const route = buildOptimizedRoute_(startAddress, stopAddresses);
  writeRouteOutput_(ss, route, startAddress);
}

function buildOptimizedRoute_(startAddress, stopAddresses) {
  const finder = Maps.newDirectionFinder()
    .setOrigin(startAddress)
    .setDestination(startAddress)
    .setMode(Maps.DirectionFinder.Mode.DRIVING)
    .setOptimizeWaypoints(true);

  stopAddresses.forEach(address => finder.addWaypoint(address));
  const directions = finder.getDirections();
  if (!directions.routes || directions.routes.length === 0) {
    throw new Error("No route found for the provided addresses.");
  }
  return directions.routes[0];
}

function writeRouteOutput_(ss, route, startAddress) {
  let sheet = ss.getSheetByName(ROUTE_CONFIG.OUTPUT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ROUTE_CONFIG.OUTPUT_SHEET_NAME);
  } else {
    sheet.clear();
  }

  const headers = [
    "Order",
    "Address",
    "Leg Distance (mi)",
    "Leg Duration (min)",
    "Cumulative Distance (mi)",
    "Cumulative Duration (min)",
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rows = [];
  let cumulativeMeters = 0;
  let cumulativeSeconds = 0;
  rows.push([0, startAddress, 0, 0, 0, 0]);

  const legs = route.legs || [];
  for (let i = 0; i < legs.length; i++) {
    const leg = legs[i];
    const meters = leg.distance ? leg.distance.value : 0;
    const seconds = leg.duration ? leg.duration.value : 0;
    cumulativeMeters += meters;
    cumulativeSeconds += seconds;
    rows.push([
      i + 1,
      leg.end_address || "",
      round1Route_(metersToMiles_(meters)),
      round1Route_(secondsToMinutes_(seconds)),
      round1Route_(metersToMiles_(cumulativeMeters)),
      round1Route_(secondsToMinutes_(cumulativeSeconds)),
    ]);
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function findRouteHeaderIndex_(headers, name) {
  const normalized = String(name || "").trim().toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim().toLowerCase() === normalized) {
      return i;
    }
  }
  return 0;
}

function metersToMiles_(meters) {
  return meters / 1609.34;
}

function secondsToMinutes_(seconds) {
  return seconds / 60;
}

function round1Route_(value) {
  return Math.round(value * 10) / 10;
}
