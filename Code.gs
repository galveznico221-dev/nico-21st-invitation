const SHEET_NAME = "RSVPs";
const HEADERS = ["name", "role", "status", "timestamp"];
const SPREADSHEET_ID_PROPERTY = "RSVP_SPREADSHEET_ID";

// Keep this roster in sync with CONFIG.guests in index.html.
const VALID_GUESTS = {
  "tita eva babatio": "Blue Bill",
  "tita tinay gadi": "Blue Bill",
  "john paul delosantos": "Blue Bill",
  "jusstin galvez": "Blue Bill",
  "gesil galvez": "Blue Bill",
  "ninang elena": "Blue Bill",
  "ninong panyot": "Blue Bill",
  "tito alan defeo": "Blue Bill",
  "ninong jay hardenero": "Blue Bill",
  "tito jojo galvez": "Blue Bill",
  "tito phet lapore": "Blue Bill",
  "tito laboy acunin": "Blue Bill",
  "tita ethang sadia": "Blue Bill",
  "tita marites senes": "Blue Bill",
  "ninong victor ilan": "Blue Bill",
  "avil tabor": "Blue Bill",
  "ninang rosalie sendon": "Blue Bill",
  "ninang sally del barrio": "Blue Bill",
  "maricel padre": "Blue Bill",
  "christian galvez": "Blue Bill",
  "kim galvez": "Blue Bill",
  "catherine babatio": "Yellow Bill",
  "petie gadi": "Yellow Bill",
  "tito erwin gadi": "Yellow Bill",
  "tito sising delosantos": "Yellow Bill",
  "tita princess gadi": "Yellow Bill",
  "tita niengkay gadi": "Yellow Bill",
  "tita emily samayo": "Yellow Bill",
  "tita emma antigua": "Yellow Bill",
  "tita yunisa pamesa": "Yellow Bill",
  "michael gadi": "Yellow Bill",
  "cecilio talla": "Yellow Bill",
  "neysa garais": "Yellow Bill",
  "emily dezon": "Yellow Bill",
  "analy almadrones": "Yellow Bill",
  "wendell maigue": "Yellow Bill",
  "telde largo": "Yellow Bill",
  "dang madera": "Yellow Bill",
  "mary jane sena": "Yellow Bill",
  "mary ann tabor": "Yellow Bill",
  "rose ann galvez": "Yellow Bill",
  "daloy pajares": "Yellow Bill",
  "kyline cuezon": "Chocolate",
  "cezyana gadi": "Chocolate",
  "alysa galvez": "Chocolate",
  "anding depio": "Chocolate",
  "ethan gadi": "Chocolate",
  "kevin gadi": "Chocolate",
  "jaypee basilla": "Chocolate",
  "ethan josh padre": "Chocolate",
  "mark lance padre": "Chocolate",
  "john carl edivor": "Chocolate",
  "rekric pajares": "Chocolate",
  "jowella gadi": "Chocolate",
  "rosemarie pestanyas": "Chocolate",
  "adan asis": "Chocolate",
  "luis maligat": "Chocolate",
  "dhane ruben": "Chocolate",
  "tonton lapore": "Chocolate",
  "tintin lapore": "Chocolate",
  "boninnay jose": "Chocolate",
  "joy joy ilan": "Chocolate",
  "jero madera": "Chocolate",
  "angel galvez": "Cake",
  "angela gadi": "Cake",
  "monica aceron": "Cake",
  "bea asis": "Cake",
  "cassandra raviz": "Cake",
  "sheila sanchez": "Cake",
  "jesielyn delosantos": "Cake",
  "emelyn pestanyas": "Cake",
  "alysa ilan": "Cake",
  "abigael acunin": "Cake",
  "lei anne arana": "Cake",
  "jaja galvez": "Cake",
  "aila basillia": "Cake",
  "roselle divine padre": "Cake",
  "bobbie krizaila teodoro": "Cake",
  "anika kim salcedo": "Cake",
  "zyra maligat": "Cake",
  "dianne gadi": "Cake",
  "ivan babatio": "Cake",
  "nicolle galvez": "Cake",
  "angie sena": "Cake",
  "ella rose padre": "Treasure & Candle",
  "aaron yanela": "Treasure & Candle",
  "paul sayno": "Treasure & Candle",
  "john andrei pateno": "Treasure & Candle",
  "john neo ebarle": "Treasure & Candle",
  "gerry balon": "Treasure & Candle",
  "beverly alava": "Treasure & Candle",
  "hannah mae era": "Treasure & Candle",
  "lourize jade daniel": "Treasure & Candle",
  "john mark ramido": "Treasure & Candle",
  "mark henson valeros": "Treasure & Candle",
  "harren pomada": "Treasure & Candle",
  "angel basilla": "Treasure & Candle",
  "tita jocy": "Treasure & Candle",
  "tita rosalie": "Treasure & Candle",
  "tita liezel saman": "Treasure & Candle",
  "tita malou defeo": "Treasure & Candle",
  "erb gadi": "Treasure & Candle",
  "april tala": "Treasure & Candle",
  "romer gamelo": "Treasure & Candle",
  "tita wilma lapore": "Treasure & Candle",
  "francis gadi": "Gift and Shots",
  "warren madera": "Gift and Shots",
  "leopardo yasis": "Gift and Shots",
  "john lloyd vibar": "Gift and Shots",
  "adrian olit": "Gift and Shots",
  "angelo gozo": "Gift and Shots",
  "nick dela rosa": "Gift and Shots",
  "kyle romero": "Gift and Shots",
  "janos rawat": "Gift and Shots",
  "juliusz emerson vasques": "Gift and Shots",
  "rhain secillano": "Gift and Shots",
  "jed ariola": "Gift and Shots",
  "jonaz sarion": "Gift and Shots",
  "fernando ocfemia": "Gift and Shots",
  "dion tabanao": "Gift and Shots",
  "emilyn tagaan": "Gift and Shots",
  "cassandra behagan": "Gift and Shots",
  "prince jheck juan": "Gift and Shots",
  "adrian gel": "Gift and Shots",
  "danica jacob": "Gift and Shots",
  "rovil cereno": "Gift and Shots"
};

function doGet(event) {
  const parameters = event && event.parameter ? event.parameter : {};
  const action = parameters.action;
  const callback = parameters.callback;
  try {
    if (action === "accepted") {
      const sheet = getSheet();
      const rows = sheet.getDataRange().getValues();
      const accepted = rows.slice(1)
        .filter((row) => String(row[2]).toLowerCase() === "yes")
        .map((row) => ({
          name: String(row[0]),
          role: String(row[1]),
          status: String(row[2]).toLowerCase()
        }));

      return jsonOutput({ accepted: accepted }, callback);
    }

    if (action === "confirm") {
      const result = saveAcceptance(parameters.name, parameters.role);
      return jsonOutput(result, callback);
    }

    return jsonOutput({ error: "Use ?action=accepted or ?action=confirm." }, callback);
  } catch (error) {
    console.error("RSVP request failed.", error);
    return jsonOutput({ error: error.message || "The RSVP request failed." }, callback);
  }
}

function doPost(event) {
  let request;
  try {
    request = JSON.parse(event.postData.contents);
  } catch (error) {
    return jsonOutput({ error: "Request body must be valid JSON." });
  }

  if (cleanText(request.status).toLowerCase() !== "yes") {
    return jsonOutput({ error: "Only a yes RSVP can be saved." });
  }
  return jsonOutput(saveAcceptance(request.name, request.role));
}

function saveAcceptance(rawName, rawRole) {
  const name = cleanText(rawName);
  const role = cleanText(rawRole);
  const expectedRole = VALID_GUESTS[normalizeName(name)];
  if (!name || !role || !expectedRole || expectedRole !== role) {
    return { error: "Guest or role is not valid." };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();
    const matchingRow = rows.findIndex((row, index) =>
      index > 0 && normalizeName(row[0]) === normalizeName(name)
    );
    const record = [name, role, "yes", new Date()];

    if (matchingRow === -1) {
      sheet.appendRow(record);
    } else {
      sheet.getRange(matchingRow + 1, 1, 1, HEADERS.length).setValues([record]);
    }

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function setupRsvpStorage() {
  const properties = PropertiesService.getScriptProperties();
  const existingId = properties.getProperty(SPREADSHEET_ID_PROPERTY);
  if (existingId) {
    const existingSpreadsheet = SpreadsheetApp.openById(existingId);
    const existingSheet = existingSpreadsheet.getSheetByName(SHEET_NAME);
    if (existingSheet) {
      Logger.log("RSVP spreadsheet: " + existingSpreadsheet.getUrl());
      return existingSpreadsheet.getUrl();
    }
  }

  const spreadsheet = SpreadsheetApp.create("Nico's 21st RSVP Responses");
  const sheet = spreadsheet.getSheets()[0];
  sheet.setName(SHEET_NAME);
  sheet.appendRow(HEADERS);
  properties.setProperty(SPREADSHEET_ID_PROPERTY, spreadsheet.getId());
  Logger.log("RSVP spreadsheet: " + spreadsheet.getUrl());
  return spreadsheet.getUrl();
}

function getSheet() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY);
  if (!spreadsheetId) {
    throw new Error("Run setupRsvpStorage() once from the Apps Script editor before deploying.");
  }
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    if (HEADERS.some((header, index) => currentHeaders[index] !== header)) {
      throw new Error("The RSVP sheet headers must be: " + HEADERS.join(", "));
    }
  }
  return sheet;
}

function normalizeName(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function cleanText(value) {
  return String(value || "").trim();
}

function jsonOutput(value, callback) {
  const json = JSON.stringify(value);
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
