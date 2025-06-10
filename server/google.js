const path = require("path");
require("dotenv").config({ path: path.resolve(__dirname, '../.env') });
const { google } = require("googleapis");

const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive.file"
];

function getOAuthClient(tokens) {
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI // Use the correct redirect URI from .env
  );

  if (tokens) {
    oauth2Client.setCredentials(tokens);
  }

  return oauth2Client;
}

function getAuthUrl() {
  const oauth2Client = getOAuthClient();
  return oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
    prompt: "consent"
  });
}

async function getTokens(code) {
  const oauth2Client = getOAuthClient();
  const { tokens } = await oauth2Client.getToken(code);
  return tokens;
}

async function createSheet(tokens, sheetTitle = "New Sheet") {
  const oauth2Client = getOAuthClient(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });

  const response = await sheets.spreadsheets.create({
    resource: { properties: { title: sheetTitle } }
  });

  return response.data;
}

async function appendToSheet(tokens, spreadsheetId, rows) {
  const oauth2Client = getOAuthClient(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: "A1",
    valueInputOption: "USER_ENTERED",
    resource: { values: rows } // <-- FIXED
  });
}

module.exports = {
  getAuthUrl,
  getTokens,
  createSheet,
  appendToSheet,
  getOAuthClient
};
