// server/index.js
const express = require("express");
const cors = require("cors");
const path = require("path");
const { google } = require('googleapis');
const fs = require('fs');
require("dotenv").config({ path: path.resolve(__dirname, '../.env') });
console.log("GOOGLE_CLIENT_ID:", process.env.GOOGLE_CLIENT_ID);
console.log("GOOGLE_CLIENT_SECRET:", process.env.GOOGLE_CLIENT_SECRET);
console.log("GOOGLE_REDIRECT_URI:", process.env.GOOGLE_REDIRECT_URI);

const {
  getAuthUrl,
  getTokens,
  createSheet,
  appendToSheet,
  getOAuthClient
} = require("./google");

const app = express();
app.use(cors());
app.use(express.json());

// In-memory token store — replace with DB for real apps
let storedTokens = null;      // Google tokens
let zohoTokens = null;        // Zoho tokens

// Google OAuth endpoint
app.get("/auth", (req, res) => {
  const url = getAuthUrl();
  console.log("Generated Google Auth URL:", url);
  res.redirect(url);
});

// OAuth2 callback - save tokens
// server/index.js
// server/index.js
app.get("/oauth2callback", async (req, res) => {
  try {
    const code = req.query.code;
    const tokens = await getTokens(code);

    if (tokens.refresh_token) {
      storedTokens = tokens;
    } else if (storedTokens) {
      tokens.refresh_token = storedTokens.refresh_token;
      storedTokens = tokens;
    } else {
      return res.status(400).json({ error: "No refresh token found. Please reauthorize." });
    }

    // After successful Google OAuth in /oauth2callback
    return res.send(`
      <html>
        <body>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: "google-auth-success" }, "*");
              window.close();
            } else {
              window.location = "/app/widget.html?google=success";
            }
          </script>
        </body>
      </html>
    `);
  } catch (err) {
    return res.send(`
      <html>
        <body>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: "google-auth-fail" }, "*");
              window.close();
            } else {
              window.location = "/app/widget.html?google=fail";
            }
          </script>
        </body>
      </html>
    `);
  }
});

// Middleware to check token availability
function ensureTokens(req, res, next) {
  if (!storedTokens) {
    return res.status(401).json({ error: "User not authenticated." });
  }
  next();
}

// Create sheet endpoint
app.post("/create-sheet", ensureTokens, async (req, res) => {
  try {
    const { sheetTitle } = req.body;
    const sheet = await createSheet(storedTokens, sheetTitle);
    res.json({ sheet });
  } catch (err) {
    res.status(500).json({ error: "Failed to create sheet", details: err.message });
  }
});

// Append to sheet endpoint
app.post("/write-data", ensureTokens, async (req, res) => {
  try {
    const { spreadsheetId, rows, chartType, xField, yField } = req.body;

    let finalRows = rows;

    // Only apply for pie/doughnut chart exports
    if (
      (chartType && (chartType.toUpperCase() === "PIE" || chartType.toUpperCase() === "DOUGHNUT")) &&
      xField && yField && Array.isArray(rows) && rows.length > 1
    ) {
      const headers = rows[0];
      const xIndex = headers.indexOf(xField);
      const yIndex = headers.indexOf(yField);
      if (xIndex !== -1 && yIndex !== -1) {
        finalRows = [headers].concat(
          rows.slice(1).map(row => {
            const newRow = [...row];
            newRow[xIndex] = String(row[xIndex]);
            const num = Number(row[yIndex]);
            newRow[yIndex] = isNaN(num) ? 0 : num;
            return newRow;
          })
        );
      }
    }

    // --- Only add header if sheet is empty ---
    const sheets = google.sheets({ version: "v4", auth: getOAuthClient(storedTokens) });
    const getRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "Sheet1"
    });
    const sheetIsEmpty = !getRes.data.values || getRes.data.values.length === 0;

    let rowsToWrite = finalRows;
    if (!sheetIsEmpty) {
      // Remove header row if appending to non-empty sheet
      rowsToWrite = finalRows.slice(1);
    }

    await appendToSheet(storedTokens, spreadsheetId, rowsToWrite);
    res.json({ message: "Data written to sheet successfully" });
  } catch (err) {
    res.status(500).json({ error: "Failed to write data", details: err.message });
  }
});

// Logout endpoint - revoke token & clear storedTokens
app.post("/logout", async (req, res) => {
  try {
    if (storedTokens && storedTokens.access_token) {
      const oauth2Client = getOAuthClient(storedTokens);
      await oauth2Client.revokeToken(storedTokens.access_token);
    }
    storedTokens = null;
    zohoTokens = null; // <-- Add this line
    res.json({ message: "Logged out successfully" });
  } catch (err) {
    res.status(500).json({ error: "Logout failed", details: err.message });
  }
});

// Zoho OAuth endpoint
app.get("/zoho/auth", (req, res) => {
  const clientId = process.env.ZOHO_CLIENT_ID;
  const redirectUri = encodeURIComponent(process.env.ZOHO_REDIRECT_URI);
  const scope = encodeURIComponent("ZohoBooks.fullaccess.all");
  const responseType = "code";
  const accessType = "offline";
  const prompt = "consent";
  const state = "custom_state";
  const zohoAuthUrl = `https://accounts.zoho.in/oauth/v2/auth?scope=${scope}&client_id=${clientId}&response_type=${responseType}&access_type=${accessType}&redirect_uri=${redirectUri}&prompt=${prompt}&state=${state}`;
  res.redirect(zohoAuthUrl);
});

// Zoho OAuth callback
app.get("/zoho/oauth2callback", async (req, res) => {
  const code = req.query.code;
  if (!code) return res.status(400).send("Missing code");

  const params = new URLSearchParams({
    code,
    client_id: process.env.ZOHO_CLIENT_ID,
    client_secret: process.env.ZOHO_CLIENT_SECRET,
    redirect_uri: process.env.ZOHO_REDIRECT_URI,
    grant_type: "authorization_code"
  });

  try {
    const fetch = (...args) => import('node-fetch').then(({default: fetch}) => fetch(...args));
    const response = await fetch("https://accounts.zoho.in/oauth/v2/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params
    });
    const data = await response.json();
    if (data.access_token) {
      // Fetch organization ID using the access token
      const orgRes = await fetch("https://www.zohoapis.in/books/v3/organizations", {
        headers: {
          Authorization: `Zoho-oauthtoken ${data.access_token}`
        }
      });
      const orgData = await orgRes.json();
      console.log("Zoho organizations API response:", JSON.stringify(orgData)); // <-- Add this line
      let orgId = null;
      if (orgData.organizations && orgData.organizations.length > 0) {
        const defaultOrg = orgData.organizations.find(o => o.is_default_org) || orgData.organizations[0];
        orgId = defaultOrg.organization_id;
      }
      // Store both tokens and orgId in memory (or session/db for multi-user)
      zohoTokens = { ...data, org_id: orgId };
      console.log("Fetched Zoho Books Organization ID:", orgId); // <-- Add this line
      // Instead of redirect, send a script to close the popup and notify parent
      return res.send(`
        <html>
          <body>
            <script>
              if (window.opener) {
                window.opener.postMessage({ type: "zoho-auth-success" }, "*");
                window.close();
              } else {
                window.location = "/app/widget.html?zoho=success";
              }
            </script>
          </body>
        </html>
      `);
    } else {
      return res.send(`
        <html>
          <body>
            <script>
              if (window.opener) {
                window.opener.postMessage({ type: "zoho-auth-fail" }, "*");
                window.close();
              } else {
                window.location = "/app/widget.html?zoho=fail";
              }
            </script>
          </body>
        </html>
      `);
    }
  } catch (err) {
    return res.redirect("/app/widget.html?zoho=fail");
  }
});

// Endpoint to check Zoho auth status
app.get("/zoho/status", (req, res) => {
  res.json({ authenticated: !!zohoTokens });
});

// New endpoint to fetch Zoho modules
app.get("/zoho/modules", async (req, res) => {
  if (!zohoTokens || !zohoTokens.access_token || !zohoTokens.org_id) {
    return res.status(401).json({ error: "Not authenticated with Zoho" });
  }
  try {
    const fetch = (...args) => import('node-fetch').then(({default: fetch}) => fetch(...args));
    const orgId = zohoTokens.org_id;
    const url = `https://www.zohoapis.in/books/v3/settings/modules?organization_id=${orgId}`;
    const response = await fetch(url, {
      headers: {
        Authorization: `Zoho-oauthtoken ${zohoTokens.access_token}`
      }
    });
    const data = await response.json();
    console.log("Zoho /settings/modules API response:", JSON.stringify(data));
    let modules = (data.modules || [])
      .filter(m => m.is_active && m.api_name)
      .map(m => ({
        value: m.api_name,
        label: m.module_name
      }));

    // Fallback: Always add "invoices" if modules is empty (for testing)
    if (modules.length === 0) {
      modules.push({ value: "invoices", label: "Invoices" });
      modules.push({ value: "contacts", label: "Contacts" });
      modules.push({ value: "items", label: "Items" });
      modules.push({ value: "salesorders", label: "Sales Orders" });
        modules.push({ value: "recurringinvoices", label: "Recurring Invoices" });
      modules.push({ value: "creditnotes", label: "Credit Notes" }); // Add this line
    }

    res.json({ modules });
  } catch (err) {
    res.status(500).json({ error: "Failed to fetch modules", details: err.message });
  }
});

// New endpoint to fetch Zoho module data
app.get("/zoho/data/:module", async (req, res) => {
  if (!zohoTokens || !zohoTokens.access_token || !zohoTokens.org_id) {
    return res.status(401).json({ error: "Not authenticated with Zoho" });
  }
  try {
    const fetch = (...args) => import('node-fetch').then(({default: fetch}) => fetch(...args));
    const orgId = zohoTokens.org_id;
    const module = req.params.module;

    // Build query string from req.query
    const params = new URLSearchParams({ organization_id: orgId });
    for (const [key, value] of Object.entries(req.query)) {
      params.append(key, value);
    }

    const url = `https://www.zohoapis.in/books/v3/${module}?${params.toString()}`;
    const response = await fetch(url, {
      headers: {
        Authorization: `Zoho-oauthtoken ${zohoTokens.access_token}`
      }
    });
    const data = await response.json();

    // Try plural, then singular, then first array property
    let result = data[module];
    if (!result) {
      result = data[`${module}`];
    }
    if (!result) {
      const singular = module.endsWith('s') ? module.slice(0, -1) : module;
      result = data[singular];
    }
    if (!result) {
      result = Object.values(data).find(v => Array.isArray(v));
    }

    // Only display required fields for each module
    let filtered = result || [];
    if (module === "invoices") {
      filtered = filtered.map(inv => ({
        invoice_id: inv.invoice_id,
        invoice_number: inv.invoice_number,
        customer_name: inv.customer_name,
        email: inv.customer_email, // often useful
        date: inv.date,
        due_date: inv.due_date,
        status: inv.status,
        total: inv.total,
        balance: inv.balance,
        currency_code: inv.currency_code,
        created_time: inv.created_time,
        last_modified_time: inv.last_modified_time
      }));
    } else if (module === "contacts") {
      filtered = filtered.map(c => ({
        contact_id: c.contact_id,
        contact_name: c.contact_name,
        company_name: c.company_name,
        contact_type: c.contact_type,
        status: c.status,
        email: c.email,
        phone: c.phone,
        outstanding_receivable_amount: c.outstanding_receivable_amount,
        unused_credits_receivable_amount: c.unused_credits_receivable_amount,
        gst_no: c.gst_no,
        place_of_contact: c.place_of_contact,
        payment_terms: c.payment_terms,
        payment_terms_label: c.payment_terms_label,
        currency_code: c.currency_code,
        created_time: c.created_time,
        last_modified_time: c.last_modified_time,
        billing_address: c.billing_address ? `${c.billing_address.address}, ${c.billing_address.city}, ${c.billing_address.state}` : "",
        shipping_address: c.shipping_address ? `${c.shipping_address.address}, ${c.shipping_address.city}, ${c.shipping_address.state}` : "",
        contact_persons: Array.isArray(c.contact_persons) ? c.contact_persons.map(p => `${p.first_name} ${p.last_name} (${p.email})`).join("; ") : ""
      }));
    } else if (module === "items") {
      filtered = filtered.map(i => ({
        item_id: i.item_id,
        name: i.name,
        description: i.description,
        rate: i.rate,
        status: i.status,
        sku: i.sku,
        product_type: i.product_type,
        quantity_available: i.quantity_available,
        created_time: i.created_time,
        last_modified_time: i.last_modified_time
      }));
    } else if (module === "salesorders") {
      filtered = filtered.map(so => ({
        salesorder_id: so.salesorder_id,
        salesorder_number: so.salesorder_number,
        date: so.date,
        status: so.status,
        customer_id: so.customer_id,
        customer_name: so.customer_name,
        reference_number: so.reference_number,
        shipment_date: so.shipment_date,
        total: so.total,
        sub_total: so.sub_total,
        currency_code: so.currency_code,
        discount: so.discount,
        discount_amount: so.discount_amount,
        shipping_charge: so.shipping_charge,
        adjustment: so.adjustment,
        salesperson_name: so.salesperson_name,
        merchant_name: so.merchant_name,
        billing_address: so.billing_address ? `${so.billing_address.address}, ${so.billing_address.city}, ${so.billing_address.state}` : "",
        shipping_address: so.shipping_address ? `${so.shipping_address.address}, ${so.shipping_address.city}, ${so.shipping_address.state}` : "",
        notes: so.notes,
        created_time: so.created_time,
        last_modified_time: so.last_modified_time,
        line_items: Array.isArray(so.line_items)
          ? so.line_items.map(item => ({
              name: item.name,
              description: item.description,
              quantity: item.quantity,
              rate: item.rate,
              total: item.item_total_inclusive_of_tax || (item.rate * item.quantity)
            }))
          : []
      }));
    } else if (module === "recurringinvoices") {
      filtered = filtered.map(ri => ({
        recurring_invoice_id: ri.recurring_invoice_id,
        recurrence_name: ri.recurrence_name,
        reference_number: ri.reference_number,
        customer_id: ri.customer_id,
        customer_name: ri.customer_name,
        currency_code: ri.currency_code,
        start_date: ri.start_date,
        end_date: ri.end_date,
        last_sent_date: ri.last_sent_date,
        next_invoice_date: ri.next_invoice_date,
        status: ri.status,
        billing_address: ri.billing_address ? `${ri.billing_address.address}, ${ri.billing_address.city}, ${ri.billing_address.state}` : "",
        shipping_address: ri.shipping_address ? `${ri.shipping_address.address}, ${ri.shipping_address.city}, ${ri.shipping_address.state}` : "",
        custom_fields: ri.custom_fields,
        line_items: Array.isArray(ri.line_items)
          ? ri.line_items.map(item => ({
              name: item.name,
              description: item.description,
              quantity: item.quantity,
              rate: item.rate,
              total: item.item_total
            }))
          : []
      }));
    } else if (module === "creditnotes") {
      filtered = filtered.map(cn => ({
        creditnote_id: cn.creditnote_id,
        creditnote_number: cn.creditnote_number,
        date: cn.date,
        status: cn.status,
        customer_id: cn.customer_id,
        customer_name: cn.customer_name,
        reference_number: cn.reference_number,
        email: cn.email,
        total: cn.total,
        balance: cn.balance,
        currency_code: cn.currency_code,
        created_time: cn.created_time,
        updated_time: cn.updated_time,
        billing_address: cn.billing_address ? `${cn.billing_address.address}, ${cn.billing_address.city}, ${cn.billing_address.state}` : "",
        shipping_address: cn.shipping_address ? `${cn.shipping_address.address}, ${cn.shipping_address.city}, ${cn.shipping_address.state}` : "",
        notes: cn.notes,
        terms: cn.terms,
        custom_fields: cn.custom_fields,
        line_items: Array.isArray(cn.line_items)
          ? cn.line_items.map(item => ({
              name: item.name,
              description: item.description,
              quantity: item.quantity,
              rate: item.rate,
              total: item.rate * item.quantity
            }))
          : []
      }));
    }

    res.json({ data: filtered });
  } catch (err) {
    res.status(500).json({ error: "Failed to fetch module data", details: err.message });
  }
});

// Helper: get or create a sheet by title, returns {sheetId, title}
async function getOrCreateSheet(sheets, spreadsheetId, title) {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  let sheet = spreadsheet.data.sheets.find(s => s.properties.title === title);
  if (!sheet) {
    const addRes = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title } } }]
      }
    });
    sheet = addRes.data.replies[0].addSheet.properties;
  } else {
    sheet = sheet.properties;
  }
  return { sheetId: sheet.sheetId, title };
}

// --- REPLACE your /write-chart endpoint with this ---
app.post('/write-chart', ensureTokens, async (req, res) => {
  try {
    const { spreadsheetId, chartType, chartName, chartData } = req.body;
    if (!spreadsheetId || !chartType || !chartName || !chartData) {
      return res.status(400).json({ error: 'Missing required chart parameters' });
    }

    const auth = getOAuthClient(storedTokens);
    const sheets = google.sheets({ version: 'v4', auth });

    // Write chartData to a dedicated sheet (e.g., "ChartData")
    const sheetTitle = "ChartData";
    const { sheetId } = await getOrCreateSheet(sheets, spreadsheetId, sheetTitle);

    // Clear the sheet first
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `${sheetTitle}!A:Z`
    });

    // Write the new chart data
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetTitle}!A1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: chartData }
    });

    // Add the chart (skip header row)
    const rowCount = chartData.length;
    if (rowCount < 2) {
      return res.status(400).json({ error: "Not enough data for chart." });
    }

    // Chart type mapping
    const chartTypeMap = {
      bar: 'COLUMN',
      line: 'LINE',
      pie: 'PIE',
      doughnut: 'PIE',      // Map unsupported to PIE
      scatter: 'SCATTER',
      radar: 'COLUMN',      // Map unsupported to COLUMN or LINE
      polarArea: 'PIE'      // Map unsupported to PIE
    };
    const sheetsChartType = chartTypeMap[req.body.chartType] || 'COLUMN';

    // Remove any existing charts on this sheet (optional, for clean-up)
    const sheetMeta = await sheets.spreadsheets.get({ spreadsheetId });
    const charts = (sheetMeta.data.sheets || []).find(s => s.properties.sheetId === sheetId)?.charts || [];
    if (charts.length) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: charts.map(c => ({ deleteEmbeddedObject: { objectId: c.chartId } }))
        }
      });
    }

    // Build chart request (place at A1 for visibility)
    let addChartRequest;
    if (sheetsChartType === 'PIE') {
      // PIE chart uses pieChart, not basicChart
      addChartRequest = {
        addChart: {
          chart: {
            spec: {
              title: chartName,
              pieChart: {
                legendPosition: 'RIGHT_LEGEND',
                domain: {
                  sourceRange: {
                    sources: [{
                      sheetId,
                      startRowIndex: 1,
                      endRowIndex: rowCount,
                      startColumnIndex: 0,
                      endColumnIndex: 1
                    }]
                  }
                },
                series: {
                  sourceRange: {
                    sources: [{
                      sheetId,
                      startRowIndex: 1,
                      endRowIndex: rowCount,
                      startColumnIndex: 1,
                      endColumnIndex: 2
                    }]
                  }
                }
              }
            },
            position: {
              overlayPosition: {
                anchorCell: {
                  sheetId,
                  rowIndex: 0,
                  columnIndex: 0
                }
              }
            }
          }
        }
      };
    } else {
      // All other supported types use basicChart
      addChartRequest = {
        addChart: {
          chart: {
            spec: {
              title: chartName,
              basicChart: {
                chartType: sheetsChartType,
                legendPosition: 'RIGHT_LEGEND',
                axis: [
                  { position: 'BOTTOM_AXIS', title: chartData[0][0] },
                  { position: 'LEFT_AXIS', title: chartData[0][1] }
                ],
                domains: [{
                  domain: {
                    sourceRange: {
                      sources: [{
                        sheetId,
                        startRowIndex: 1,
                        endRowIndex: rowCount,
                        startColumnIndex: 0,
                        endColumnIndex: 1
                      }]
                    }
                  }
                }],
                series: [{
                  series: {
                    sourceRange: {
                      sources: [{
                        sheetId,
                        startRowIndex: 1,
                        endRowIndex: rowCount,
                        startColumnIndex: 1,
                        endColumnIndex: 2
                      }]
                    }
                  },
                  targetAxis: 'LEFT_AXIS'
                }]
              }
            },
            position: {
              overlayPosition: {
                anchorCell: {
                  sheetId,
                  rowIndex: 0,
                  columnIndex: 0
                }
              }
            }
          }
        }
      };
    }

    // <-- Place the batchUpdate call here:
    const chartResponse = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [addChartRequest]
      }
    });
    console.log("Chart batchUpdate response:", JSON.stringify(chartResponse.data));

    res.json({ message: "Chart created successfully!" });
  } catch (err) {
    console.error("write-chart error:", err);
    res.status(500).json({ error: err.message });
  }
});

// New endpoint to list Google Sheets
app.get("/google/sheets", async (req, res) => {
  try {
    if (!storedTokens) {
      return res.status(401).json({ error: "Not authenticated with Google" });
    }
    const auth = getOAuthClient(storedTokens); // getOAuthClient should accept tokens
    const drive = google.drive({ version: "v3", auth });
    const result = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
      fields: "files(id, name)",
      pageSize: 50
    });
    res.json({ sheets: result.data.files });
  } catch (err) {
    console.error("Error in /google/sheets:", err);
    res.status(500).json({ error: "Failed to list sheets", details: err.message });
  }
});

// New endpoint to get sheet row count
app.get("/sheet-row-count", ensureTokens, async (req, res) => {
  const { spreadsheetId } = req.query;
  const sheets = google.sheets({ version: "v4", auth: getOAuthClient(storedTokens) });
  const getRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "Sheet1"
  });
  const rowCount = getRes.data.values ? getRes.data.values.length : 0;
  res.json({ rowCount });
});

// Serve frontend and start server (same as before)...

app.use("/app", express.static(path.join(__dirname, "../app")));
app.get("/", (req, res) => res.redirect("/app/widget.html"));

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`✅ Server running on http://localhost:${PORT}`);
});
