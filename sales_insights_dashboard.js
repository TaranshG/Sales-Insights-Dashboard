/*
 * Sales Insights Dashboard Apps Script
 * =====================================
 *
 * This script connects Google Sheets to BigQuery to pull and display key sales metrics:
 *   1. Daily revenue trends
 *   2. Average order value over time
 *   3. Month-over-month revenue growth rates
 *
 * Prerequisites:
 *   • BigQuery Advanced Service enabled in Apps Script (Extensions → Apps Script → Services → BigQuery API)
 *   • BigQuery API enabled in your Google Cloud project
 *   • Update PROJECT_ID and DATASET_ID constants below
 */

const PROJECT_ID = 'YOUR_PROJECT_ID';
const DATASET_ID = 'YOUR_DATASET_ID';
const TABLE_ID   = 'sales';  // e.g., `${DATASET_ID}.sales`

/**
 * Adds a custom “Sales Dashboard” menu to the Google Sheets UI when the spreadsheet opens.
 * Provides quick access to refresh individual reports or all at once.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sales Dashboard')
    .addItem('Refresh All', 'refreshAll')
    .addSeparator()
    .addItem('Revenue Trends', 'populateRevenueTrends')
    .addItem('Average Order Value', 'populateAverageOrderValue')
    .addItem('Growth Metrics', 'populateGrowthMetrics')
    .addToUi();
}

/**
 * Refreshes all dashboard sections by re-running each data-population function.
 */
function refreshAll() {
  populateRevenueTrends();
  populateAverageOrderValue();
  populateGrowthMetrics();
}

/**
 * Executes a SQL query against BigQuery and returns the results as a 2D array
 * including a header row.
 *
 * @param {string} sql - The standard SQL query string to execute.
 * @returns {Array.<Array<string|number>>} 2D array: [ [col1, col2, ...], [val1, val2, ...], ... ]
 */
function runBigQuery(sql) {
  const request = { query: sql, useLegacySql: false };
  let queryResults = BigQuery.Jobs.query(request, PROJECT_ID);
  const jobId = queryResults.jobReference.jobId;

  // Poll until the query job completes
  let waitMs = 1000;
  while (!queryResults.jobComplete) {
    Utilities.sleep(waitMs);
    waitMs *= 1.5;
    queryResults = BigQuery.Jobs.getQueryResults(PROJECT_ID, jobId);
  }

  // If no rows returned, return only the header
  const fields = queryResults.schema.fields;
  const headers = fields.map(field => field.name);
  const rows = queryResults.rows || [];

  // Extract values from each row
  const dataRows = rows.map(row => row.f.map(cell => cell.v));
  return [headers].concat(dataRows);
}

/**
 * Clears an existing sheet (or creates it) and writes the provided data array.
 *
 * @param {string} sheetName - Name of the sheet to write data into.
 * @param {Array.<Array>} data - 2D array including header row.
 */
function writeToSheet(sheetName, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.clear();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

/**
 * Queries daily total revenue and populates the "Revenue Trends" sheet.
 */
function populateRevenueTrends() {
  const sql = `
    SELECT
      DATE(order_date) AS date,
      SUM(revenue) AS total_revenue
    FROM \`${PROJECT_ID}.${DATASET_ID}.${TABLE_ID}\`
    GROUP BY date
    ORDER BY date;
  `;
  const data = runBigQuery(sql);
  writeToSheet('Revenue Trends', data);
}

/**
 * Queries daily average order value (AOV) and populates the "Average Order Value" sheet.
 */
function populateAverageOrderValue() {
  const sql = `
    SELECT
      DATE(order_date) AS date,
      ROUND(AVG(order_value), 2) AS avg_order_value
    FROM \`${PROJECT_ID}.${DATASET_ID}.${TABLE_ID}\`
    GROUP BY date
    ORDER BY date;
  `;
  const data = runBigQuery(sql);
  writeToSheet('Average Order Value', data);
}

/**
 * Calculates month-over-month revenue growth and populates the "Growth Metrics" sheet.
 * Uses a subquery and window function to compare each month's revenue to the previous month.
 */
function populateGrowthMetrics() {
  const sql = `
    WITH monthly AS (
      SELECT
        EXTRACT(YEAR FROM order_date) AS yr,
        EXTRACT(MONTH FROM order_date) AS mth,
        SUM(revenue) AS revenue
      FROM \`${PROJECT_ID}.${DATASET_ID}.${TABLE_ID}\`
      GROUP BY yr, mth
    ),
    ranked AS (
      SELECT
        yr,
        mth,
        revenue,
        LAG(revenue) OVER (ORDER BY yr, mth) AS prev_revenue
      FROM monthly
    )
    SELECT
      CONCAT(yr, '-', LPAD(mth, 2, '0')) AS year_month,
      revenue,
      ROUND(SAFE_DIVIDE(revenue - prev_revenue, prev_revenue) * 100, 2) AS mom_growth_pct
    FROM ranked
    ORDER BY yr, mth;
  `;
  const data = runBigQuery(sql);
  writeToSheet('Growth Metrics', data);
}
