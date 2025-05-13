/**
 * ---------------------------------------------------------------
 * SpotDev – HubSpot Form Submission Exporter
 * ---------------------------------------------------------------
 * Purpose
 * -------
 *  • Retrieves all HubSpot forms whose *name* contains a given
 *    search phrase (REQUEST_TERM).
 *  • For each matching form, fetches raw submissions for the
 *    last N days (DAYS_BACK, default 30).
 *  • Writes the results to an Excel workbook
 *    (file: form-submissions.xlsx) with the following columns:
 *        - Form GUID
 *        - Form Name
 *        - Form Submission ID
 *        - Time Submitted (ISO 8601)
 *        - Contact Email Address
 *        - Page converted on
 *        - All Properties (JSON-stringified array of every field)
 *
 * Prerequisites
 * -------------
 *   • Node.js v18+ (for native `fetch`, though we use Axios here)
 *   • `npm i axios xlsx dotenv`
 *   • A HubSpot *private-app token* with at least:
 *       – forms (read)
 *       – business-intelligence (read)
 *
 * Environment Variables
 * ---------------------
 *   HUBSPOT_PAT   – REQUIRED – your private-app token
 *   REQUEST_TERM  – OPTIONAL – text to match in form names
 *                                 (default 'exprom')
 *   DAYS_BACK     – OPTIONAL – look-back window in days
 *                                 (default 30)
 *
 * Usage
 * -----
 *   HUBSPOT_PAT=pat_xxx node export-forms.js
 *
 *   # override defaults
 *   REQUEST_TERM="Contact" DAYS_BACK=90 node export-forms.js
 *
 * Notes
 * -----
 *   • The script is intentionally single-threaded (one HTTP request
 *     in flight at any time) to stay well within HubSpot’s
 *     rate limits.  Parallelism is straightforward to add if ever
 *     required – drop in p-limit or Promise.all with a throttle.
 *   • Fail-fast strategy: any HTTP error bubbles and crashes the
 *     process, so CI/CD tasks fail loudly.
 *   • Written in *vanilla JavaScript* and Axios to respect standard SpotDev coding standards.
 */

import axios  from 'axios';
import xlsx   from 'xlsx';
import dotenv from 'dotenv';

/* -------------------------------------------------------------------------
 * 1. Environment & Configuration
 * ---------------------------------------------------------------------- */

dotenv.config();                           // Load .env (if present)

/** HubSpot private-app token (required) */
const HUBSPOT_PAT = process.env.HUBSPOT_PAT;
if (!HUBSPOT_PAT) {
  throw new Error('Environment variable HUBSPOT_PAT is missing.');
}

/** Case-insensitive text used to filter form names */
const REQUEST_TERM = (process.env.REQUEST_TERM || 'exprom').toLowerCase();

/** How many days of submissions we retain */
const DAYS_BACK = Number(process.env.DAYS_BACK) || 30;

/** UNIX timestamp that defines the earliest acceptable submission */
const CUTOFF_TS = Date.now() - DAYS_BACK * 24 * 60 * 60 * 1_000;

/* -------------------------------------------------------------------------
 * 2. Axios instance (shared by all requests)
 * ---------------------------------------------------------------------- */

const hubspot = axios.create({
  baseURL : 'https://api.hubapi.com',
  headers : { Authorization: `Bearer ${HUBSPOT_PAT}` },
  timeout : 10_000,        // 10-second network timeout
});

/* -------------------------------------------------------------------------
 * 3. Helper Functions
 * ---------------------------------------------------------------------- */

/**
 * Fetch *every* form in the portal.
 * Endpoint: GET /marketing/v3/forms
 * HubSpot paginates with `after`; we loop until exhausted.
 *
 * @returns {Promise<Array>} – array of form objects
 */
async function getAllForms() {
  const forms = [];
  let after;                // undefined on first request

  do {
    const { data } = await hubspot.get('/marketing/v3/forms', {
      params: { limit: 100, after },
    });

    forms.push(...data.results);
    after = data.paging?.next?.after;   // undefined/null when no more pages
  } while (after);

  return forms;
}

/**
 * Generator that yields *each* submission for a single form.
 * Endpoint: GET /form-integrations/v1/submissions/forms/{formId}
 * HubSpot paginates with `offset`; we loop until finished.
 *
 * @param   {string} formId                     – GUID of the form
 * @yields  {object} submission object (raw API payload)
 */
async function* getSubmissions(formId) {
  let offset = 0;           // HubSpot returns 0 when no next page
  const PAGE_SIZE = 50;     // chosen for safety vs. rate limits

  do {
    const { data } = await hubspot.get(
      `/form-integrations/v1/submissions/forms/${formId}`,
      { params: { limit: PAGE_SIZE, offset } },
    );

    for (const submission of data.results) yield submission;

    offset = data.offset ?? 0;
  } while (offset);
}

/**
 * Pulls the primary **email** value from a submission’s `values` array.
 * Looks specifically for the CRM contact property (objectTypeId '0-1').
 *
 * @param   {Array} values – submission.values array
 * @returns {string}       – email address or '' when absent
 */
function extractEmail(values) {
  const match = values?.find(
    v => v.name === 'email' && v.objectTypeId === '0-1',
  );
  return match?.value || '';
}

/* -------------------------------------------------------------------------
 * 4. Main Flow
 * ---------------------------------------------------------------------- */

(async () => {
  console.log(`⏳ Fetching all forms...`);
  const allForms = await getAllForms();

  /** Forms whose *name* contains REQUEST_TERM (case-insensitive) */
  const targetForms = allForms.filter(f =>
    (f.name || '').toLowerCase().includes(REQUEST_TERM),
  );

  if (!targetForms.length) {
    console.log(`❌ No forms found containing “${REQUEST_TERM}”.`);
    return;
  }

  console.log(
    `🔍 Found ${targetForms.length} matching form(s). ` +
    `Collecting submissions from the last ${DAYS_BACK} day(s)...`,
  );

  const rows = [];          // array of row objects for Excel

  /** Iterate through each form then each submission */
  for (const form of targetForms) {
    for await (const sub of getSubmissions(form.id)) {
      // Stop iterating this form once we hit submissions older than the cutoff
      if (sub.submittedAt < CUTOFF_TS) break;

      rows.push({
        'Form GUID'            : form.id,
        'Form Name'            : form.name,
        'Form Submission ID'   : sub.conversionId,
        'Time Submitted'       : new Date(sub.submittedAt).toISOString(),
        'Contact Email Address': extractEmail(sub.values),
        'Page converted on'    : sub.pageUrl,
        'All Properties'       : JSON.stringify(sub.values), // preserve entire array
      });
    }
  }

  if (!rows.length) {
    console.log(`⚠️  No submissions in the last ${DAYS_BACK} day(s).`);
    return;
  }

  /* -----------------------------------------------------------------------
   * 5. Build & Write the Excel Workbook
   * -------------------------------------------------------------------- */

  const workbook  = xlsx.utils.book_new();
  const worksheet = xlsx.utils.json_to_sheet(rows, {
    header: [
      'Form GUID',
      'Form Name',
      'Form Submission ID',
      'Time Submitted',
      'Contact Email Address',
      'Page converted on',
      'All Properties',
    ],
  });

  xlsx.utils.book_append_sheet(workbook, worksheet, 'Submissions');
  xlsx.writeFile(workbook, 'form-submissions.xlsx');

  console.log(
    `✅ Export complete – form-submissions.xlsx ` +
    `with ${rows.length} row(s).`,
  );
})().catch(err => {
  // Any unhandled rejection lands here
  console.error('❌ Fatal error:', err.message);
  process.exit(1);
});