// ============================================================
// Rise8 PVD Office — Feedback Form → Google Sheets
// Deploy this from pvd@rise8.us Google account
// Script > Deploy > New deployment > Web app
//   - Execute as: Me (pvd@rise8.us)
//   - Who has access: Anyone
// Copy the deployment URL into the feedback form
// ============================================================

const SHEET_NAME = 'Responses';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Create sheet + header row if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS.map(h => h.label));
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#3B5145')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    // Build the row in header order
    const row = HEADERS.map(h => {
      const val = data[h.key];
      if (Array.isArray(val)) return val.join(', ');
      return val ?? '';
    });

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Column definitions — order here = column order in the sheet
const HEADERS = [
  { key: 'submittedAt',         label: 'Submitted At' },
  { key: 'respondentName',      label: 'Name' },
  { key: 'teamName',            label: 'Team' },
  { key: 'visitDates',          label: 'Visit Dates' },
  { key: 'visitPurpose',        label: 'Purpose' },
  { key: 'groupSize',           label: 'Group Size' },

  // Office
  { key: 'officeStars',         label: 'Office — Overall (★)' },
  { key: 'officeHighlights',    label: 'Office — What Landed' },
  { key: 'officeImprovements',  label: 'Office — Improvements' },
  { key: 'scaleComfort',        label: 'Scale — Comfort' },
  { key: 'scaleCollab',         label: 'Scale — Collaboration' },
  { key: 'scaleCreative',       label: 'Scale — Creative Inspiration' },
  { key: 'scaleProductive',     label: 'Scale — Productivity' },
  { key: 'vibesChecked',        label: 'Vibes Selected' },
  { key: 'officeWishList',      label: 'Office — Wish List' },

  // Logistics
  { key: 'restaurants',         label: 'Restaurants (name | meal | stars)' },
  { key: 'foodHighlights',      label: 'Food Highlights' },
  { key: 'hotelName',           label: 'Hotel' },
  { key: 'hotelStars',          label: 'Hotel — Rating (★)' },
  { key: 'hotelNotes',          label: 'Hotel — Notes' },
  { key: 'activitiesDone',      label: 'Activities Done' },
  { key: 'activityHighlights',  label: 'Activity Highlights' },
  { key: 'scaleTransit',        label: 'Scale — Getting Around' },
  { key: 'transitNotes',        label: 'Transit Notes' },

  // Programming
  { key: 'programStars',        label: 'Programming — Overall (★)' },
  { key: 'programHighlights',   label: 'Programming — What Landed' },
  { key: 'programFlatSpots',    label: 'Programming — What Fell Flat' },
  { key: 'scaleAgenda',         label: 'Scale — Agenda Match' },
  { key: 'programWishList',     label: 'Programming — Wish List' },
  { key: 'programOther',        label: 'Programming — Other' },

  // Final
  { key: 'wouldReturn',         label: 'Would Return?' },
  { key: 'tipForNextTeam',      label: 'Tip for Next Team' },
  { key: 'anythingElse',        label: 'Anything Else' },
];

// Handle GET requests (useful for testing the endpoint is live)
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'live', message: 'PVD feedback endpoint is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
