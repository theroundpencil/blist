/**
 * BLIST - Book List Tracker
 * Main Apps Script backend
 * Deploy as: Web App → Execute as "User accessing the web app"
 */

// ── Constants ──
const SHEET_NAME = 'Blist_Books';
const PROPS_KEY = 'blist_sheet_id';
const FOLDER_KEY = 'blist_folder_id';
const FOLDER_NAME = 'Library';
const HEADERS = [
  'ID', 'Title', 'Author', 'ISBN', 'CoverURL', 'Description',
  'Format', 'Location', 'Ownership', 'Status',
  'Rating', 'Recommend', 'Notes', 'DateAdded', 'DateModified'
];

/** Normalize a book title for duplicate comparison:
 *  lowercase, strip leading articles (the/a/an), remove punctuation, collapse whitespace */
function normalizeTitle_(title) {
  if (!title) return '';
  return String(title)
    .toLowerCase()
    .replace(/^(the|a|an)\s+/i, '')  // strip leading articles
    .replace(/[^a-z0-9\s]/g, '')     // remove punctuation
    .replace(/\s+/g, ' ')            // collapse whitespace
    .trim();
}

// ── Web App Entry ──
function doGet(e) {
  // JSON API endpoint for widgets: ?action=stats
  if (e && e.parameter && e.parameter.action === 'stats') {
    const stats = getDashboardStats();
    return ContentService.createTextOutput(JSON.stringify(stats))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Blist – Book List Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/** Get a public URL for the app icon (creates it in Library folder if needed) */
function getIconUrl() {
  const userProps = PropertiesService.getUserProperties();
  let iconUrl = userProps.getProperty('blist_icon_url');
  if (iconUrl) {
    try {
      const fileId = iconUrl.split('id=')[1] || iconUrl.split('/d/')[1].split('/')[0];
      DriveApp.getFileById(fileId);
      return iconUrl;
    } catch(e) { /* recreate */ }
  }

  // 192x192 PNG icon — open book with red bookmark on yellow bg
  const b64 = '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';
  const bytes = Utilities.base64Decode(b64);
  const blob = Utilities.newBlob(bytes, 'image/png', 'blist-icon-192.png');

  // Save to Library folder
  const folder = getOrCreateFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  iconUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
  userProps.setProperty('blist_icon_url', iconUrl);
  return iconUrl;
}

// ── Banner dismiss (stored server-side since localStorage doesn't work in Apps Script iframe) ──

function setBannerDismissed() {
  PropertiesService.getUserProperties().setProperty('blist_banner_dismissed', '1');
}

function isBannerDismissed() {
  return PropertiesService.getUserProperties().getProperty('blist_banner_dismissed') === '1';
}

// ── Sheet Management ──

/** Get or create the "Library" folder in user's Drive */
function getOrCreateFolder_() {
  const userProps = PropertiesService.getUserProperties();
  let folderId = userProps.getProperty(FOLDER_KEY);

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      userProps.deleteProperty(FOLDER_KEY);
    }
  }

  // Check if folder already exists in root
  const folders = DriveApp.getRootFolder().getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) {
    const folder = folders.next();
    userProps.setProperty(FOLDER_KEY, folder.getId());
    return folder;
  }

  // Create new folder
  const folder = DriveApp.createFolder(FOLDER_NAME);
  userProps.setProperty(FOLDER_KEY, folder.getId());
  return folder;
}

/** Get or create the user's Blist sheet inside Library folder */
function getOrCreateSheet_() {
  const userProps = PropertiesService.getUserProperties();
  let sheetId = userProps.getProperty(PROPS_KEY);

  // Try to open existing sheet
  if (sheetId) {
    try {
      const ss = SpreadsheetApp.openById(sheetId);
      return ss.getSheetByName(SHEET_NAME) || createBlistSheet_(ss);
    } catch (e) {
      userProps.deleteProperty(PROPS_KEY);
    }
  }

  // Create new spreadsheet and move to Library folder
  const ss = SpreadsheetApp.create('Blist – My Book List');
  const file = DriveApp.getFileById(ss.getId());
  const folder = getOrCreateFolder_();
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file); // Remove from root
  userProps.setProperty(PROPS_KEY, ss.getId());
  return createBlistSheet_(ss);
}

/** Set up the sheet with headers and formatting */
function createBlistSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getActiveSheet().setName(SHEET_NAME);
  }
  // Write headers
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#FFD60A')
    .setFontColor('#1a1a1a');
  sheet.setFrozenRows(1);
  return sheet;
}

// ── Book Search ──
// Primary: Open Library API (free, no auth/key needed)
// Fallback: Google Books API with OAuth token

function searchBooks(query) {
  // Try Open Library first (always works, no auth)
  try {
    const results = searchOpenLibrary_(query);
    if (results.length > 0) return results;
  } catch (e) {
    Logger.log('Open Library error: ' + e.message);
  }

  // Fallback: Google Books with OAuth
  try {
    return searchGoogleBooks_(query);
  } catch (e2) {
    Logger.log('Google Books fallback error: ' + e2.message);
    return [];
  }
}

/** Open Library search — free, no key needed */
function searchOpenLibrary_(query) {
  const url = 'https://openlibrary.org/search.json?q=' +
    encodeURIComponent(query) + '&limit=8&fields=key,title,author_name,isbn,cover_i,first_publish_year,number_of_pages_median,subject';
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = response.getResponseCode();
  if (code !== 200) {
    Logger.log('Open Library HTTP ' + code);
    return [];
  }

  const data = JSON.parse(response.getContentText());
  if (!data.docs || data.docs.length === 0) return [];

  return data.docs.map(doc => ({
    googleId: doc.key || '',
    title: doc.title || 'Unknown',
    author: (doc.author_name || []).join(', ') || 'Unknown',
    isbn: (doc.isbn || [])[0] || '',
    coverUrl: doc.cover_i
      ? 'https://covers.openlibrary.org/b/id/' + doc.cover_i + '-L.jpg'
      : '',
    description: (doc.subject || []).slice(0, 5).join(', ') || '',
    pageCount: doc.number_of_pages_median || 0,
    publishedDate: doc.first_publish_year ? String(doc.first_publish_year) : ''
  }));
}

/** Google Books with OAuth token (fallback) */
function searchGoogleBooks_(query) {
  const token = ScriptApp.getOAuthToken();
  const url = 'https://www.googleapis.com/books/v1/volumes?q=' +
    encodeURIComponent(query) + '&maxResults=8&printType=books';
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { 'Authorization': 'Bearer ' + token }
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    Logger.log('Google Books HTTP ' + code + ': ' + response.getContentText().substring(0, 200));
    return [];
  }

  const data = JSON.parse(response.getContentText());
  if (!data.items) return [];
  return data.items.map(parseGoogleBookItem_);
}

/** Parser for Google Books volume item */
function parseGoogleBookItem_(item) {
  const info = item.volumeInfo || {};
  const isbn = (info.industryIdentifiers || [])
    .find(id => id.type === 'ISBN_13' || id.type === 'ISBN_10');
  return {
    googleId: item.id,
    title: info.title || 'Unknown',
    author: (info.authors || []).join(', ') || 'Unknown',
    isbn: isbn ? isbn.identifier : '',
    coverUrl: info.imageLinks
      ? (info.imageLinks.thumbnail || '').replace('http://', 'https://')
      : '',
    description: (info.description || '').substring(0, 300),
    pageCount: info.pageCount || 0,
    publishedDate: info.publishedDate || ''
  };
}

// ── CRUD Operations ──

/** Get all books from the sheet */
function getAllBooks() {
  const sheet = getOrCreateSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Only headers

  return data.slice(1).map(row => ({
    id: String(row[0] || ''),
    title: String(row[1] || ''),
    author: String(row[2] || ''),
    isbn: String(row[3] || ''),
    coverUrl: String(row[4] || ''),
    description: String(row[5] || '').substring(0, 200),
    format: String(row[6] || ''),
    location: String(row[7] || ''),
    ownership: String(row[8] || ''),
    status: String(row[9] || ''),
    rating: Number(row[10]) || 0,
    recommend: String(row[11] || ''),
    notes: String(row[12] || ''),
    dateAdded: row[13] instanceof Date ? row[13].toISOString() : String(row[13] || ''),
    dateModified: row[14] instanceof Date ? row[14].toISOString() : String(row[14] || '')
  }));
}

/** Search books in the sheet by title, author, or ISBN */
function searchMyBooks(query) {
  if (!query || !String(query).trim()) return getAllBooks();
  const q = String(query).toLowerCase().trim();
  const allBooks = getAllBooks();
  return allBooks.filter(b =>
    String(b.title || '').toLowerCase().indexOf(q) !== -1 ||
    String(b.author || '').toLowerCase().indexOf(q) !== -1 ||
    String(b.isbn || '').toLowerCase().indexOf(q) !== -1
  );
}

/** Add a new book (with duplicate check) */
function addBook(bookData) {
  const sheet = getOrCreateSheet_();

  // Duplicate check — ISBN match (strongest) or normalized title match
  const data = sheet.getDataRange().getValues();
  const newIsbn = String(bookData.isbn || '').trim();
  const newTitle = normalizeTitle_(bookData.title);

  for (let i = 1; i < data.length; i++) {
    // ISBN match (if both have one)
    if (newIsbn && newIsbn.length >= 10) {
      const existingIsbn = String(data[i][3] || '').trim();
      if (existingIsbn === newIsbn) {
        return { success: false, error: 'duplicate', message: '"' + bookData.title + '" is already in your Blist (ISBN match)' };
      }
    }
    // Normalized title match
    const existingTitle = normalizeTitle_(String(data[i][1] || ''));
    if (newTitle && existingTitle && newTitle === existingTitle) {
      return { success: false, error: 'duplicate', message: '"' + bookData.title + '" is already in your Blist' };
    }
  }

  const id = Utilities.getUuid();
  const now = new Date().toISOString();

  sheet.appendRow([
    id,
    bookData.title,
    bookData.author,
    bookData.isbn || '',
    bookData.coverUrl || '',
    bookData.description || '',
    bookData.format || '',
    bookData.location || '',
    bookData.ownership || '',
    bookData.status || 'To Read',
    bookData.rating || 0,
    bookData.recommend || '',
    bookData.notes || '',
    now,
    now
  ]);

  return { success: true, id: id };
}

/** Update an existing book by ID */
function updateBook(bookData) {
  const sheet = getOrCreateSheet_();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bookData.id) {
      const row = i + 1;
      const now = new Date().toISOString();
      sheet.getRange(row, 2, 1, 13).setValues([[
        bookData.title,
        bookData.author,
        bookData.isbn || '',
        bookData.coverUrl || '',
        bookData.description || '',
        bookData.format || '',
        bookData.location || '',
        bookData.ownership || '',
        bookData.status || 'To Read',
        bookData.rating || 0,
        bookData.recommend || '',
        bookData.notes || '',
        now
      ]]);
      return { success: true };
    }
  }
  return { success: false, error: 'Book not found' };
}

/** Delete a book by ID */
function deleteBook(bookId) {
  const sheet = getOrCreateSheet_();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bookId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Book not found' };
}

/** Get dashboard stats */
function getDashboardStats() {
  const books = getAllBooks();
  const total = books.length;

  const byStatus = { 'To Read': 0, 'Reading': 0, 'Read': 0, "Don't Read": 0 };
  const byFormat = { 'Physical': 0, 'Digital': 0, 'Both': 0 };
  const byOwnership = { 'Own': 0, 'Buy': 0 };
  const byLocation = {};
  let totalRating = 0;
  let ratedCount = 0;
  let recommendCount = 0;

  books.forEach(b => {
    // Status
    if (byStatus.hasOwnProperty(b.status)) byStatus[b.status]++;

    // Format
    if (byFormat.hasOwnProperty(b.format)) byFormat[b.format]++;

    // Ownership
    if (byOwnership.hasOwnProperty(b.ownership)) byOwnership[b.ownership]++;

    // Location (multi-value)
    if (b.location) {
      b.location.split(',').forEach(loc => {
        loc = loc.trim();
        if (loc) byLocation[loc] = (byLocation[loc] || 0) + 1;
      });
    }

    // Rating
    if (b.rating && b.rating > 0) {
      totalRating += Number(b.rating);
      ratedCount++;
    }

    // Recommend
    if (b.recommend === 'Yes') recommendCount++;
  });

  return {
    total,
    byStatus,
    byFormat,
    byOwnership,
    byLocation,
    avgRating: ratedCount > 0 ? (totalRating / ratedCount).toFixed(1) : 0,
    recommendCount,
    ratedCount,
    randomBook: getRandomBookWithCover_(books)
  };
}

/** Pick a random book that has a cover image */
function getRandomBookWithCover_(books) {
  const withCovers = books.filter(b => b.coverUrl && b.coverUrl.length > 10);
  if (withCovers.length === 0) return null;
  // Hour-based seed: changes every hour, stable within the hour
  const hourSeed = Math.floor(Date.now() / 3600000);
  const idx = hourSeed % withCovers.length;
  const b = withCovers[idx];
  return { title: b.title, author: b.author, coverUrl: b.coverUrl, status: b.status };
}

// ── Widget Data (separate public sheet for Scriptable) ──

/** Write stats JSON to a separate "Blist Widget" spreadsheet.
 *  This spreadsheet is shared publicly — your main book data stays private.
 *  Run on an hourly trigger via setupWidgetTrigger(). */
function updateWidgetData() {
  const stats = getDashboardStats();
  const ws = getOrCreateWidgetSheet_();

  ws.getRange('A2').setValue(JSON.stringify(stats));
  ws.getRange('B2').setValue(new Date().toISOString());
}

/** Get or create the separate widget spreadsheet */
function getOrCreateWidgetSheet_() {
  const userProps = PropertiesService.getUserProperties();
  let ssId = userProps.getProperty('blist_widget_ss_id');

  if (ssId) {
    try {
      const ss = SpreadsheetApp.openById(ssId);
      return ss.getSheets()[0];
    } catch(e) { /* recreate if deleted */ }
  }

  // Create new spreadsheet in Library folder
  const folder = getOrCreateFolder_();
  const ss = SpreadsheetApp.create('Blist Widget Data');
  const file = DriveApp.getFileById(ss.getId());

  // Move to Library folder
  file.moveTo(folder);

  // Share publicly as view-only
  file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

  // Set up the sheet
  const ws = ss.getSheets()[0];
  ws.setName('Widget');
  ws.getRange('A1').setValue('stats_json');
  ws.getRange('B1').setValue('updated_at');

  // Store the ID
  userProps.setProperty('blist_widget_ss_id', ss.getId());

  Logger.log('Widget spreadsheet created: ' + ss.getUrl());
  Logger.log('Sheet ID for Scriptable: ' + ss.getId());

  return ws;
}

/** Run once to set up the hourly trigger. Logs the Sheet ID you need for Scriptable. */
function setupWidgetTrigger() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'updateWidgetData') ScriptApp.deleteTrigger(t);
  });

  // Create hourly trigger
  ScriptApp.newTrigger('updateWidgetData')
    .timeBased()
    .everyHours(1)
    .create();

  // Run once immediately
  updateWidgetData();

  // Log the Sheet ID
  const ssId = PropertiesService.getUserProperties().getProperty('blist_widget_ss_id');
  Logger.log('=== WIDGET SETUP COMPLETE ===');
  Logger.log('Sheet ID for Scriptable: ' + ssId);
  Logger.log('The spreadsheet is already shared publicly (view-only).');
  Logger.log('Paste this Sheet ID on line 16 of Blist-Widget.js');
}

/** Get user email for display */
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}
