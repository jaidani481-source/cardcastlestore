/**
 * CardCastle Store review API.
 * Bind this script to the reviews spreadsheet, then deploy it as a web app.
 * Reviews are added with Status=Pending. Set Status to Approved in the sheet
 * to publish a review on the store website.
 */
const SHEET_NAME = 'Reviews';
const MEDIA_DRAFTS_SHEET_NAME = 'Review Media Drafts';
const PHOTO_FOLDER_NAME = 'CardCastle Review Photos';
const MEDIA_PREVIEW_COUNT = 10;
const HEADERS = [
  'Created At', 'Name', 'Phone', 'Rating', 'Comment', 'Photo URLs', 'Status', 'Review ID', 'Product', 'Video URLs', 'Group', 'Customer Media', 'Watch Videos',
  ...Array.from({ length: MEDIA_PREVIEW_COUNT }, (_, i) => 'Media ' + (i + 1) + ' Preview'),
  ...Array.from({ length: MEDIA_PREVIEW_COUNT }, (_, i) => 'Media ' + (i + 1) + ' Open')
];

function doGet(e) {
  const reviews = getApprovedReviews_();
  const callback = e && e.parameter && e.parameter.callback;
  const json = JSON.stringify(reviews);

  // Apps Script ContentService cannot set arbitrary CORS response headers.
  // JSONP lets the public website read approved reviews without a CORS fetch.
  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return output_('/* Invalid callback */', ContentService.MimeType.JAVASCRIPT);
    }
    return output_(callback + '(' + json + ');', ContentService.MimeType.JAVASCRIPT);
  }
  return output_(json, ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents;
    if (!body) throw new Error('Missing review data.');
    const review = JSON.parse(body);
    if (review.action === 'mediaDraft') {
      return output_(JSON.stringify(createMediaDraft_(review)), ContentService.MimeType.JSON);
    }
    if (review.action === 'thinkFastBatchDrafts') {
      return output_(JSON.stringify(createThinkFastBatchDrafts_(review)), ContentService.MimeType.JSON);
    }
    const name = cleanText_(review.name, 120);
    const phone = cleanText_(review.phone, 40);
    const comment = cleanText_(review.comment, 3000);
    const product = cleanText_(review.product || review.deck || review.game, 160);
    const rating = Number(review.rating);
    if (!name || !phone || !comment || !Number.isInteger(rating) || rating < 1 || rating > 5) {
      throw new Error('Please provide a name, phone, comment, and a rating from 1 to 5.');
    }

    const photos = asArray_(review.photo);
    const videos = asArray_(review.video);
    const media = photos.map(value => saveMedia_(value, 'image')).concat(videos.map(value => saveMedia_(value, 'video'))).filter(Boolean);
    const photoUrls = media.filter(item => item.type === 'image').map(item => item.url);
    const videoUrls = media.filter(item => item.type === 'video').map(item => item.url);
    const sheet = getSheet_();
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const rowValues = {
        'Created At': new Date(), Name: name, Phone: phone, Rating: rating, Comment: comment,
        'Photo URLs': JSON.stringify(photoUrls), 'Video URLs': JSON.stringify(videoUrls), Status: 'Pending',
        'Review ID': cleanText_(review.id, 100) || Utilities.getUuid(), Product: product
      };
      appendByHeaders_(sheet, rowValues);
      writeMediaCells_(sheet, sheet.getLastRow(), media, HEADERS);
    } finally {
      lock.releaseLock();
    }

    // With a no-cors browser POST this response is intentionally opaque.
    return output_(JSON.stringify({ ok: true, status: 'Pending', photoCount: photoUrls.length, videoCount: videoUrls.length }), ContentService.MimeType.JSON);
  } catch (error) {
    console.error(error);
    return output_(JSON.stringify({ ok: false, error: String(error.message || error) }), ContentService.MimeType.JSON);
  }
}

function getApprovedReviews_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(value => String(value).trim());
  const col = {};
  headers.forEach((header, index) => { col[header] = index; });

  return values.slice(1)
    .filter(row => String(row[col.Status] || '').trim().toLowerCase() === 'approved')
    .map(row => {
      let photos = [];
      let videos = [];
      try { photos = JSON.parse(row[col['Photo URLs']] || '[]'); } catch (_) {}
      try { videos = JSON.parse(row[col['Video URLs']] || '[]'); } catch (_) {}
      return {
        id: String(row[col['Review ID']] || ''),
        name: String(row[col.Name] || ''),
        phone: String(row[col.Phone] || ''),
        rating: Number(row[col.Rating]) || 0,
        comment: String(row[col.Comment] || ''),
        product: String(row[col.Product] || ''),
        photo: photos,
        video: videos,
        date: row[col['Created At']] instanceof Date
          ? Utilities.formatDate(row[col['Created At']], Session.getScriptTimeZone(), 'dd MMM yyyy')
          : String(row[col['Created At']] || ''),
        status: 'approved'
      };
    });
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Bind this Apps Script project to the reviews spreadsheet.');
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  ensureHeaders_(sheet, HEADERS);
  const col = headerMap_(sheet);
  ['Photo URLs', 'Video URLs'].forEach(name => { if (col[name]) sheet.hideColumns(col[name]); });
  return sheet;
}

function createMediaDraft_(review) {
  const product = cleanText_(review.product || review.deck || review.game, 160);
  const group = cleanText_(review.group, 40);
  const mediaValues = asArray_(review.media);
  if (product !== 'Think Fast' || group !== '9' || !mediaValues.length || mediaValues.length > 5) {
    throw new Error('A media draft needs a product, group label, and 1–' + MEDIA_PREVIEW_COUNT + ' files.');
  }
  if (mediaValues.some(value => typeof value !== 'string' || value.length > 7000000) || mediaValues.reduce((sum, value) => sum + String(value || '').length, 0) > 18000000) {
    throw new Error('The media draft is too large. Keep each file under 5 MB and the total under 13 MB.');
  }
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Bind this Apps Script project to the reviews spreadsheet.');
  let sheet = spreadsheet.getSheetByName(MEDIA_DRAFTS_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(MEDIA_DRAFTS_SHEET_NAME);
  const headers = [
    'Created At', 'Product', 'Group', 'Owner Comment', 'Status', 'Media Files (internal)',
    ...Array.from({ length: MEDIA_PREVIEW_COUNT }, (_, i) => 'Media ' + (i + 1) + ' Preview'),
    ...Array.from({ length: MEDIA_PREVIEW_COUNT }, (_, i) => 'Media ' + (i + 1) + ' Open')
  ];
  ensureHeaders_(sheet, headers);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const col = headerMap_(sheet);
    const existingRows = sheet.getDataRange().getValues();
    const existingGroup = existingRows.slice(1).some(row =>
      String(row[col.Product - 1] || '') === product && String(row[col.Group - 1] || '') === group
    );
    if (existingGroup) throw new Error('A draft for this product and group already exists.');

    // Draft media stays private in Drive. The sheet embeds previews from the
    // file blobs, so the owner can review them without publishing public URLs.
    const media = mediaValues.map(value => saveMedia_(value, 'auto', false)).filter(Boolean);
    if (!media.length) throw new Error('No supported photo or video files were received.');
    const rowValues = {
      'Created At': new Date(), Product: product, Group: group, 'Owner Comment': '', Status: 'Draft',
      'Media Files (internal)': JSON.stringify(media)
    };
    const row = appendByHeaders_(sheet, rowValues);
    writeMediaCells_(sheet, row, media, headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
    sheet.hideColumns(headerMap_(sheet)['Media Files (internal)']);
    return { ok: true, status: 'Draft', product: product, group: group, mediaCount: media.length, sheet: MEDIA_DRAFTS_SHEET_NAME };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Imports the prepared Think Fast media as one Draft row per customer review.
 * Keep customer name, phone, rating, and comment blank so the owner can fill
 * them in after checking the original review details.
 * Payload shape: {action:'thinkFastBatchDrafts', groups:[{group:'1',media:[data URLs...]}]}
 */
function createThinkFastBatchDrafts_(request) {
  const groups = asArray_(request.groups);
  if (groups.length !== 10) throw new Error('Expected exactly 10 Think Fast review groups.');
  const seen = {};
  let totalBytes = 0;
  groups.forEach(group => {
    const id = String(group.group || '');
    const media = asArray_(group.media);
    if (!/^(?:[1-9]|10)$/.test(id) || seen[id]) throw new Error('Review group labels must be unique from 1 to 10.');
    if (!media.length || media.length > 5) throw new Error('Each review group must contain 1 to 5 media files.');
    seen[id] = true;
    media.forEach(value => {
      if (typeof value !== 'string' || !/^data:(?:image\/[\w.+-]+|video\/(?:mp4|webm|quicktime));base64,/i.test(value)) {
        throw new Error('A group contains an unsupported image or video.');
      }
      totalBytes += value.length;
    });
  });
  if (totalBytes > 18000000) throw new Error('The complete media upload is too large (18 MB maximum).');

  const sheet = getSheet_();
  ensureReviewMediaLayout_(sheet);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const col = headerMap_(sheet);
    const result = [];
    groups.sort((a, b) => Number(a.group) - Number(b.group)).forEach(group => {
      const groupId = String(group.group);
      const media = asArray_(group.media).map(value => saveMedia_(value, 'auto', false)).filter(Boolean);
      if (!media.length) throw new Error('No supported media received for Group ' + groupId + '.');
      const reviewId = 'think-fast-media-group-' + groupId;
      const values = {
        'Created At': new Date(), Name: '', Phone: '', Rating: '', Comment: '',
        'Photo URLs': JSON.stringify(media.filter(item => item.type === 'image').map(item => item.url)),
        'Video URLs': JSON.stringify(media.filter(item => item.type === 'video').map(item => item.url)),
        Status: 'Draft', 'Review ID': reviewId, Product: 'Think Fast', Group: groupId
      };
      let row = findReviewRow_(sheet, col, reviewId, 'Think Fast', groupId);
      if (row) {
        sheet.getRange(row, 1, 1, sheet.getLastColumn()).clearContent();
        writeByHeaders_(sheet, row, values);
      } else {
        row = appendByHeaders_(sheet, values);
      }
      writeCompactMedia_(sheet, row, media, headerMap_(sheet));
      result.push({ group: groupId, row: row, photos: media.filter(item => item.type === 'image').length, videos: media.filter(item => item.type === 'video').length });
    });
    return { ok: true, status: 'Draft', imported: result.length, groups: result };
  } finally {
    lock.releaseLock();
  }
}

function findReviewRow_(sheet, col, reviewId, product, group) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][col['Review ID'] - 1] || '') === reviewId ||
        (String(rows[i][col.Product - 1] || '') === product && String(rows[i][col.Group - 1] || '') === group)) return i + 1;
  }
  return 0;
}

function writeByHeaders_(sheet, row, valuesByHeader) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value).trim());
  headers.forEach((header, index) => {
    if (Object.prototype.hasOwnProperty.call(valuesByHeader, header)) sheet.getRange(row, index + 1).setValue(valuesByHeader[header]);
  });
}

function ensureReviewMediaLayout_(sheet) {
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value).trim());
  let col = {};
  headers.forEach((header, index) => { col[header] = index + 1; });
  if (col.Product && col.Group && col['Customer Media'] && col['Watch Videos'] && col.Group > col.Product + 1) {
    sheet.moveColumns(sheet.getRange(1, col.Group, 1, 3), col.Product + 1);
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value).trim());
    col = {};
    headers.forEach((header, index) => { col[header] = index + 1; });
  }
  // Put five individual media cells directly beside Group. A separate cell
  // per preview keeps images inside their review row instead of spilling out
  // of one wide Customer Media cell.
  const firstFive = Array.from({ length: 5 }, (_, i) => col['Media ' + (i + 1) + ' Preview']);
  if (firstFive.every(Boolean)) {
    const expectedStart = col.Group + 1;
    if (firstFive[0] !== expectedStart || firstFive.some((value, i) => value !== firstFive[0] + i)) {
      sheet.moveColumns(sheet.getRange(1, firstFive[0], 1, 5), expectedStart);
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value).trim());
      col = {};
      headers.forEach((header, index) => { col[header] = index + 1; });
    }
  }
  headers.forEach((header, index) => {
    if (/^Media (?:[6-9]|10) Preview$/.test(header) || /^Media \d+ Open$/.test(header) || header === 'Customer Media') sheet.hideColumns(index + 1);
  });
  for (let i = 1; i <= 5; i++) {
    sheet.showColumns(col['Media ' + i + ' Preview']);
    sheet.setColumnWidth(col['Media ' + i + ' Preview'], 112);
  }
  if (col['Watch Videos']) sheet.setColumnWidth(col['Watch Videos'], 150);
  sheet.setFrozenRows(1);
  return col;
}

/** Rebuilds all Think Fast draft previews into five aligned media cells per row. */
function repairThinkFastDraftRows() {
  const sheet = getSheet_();
  const col = ensureReviewMediaLayout_(sheet);
  const manifestsByGroup = {};
  const drafts = sheet.getParent().getSheetByName(MEDIA_DRAFTS_SHEET_NAME);
  if (drafts && drafts.getLastRow() > 1) {
    const draftValues = drafts.getDataRange().getValues();
    const draftHeaders = draftValues[0].map(value => String(value).trim());
    const draftCol = {};
    draftHeaders.forEach((header, index) => { draftCol[header] = index; });
    draftValues.slice(1).forEach(draft => {
      if (String(draft[draftCol.Product] || '').trim() !== 'Think Fast') return;
      const group = String(draft[draftCol.Group] || '').trim();
      const manifest = parseJsonArray_(draft[draftCol['Media Files (internal)']]);
      if (group && manifest.length && (!manifestsByGroup[group] || manifest.length > manifestsByGroup[group].length)) {
        manifestsByGroup[group] = manifest;
      }
    });
  }
  const rows = sheet.getDataRange().getValues();
  let repaired = 0;
  for (let r = 1; r < rows.length; r++) {
    if (String(rows[r][col.Product - 1] || '').trim() !== 'Think Fast') continue;
    const group = String(rows[r][col.Group - 1] || '').trim();
    if (!/^(?:[1-9]|10)$/.test(group)) continue;
    const photos = parseJsonArray_(rows[r][col['Photo URLs'] - 1]);
    const videos = parseJsonArray_(rows[r][col['Video URLs'] - 1]);
    let media = photos.map(url => ({ type: 'image', url: url, fileId: driveFileId_(url) }))
      .concat(videos.map(url => ({ type: 'video', url: url, fileId: driveFileId_(url) })));
    // Use the legacy manifest only when the review row has no media URLs.
    // A longer Group 9 manifest may contain assets from other reviews.
    if (!media.length && manifestsByGroup[group]) media = manifestsByGroup[group];
    if (!media.length) continue;
    writeCompactMedia_(sheet, r + 1, media, col);
    repaired++;
  }
  return { ok: true, repaired: repaired };
}

/** Makes Drive media publicly viewable only for rows the owner approved. */
function publishApprovedReviewMedia() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { approvedRows: 0, filesShared: 0 };

  const headers = values[0].map(value => String(value).trim());
  const col = {};
  headers.forEach((header, index) => { col[header] = index; });

  let approvedRows = 0;
  let filesShared = 0;
  const sharedIds = {};

  values.slice(1).forEach(row => {
    if (String(row[col.Status] || '').trim().toLowerCase() !== 'approved') return;
    approvedRows++;

    ['Photo URLs', 'Video URLs'].forEach(header => {
      parseJsonArray_(row[col[header]]).forEach(url => {
        const fileId = driveFileId_(url);
        if (!fileId || sharedIds[fileId]) return;

        DriveApp.getFileById(fileId).setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW
        );
        sharedIds[fileId] = true;
        filesShared++;
      });
    });
  });

  return { approvedRows: approvedRows, filesShared: filesShared };
}

/**
 * Run this once Drive has finished processing uploaded videos. It fills the
 * pending video preview cells in Review Media Drafts with embedded thumbnails.
 * The actual videos remain in Drive and are opened with the adjacent Play link.
 */
function refreshMediaDraftPreviews() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet && spreadsheet.getSheetByName(MEDIA_DRAFTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { refreshed: 0, pending: 0 };

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(value => String(value).trim());
  const col = {};
  headers.forEach((header, index) => { col[header] = index + 1; });
  const fileCol = col['Media Files (internal)'];
  if (!fileCol) throw new Error('Media Files (internal) column is missing.');
  let refreshed = 0;
  let pending = 0;

  for (let row = 2; row <= values.length; row++) {
    let media = [];
    try { media = JSON.parse(values[row - 1][fileCol - 1] || '[]'); } catch (_) {}
    media.forEach((item, index) => {
      if (item.type !== 'video' || !item.fileId) return;
      const previewCol = col['Media ' + (index + 1) + ' Preview'];
      if (!previewCol) return;
      const cellValue = String(sheet.getRange(row, previewCol).getDisplayValue() || '');
      if (cellValue && cellValue.indexOf('Video preview pending') === -1) return;
      const thumbnail = DriveApp.getFileById(item.fileId).getThumbnail();
      if (!thumbnail) { pending++; return; }

      sheet.getRange(row, previewCol).clearContent();
      const image = sheet.insertImage(thumbnail, previewCol, row);
      image.setWidth(100).setHeight(100);
      image.setAltTextTitle('Video preview');
      image.setAltTextDescription('Use the Play link beside this preview to watch the video.');
      refreshed++;
    });
  }
  return { refreshed: refreshed, pending: pending };
}

/**
 * Copies existing media draft rows into the main Reviews tab, preserving the
 * photos/videos as movable sheet images. Name and comment stay blank for the
 * owner to fill in. Re-running this function will not create duplicate rows.
 */
function moveMediaDraftsToReviews() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const source = spreadsheet && spreadsheet.getSheetByName(MEDIA_DRAFTS_SHEET_NAME);
  if (!source || source.getLastRow() < 2) return { moved: 0 };
  const target = getSheet_();
  const sourceValues = source.getDataRange().getValues();
  const sourceHeaders = sourceValues[0].map(value => String(value).trim());
  const sourceCol = {};
  sourceHeaders.forEach((header, index) => { sourceCol[header] = index; });
  if (sourceCol['Media Files (internal)'] == null) throw new Error('Media Files (internal) column is missing.');

  const targetHeaders = target.getRange(1, 1, 1, target.getLastColumn()).getValues()[0].map(value => String(value).trim());
  const targetCol = {};
  targetHeaders.forEach((header, index) => { targetCol[header] = index; });
  let moved = 0;
  for (let i = 1; i < sourceValues.length; i++) {
    const row = sourceValues[i];
    const status = String(row[sourceCol.Status] || '').trim().toLowerCase();
    if (status !== 'draft') continue;
    let media = [];
    try { media = JSON.parse(row[sourceCol['Media Files (internal)']] || '[]'); } catch (_) {}
    if (!media.length) continue;

    const product = String(row[sourceCol.Product] || 'Think Fast');
    const group = String(row[sourceCol.Group] || '');
    const reviewId = 'media-draft-' + product.toLowerCase().replace(/[^a-z0-9]+/g, '-') + '-' + group;
    const reviewRows = target.getDataRange().getValues();
    const alreadyMoved = reviewRows.slice(1).some(reviewRow =>
      String(reviewRow[targetCol['Review ID']] || '') === reviewId
    );
    if (alreadyMoved) {
      source.getRange(i + 1, sourceCol.Status + 1).setValue('Moved to Reviews');
      continue;
    }

    const photos = media.filter(item => item.type === 'image').map(item => item.url);
    const videos = media.filter(item => item.type === 'video').map(item => item.url);
    const reviewValues = {
      'Created At': row[sourceCol['Created At']] || new Date(),
      Name: '', Phone: '', Rating: '', Comment: '',
      'Photo URLs': JSON.stringify(photos), 'Video URLs': JSON.stringify(videos),
      Status: 'Draft', 'Review ID': reviewId, Product: product, Group: group
    };
    const reviewRow = appendByHeaders_(target, reviewValues);
    writeMediaCells_(target, reviewRow, media, HEADERS);
    source.getRange(i + 1, sourceCol.Status + 1).setValue('Moved to Reviews');
    moved++;
  }
  return { moved: moved };
}

/** One-time cleanup: put the Think Fast media beside its review fields. */
function cleanThinkFastGroup9Review() {
  const sheet = getSheet_();
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value).trim());
  let col = {};
  headers.forEach((header, index) => { col[header] = index + 1; });

  // Move Group and both compact media columns next to Product, even if the
  // older sheet version appended Group after many legacy helper columns.
  if (col.Product && col.Group && col['Customer Media'] && col['Watch Videos'] && col.Group > col.Product + 1) {
    sheet.moveColumns(sheet.getRange(1, col.Group, 1, 3), col.Product + 1);
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value).trim());
    col = {};
    headers.forEach((header, index) => { col[header] = index + 1; });
  }

  // Remove rows that are only accidental copies of the header labels.
  for (let row = sheet.getLastRow(); row >= 2; row--) {
    const name = String(sheet.getRange(row, col.Name).getDisplayValue()).trim();
    const phone = String(sheet.getRange(row, col.Phone).getDisplayValue()).trim();
    const rating = String(sheet.getRange(row, col.Rating).getDisplayValue()).trim();
    const comment = String(sheet.getRange(row, col.Comment).getDisplayValue()).trim();
    const status = String(sheet.getRange(row, col.Status).getDisplayValue()).trim();
    const id = String(sheet.getRange(row, col['Review ID']).getDisplayValue()).trim();
    const product = String(sheet.getRange(row, col.Product).getDisplayValue()).trim();
    if (name === 'Name' && phone === 'Phone' && rating === 'Rating' && comment === 'Comment' && status === 'Status' && id === 'Review ID' && product === 'Product') {
      sheet.deleteRow(row);
    }
  }

  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value).trim());
  col = {};
  headers.forEach((header, index) => { col[header] = index + 1; });
  let targetRow = 0;
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    if (String(sheet.getRange(row, col.Product).getDisplayValue()).trim() === 'Think Fast' &&
        String(sheet.getRange(row, col.Group).getDisplayValue()).trim() === '9') {
      targetRow = row;
      break;
    }
  }
  if (!targetRow) throw new Error('Could not find the Think Fast, Group 9 review row.');

  let photos = parseJsonArray_(sheet.getRange(targetRow, col['Photo URLs']).getValue());
  let videos = parseJsonArray_(sheet.getRange(targetRow, col['Video URLs']).getValue());
  let media = photos.map(url => ({ type: 'image', url: url, fileId: driveFileId_(url) }))
    .concat(videos.map(url => ({ type: 'video', url: url, fileId: driveFileId_(url) })));

  // The original draft row retains the full media manifest. Use it when the
  // review row's URL cells were blank or incomplete in an earlier migration.
  const draftSheet = sheet.getParent().getSheetByName(MEDIA_DRAFTS_SHEET_NAME);
  if (draftSheet && draftSheet.getLastRow() > 1) {
    const draftValues = draftSheet.getDataRange().getValues();
    const draftHeaders = draftValues[0].map(value => String(value).trim());
    const draftCol = {};
    draftHeaders.forEach((header, index) => { draftCol[header] = index; });
    for (let i = 1; i < draftValues.length; i++) {
      const draft = draftValues[i];
      if (String(draft[draftCol.Product] || '') === 'Think Fast' && String(draft[draftCol.Group] || '') === '9') {
        const manifest = parseJsonArray_(draft[draftCol['Media Files (internal)']]);
        if (manifest.length) media = manifest;
        break;
      }
    }
  }
  if (!media.length) throw new Error('No media files were found for Think Fast, Group 9.');
  photos = media.filter(item => item.type === 'image').map(item => item.url);
  videos = media.filter(item => item.type === 'video').map(item => item.url);
  writeCompactMedia_(sheet, targetRow, media, col);

  // Hide the old one-column-per-asset fields after the compact strip is built.
  headers.forEach((header, index) => {
    if (/^Media \d+ (Preview|Open)$/.test(header)) sheet.hideColumns(index + 1);
  });
  sheet.setFrozenRows(1);
  return { ok: true, reviewRow: targetRow, photoCount: photos.length, videoCount: videos.length };
}

function writeCompactMedia_(sheet, row, media, col) {
  const watchCol = col['Watch Videos'];
  const previewCols = Array.from({ length: 5 }, (_, i) => col['Media ' + (i + 1) + ' Preview']);
  if (previewCols.some(value => !value) || !watchCol) throw new Error('Review media preview or Watch Videos columns are missing.');
  sheet.getImages().forEach(image => {
    const anchor = image.getAnchorCell();
    if (anchor.getRow() === row) image.remove();
  });

  const videos = [];
  previewCols.forEach(column => sheet.getRange(row, column).clearContent());
  media.slice(0, previewCols.length).forEach((item, mediaIndex) => {
    const previewCol = previewCols[mediaIndex];
    if (!item.fileId) {
      sheet.getRange(row, previewCol).setValue(item.type === 'video' ? 'Video preview unavailable' : 'Preview unavailable');
      return;
    }
    const file = DriveApp.getFileById(item.fileId);
    const blob = item.type === 'video' ? file.getThumbnail() : file.getBlob();
    if (item.type === 'video') videos.push({ url: driveFileUrl_(item.url), index: videos.length + 1 });
    if (!blob) {
      sheet.getRange(row, previewCol).setValue(item.type === 'video' ? 'Video preview pending' : 'Preview unavailable');
      return;
    }
    const image = sheet.insertImage(blob, previewCol, row, 6, 4);
    image.setWidth(96).setHeight(96);
    image.setAltTextTitle(item.type === 'video' ? 'Video preview — use Watch Videos link' : 'Customer review photo');
    image.setAltTextDescription(item.type === 'video' ? 'This is a preview frame. Play the original video from the Watch Videos cell.' : 'Customer photo preview.');
  });

  sheet.setRowHeight(row, 112);

  const watchCell = sheet.getRange(row, watchCol);
  watchCell.clearContent();
  if (videos.length) {
    const labels = videos.map(video => '▶ Watch video ' + video.index);
    const text = labels.join('\n');
    const builder = SpreadsheetApp.newRichTextValue().setText(text);
    let start = 0;
    videos.forEach((video, index) => {
      builder.setLinkUrl(start, start + labels[index].length, video.url);
      start += labels[index].length + 1;
    });
    watchCell.setRichTextValue(builder.build());
  }
  watchCell.setVerticalAlignment('middle').setWrap(true);
}

function parseJsonArray_(value) {
  try {
    const parsed = JSON.parse(value || '[]');
    return Array.isArray(parsed) ? parsed : [];
  } catch (_) {
    return [];
  }
}

function saveMedia_(value, expectedType, sharePublic) {
  if (typeof value !== 'string' || !value) return '';
  if (/^https?:\/\//i.test(value)) {
    const type = /^https?:\/\/drive\.google\.com\/file\/d\//i.test(value) && /\.(mp4|webm|mov)(?:[?#]|$)/i.test(value) ? 'video' : expectedType === 'video' ? 'video' : 'image';
    const id = driveFileId_(value);
    const previewUrl = id ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400' : value;
    return { type: type, url: type === 'video' ? value : previewUrl, previewUrl: previewUrl, openUrl: driveFileUrl_(value), fileId: id };
  }
  const match = value.match(/^data:(image\/[a-zA-Z0-9.+-]+|video\/(?:mp4|webm|quicktime));base64,([\s\S]+)$/i);
  if (!match) return '';
  const type = match[1].toLowerCase().indexOf('video/') === 0 ? 'video' : 'image';
  const bytes = Utilities.base64Decode(match[2]);
  const extension = (match[1].split('/')[1] || 'jpg').replace('jpeg', 'jpg').replace('quicktime', 'mov');
  const blob = Utilities.newBlob(bytes, match[1], 'review-' + Utilities.getUuid() + '.' + extension);
  const folder = getPhotoFolder_();
  const file = folder.createFile(blob);
  if (sharePublic !== false) file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';
  return {
    type: type,
    url: type === 'video' ? file.getUrl() : thumbnailUrl,
    previewUrl: thumbnailUrl,
    openUrl: file.getUrl(),
    fileId: file.getId()
  };
}

function writeMediaCells_(sheet, row, media, headers) {
  const col = headerMap_(sheet);
  if (sheet.getName() === SHEET_NAME && col['Customer Media'] && col['Watch Videos']) {
    writeCompactMedia_(sheet, row, media, col);
    return;
  }
  media.slice(0, MEDIA_PREVIEW_COUNT).forEach((item, i) => {
    const previewCol = col['Media ' + (i + 1) + ' Preview'];
    const openCol = col['Media ' + (i + 1) + ' Open'];
    if (!previewCol || !openCol) return;
    let preview = null;
    if (item.fileId) {
      const file = DriveApp.getFileById(item.fileId);
      preview = item.type === 'video' ? file.getThumbnail() : file.getBlob();
    }
    if (preview) {
      const image = sheet.insertImage(preview, previewCol, row);
      image.setWidth(100).setHeight(100);
      image.setAltTextTitle(item.type === 'video' ? 'Video preview' : 'Customer photo');
      image.setAltTextDescription(item.type === 'video' ? 'Open the Play link to watch this video.' : 'Customer photo preview.');
    } else {
      sheet.getRange(row, previewCol).setValue(item.type === 'video' ? 'Video preview pending' : 'Preview unavailable');
    }
    sheet.getRange(row, openCol).setFormula('=HYPERLINK("' + item.openUrl + '","' + (item.type === 'video' ? '▶ Play' : 'Open') + '")');
    sheet.setColumnWidth(previewCol, 112);
    sheet.setColumnWidth(openCol, 78);
  });
  sheet.setRowHeight(row, 112);
}

function refreshReviewMediaPreviews() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  const col = {};
  values[0].forEach((value, i) => { col[String(value).trim()] = i; });
  for (let r = 1; r < values.length; r++) {
    const media = [];
    [['Photo URLs', 'image'], ['Video URLs', 'video']].forEach(([header, type]) => {
      let urls = [];
      try { urls = JSON.parse(values[r][col[header]] || '[]'); } catch (_) {}
      urls.forEach(url => media.push({ type: type, url: url, previewUrl: driveThumbnailUrl_(url), openUrl: driveFileUrl_(url), fileId: driveFileId_(url) }));
    });
    if (media.length) writeMediaCells_(sheet, r + 1, media, HEADERS);
  }
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return;
  }
  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value).trim());
  headers.forEach(header => {
    if (existing.indexOf(header) === -1) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(header).setFontWeight('bold');
      existing.push(header);
    }
  });
}

function headerMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value).trim());
  const map = {};
  headers.forEach((header, i) => { map[header] = i + 1; });
  return map;
}

function appendByHeaders_(sheet, valuesByHeader) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(value => String(value).trim());
  const rowValues = headers.map(header => Object.prototype.hasOwnProperty.call(valuesByHeader, header) ? valuesByHeader[header] : '');
  sheet.appendRow(rowValues);
  return sheet.getLastRow();
}

function asArray_(value) {
  return Array.isArray(value) ? value : (value ? [value] : []);
}

function driveFileId_(value) {
  const match = String(value || '').match(/[?&]id=([\w-]+)|\/d\/([\w-]+)/);
  return match ? (match[1] || match[2]) : '';
}

function driveThumbnailUrl_(value) {
  const id = driveFileId_(value);
  return id ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400' : String(value || '');
}

function driveFileUrl_(value) {
  const id = driveFileId_(value);
  return id ? 'https://drive.google.com/file/d/' + id + '/view' : String(value || '');
}

function getPhotoFolder_() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function cleanText_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength);
}

function output_(text, mimeType) {
  return ContentService.createTextOutput(text).setMimeType(mimeType);
}
