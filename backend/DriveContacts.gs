/* =========================================================
   THE PLEASURE DISPATCH
   DriveContacts.gs

   ADD THESE FUNCTIONS TO THE SAME GOOGLE APPS SCRIPT PROJECT
   THAT SERVES DRIVE_API_URL.

   It expects your existing FOLDER_ID constant and getFolder()
   function to remain in that project.
========================================================= */

function getContactCsvFiles_() {
  const folder = getFolder();
  const files = folder.getFiles();
  const result = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const mime = String(file.getMimeType() || '').toLowerCase();

    if (mime === 'text/csv' || /\.csv$/i.test(name)) {
      result.push({
        fileId: file.getId(),
        fileName: name,
        sizeBytes: file.getSize(),
        updatedAt: file.getLastUpdated().toISOString()
      });
    }
  }

  result.sort(function(a, b) {
    return new Date(b.updatedAt) - new Date(a.updatedAt);
  });

  return result;
}

function getContactsFromCsv_(fileId) {
  if (!fileId) throw new Error('Missing contact CSV file ID.');

  const file = DriveApp.getFileById(String(fileId));
  const folder = getFolder();

  if (file.getMimeType() !== MimeType.CSV &&
      file.getMimeType() !== 'text/csv') {
    throw new Error('Selected file is not a CSV.');
  }

  if (file.getParents().hasNext()) {
    let belongsToFolder = false;
    const parents = file.getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() === folder.getId()) {
        belongsToFolder = true;
        break;
      }
    }
    if (!belongsToFolder) throw new Error('Selected CSV is not in the Dispatch Drive folder.');
  }

  const csvText = file.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/, '');
  const rows = Utilities.parseCsv(csvText);

  if (!rows || !rows.length) return [];

  const headers = rows[0].map(function(header) {
    return normalizeContactHeader_(header);
  });

  const emailIndex = findHeaderIndex_(headers, [
    'email', 'emailaddress', 'email_address', 'e-mail', 'e_mail'
  ]);

  const nameIndex = findHeaderIndex_(headers, [
    'name', 'fullname', 'full_name', 'contactname', 'contact_name'
  ]);

  const firstNameIndex = findHeaderIndex_(headers, [
    'firstname', 'first_name', 'givenname', 'given_name'
  ]);

  const lastNameIndex = findHeaderIndex_(headers, [
    'lastname', 'last_name', 'surname', 'familyname', 'family_name'
  ]);

  const statusIndex = findHeaderIndex_(headers, [
    'status', 'subscriptionstatus', 'subscription_status', 'subscribed', 'active'
  ]);

  if (emailIndex < 0) {
    throw new Error('CSV must contain an Email column.');
  }

  const contacts = [];
  const seen = {};

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const email = String(row[emailIndex] || '').trim().toLowerCase();
    if (!email || !isValidEmail_(email)) continue;
    if (seen[email]) continue;
    seen[email] = true;

    let name = '';
    if (nameIndex >= 0) {
      name = String(row[nameIndex] || '').trim();
    }
    if (!name) {
      const first = firstNameIndex >= 0 ? String(row[firstNameIndex] || '').trim() : '';
      const last = lastNameIndex >= 0 ? String(row[lastNameIndex] || '').trim() : '';
      name = (first + ' ' + last).trim();
    }

    const rawStatus = statusIndex >= 0
      ? String(row[statusIndex] || '').trim().toLowerCase()
      : 'active';

    const active = !isUnsubscribedStatus_(rawStatus);

    contacts.push({
      id: email,
      name: name || email,
      email: email,
      status: active ? 'active' : 'unsubscribed',
      active: active
    });
  }

  return contacts;
}

function normalizeContactHeader_(header) {
  return String(header || '')
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, '');
}

function findHeaderIndex_(headers, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const target = normalizeContactHeader_(candidates[i]);
    const index = headers.indexOf(target);
    if (index >= 0) return index;
  }
  return -1;
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function isUnsubscribedStatus_(status) {
  const value = String(status || '').toLowerCase().trim();
  return [
    'unsubscribed',
    'unsubscribe',
    'inactive',
    'opted_out',
    'opted-out',
    'opted out',
    'false',
    'no'
  ].indexOf(value) >= 0;
}

function unsubscribeContact_(email) {
  const targetEmail = String(email || '').trim().toLowerCase();
  if (!isValidEmail_(targetEmail)) throw new Error('A valid email address is required.');

  const folder = getFolder();
  const files = folder.getFiles();
  let changed = false;
  let found = false;

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== MimeType.CSV && !/\.csv$/i.test(file.getName())) continue;

    const csvText = file.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/, '');
    const rows = Utilities.parseCsv(csvText);
    if (!rows || !rows.length) continue;

    const headers = rows[0].map(normalizeContactHeader_);
    const emailIndex = findHeaderIndex_(headers, ['email','emailaddress','email_address','e-mail','e_mail']);
    if (emailIndex < 0) continue;

    const statusIndex = findHeaderIndex_(headers, ['status','subscriptionstatus','subscription_status','subscribed','active']);
    if (statusIndex < 0) continue;

    for (let i = 1; i < rows.length; i++) {
      const rowEmail = String((rows[i] || [])[emailIndex] || '').trim().toLowerCase();
      if (rowEmail === targetEmail) {
        found = true;
        rows[i][statusIndex] = 'unsubscribed';
        changed = true;
      }
    }

    if (changed) {
      const output = rows.map(function(row) {
        return row.map(csvEscape_).join(',');
      }).join('\n');
      file.setContent(output);
      break;
    }
  }

  return { found: found, changed: changed };
}

function csvEscape_(value) {
  const text = String(value == null ? '' : value);
  if (/[",\n\r]/.test(text)) return '"' + text.replace(/"/g, '""') + '"';
  return text;
}

/* =========================================================
   ADD THESE BRANCHES INSIDE YOUR EXISTING doGet(e)

   Put them before your normal status/list response.
========================================================= */

function contactGetAction_(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'contactFiles') {
    return jsonResponse({
      status: 'success',
      files: getContactCsvFiles_()
    });
  }

  if (action === 'contacts') {
    return jsonResponse({
      status: 'success',
      contacts: getContactsFromCsv_(e.parameter.fileId)
    });
  }

  if (action === 'unsubscribe') {
    const result = unsubscribeContact_(e.parameter.email);
    return jsonResponse({
      status: 'unsubscribed',
      found: result.found,
      changed: result.changed
    });
  }

  return null;
}
