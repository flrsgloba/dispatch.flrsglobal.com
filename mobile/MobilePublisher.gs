/**
 * THE PLEASURE DISPATCH
 * MobilePublisher.gs
 *
 * Thin Google Apps Script bridge for the mobile Dispatch Publisher.
 *
 * Mobile -> this web app -> Main Dispatch Publisher -> GitHub/site.
 *
 * IMPORTANT:
 * 1. Deploy this file as a separate Apps Script Web App.
 * 2. Set MAIN_PUBLISHER_URL to the deployed URL of the main Dispatch publisher.
 * 3. Set MOBILE_PUBLISH_KEY in Script Properties. Do not put the key in this file.
 * 4. The main publisher remains the system of record for publishing.
 */

const MAIN_PUBLISHER_URL =
  'https://script.google.com/macros/s/AKfycbxY93Vr1Zuij1sIKM7X0sgmyT5ipFnufnYGUrw6DqSAQL8QQYM6juVkRszGf-QdRKMEWQ/exec';

const MOBILE_PUBLISH_KEY_PROPERTY = 'MOBILE_PUBLISH_KEY';
const MAX_BODY_BYTES = 9000000;

function doGet() {
  return json_({
    success: true,
    service: 'Pleasure Dispatch Mobile Publisher',
    status: 'online',
    version: '1.0.0'
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json_({ success: false, error: 'Empty request body.' });
    }

    if (e.postData.contents.length > MAX_BODY_BYTES) {
      return json_({ success: false, error: 'Request is too large.' });
    }

    const request = parseJson_(e.postData.contents);
    if (!request) {
      return json_({ success: false, error: 'Invalid JSON payload.' });
    }

    const action = String(request.action || '').toLowerCase();

    if (action === 'health') {
      return json_({
        success: true,
        service: 'Pleasure Dispatch Mobile Publisher',
        status: 'online',
        mainPublisherConfigured: Boolean(MAIN_PUBLISHER_URL)
      });
    }

    if (action !== 'publish') {
      return json_({
        success: false,
        error: 'Unsupported action. Use action: publish.'
      });
    }

    authenticate_(request.secret);

    const payload = sanitizePublishPayload_(request);

    // Never forward the mobile authentication secret to the main publisher
    // unless the main publisher explicitly requires the same secret. The
    // bridge authenticates the mobile client independently.
    delete payload.secret;

    const response = forwardToMainPublisher_(payload);

    return json_({
      success: true,
      service: 'Pleasure Dispatch Mobile Publisher',
      upstream: normalizeUpstreamResponse_(response)
    });

  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return json_({
      success: false,
      error: err && err.message ? err.message : String(err)
    });
  }
}

function authenticate_(provided) {
  const expected = PropertiesService
    .getScriptProperties()
    .getProperty(MOBILE_PUBLISH_KEY_PROPERTY);

  if (!expected) {
    throw new Error(
      'Mobile publisher is not configured. Set MOBILE_PUBLISH_KEY in Script Properties.'
    );
  }

  if (!provided || String(provided) !== String(expected)) {
    throw new Error('Unauthorized publishing request.');
  }
}

function sanitizePublishPayload_(request) {
  const allowed = [
    'action',
    'edition',
    'editionLabel',
    'date',
    'title',
    'subtitle',
    'subject',
    'reflection',
    'html',
    'searchText',
    'publicUrl',
    'dispatchUrl'
  ];

  const payload = {};

  allowed.forEach(function (key) {
    if (Object.prototype.hasOwnProperty.call(request, key)) {
      payload[key] = request[key];
    }
  });

  if (!payload.edition) throw new Error('Edition is required.');
  if (!payload.title) throw new Error('Title is required.');
  if (!payload.html) throw new Error('Dispatch HTML is required.');

  return payload;
}

function forwardToMainPublisher_(payload) {
  const response = UrlFetchApp.fetch(MAIN_PUBLISHER_URL, {
    method: 'post',
    contentType: 'text/plain;charset=utf-8',
    payload: JSON.stringify(payload),
    followRedirects: true,
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(
      'Main Dispatch publisher returned HTTP ' + code +
      (text ? ': ' + text.slice(0, 500) : '.')
    );
  }

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (err) {
    throw new Error(
      'Main Dispatch publisher returned a non-JSON response: ' + text.slice(0, 500)
    );
  }

  if (parsed && (parsed.success === false || parsed.status === 'error')) {
    throw new Error(parsed.error || parsed.message || 'Main Dispatch publisher rejected the request.');
  }

  return parsed;
}

function normalizeUpstreamResponse_(response) {
  if (!response) return {};

  return {
    success: response.success !== false && response.status !== 'error',
    status: response.status || null,
    url: response.url || response.publicUrl || response.dispatchUrl || null,
    edition: response.edition || null,
    message: response.message || null
  };
}

function parseJson_(text) {
  try {
    return JSON.parse(text);
  } catch (err) {
    return null;
  }
}

function json_(object) {
  return ContentService
    .createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}
