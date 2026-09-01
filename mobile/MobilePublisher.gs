/**
 * THE PLEASURE DISPATCH
 * MobilePublisher.gs
 *
 * Mobile -> this web app -> Main Dispatch Publisher -> GitHub/site.
 *
 * Authentication is deliberately split:
 *   DRIVE_PUBLISH_KEY authenticates the mobile client to this bridge.
 *   PUBLISH_SECRET authenticates this bridge to the main Publisher.
 */

// CURRENT MAIN DISPATCH PUBLISHER. This is NOT the Drive API.
const MAIN_PUBLISHER_URL =
  'https://script.google.com/macros/s/AKfycbwpdElO35PiRlfhLvzISgpT3rtcxz8Iv5wewQoqvQJvC7yP02xN6UqrAjwPjfNHBv0T/exec';

const MOBILE_PUBLISH_KEY_PROPERTY = 'DRIVE_PUBLISH_KEY';
const UPSTREAM_PUBLISH_SECRET_PROPERTY = 'PUBLISH_SECRET';
const MAX_BODY_BYTES = 9000000;

function doGet() {
  return json_({
    success: true,
    service: 'Pleasure Dispatch Mobile Publisher',
    status: 'online',
    version: '1.4.0'
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
        mainPublisherConfigured: Boolean(MAIN_PUBLISHER_URL),
        authenticationConfigured: Boolean(
          PropertiesService.getScriptProperties().getProperty(MOBILE_PUBLISH_KEY_PROPERTY)
        ) && Boolean(
          PropertiesService.getScriptProperties().getProperty(UPSTREAM_PUBLISH_SECRET_PROPERTY)
        )
      });
    }

    if (action !== 'publish') {
      return json_({ success: false, error: 'Unsupported action. Use action: publish.' });
    }

    // Step 1: authenticate the mobile client with DRIVE_PUBLISH_KEY.
    authenticateMobile_(request.secret);

    // Step 2: create a publisher payload using the server-side PUBLISH_SECRET.
    // The mobile client never needs to know or send the main publisher secret.
    const payload = buildPublisherPayload_(request);
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

function authenticateMobile_(provided) {
  const expected = PropertiesService
    .getScriptProperties()
    .getProperty(MOBILE_PUBLISH_KEY_PROPERTY);

  if (!expected) {
    throw new Error('Mobile publisher is not configured. Set DRIVE_PUBLISH_KEY in Script Properties.');
  }

  if (!provided || String(provided) !== String(expected)) {
    throw new Error('Publishing key rejected.');
  }
}

function buildPublisherPayload_(request) {
  const payload = {};

  Object.keys(request).forEach(function(key) {
    payload[key] = request[key];
  });

  const publishSecret = PropertiesService
    .getScriptProperties()
    .getProperty(UPSTREAM_PUBLISH_SECRET_PROPERTY);

  if (!publishSecret) {
    throw new Error('Mobile publisher is not configured. Set PUBLISH_SECRET in Script Properties.');
  }

  payload.action = 'publish';
  payload.secret = publishSecret;

  if (!payload.edition) throw new Error('Edition is required.');
  if (!payload.title) throw new Error('Title is required.');
  if (!payload.html) throw new Error('Dispatch HTML is required.');

  return payload;
}

function forwardToMainPublisher_(payload) {
  const response = UrlFetchApp.fetch(MAIN_PUBLISHER_URL, {
    method: 'post',
    contentType: 'application/json;charset=utf-8',
    payload: JSON.stringify(payload),
    followRedirects: true,
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(
      'Main Dispatch publisher returned HTTP ' + code +
      (text ? ': ' + text.slice(0, 800) : '.')
    );
  }

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (err) {
    throw new Error(
      'Main Dispatch publisher returned a non-JSON response: ' + text.slice(0, 800)
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
