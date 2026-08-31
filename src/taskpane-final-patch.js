/* =========================================================
   THE PLEASURE DISPATCH
   taskpane-final-patch.js
   Final image/preview stabilization layer

   Goals:
   - One authoritative Drive upload endpoint.
   - Always store browser-processed images as real JPEG files.
   - Never report an image as ready until its Drive thumbnail loads.
   - Keep Drive as the archive and the existing public site/lightbox.
   - Avoid data: and cid: image sources in the editor preview.
   - Return actionable diagnostics when the Drive bridge rejects a request.
========================================================= */
(function () {
  "use strict";

  const CURRENT_DRIVE_API =
    "https://script.google.com/macros/s/AKfycbyAodEuYNjMNIRf-6RhBLRaRCYaXtgTSDqYTpD3X7-YyHGVmXt5TQBC6v40joHDklTm0w/exec";

  const DRIVE_THUMBNAIL = function (fileId) {
    return "https://drive.google.com/thumbnail?id=" +
      encodeURIComponent(fileId) +
      "&sz=w1800";
  };

  const DRIVE_FILE = function (fileId) {
    return "https://drive.google.com/file/d/" +
      encodeURIComponent(fileId) +
      "/view";
  };

  function verifyImage(url) {
    return new Promise(function (resolve) {
      if (!url || /^cid:/i.test(url) || /^data:/i.test(url)) {
        resolve(false);
        return;
      }

      const image = new Image();
      let complete = false;

      function finish(ok) {
        if (complete) return;
        complete = true;
        resolve(ok);
      }

      image.onload = function () {
        finish(!!(image.naturalWidth && image.naturalHeight));
      };

      image.onerror = function () {
        finish(false);
      };

      setTimeout(function () {
        finish(false);
      }, 10000);

      image.src = url;
    });
  }

  function safeBase64(dataUrl) {
    const text = String(dataUrl || "");
    const comma = text.indexOf(",");
    return comma === -1 ? text : text.substring(comma + 1);
  }

  function safeName(name) {
    return String(name || "image")
      .replace(/[\\/:*?"<>|#%{}\[\]]/g, "_")
      .replace(/\s+/g, "_")
      .replace(/\.[^.]+$/, "")
      .substring(0, 90);
  }

  function editionName() {
    const input = document.getElementById("edition");
    return String(input ? input.value : "001")
      .trim()
      .replace(/[^a-zA-Z0-9_-]/g, "_") || "001";
  }

  function describeBridgeFailure(result, httpStatus, rawText) {
    if (result && result.code) {
      const code = String(result.code);
      const message = result.message ? String(result.message) : "";

      if (code === "INVALID_PUBLISH_KEY") {
        return "Google Drive rejected the publishing key. Confirm PUBLISH_KEY matches the Drive bridge Script Property and the add-in key.";
      }

      if (code === "UNKNOWN_ACTION") {
        return "Google Drive received the request but did not recognize the upload action. The add-in may be using an outdated Drive bridge deployment.";
      }

      if (code === "NO_POST_DATA") {
        return "Google Drive received no POST body.";
      }

      if (code === "INVALID_JSON") {
        return "Google Drive could not read the upload JSON payload.";
      }

      if (message) {
        return message;
      }
    }

    if (result && result.message) {
      return String(result.message);
    }

    if (result && result.status) {
      return "Google Drive returned status '" +
        String(result.status) +
        "' without a usable fileId.";
    }

    if (rawText) {
      return "Google Drive returned an unexpected response: " +
        String(rawText).substring(0, 500);
    }

    return "Google Drive did not create the image" +
      (httpStatus ? " (HTTP " + httpStatus + ")" : ".");
  }

  async function stableUpload(dataUrl, originalFileName) {
    if (!dataUrl || !/^data:image\//i.test(String(dataUrl))) {
      throw new Error("The image data was empty or was not a browser image.");
    }

    const publishKey =
      typeof PUBLISH_KEY !== "undefined"
        ? String(PUBLISH_KEY || "")
        : "";

    if (!publishKey) {
      throw new Error("The Outlook add-in has no PUBLISH_KEY configured.");
    }

    const payload = {
      action: "upload",
      publishKey: publishKey,
      fileName:
        "PD_" + editionName() + "_" + Date.now() + "_" + safeName(originalFileName) + ".jpg",
      mimeType: "image/jpeg",
      fileContent: safeBase64(dataUrl)
    };

    if (!payload.fileContent || payload.fileContent.length < 100) {
      throw new Error("The compressed image data was empty.");
    }

    console.log("[Pleasure Dispatch] Drive upload endpoint:", CURRENT_DRIVE_API);
    console.log("[Pleasure Dispatch] Drive upload filename:", payload.fileName);
    console.log("[Pleasure Dispatch] Drive upload payload size:", payload.fileContent.length);

    let response;

    try {
      response = await fetch(CURRENT_DRIVE_API, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8"
        },
        body: JSON.stringify(payload)
      });
    } catch (error) {
      throw new Error(
        "Could not connect to the Google Drive bridge: " +
        (error && error.message ? error.message : String(error))
      );
    }

    const responseText = await response.text();
    let result = null;

    try {
      result = responseText ? JSON.parse(responseText) : null;
    } catch (error) {
      console.error("[Pleasure Dispatch] Drive returned non-JSON:", responseText);
      throw new Error(
        "Google Drive returned an unreadable response (HTTP " +
        response.status + ")."
      );
    }

    /*
     * Do not log the publish key or image payload.
     * Log the complete response shape needed to diagnose bridge/deployment issues.
     */
    console.log("[Pleasure Dispatch] Drive upload HTTP status:", response.status);
    console.log("[Pleasure Dispatch] Drive upload response status:", result && result.status);
    console.log("[Pleasure Dispatch] Drive upload response code:", result && result.code);
    console.log("[Pleasure Dispatch] Drive upload response message:", result && result.message);
    console.log("[Pleasure Dispatch] Drive upload response fileId:", result && result.fileId);
    console.log("[Pleasure Dispatch] Drive upload response mimeType:", result && result.mimeType);
    console.log("[Pleasure Dispatch] Drive upload response sizeBytes:", result && result.sizeBytes);
    console.log("[Pleasure Dispatch] Drive upload response:", result);

    if (!response.ok) {
      throw new Error(
        describeBridgeFailure(result, response.status, responseText) +
        " (HTTP " + response.status + ")."
      );
    }

    /*
     * The current Drive bridge returns:
     *   { status: "created", fileId: "...", ... }
     *
     * Accept only a real Drive file ID as proof that the upload succeeded.
     * This also protects us if a deployment returns a different success label.
     */
    if (!result || !result.fileId) {
      throw new Error(
        describeBridgeFailure(result, response.status, responseText)
      );
    }

    if (result.status !== "created") {
      throw new Error(
        "Google Drive returned a fileId but an unexpected status ('" +
        String(result.status || "[empty]") +
        "'). The image may have been created, but the bridge response is not the expected upload response."
      );
    }

    const fileId = result.fileId;
    const imageUrl = result.imageUrl || DRIVE_THUMBNAIL(fileId);
    const driveUrl = result.driveUrl || DRIVE_FILE(fileId);

    if (result.mimeType && result.mimeType !== "image/jpeg") {
      throw new Error(
        "Drive created the image with MIME type " + result.mimeType + " instead of image/jpeg."
      );
    }

    if (!result.sizeBytes || Number(result.sizeBytes) < 100) {
      throw new Error("Drive created an image file with an invalid or empty size.");
    }

    const ready = await verifyImage(imageUrl);

    if (!ready) {
      throw new Error(
        "Drive created the image, but its thumbnail is not accessible yet."
      );
    }

    return Object.assign({}, result, {
      fileId: fileId,
      imageUrl: imageUrl,
      driveUrl: driveUrl,
      lightboxUrl:
        (typeof buildLightboxUrl === "function")
          ? buildLightboxUrl(fileId)
          : driveUrl,
      previewReady: true
    });
  }

  /* Replace the production patch's upload implementation. */
  window.uploadToDriveOnce = stableUpload;

  /* Make Drive URLs consistently generated from the authoritative endpoint. */
  window.PleasureDispatchImageBridge = {
    driveApiUrl: CURRENT_DRIVE_API,
    buildDriveImageUrl: DRIVE_THUMBNAIL,
    buildDriveFileUrl: DRIVE_FILE,
    verifyImage: verifyImage,
    upload: stableUpload
  };

  /* Preview guard: cid: and data: must never be rendered by the task pane. */
  const originalRenderHeroPreview = window.renderHeroPreview;
  if (typeof originalRenderHeroPreview === "function") {
    window.renderHeroPreview = function (key, url) {
      if (/^(cid:|data:)/i.test(String(url || ""))) {
        const preview = document.getElementById(key + "Preview");
        if (preview) {
          preview.className = "preview hero-preview error";
          preview.textContent = "Image is not available as a browser preview.";
        }
        return;
      }
      return originalRenderHeroPreview.apply(this, arguments);
    };
  }

  const originalRenderModulePreview = window.renderModulePreview;
  if (typeof originalRenderModulePreview === "function") {
    window.renderModulePreview = function (item, url) {
      if (/^(cid:|data:)/i.test(String(url || ""))) {
        const preview = item && item.querySelector
          ? item.querySelector(".module-preview")
          : null;
        if (preview) {
          preview.className = "preview module-preview error";
          preview.textContent = "Image is not available as a browser preview.";
        }
        return;
      }
      return originalRenderModulePreview.apply(this, arguments);
    };
  }

  console.log("[Pleasure Dispatch] Final image/preview patch loaded.");
})();
