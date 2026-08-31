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

  async function stableUpload(dataUrl, originalFileName) {
    const payload = {
      action: "upload",
      publishKey: typeof PUBLISH_KEY !== "undefined" ? PUBLISH_KEY : "",
      fileName:
        "PD_" + editionName() + "_" + Date.now() + "_" + safeName(originalFileName) + ".jpg",
      mimeType: "image/jpeg",
      fileContent: safeBase64(dataUrl)
    };

    const response = await fetch(CURRENT_DRIVE_API, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      throw new Error("Google Drive returned HTTP " + response.status + ".");
    }

    const result = await response.json();

    if (!result || result.status !== "created" || !result.fileId) {
      throw new Error(
        result && result.message
          ? result.message
          : "Google Drive did not create the image."
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
