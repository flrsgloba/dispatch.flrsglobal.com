/* =========================================================
   THE PLEASURE DISPATCH
   FLRSGLOBAL
   taskpane.js — clean consolidated build
========================================================= */

const DRIVE_API_URL =
  "https://script.google.com/macros/s/AKfycbyTLvYbe1O_BbzsH09UMSZdbY9_XZXga-TbSPkR3UclT3Qhlaj7gy5yhXPA_UpE6Fu7tw/exec";

const LIGHTBOX_BASE_URL =
  "https://flrsgloba.github.io/dispatch.flrsglobal.com/lightbox.html";

const LOOP_ASSET_URL =
  "https://flrsgloba.github.io/dispatch.flrsglobal.com/assets/pleasure-loop.png";

const COLORS = {
  background: "#303030",
  surface: "#595959",
  text: "#F2EEE5",
  secondary: "#C9C3B8",
  rule: "#777777",
  accent: "#D8D0C3"
};

const IMAGE_MAX_WIDTH = 1800;
const IMAGE_MAX_HEIGHT = 1800;
const IMAGE_QUALITY = 0.82;
const MAX_SOURCE_IMAGE_MB = 40;
const DRIVE_IMAGE_WIDTH = 1800;
const BUILD_TIMEOUT_MS = 60000;

let blockCounter = 0;
let pleasureCounter = 0;
let initialized = false;

Office.onReady(function () {
  initializeDispatch();
});

function initializeDispatch() {
  if (initialized) return;
  initialized = true;

  bindStaticControls();
  bindDelegatedControls();
  setupHero("hero1");
  setupHero("hero2");

  const pleasureRows = $("pleasureRows");
  if (pleasureRows && pleasureRows.children.length === 0) {
    addPleasureRow("Coffee", "");
    addPleasureRow("Art", "");
    addPleasureRow("Object", "");
  }

  const imageBlocks = $("imageBlocks");
  if (imageBlocks && imageBlocks.children.length === 0) {
    addImageBlock();
  }

  setStatus("The Pleasure Dispatch is ready.");
}

function $(id) {
  return document.getElementById(id);
}

function val(id) {
  const el = $(id);
  return el ? el.value.trim() : "";
}

function esc(text) {
  return String(text || "").replace(/[&<>"']/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[c];
  });
}

function attr(text) {
  return esc(text);
}

function setStatus(message) {
  const status = $("status");
  if (status) status.textContent = message;
  console.log("[Pleasure Dispatch]", message);
}

function paragraph(text) {
  if (!text) return "";
  return '<p style="font:17px/1.75 Garamond,Georgia,Times New Roman,serif;color:' + COLORS.text + ';margin:0 0 24px;">' +
    esc(text).replace(/\n/g, "<br>") +
    "</p>";
}

function buildDriveImageUrl(fileId) {
  return "https://drive.google.com/thumbnail?id=" + encodeURIComponent(fileId) + "&sz=w" + DRIVE_IMAGE_WIDTH;
}

function buildDriveFileUrl(fileId) {
  return "https://drive.google.com/file/d/" + encodeURIComponent(fileId) + "/view";
}

function buildLightboxUrl(fileId, caption) {
  let url = LIGHTBOX_BASE_URL + "?id=" + encodeURIComponent(fileId);
  if (caption) url += "&caption=" + encodeURIComponent(caption);
  return url;
}

function editionSafe() {
  return (val("edition") || "001").replace(/[^a-zA-Z0-9_-]/g, "_");
}

function cleanFileName(name) {
  return String(name || "image")
    .replace(/[\\/:*?"<>|#%{}\[\]]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/\.[^.]+$/, "")
    .substring(0, 80);
}

function stripDataUrl(dataUrl) {
  const comma = dataUrl.indexOf(",");
  return comma === -1 ? dataUrl : dataUrl.substring(comma + 1);
}

function blobToDataUrl(blob, callback) {
  const reader = new FileReader();
  reader.onload = function () { callback(reader.result); };
  reader.onerror = function () { callback(null); };
  reader.readAsDataURL(blob);
}

function compressImage(file, callback) {
  if (!file) {
    callback(null, new Error("No image selected."));
    return;
  }

  const sizeMB = file.size / (1024 * 1024);
  if (sizeMB > MAX_SOURCE_IMAGE_MB) {
    callback(null, new Error("That image is larger than " + MAX_SOURCE_IMAGE_MB + " MB."));
    return;
  }

  if (!/^image\/(jpeg|jpg|png|webp)$/.test(file.type || "")) {
    callback(null, new Error("Please choose a JPG, PNG, or WebP image."));
    return;
  }

  const reader = new FileReader();
  reader.onload = function () {
    const image = new Image();
    image.onload = function () {
      let width = image.naturalWidth;
      let height = image.naturalHeight;
      const scale = Math.min(1, IMAGE_MAX_WIDTH / width, IMAGE_MAX_HEIGHT / height);
      width = Math.round(width * scale);
      height = Math.round(height * scale);

      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext("2d");
      context.fillStyle = "#ffffff";
      context.fillRect(0, 0, width, height);
      context.drawImage(image, 0, 0, width, height);

      canvas.toBlob(function (blob) {
        if (!blob) {
          callback(null, new Error("Image compression failed."));
          return;
        }
        blobToDataUrl(blob, function (dataUrl) {
          callback({ dataUrl: dataUrl, blob: blob, width: width, height: height }, null);
        });
      }, "image/jpeg", IMAGE_QUALITY);
    };
    image.onerror = function () {
      callback(null, new Error("The image could not be decoded."));
    };
    image.src = reader.result;
  };
  reader.onerror = function () {
    callback(null, new Error("The selected image could not be read."));
  };
  reader.readAsDataURL(file);
}

async function uploadToDrive(dataUrl, originalFileName) {
  const payload = {
    action: "upload",
    fileName: "PD_" + editionSafe() + "_" + Date.now() + "_" + cleanFileName(originalFileName),
    mimeType: "image/jpeg",
    fileContent: stripDataUrl(dataUrl)
  };

  setStatus("Saving image to Google Drive…");

  let response;
  try {
    response = await fetch(DRIVE_API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    });
  } catch (error) {
    throw new Error("Could not connect to Google Drive: " + error.message);
  }

  if (!response.ok) {
    throw new Error("Google Drive returned HTTP " + response.status + ".");
  }

  let result;
  try {
    result = await response.json();
  } catch (error) {
    throw new Error("Google Drive returned an unreadable response.");
  }

  if (!result || result.status !== "created") {
    throw new Error(result && result.message ? result.message : "Google Drive did not create the image.");
  }

  result.imageUrl = buildDriveImageUrl(result.fileId);
  result.driveUrl = buildDriveFileUrl(result.fileId);
  result.lightboxUrl = buildLightboxUrl(result.fileId);

  return result;
}

/* =========================================================
   CONTROLS
========================================================= */

function bindStaticControls() {
  const preview = $("previewBtn");
  const build = $("insertBtn");
  const addPleasure = $("addPleasure");
  let addImage = $("addImageBlock");

  if (preview) {
    preview.onclick = function (e) {
      e.preventDefault();
      previewNewsletter();
    };
  }

  if (build) {
    build.onclick = function (e) {
      e.preventDefault();
      buildInOutlook();
    };
  }

  if (addPleasure) {
    addPleasure.onclick = function (e) {
      e.preventDefault();
      addPleasureRow("", "");
      setStatus("Pleasure Note added.");
    };
  }

  if (!addImage) {
    const candidates = document.querySelectorAll("button,a,[role='button']");
    for (let i = 0; i < candidates.length; i++) {
      if ((candidates[i].textContent || "").toLowerCase().includes("add image block")) {
        addImage = candidates[i];
        break;
      }
    }
  }

  if (addImage) {
    addImage.onclick = function (e) {
      e.preventDefault();
      addImageBlock();
    };
  }
}

function bindDelegatedControls() {
  document.addEventListener("click", function (event) {
    const button = event.target.closest("button,a,[role='button']");
    if (!button) return;

    if (
      button.id === "previewBtn" ||
      button.id === "insertBtn" ||
      button.id === "addImageBlock" ||
      button.id === "addPleasure"
    ) return;

    if (button.classList.contains("move-up")) {
      event.preventDefault();
      moveImageBlock(button.closest(".image-block"), -1);
      return;
    }

    if (button.classList.contains("move-down")) {
      event.preventDefault();
      moveImageBlock(button.closest(".image-block"), 1);
      return;
    }

    if (button.classList.contains("duplicate")) {
      event.preventDefault();
      duplicateBlock(button.closest(".image-block"));
      return;
    }

    if (button.classList.contains("remove")) {
      event.preventDefault();
      const block = button.closest(".image-block");
      if (block) {
        block.remove();
        renumberImageBlocks();
        setStatus("Image Module removed.");
      }
      return;
    }

    if (button.classList.contains("remove-image")) {
      event.preventDefault();
      removeImageItem(button.closest(".image-item"));
      return;
    }

    if (button.classList.contains("url-item")) {
      event.preventDefault();
      const item = button.closest(".image-item");
      const input = item && item.querySelector(".module-url");
      if (input) input.focus();
    }
  });

  document.addEventListener("change", function (event) {
    if (event.target.matches(".layout-select")) {
      const block = event.target.closest(".image-block");
      if (block) {
        block.dataset.layout = event.target.value;
        syncImageInputs(block);
        setStatus("Layout changed to " + event.target.options[event.target.selectedIndex].text + ".");
      }
      return;
    }

    if (event.target.matches(".module-file-input")) {
      const item = event.target.closest(".image-item");
      const file = event.target.files && event.target.files[0];
      if (item && file) handleModuleUpload(item, file);
    }
  });

  document.addEventListener("input", function (event) {
    if (event.target.matches(".module-url")) {
      const item = event.target.closest(".image-item");
      if (item) {
        delete item.dataset.fileId;
        delete item.dataset.fullUrl;
        delete item.dataset.lightboxUrl;
        renderModulePreview(item, event.target.value.trim());
      }
    }

    if (event.target.id === "hero1Url") renderHeroPreview("hero1", event.target.value.trim());
    if (event.target.id === "hero2Url") renderHeroPreview("hero2", event.target.value.trim());
  });
}

/* =========================================================
   HEROES
========================================================= */

function setupHero(key) {
  const fileInput = $(key + "File");
  const urlInput = $(key + "Url");

  if (fileInput) {
    fileInput.addEventListener("change", function () {
      const file = fileInput.files && fileInput.files[0];
      if (file) handleHeroUpload(key, file);
    });
  }

  if (urlInput) {
    urlInput.addEventListener("input", function () {
      renderHeroPreview(key, urlInput.value.trim());
    });
  }
}

function handleHeroUpload(key, file) {
  setStatus("Optimizing hero image…");
  compressImage(file, async function (result, error) {
    if (error) {
      setStatus(error.message);
      return;
    }

    const input = $(key + "Url");
    const preview = $(key + "Preview");

    input.value = result.dataUrl;
    renderHeroPreview(key, result.dataUrl);

    try {
      const drive = await uploadToDrive(result.dataUrl, file.name);
      input.value = drive.imageUrl;
      if (preview) {
        preview.dataset.fileId = drive.fileId;
        preview.dataset.fullUrl = drive.lightboxUrl;
        preview.dataset.lightboxUrl = drive.lightboxUrl;
      }
      renderHeroPreview(key, drive.imageUrl);
      setStatus(key === "hero1" ? "✓ Hero Image 01 saved to Drive." : "✓ Hero Image 02 saved to Drive.");
    } catch (error) {
      setStatus("Preview ready. Drive upload failed: " + error.message);
    }
  });
}

function renderHeroPreview(key, url) {
  const preview = $(key + "Preview");
  if (!preview) return;

  if (!url) {
    preview.className = "preview hero-preview empty";
    preview.textContent = "No hero image selected";
    return;
  }

  preview.className = "preview hero-preview";
  preview.innerHTML = "";
  const image = new Image();
  image.onload = function () { preview.appendChild(image); };
  image.onerror = function () {
    preview.className = "preview hero-preview error";
    preview.textContent = "Image could not be loaded.";
  };
  image.src = url;
}

/* =========================================================
   MODULAR IMAGES
========================================================= */

function handleModuleUpload(item, file) {
  setStatus("Optimizing modular image…");

  compressImage(file, async function (result, error) {
    if (error) {
      setStatus(error.message);
      return;
    }

    const input = item.querySelector(".module-url");
    if (!input) {
      setStatus("Could not find the modular image field.");
      return;
    }

    input.value = result.dataUrl;
    renderModulePreview(item, result.dataUrl);

    try {
      const drive = await uploadToDrive(result.dataUrl, file.name);

      item.dataset.fileId = drive.fileId;
      item.dataset.imageUrl = drive.imageUrl;
      item.dataset.fullUrl = drive.lightboxUrl;
      item.dataset.lightboxUrl = drive.lightboxUrl;

      input.value = drive.imageUrl;
      renderModulePreview(item, drive.imageUrl);

      setStatus("✓ Modular image saved to Drive.");
    } catch (error) {
      setStatus("Preview ready. Drive upload failed: " + error.message);
    }
  });
}

function renderModulePreview(item, url) {
  const preview = item.querySelector(".module-preview");
  if (!preview) return;

  if (!url) {
    preview.className = "preview module-preview empty";
    preview.textContent = "No image selected";
    return;
  }

  preview.className = "preview module-preview";
  preview.innerHTML = "";

  const image = new Image();
  image.onload = function () { preview.appendChild(image); };
  image.onerror = function () {
    preview.className = "preview module-preview error";
    preview.textContent = "Image could not be loaded.";
  };
  image.src = url;
}

/* =========================================================
   IMAGE BLOCKS
========================================================= */

function getImageBlocks() {
  const container = $("imageBlocks");
  if (!container) return [];
  return Array.from(container.children).filter(function (element) {
    return element.classList.contains("image-block");
  });
}

function addImageBlock() {
  const container = $("imageBlocks");
  if (!container) {
    setStatus("Image module container not found.");
    return;
  }

  blockCounter++;

  const block = document.createElement("article");
  block.className = "image-block";
  block.dataset.id = String(blockCounter);
  block.dataset.layout = "full";

  block.innerHTML =
    '<div class="block-top">' +
      '<span class="block-number">Image Module ' + blockCounter + '</span>' +
      '<select class="layout-select" aria-label="Image layout">' +
        '<option value="full">Full Width</option>' +
        '<option value="two">Two Up</option>' +
        '<option value="three">Three Up</option>' +
        '<option value="four">Four Up</option>' +
      '</select>' +
    '</div>' +
    '<div class="image-items one"></div>' +
    '<div class="block-actions">' +
      '<button type="button" class="move-up">↑ Move Up</button>' +
      '<button type="button" class="move-down">↓ Move Down</button>' +
      '<button type="button" class="duplicate">Duplicate</button>' +
      '<button type="button" class="remove">Remove</button>' +
    '</div>';

  container.appendChild(block);
  syncImageInputs(block);
  renumberImageBlocks();
  setStatus("✓ Image Module added.");
}

function renumberImageBlocks() {
  getImageBlocks().forEach(function (block, index) {
    const number = block.querySelector(".block-number");
    if (number) number.textContent = "Image Module " + (index + 1);
    block.dataset.position = String(index + 1);
  });
}

function moveImageBlock(block, direction) {
  const container = $("imageBlocks");
  const blocks = getImageBlocks();
  if (!container || !block) return;

  const index = blocks.indexOf(block);
  const targetIndex = index + direction;

  if (index === -1) return;

  if (targetIndex < 0 || targetIndex >= blocks.length) {
    setStatus(direction < 0 ? "Already at the top." : "Already at the bottom.");
    return;
  }

  const target = blocks[targetIndex];

  if (direction < 0) container.insertBefore(block, target);
  else container.insertBefore(block, target.nextSibling);

  renumberImageBlocks();
  setStatus(direction < 0 ? "✓ Image Module moved up." : "✓ Image Module moved down.");
}

function syncImageInputs(block) {
  const layout = block.dataset.layout || "full";
  const count = layout === "full" ? 1 : layout === "two" ? 2 : layout === "three" ? 3 : 4;
  const wrap = block.querySelector(".image-items");
  if (!wrap) return;

  wrap.className = "image-items " + (layout === "full" ? "one" : layout);

  while (wrap.children.length > count) wrap.lastElementChild.remove();
  while (wrap.children.length < count) addImageItem(wrap, wrap.children.length + 1, {});

  renumberImageItems(wrap);
}

function addImageItem(wrap, number, preset) {
  preset = preset || {};

  const item = document.createElement("div");
  item.className = "image-item";

  item.innerHTML =
    '<label for="moduleFile_' + blockCounter + '_' + number + '">Image ' + number + '</label>' +
    '<label class="upload-button" for="moduleFile_' + blockCounter + '_' + number + '">Upload</label>' +
    '<input id="moduleFile_' + blockCounter + '_' + number + '" name="moduleFile_' + blockCounter + '_' + number + '" type="file" class="module-file-input" accept="image/jpeg,image/png,image/webp" hidden>' +
    '<button type="button" class="secondary url-item">Image URL</button>' +
    '<label for="moduleUrl_' + blockCounter + '_' + number + '" class="sr-only">Image URL</label>' +
    '<input id="moduleUrl_' + blockCounter + '_' + number + '" name="moduleUrl_' + blockCounter + '_' + number + '" class="module-url source-url" value="' + attr(preset.url || "") + '" placeholder="Paste direct image URL">' +
    '<div class="preview module-preview empty">No image selected</div>' +
    '<label for="moduleCaption_' + blockCounter + '_' + number + '" class="sr-only">Caption</label>' +
    '<input id="moduleCaption_' + blockCounter + '_' + number + '" name="moduleCaption_' + blockCounter + '_' + number + '" class="module-caption" value="' + attr(preset.caption || "") + '" placeholder="Caption">' +
    '<button type="button" class="remove-image">Remove image</button>';

  wrap.appendChild(item);

  if (preset.clickUrl) item.dataset.fullUrl = preset.clickUrl;
  if (preset.url) renderModulePreview(item, preset.url);
}

function renumberImageItems(wrap) {
  Array.from(wrap.children).forEach(function (item, index) {
    const label = item.querySelector("label:first-child");
    if (label) label.textContent = "Image " + (index + 1);
  });
}

function removeImageItem(item) {
  if (!item) return;
  const wrap = item.parentElement;
  if (!wrap) return;

  if (wrap.children.length === 1) {
    const url = item.querySelector(".module-url");
    const caption = item.querySelector(".module-caption");
    if (url) url.value = "";
    if (caption) caption.value = "";
    delete item.dataset.fileId;
    delete item.dataset.fullUrl;
    delete item.dataset.lightboxUrl;
    renderModulePreview(item, "");
    setStatus("Image removed.");
    return;
  }

  item.remove();
  renumberImageItems(wrap);
  setStatus("Image removed.");
}

function duplicateBlock(block) {
  if (!block) return;

  const sourceLayout = block.dataset.layout || "full";
  const sourceItems = Array.from(block.querySelectorAll(".image-item")).map(function (item) {
    const url = item.querySelector(".module-url");
    const caption = item.querySelector(".module-caption");
    return {
      url: url ? url.value.trim() : "",
      caption: caption ? caption.value.trim() : "",
      clickUrl: item.dataset.fullUrl || item.dataset.lightboxUrl || ""
    };
  });

  const container = $("imageBlocks");
  const oldCounter = blockCounter;
  addImageBlock();
  const clone = getImageBlocks()[getImageBlocks().length - 1];

  clone.dataset.layout = sourceLayout;
  syncImageInputs(clone);

  const cloneItems = clone.querySelectorAll(".image-item");
  sourceItems.forEach(function (data, index) {
    if (!cloneItems[index]) return;
    const url = cloneItems[index].querySelector(".module-url");
    const caption = cloneItems[index].querySelector(".module-caption");
    if (url) url.value = data.url;
    if (caption) caption.value = data.caption;
    if (data.clickUrl) cloneItems[index].dataset.fullUrl = data.clickUrl;
    if (data.url) renderModulePreview(cloneItems[index], data.url);
  });

  if (container && clone.previousElementSibling !== block) {
    container.insertBefore(clone, block.nextSibling);
  }

  if (blockCounter < oldCounter + 1) blockCounter = oldCounter + 1;
  renumberImageBlocks();
  setStatus("✓ Image Module duplicated.");
}

/* =========================================================
   NEWSLETTER COLLECTION
========================================================= */

function collectPleasureNotes() {
  return Array.from(document.querySelectorAll(".pleasure-row"))
    .map(function (row) {
      const label = row.querySelector(".pleasure-label");
      const note = row.querySelector(".pleasure-value");
      return {
        label: label ? label.value.trim() : "",
        value: note ? note.value.trim() : ""
      };
    })
    .filter(function (item) { return item.label || item.value; });
}

function addPleasureRow(labelValue, noteValue) {
  const container = $("pleasureRows");
  if (!container) return;

  pleasureCounter++;
  const rowId = "pleasureRow_" + pleasureCounter;
  const labelId = rowId + "_label";
  const valueId = rowId + "_value";

  const row = document.createElement("div");
  row.className = "pleasure-row";
  row.innerHTML =
    '<label for="' + labelId + '" class="sr-only">Category</label>' +
    '<input id="' + labelId + '" name="' + labelId + '" class="pleasure-label" value="' + attr(labelValue || "") + '" placeholder="Category">' +
    '<label for="' + valueId + '" class="sr-only">Pleasure Note</label>' +
    '<input id="' + valueId + '" name="' + valueId + '" class="pleasure-value" value="' + attr(noteValue || "") + '" placeholder="What has held your attention?">' +
    '<button type="button" aria-label="Remove pleasure note">×</button>';

  row.querySelector("button").addEventListener("click", function () {
    row.remove();
    setStatus("Pleasure Note removed.");
  });

  container.appendChild(row);
}

function buildPleasureNotesHtml(notes) {
  if (!notes.length) return "";
  return '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0"><tbody>' +
    notes.map(function (note) {
      return '<tr>' +
        '<td width="120" style="padding:5px 18px 5px 0;color:' + COLORS.secondary + ';font:10px Arial,sans-serif;text-transform:uppercase;letter-spacing:1px;vertical-align:top;">' + esc(note.label) + '</td>' +
        '<td style="padding:5px 0;color:' + COLORS.text + ';font:16px/1.5 Garamond,Georgia,serif;vertical-align:top;">' + esc(note.value) + '</td>' +
      '</tr>';
    }).join("") +
    '</tbody></table>';
}

function collectImageBlocks() {
  return getImageBlocks().map(function (block) {
    return {
      layout: block.dataset.layout || "full",
      items: Array.from(block.querySelectorAll(".image-item")).map(function (item) {
        const url = item.querySelector(".module-url");
        const caption = item.querySelector(".module-caption");
        return {
          url: url ? url.value.trim() : "",
          clickUrl: item.dataset.lightboxUrl || item.dataset.fullUrl || "",
          caption: caption ? caption.value.trim() : ""
        };
      }).filter(function (item) {
        return item.url || item.caption;
      })
    };
  }).filter(function (block) {
    return block.items.length;
  });
}

/* =========================================================
   EMAIL HTML
========================================================= */

function emailImage(item) {
  if (!item || !item.url) return "";
  return '<a href="' + attr(item.clickUrl || item.url) + '" target="_blank" style="text-decoration:none;">' +
    '<img src="' + attr(item.url) + '" alt="" width="100%" style="display:block;width:100%;height:auto;border:0;">' +
    '</a>' +
    (item.caption ? '<div style="padding-top:7px;color:' + COLORS.secondary + ';font:12px/1.4 Arial,sans-serif;">' + esc(item.caption) + '</div>' : "");
}

function moduleHtml(block) {
  if (!block || !block.items.length) return "";

  if (block.layout === "full") {
    return '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="margin:28px 0;background:' + COLORS.surface + ';"><tr><td style="padding:8px;">' +
      emailImage(block.items[0]) +
      '</td></tr></table>';
  }

  const columns = block.layout === "two" ? 2 : block.layout === "three" ? 3 : 4;
  const width = Math.floor(100 / columns);

  return '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="margin:28px 0;table-layout:fixed;"><tr>' +
    block.items.slice(0, columns).map(function (item) {
      return '<td width="' + width + '%" valign="top" style="width:' + width + '%;padding:3px;background:' + COLORS.surface + ';vertical-align:top;">' +
        emailImage(item) +
      '</td>';
    }).join("") +
    '</tr></table>';
}

function buildNewsletterHtml() {
  const hero1 = val("hero1Url");
  const hero2 = val("hero2Url");
  const hero1Preview = $("hero1Preview");
  const hero2Preview = $("hero2Preview");
  const modules = collectImageBlocks();

  const hero1Click = hero1Preview && (hero1Preview.dataset.lightboxUrl || hero1Preview.dataset.fullUrl) || hero1;
  const hero2Click = hero2Preview && (hero2Preview.dataset.lightboxUrl || hero2Preview.dataset.fullUrl) || hero2;

  const hero1Html = hero1 ? emailImage({url: hero1, clickUrl: hero1Click, caption: val("hero1Caption")}) : "";
  const hero2Html = hero2 ? emailImage({url: hero2, clickUrl: hero2Click, caption: val("hero2Caption")}) : "";
  const moduleHtml = modules.map(moduleHtml).join("");

  const invite = (val("inviteTitle") || val("inviteText") || val("ctaUrl")) ?
    '<div style="margin:38px 0 10px;color:' + COLORS.secondary + ';font:10px Arial,sans-serif;letter-spacing:1.5px;text-transform:uppercase;">05 — AN INVITATION</div>' +
    (val("inviteTitle") ? '<div style="color:' + COLORS.text + ';font:27px/1.15 Garamond,Georgia,serif;margin:0 0 10px;">' + esc(val("inviteTitle")) + '</div>' : "") +
    paragraph(val("inviteText")) +
    (val("ctaUrl") ? '<p><a href="' + attr(val("ctaUrl")) + '" style="color:' + COLORS.text + ';font:10px Arial,sans-serif;letter-spacing:1px;text-transform:uppercase;">' + esc(val("ctaLabel") || "INQUIRE") + '</a></p>' : "") : "";

  const notes = buildPleasureNotesHtml(collectPleasureNotes());

  return '<div style="margin:0;padding:0;background:' + COLORS.background + ';color:' + COLORS.text + ';">' +
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0"><tr><td align="center" style="padding:20px 10px;">' +
    '<table role="presentation" width="680" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:680px;background:' + COLORS.background + ';">' +

    '<tr><td style="padding:34px 36px 20px;border-bottom:1px solid ' + COLORS.rule + ';">' +
      '<div style="color:' + COLORS.text + ';font:10px Arial,sans-serif;letter-spacing:2px;">FLRSGLOBAL</div>' +
      '<div style="color:' + COLORS.secondary + ';font:10px Arial,sans-serif;letter-spacing:1.4px;margin-top:8px;">FROM THE STUDIO OF FREDDIE L. RANKIN II</div>' +
      '<div style="color:' + COLORS.text + ';font:400 48px/1 Garamond,Georgia,serif;margin:18px 0 10px;">The Pleasure Dispatch</div>' +
      '<div style="color:' + COLORS.secondary + ';font:10px Arial,sans-serif;letter-spacing:1.2px;">' + esc(val("edition")) + (val("date") ? ' · ' + esc(val("date")) : "") + (val("title") ? ' · ' + esc(val("title")) : "") + '</div>' +
    '</td></tr>' +

    '<tr><td style="padding:24px 36px 42px;">' +
      '<div style="text-align:center;padding:2px 0 26px;"><img src="' + LOOP_ASSET_URL + '" alt="The Loop" width="92" style="display:inline-block;width:92px;height:auto;border:0;"></div>' +
      (val("subtitle") ? '<div style="color:' + COLORS.secondary + ';font:18px/1.45 Garamond,Georgia,serif;margin:0 0 24px;">' + esc(val("subtitle")) + '</div>' : "") +
      (hero1Html ? '<div style="margin-bottom:8px;">' + hero1Html + '</div>' : "") +
      '<div style="color:' + COLORS.secondary + ';font:10px Arial,sans-serif;letter-spacing:1.5px;margin:30px 0 10px;">01 — A REFLECTION</div>' +
      paragraph(val("reflection")) +
      '<div style="color:' + COLORS.secondary + ';font:10px Arial,sans-serif;letter-spacing:1.5px;margin:36px 0 10px;">02 — THE WORK</div>' +
      paragraph(val("workText")) +
      moduleHtml +
      '<div style="color:' + COLORS.secondary + ';font:10px Arial,sans-serif;letter-spacing:1.5px;margin:36px 0 10px;">03 — STUDIO NOTES</div>' +
      paragraph(val("studioText")) +
      (hero2Html ? '<div style="margin-top:8px;">' + hero2Html + '</div>' : "") +
      '<div style="color:' + COLORS.secondary + ';font:10px Arial,sans-serif;letter-spacing:1.5px;margin:36px 0 10px;">04 — PLEASURE NOTES</div>' +
      '<div style="color:' + COLORS.text + ';font:19px/1.4 Garamond,Georgia,serif;margin-bottom:8px;">An offering of what has held my attention.</div>' +
      notes +
      invite +
      '<div style="color:' + COLORS.secondary + ';font:10px Arial,sans-serif;letter-spacing:1.5px;margin:36px 0 10px;">06 — A QUESTION</div>' +
      '<div style="color:' + COLORS.text + ';font:25px/1.35 Garamond,Georgia,serif;margin-bottom:30px;">' + esc(val("question")) + '</div>' +
      '<div style="text-align:center;margin:20px 0 12px;"><img src="' + LOOP_ASSET_URL + '" alt="The Loop" width="64" style="display:inline-block;width:64px;height:auto;border:0;"></div>' +
      '<div style="text-align:center;color:' + COLORS.secondary + ';font:italic 15px/1.4 Garamond,Georgia,serif;">Pleasure is the desire to return.</div>' +
    '</td></tr>' +

    '<tr><td style="padding:16px 36px 30px;border-top:1px solid ' + COLORS.rule + ';color:' + COLORS.secondary + ';font:10px Arial,sans-serif;letter-spacing:1px;">THE PLEASURE DISPATCH · BY FLRSGLOBAL</td></tr>' +
    '</table></td></tr></table></div>';
}

/* =========================================================
   BUILD IN OUTLOOK
   Uses setSelectedDataAsync instead of body.setAsync.
========================================================= */

function updateBuildProgress(current, message) {
  setStatus(current + " / 4  " + message);
}

function getAsyncError(result) {
  if (result && result.error && result.error.message) return result.error.message;
  if (result && result.error && result.error.name) return result.error.name;
  return "Unknown Outlook error.";
}

function buildSubject() {
  return "The Pleasure Dispatch — " +
    (val("edition") || "No. 001") + ": " +
    (val("title") || "A Note on Pleasure");
}

function buildInOutlook() {
  updateBuildProgress(0, "Starting…");

  if (typeof Office === "undefined") {
    setStatus("Build failed — Office.js is unavailable.");
    return;
  }

  const item = Office.context && Office.context.mailbox && Office.context.mailbox.item;

  if (!item) {
    setStatus("Build failed — open The Pleasure Dispatch from a new Outlook message.");
    return;
  }

  if (!item.body || typeof item.body.setSelectedDataAsync !== "function") {
    setStatus("Build failed — this Outlook compose surface does not support HTML insertion here.");
    return;
  }

  updateBuildProgress(1, "Generating newsletter…");

  let html;
  try {
    html = buildNewsletterHtml();
  } catch (error) {
    setStatus("Build failed — " + error.message);
    console.error(error);
    return;
  }

  updateBuildProgress(2, "Checking image sources…");

  const modules = collectImageBlocks();
  let imageCount = modules.reduce(function (sum, block) {
    return sum + block.items.length;
  }, 0);
  if (val("hero1Url")) imageCount++;
  if (val("hero2Url")) imageCount++;

  updateBuildProgress(
    2,
    imageCount + " image" + (imageCount === 1 ? "" : "s") + " ready…"
  );

  const subject = buildSubject();

  updateBuildProgress(3, "Preparing Outlook message…");

  let finished = false;
  const timeout = setTimeout(function () {
    if (finished) return;
    finished = true;
    setStatus("Build timed out — Outlook did not respond after 60 seconds.");
  }, BUILD_TIMEOUT_MS);

  /* Set subject first. */
  if (item.subject && typeof item.subject.setAsync === "function") {
    item.subject.setAsync(subject, function (subjectResult) {
      if (finished) return;

      if (
        subjectResult &&
        subjectResult.status !== Office.AsyncResultStatus.Succeeded
      ) {
        finished = true;
        clearTimeout(timeout);
        setStatus("Build failed setting subject: " + getAsyncError(subjectResult));
        return;
      }

      insertNewsletterAtCursor();
    });
  } else {
    insertNewsletterAtCursor();
  }

  function insertNewsletterAtCursor() {
    updateBuildProgress(3, "Inserting newsletter into Outlook…");

    item.body.setSelectedDataAsync(
      html,
      {
        coercionType: Office.CoercionType.Html
      },
      function (result) {
        if (finished) return;

        finished = true;
        clearTimeout(timeout);

        if (
          result &&
          result.status === Office.AsyncResultStatus.Succeeded
        ) {
          updateBuildProgress(4, "✓ Complete — Dispatch built in Outlook.");
        } else {
          setStatus(
            "Build failed inserting newsletter: " +
            getAsyncError(result)
          );
          console.error("setSelectedDataAsync result:", result);
        }
      }
    );
  }
}

/* =========================================================
   PREVIEW
========================================================= */

function previewNewsletter() {
  setStatus("Opening Dispatch preview…");

  let html;
  try {
    html = buildNewsletterHtml();
  } catch (error) {
    setStatus("Preview error: " + error.message);
    return;
  }

  const win = window.open("", "_blank");
  if (!win) {
    setStatus("Preview was blocked by the browser.");
    return;
  }

  win.document.open();
  win.document.write(
    "<!doctype html><html><head><meta charset='utf-8'><title>The Pleasure Dispatch — FLRSGLOBAL</title></head>" +
    "<body style='margin:0;background:" + COLORS.background + ";'>" +
    html +
    "</body></html>"
  );
  win.document.close();
  setStatus("✓ Preview opened.");
}
