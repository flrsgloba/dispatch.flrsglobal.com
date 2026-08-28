let blockCounter = 0;
let pleasureCounter = 0;

Office.onReady(function () {
  const insertBtn = document.getElementById("insertBtn");
  const previewBtn = document.getElementById("previewBtn");
  const addImageBlockBtn = document.getElementById("addImageBlock");
  const addPleasureBtn = document.getElementById("addPleasure");

  if (insertBtn) {
    insertBtn.addEventListener("click", buildInOutlook);
  }

  if (previewBtn) {
    previewBtn.addEventListener("click", preview);
  }

  if (addImageBlockBtn) {
    addImageBlockBtn.addEventListener("click", function () {
      addImageBlock();
    });
  }

  if (addPleasureBtn) {
    addPleasureBtn.addEventListener("click", function () {
      addPleasureRow();
    });
  }

  setupHero("hero1");
  setupHero("hero2");

  document.querySelectorAll("[data-url]").forEach(function (button) {
    button.addEventListener("click", function () {
      const key = button.dataset.url;
      const input = document.getElementById(key + "Url");

      if (input) {
        input.focus();
        setStatus("Paste the image URL into the field above.");
      }
    });
  });

  document.querySelectorAll("[data-drive]").forEach(function (button) {
    button.addEventListener("click", function () {
      openDriveChooser(button.dataset.drive);
    });
  });

  addPleasureRow("Coffee", "");
  addPleasureRow("Art", "");
  addPleasureRow("Object", "");

  addImageBlock();
});


/* -----------------------------------------------------------
   BASIC HELPERS
----------------------------------------------------------- */

function value(id) {
  const element = document.getElementById(id);

  if (!element) {
    return "";
  }

  return element.value.trim();
}


function escapeHtml(text) {
  return String(text || "").replace(/[&<>"']/g, function (character) {
    const entities = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;"
    };

    return entities[character];
  });
}


function escapeAttribute(text) {
  return escapeHtml(text);
}


function paragraph(text) {
  if (!text) {
    return "";
  }

  return (
    '<p style="' +
    'font:17px/1.7 Garamond, Georgia, \\'Times New Roman\\', serif;' +
    'margin:0 0 24px;color:#151515;">' +
    escapeHtml(text).replace(/\n/g, "<br>") +
    "</p>"
  );
}


function setStatus(message) {
  const status = document.getElementById("status");

  if (status) {
    status.textContent = message;
  }
}


/* -----------------------------------------------------------
   HERO IMAGE SYSTEM
----------------------------------------------------------- */

function setupHero(key) {
  const fileInput = document.getElementById(key + "File");
  const urlInput = document.getElementById(key + "Url");

  if (fileInput) {
    fileInput.addEventListener("change", function (event) {
      const file = event.target.files && event.target.files[0];

      if (!file) {
        return;
      }

      handleLocalImage(file, function (dataUrl) {
        if (urlInput) {
          urlInput.value = dataUrl;
        }

        renderHeroPreview(key, dataUrl);

        setStatus(
          "Desktop image loaded into the preview. " +
          "For a production newsletter, the image should be hosted publicly."
        );
      });
    });
  }

  if (urlInput) {
    urlInput.addEventListener("input", function () {
      renderHeroPreview(key, urlInput.value.trim());
    });
  }
}


function handleLocalImage(file, callback) {
  if (!file.type || !file.type.match(/^image\/(jpeg|png|webp)$/)) {
    setStatus("Please choose a JPG, PNG, or WebP image.");
    return;
  }

  const reader = new FileReader();

  reader.onload = function () {
    callback(reader.result);
  };

  reader.onerror = function () {
    setStatus("The image could not be read.");
  };

  reader.readAsDataURL(file);
}


function renderHeroPreview(key, url) {
  const preview = document.getElementById(key + "Preview");

  if (!preview) {
    return;
  }

  if (!url) {
    preview.className = "preview hero-preview empty";
    preview.textContent = "No hero image selected";
    return;
  }

  preview.className = "preview hero-preview";
  preview.innerHTML = "";

  const image = new Image();

  image.onload = function () {
    preview.appendChild(image);
  };

  image.onerror = function () {
    preview.className = "preview hero-preview error";
    preview.textContent =
      "This image could not be loaded. Please use a direct HTTPS image URL.";
  };

  image.src = url;
}


/* -----------------------------------------------------------
   PLEASURE NOTES
----------------------------------------------------------- */

function addPleasureRow(labelValue, noteValue) {
  pleasureCounter++;

  const container = document.getElementById("pleasureRows");

  if (!container) {
    return;
  }

  const row = document.createElement("div");

  row.className = "pleasure-row";
  row.dataset.id = String(pleasureCounter);

  row.innerHTML =
    '<input class="pleasure-label" value="' +
    escapeAttribute(labelValue || "") +
    '" placeholder="Category">' +

    '<input class="pleasure-value" value="' +
    escapeAttribute(noteValue || "") +
    '" placeholder="What has held your attention?">' +

    '<button type="button" aria-label="Remove pleasure note">×</button>';

  const removeButton = row.querySelector("button");

  if (removeButton) {
    removeButton.addEventListener("click", function () {
      row.remove();
    });
  }

  container.appendChild(row);
}


function collectPleasureNotes() {
  const rows = document.querySelectorAll(".pleasure-row");

  return Array.from(rows)
    .map(function (row) {
      const label = row.querySelector(".pleasure-label");
      const valueField = row.querySelector(".pleasure-value");

      return {
        label: label ? label.value.trim() : "",
        value: valueField ? valueField.value.trim() : ""
      };
    })
    .filter(function (item) {
      return item.label || item.value;
    });
}


function buildPleasureNotesHtml(notes) {
  if (!notes.length) {
    return "";
  }

  return (
    '<table role="presentation" cellspacing="0" cellpadding="0" border="0" ' +
    'style="border-collapse:collapse;width:100%;margin:4px 0 30px;">' +
    notes
      .map(function (note) {
        return (
          "<tr>" +
          '<td style="' +
          'width:120px;' +
          'vertical-align:top;' +
          'padding:5px 18px 5px 0;' +
          'font:10px Arial,Helvetica,sans-serif;' +
          'letter-spacing:1px;' +
          'text-transform:uppercase;' +
          'color:#777;">' +
          escapeHtml(note.label) +
          "</td>" +
          '<td style="' +
          'vertical-align:top;' +
          'padding:5px 0;' +
          'font:16px/1.5 Garamond,Georgia,\\'Times New Roman\\',serif;' +
          'color:#151515;">' +
          escapeHtml(note.value) +
          "</td>" +
          "</tr>"
        );
      })
      .join("") +
    "</table>"
  );
}


/* -----------------------------------------------------------
   MODULAR IMAGE BLOCKS
----------------------------------------------------------- */

function addImageBlock() {
  blockCounter++;

  const container = document.getElementById("imageBlocks");

  if (!container) {
    return;
  }

  const block = document.createElement("article");

  block.className = "image-block";
  block.dataset.id = String(blockCounter);
  block.dataset.layout = "full";

  block.innerHTML =
    '<div class="block-top">' +

    '<span class="block-number">Image Module ' +
    blockCounter +
    "</span>" +

    '<select class="layout-select" aria-label="Image layout">' +
    '<option value="full">Full Width</option>' +
    '<option value="two">Two Up</option>' +
    '<option value="three">Three Up</option>' +
    '<option value="four">Four Up</option>' +
    "</select>" +

    "</div>" +

    '<div class="image-items one"></div>' +

    '<div class="block-actions">' +

    '<button type="button" class="move-up">↑ Move Up</button>' +
    '<button type="button" class="move-down">↓ Move Down</button>' +
    '<button type="button" class="duplicate">Duplicate</button>' +
    '<button type="button" class="remove">Remove</button>' +

    "</div>";

  const layoutSelect = block.querySelector(".layout-select");

  if (layoutSelect) {
    layoutSelect.addEventListener("change", function () {
      block.dataset.layout = layoutSelect.value;
      syncImageInputs(block);
    });
  }

  const moveUp = block.querySelector(".move-up");
  const moveDown = block.querySelector(".move-down");
  const duplicate = block.querySelector(".duplicate");
  const remove = block.querySelector(".remove");

  if (moveUp) {
    moveUp.addEventListener("click", function () {
      moveBlock(block, -1);
    });
  }

  if (moveDown) {
    moveDown.addEventListener("click", function () {
      moveBlock(block, 1);
    });
  }

  if (duplicate) {
    duplicate.addEventListener("click", function () {
      duplicateBlock(block);
    });
  }

  if (remove) {
    remove.addEventListener("click", function () {
      block.remove();
    });
  }

  container.appendChild(block);

  syncImageInputs(block);
}


function syncImageInputs(block) {
  const layout = block.dataset.layout || "full";

  let count = 1;

  if (layout === "two") {
    count = 2;
  }

  if (layout === "three") {
    count = 3;
  }

  if (layout === "four") {
    count = 4;
  }

  const wrap = block.querySelector(".image-items");

  if (!wrap) {
    return;
  }

  let className = "one";

  if (layout === "two") {
    className = "two";
  }

  if (layout === "three") {
    className = "three";
  }

  if (layout === "four") {
    className = "four";
  }

  wrap.className = "image-items " + className;

  const currentItems = Array.from(wrap.querySelectorAll(".image-item"));

  while (currentItems.length > count) {
    const item = currentItems.pop();

    if (item) {
      item.remove();
    }
  }

  while (wrap.children.length < count) {
    addImageItem(wrap, wrap.children.length + 1);
  }

  renumberImageItems(wrap);
}


function addImageItem(wrap, number, preset) {
  preset = preset || {};

  const item = document.createElement("div");

  item.className = "image-item";

  item.innerHTML =
    '<label>Image ' +
    number +
    "</label>" +

    '<label class="upload-button">Upload' +
    '<input type="file" accept="image/jpeg,image/png,image/webp">' +
    "</label>" +

    '<button type="button" class="secondary drive-item">Google Drive</button>' +
    '<button type="button" class="secondary url-item">Image URL</button>' +

    '<input class="module-url source-url" value="' +
    escapeAttribute(preset.url || "") +
    '" placeholder="Paste direct image URL">' +

    '<div class="preview module-preview empty">No image selected</div>' +

    '<input class="module-caption" value="' +
    escapeAttribute(preset.caption || "") +
    '" placeholder="Caption">' +

    '<button type="button" class="remove-image">Remove image</button>';

  wireImageItem(item);

  wrap.appendChild(item);

  if (preset.url) {
    renderModulePreview(item, preset.url);
  }
}


function wireImageItem(item) {
  const fileInput = item.querySelector('input[type="file"]');
  const urlInput = item.querySelector(".module-url");
  const driveButton = item.querySelector(".drive-item");
  const urlButton = item.querySelector(".url-item");
  const removeButton = item.querySelector(".remove-image");

  if (fileInput) {
    fileInput.addEventListener("change", function (event) {
      const file = event.target.files && event.target.files[0];

      if (!file) {
        return;
      }

      handleLocalImage(file, function (dataUrl) {
        if (urlInput) {
          urlInput.value = dataUrl;
        }

        renderModulePreview(item, dataUrl);

        setStatus(
          "Desktop image loaded into the preview. " +
          "Google Drive hosting will provide persistent email assets."
        );
      });
    });
  }

  if (urlInput) {
    urlInput.addEventListener("input", function () {
      renderModulePreview(item, urlInput.value.trim());
    });
  }

  if (driveButton) {
    driveButton.addEventListener("click", function () {
      openDriveChooser("module", item);
    });
  }

  if (urlButton) {
    urlButton.addEventListener("click", function () {
      if (urlInput) {
        urlInput.focus();
      }
    });
  }

  if (removeButton) {
    removeButton.addEventListener("click", function () {
      const wrap = item.parentElement;

      if (!wrap) {
        item.remove();
        return;
      }

      if (wrap.children.length === 1) {
        if (urlInput) {
          urlInput.value = "";
        }

        const caption = item.querySelector(".module-caption");

        if (caption) {
          caption.value = "";
        }

        renderModulePreview(item, "");
        return;
      }

      item.remove();
      renumberImageItems(wrap);
    });
  }
}


function renderModulePreview(item, url) {
  const preview = item.querySelector(".module-preview");

  if (!preview) {
    return;
  }

  if (!url) {
    preview.className = "preview module-preview empty";
    preview.textContent = "No image selected";
    return;
  }

  preview.className = "preview module-preview";
  preview.innerHTML = "";

  const image = new Image();

  image.onload = function () {
    preview.appendChild(image);
  };

  image.onerror = function () {
    preview.className = "preview module-preview error";
    preview.textContent =
      "Image could not be loaded. Please use a direct HTTPS image URL.";
  };

  image.src = url;
}


function renumberImageItems(wrap) {
  const items = wrap.querySelectorAll(".image-item");

  Array.from(items).forEach(function (item, index) {
    const label = item.querySelector("label");

    if (label) {
      label.textContent = "Image " + (index + 1);
    }
  });
}


function moveBlock(block, direction) {
  const parent = block.parentElement;

  if (!parent) {
    return;
  }

  const blocks = Array.from(parent.children);
  const currentIndex = blocks.indexOf(block);

  if (currentIndex === -1) {
    return;
  }

  const newIndex = currentIndex + direction;

  if (newIndex < 0 || newIndex >= blocks.length) {
    return;
  }

  if (direction < 0) {
    parent.insertBefore(block, blocks[newIndex]);
  } else {
    parent.insertBefore(blocks[newIndex], block);
  }
}


function duplicateBlock(block) {
  const layout = block.dataset.layout || "full";
  const sourceItems = Array.from(
    block.querySelectorAll(".image-item")
  ).map(function (item) {
    const url = item.querySelector(".module-url");
    const caption = item.querySelector(".module-caption");

    return {
      url: url ? url.value.trim() : "",
      caption: caption ? caption.value.trim() : ""
    };
  });

  blockCounter++;

  const container = document.getElementById("imageBlocks");

  if (!container) {
    return;
  }

  const clone = document.createElement("article");

  clone.className = "image-block";
  clone.dataset.id = String(blockCounter);
  clone.dataset.layout = layout;

  clone.innerHTML =
    '<div class="block-top">' +
    '<span class="block-number">Image Module ' +
    blockCounter +
    "</span>" +
    '<select class="layout-select" aria-label="Image layout">' +
    '<option value="full">Full Width</option>' +
    '<option value="two">Two Up</option>' +
    '<option value="three">Three Up</option>' +
    '<option value="four">Four Up</option>' +
    "</select>" +
    "</div>" +
    '<div class="image-items"></div>' +
    '<div class="block-actions">' +
    '<button type="button" class="move-up">↑ Move Up</button>' +
    '<button type="button" class="move-down">↓ Move Down</button>' +
    '<button type="button" class="duplicate">Duplicate</button>' +
    '<button type="button" class="remove">Remove</button>' +
    "</div>";

  const layoutSelect = clone.querySelector(".layout-select");

  if (layoutSelect) {
    layoutSelect.value = layout;

    layoutSelect.addEventListener("change", function () {
      clone.dataset.layout = layoutSelect.value;
      syncImageInputs(clone);
    });
  }

  const imageWrap = clone.querySelector(".image-items");

  const count = layout === "full"
    ? 1
    : layout === "two"
      ? 2
      : layout === "three"
        ? 3
        : 4;

  imageWrap.className =
    "image-items " + (layout === "full" ? "one" : layout);

  for (let i = 0; i < count; i++) {
    addImageItem(
      imageWrap,
      i + 1,
      sourceItems[i] || {}
    );
  }

  clone.querySelector(".move-up").addEventListener("click", function () {
    moveBlock(clone, -1);
  });

  clone.querySelector(".move-down").addEventListener("click", function () {
    moveBlock(clone, 1);
  });

  clone.querySelector(".duplicate").addEventListener("click", function () {
    duplicateBlock(clone);
  });

  clone.querySelector(".remove").addEventListener("click", function () {
    clone.remove();
  });

  container.insertBefore(clone, block.nextSibling);
}


/* -----------------------------------------------------------
   GOOGLE DRIVE PLACEHOLDER
----------------------------------------------------------- */

function openDriveChooser(target, item) {
  setStatus(
    "Google Drive selection will be connected through your Google Workspace backend. " +
    "For now, use Desktop Upload or Image URL."
  );
}


/* -----------------------------------------------------------
   COLLECT MODULAR IMAGES
----------------------------------------------------------- */

function collectImageBlocks() {
  const blocks = document.querySelectorAll(".image-block");

  return Array.from(blocks)
    .map(function (block) {
      return {
        layout: block.dataset.layout || "full",
        items: Array.from(
          block.querySelectorAll(".image-item")
        )
          .map(function (item) {
            const url = item.querySelector(".module-url");
            const caption = item.querySelector(".module-caption");

            return {
              url: url ? url.value.trim() : "",
              caption: caption ? caption.value.trim() : ""
            };
          })
          .filter(function (item) {
            return item.url || item.caption;
          })
      };
    })
    .filter(function (block) {
      return block.items.length > 0;
    });
}


/* -----------------------------------------------------------
   EMAIL IMAGE MODULES
----------------------------------------------------------- */

function buildImageModuleHtml(block) {
  if (!block || !block.items || !block.items.length) {
    return "";
  }

  const layout = block.layout || "full";

  if (layout === "full") {
    return buildFullWidthImage(block.items[0]);
  }

  if (layout === "two") {
    return buildImageRow(block.items, 2);
  }

  if (layout === "three") {
    return buildImageRow(block.items, 3);
  }

  if (layout === "four") {
    return buildFourImageGrid(block.items);
  }

  return "";
}


function buildFullWidthImage(item) {
  if (!item || !item.url) {
    return "";
  }

  return (
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" ' +
    'style="border-collapse:collapse;margin:30px 0;">' +
    "<tr>" +
    '<td style="padding:0;">' +

    '<img src="' +
    escapeAttribute(item.url) +
    '" alt="" ' +
    'style="display:block;width:100%;height:auto;border:0;">' +

    (item.caption
      ? '<div style="font:12px/1.45 Arial,Helvetica,sans-serif;color:#777;margin-top:7px;">' +
        escapeHtml(item.caption) +
        "</div>"
      : "") +

    "</td>" +
    "</tr>" +
    "</table>"
  );
}


function buildImageRow(items, columns) {
  const usable = items.slice(0, columns);

  const width = Math.floor(100 / columns);

  const cells = usable
    .map(function (item) {
      return (
        '<td width="' +
        width +
        '%" valign="top" style="padding:3px;vertical-align:top;">' +

        (item.url
          ? '<img src="' +
            escapeAttribute(item.url) +
            '" alt="" style="display:block;width:100%;height:auto;border:0;">'
          : "") +

        (item.caption
          ? '<div style="font:11px/1.4 Arial,Helvetica,sans-serif;color:#777;padding-top:6px;">' +
            escapeHtml(item.caption) +
            "</div>"
          : "") +

        "</td>"
      );
    })
    .join("");

  return (
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" ' +
    'style="border-collapse:collapse;margin:28px 0;">' +
    "<tr>" +
    cells +
    "</tr>" +
    "</table>"
  );
}


function buildFourImageGrid(items) {
  const usable = items.slice(0, 4);

  let firstRow = "";
  let secondRow = "";

  if (usable[0]) {
    firstRow += buildGridCell(usable[0]);
  }

  if (usable[1]) {
    firstRow += buildGridCell(usable[1]);
  }

  if (usable[2]) {
    secondRow += buildGridCell(usable[2]);
  }

  if (usable[3]) {
    secondRow += buildGridCell(usable[3]);
  }

  return (
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" ' +
    'style="border-collapse:collapse;margin:28px 0;">' +

    "<tr>" +
    firstRow +
    "</tr>" +

    (secondRow
      ? "<tr>" + secondRow + "</tr>"
      : "") +

    "</table>"
  );
}


function buildGridCell(item) {
  return (
    '<td width="50%" valign="top" style="width:50%;padding:3px;vertical-align:top;">' +

    (item.url
      ? '<img src="' +
        escapeAttribute(item.url) +
        '" alt="" style="display:block;width:100%;height:auto;border:0;">'
      : "") +

    (item.caption
      ? '<div style="font:11px/1.4 Arial,Helvetica,sans-serif;color:#777;padding-top:6px;">' +
        escapeHtml(item.caption) +
        "</div>"
      : "") +

    "</td>"
  );
}


/* -----------------------------------------------------------
   BUILD COMPLETE NEWSLETTER
----------------------------------------------------------- */

function buildNewsletterHtml() {
  const edition = value("edition");
  const date = value("date");
  const title = value("title");
  const subtitle = value("subtitle");

  const hero1 = value("hero1Url");
  const hero1Caption = value("hero1Caption");

  const hero2 = value("hero2Url");
  const hero2Caption = value("hero2Caption");

  const reflection = value("reflection");
  const workText = value("workText");
  const studioText = value("studioText");

  const inviteTitle = value("inviteTitle");
  const inviteText = value("inviteText");

  const ctaLabel = value("ctaLabel") || "INQUIRE";
  const ctaUrl = value("ctaUrl");

  const question = value("question");

  const pleasureNotes = collectPleasureNotes();
  const imageBlocks = collectImageBlocks();

  let hero1Html = "";

  if (hero1) {
    hero1Html =
      '<img src="' +
      escapeAttribute(hero1) +
      '" alt="" ' +
      'style="display:block;width:100%;height:auto;border:0;margin:0 0 10px;">' +

      (hero1Caption
        ? '<div style="font:12px/1.45 Arial,Helvetica,sans-serif;color:#777;margin:0 0 32px;">' +
          escapeHtml(hero1Caption) +
          "</div>"
        : "");
  }

  let hero2Html = "";

  if (hero2) {
    hero2Html =
      '<img src="' +
      escapeAttribute(hero2) +
      '" alt="" ' +
      'style="display:block;width:100%;height:auto;border:0;margin:0 0 10px;">' +

      (hero2Caption
        ? '<div style="font:12px/1.45 Arial,Helvetica,sans-serif;color:#777;margin:0 0 32px;">' +
          escapeHtml(hero2Caption) +
          "</div>"
        : "");
  }

  let invitationHtml = "";

  if (inviteTitle || inviteText || ctaUrl) {
    invitationHtml =
      '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:40px 0 11px;">' +
      "05 — AN INVITATION" +
      "</div>" +

      (inviteTitle
        ? '<div style="font:27px/1.15 Garamond,Georgia,\\'Times New Roman\\',serif;margin:0 0 10px;">' +
          escapeHtml(inviteTitle) +
          "</div>"
        : "") +

      paragraph(inviteText) +

      (ctaUrl
        ? '<div style="padding:0 0 25px;">' +
          '<a href="' +
          escapeAttribute(ctaUrl) +
          '" style="' +
          'display:inline-block;' +
          'background:#151515;' +
          'color:#fff;' +
          'text-decoration:none;' +
          'padding:12px 18px;' +
          'font:10px Arial,Helvetica,sans-serif;' +
          'letter-spacing:1.2px;' +
          '">' +
          escapeHtml(ctaLabel) +
          "</a>" +
          "</div>"
        : "");
  }

  const dateLine = [
    edition,
    date,
    title
  ]
    .filter(Boolean)
    .map(escapeHtml)
    .join(" · ");

  return (
    '<div style="margin:0;padding:0;background:#f4f0e8;color:#151515;">' +

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" ' +
    'style="border-collapse:collapse;background:#f4f0e8;">' +

    "<tr>" +

    '<td align="center" style="padding:28px 12px;">' +

    '<table role="presentation" width="680" cellspacing="0" cellpadding="0" border="0" ' +
    'style="border-collapse:collapse;width:100%;max-width:680px;background:#fffdf8;">' +

    "<tr>" +

    '<td style="padding:42px 42px 20px;">' +

    '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:2px;">' +
    "FLRS GLOBAL" +
    "</div>" +

    '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.4px;color:#777;margin-top:8px;">' +
    "FROM THE STUDIO OF FREDDIE L. RANKIN II" +
    "</div>" +

    '<h1 style="font:400 50px/0.96 Garamond,Georgia,\\'Times New Roman\\',serif;margin:20px 0 10px;">' +
    "The Pleasure Dispatch" +
    "</h1>" +

    '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.3px;color:#777;border-bottom:1px solid #151515;padding-bottom:20px;">' +
    dateLine +
    "</div>" +

    (subtitle
      ? '<div style="font:18px/1.45 Garamond,Georgia,\\'Times New Roman\\',serif;margin-top:20px;">' +
        escapeHtml(subtitle) +
        "</div>"
      : "") +

    "</td>" +
    "</tr>" +

    '<tr><td style="padding:0 42px;">' +

    '<div style="text-align:center;font:27px Garamond,Georgia,serif;margin:0 0 20px;">◒</div>' +

    hero1Html +

    '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:30px 0 11px;">' +
    "01 — A REFLECTION" +
    "</div>" +

    paragraph(reflection) +

    '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +
    "02 — THE WORK" +
    "</div>" +

    paragraph(workText) +

    imageBlocks
      .map(function (block) {
        return buildImageModuleHtml(block);
      })
      .join("") +

    '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +
    "03 — STUDIO NOTES" +
    "</div>" +

    paragraph(studioText) +

    hero2Html +

    '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +
    "04 — PLEASURE NOTES" +
    "</div>" +

    '<div style="font:19px/1.4 Garamond,Georgia,\\'Times New Roman\\',serif;margin:0 0 8px;">' +
    "An offering of what has held my attention." +
    "</div>" +

    buildPleasureNotesHtml(pleasureNotes) +

    invitationHtml +

    '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +
    "06 — A QUESTION" +
    "</div>" +

    '<div style="font:25px/1.35 Garamond,Georgia,\\'Times New Roman\\',serif;margin:0 0 36px;">' +
    escapeHtml(question) +
    "</div>" +

    '<div style="text-align:center;font:27px Garamond,Georgia,serif;margin:20px 0 32px;">◒</div>' +

    "</td></tr>" +

    '<tr><td style="padding:18px 42px 34px;border-top:1px solid #151515;">' +

    '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.1px;color:#777;">' +
    "THE PLEASURE DISPATCH · BY FLRS GLOBAL" +
    "</div>" +

    "</td></tr>" +

    "</table>" +

    "</td>" +

    "</tr>" +

    "</table>" +

    "</div>"
  );
}


/* -----------------------------------------------------------
   OUTLOOK
----------------------------------------------------------- */

function buildSubject() {
  const edition = value("edition") || "No. 001";
  const title = value("title") || "A Note on Pleasure";

  return (
    "The Pleasure Dispatch — " +
    edition +
    ": " +
    title
  );
}


function buildInOutlook() {
  setStatus("Building The Pleasure Dispatch in Outlook…");

  if (
    typeof Office === "undefined" ||
    !Office.context ||
    !Office.context.mailbox ||
    !Office.context.mailbox.item
  ) {
    setStatus(
      "Open The Pleasure Dispatch from a new Outlook message. " +
      "The composer is not currently connected to an Outlook message."
    );

    return;
  }

  const item = Office.context.mailbox.item;

  if (!item.body || typeof item.body.setAsync !== "function") {
    setStatus(
      "Outlook did not provide access to the message body. " +
      "Please open a new message and try again."
    );

    return;
  }

  const html = buildNewsletterHtml();
  const subject = buildSubject();

  item.subject.setAsync(subject, function (subjectResult) {
    if (
      subjectResult &&
      subjectResult.status !== Office.AsyncResultStatus.Succeeded
    ) {
      setStatus(
        "The subject could not be updated: " +
        (
          subjectResult.error
            ? subjectResult.error.message
            : "Unknown Outlook error."
        )
      );

      return;
    }

    item.body.setAsync(
      html,
      {
        coercionType: Office.CoercionType.Html
      },
      function (bodyResult) {
        if (
          bodyResult &&
          bodyResult.status === Office.AsyncResultStatus.Succeeded
        ) {
          setStatus(
            "Dispatch built in Outlook. Review the email before sending."
          );
        } else {
          setStatus(
            "The Dispatch could not be inserted into Outlook: " +
            (
              bodyResult &&
              bodyResult.error
                ? bodyResult.error.message
                : "Unknown Outlook error."
            )
          );
        }
      }
    );
  });
}


/* -----------------------------------------------------------
   PREVIEW
----------------------------------------------------------- */

function preview() {
  const html = buildNewsletterHtml();

  const previewWindow = window.open("", "_blank");

  if (!previewWindow) {
    setStatus(
      "Preview was blocked. Allow pop-ups for this Outlook add-in."
    );

    return;
  }

  previewWindow.document.open();

  previewWindow.document.write(
    "<!doctype html>" +
    '<html><head><meta charset="utf-8">' +
    "<title>The Pleasure Dispatch Preview</title>" +
    "</head><body style='margin:0;'>" +
    html +
    "</body></html>"
  );

  previewWindow.document.close();
}
