/* =========================================================
   THE PLEASURE DISPATCH
   v1.3 — Consolidated Taskpane
========================================================= */

const DRIVE_API_URL =
  "https://script.google.com/macros/s/AKfycbyTLvYbe1O_BbzsH09UMSZdbY9_XZXga-TbSPkR3UclT3Qhlaj7gy5yhXPA_UpE6Fu7tw/exec";

const IMAGE_MAX_WIDTH = 1800;
const IMAGE_MAX_HEIGHT = 1800;
const IMAGE_QUALITY = 0.82;
const MAX_SOURCE_IMAGE_MB = 40;
const DRIVE_IMAGE_WIDTH = 1800;

let blockCounter = 0;
let pleasureCounter = 0;
let initialized = false;


/* =========================================================
   OFFICE INITIALIZATION
========================================================= */

Office.onReady(function () {

  if (initialized) {
    return;
  }

  initialized = true;

  initializeDispatch();

});


function initializeDispatch() {

  bindStaticControls();
  bindDelegatedControls();

  setupHero("hero1");
  setupHero("hero2");

  /*
   * Only create defaults if the editor is empty.
   */
  const pleasureRows =
    document.getElementById("pleasureRows");

  if (
    pleasureRows &&
    pleasureRows.children.length === 0
  ) {

    addPleasureRow("Coffee", "");
    addPleasureRow("Art", "");
    addPleasureRow("Object", "");

  }


  const imageBlocks =
    document.getElementById("imageBlocks");

  if (
    imageBlocks &&
    imageBlocks.children.length === 0
  ) {

    addImageBlock();

  }


  setStatus(
    "The Pleasure Dispatch is ready."
  );

}


/* =========================================================
   STATIC CONTROLS
========================================================= */

function bindStaticControls() {

  const preview =
    document.getElementById(
      "previewBtn"
    );

  const build =
    document.getElementById(
      "insertBtn"
    );

  const addImage =
    document.getElementById(
      "addImageBlock"
    );

  const addPleasure =
    document.getElementById(
      "addPleasure"
    );


  if (preview) {

    preview.addEventListener(
      "click",
      function (event) {

        event.preventDefault();

        previewNewsletter();

      }
    );

  }


  if (build) {

    build.addEventListener(
      "click",
      function (event) {

        event.preventDefault();

        buildInOutlook();

      }
    );

  }


  if (addImage) {

    addImage.addEventListener(
      "click",
      function (event) {

        event.preventDefault();

        addImageBlock();

      }
    );

  }


  if (addPleasure) {

    addPleasure.addEventListener(
      "click",
      function (event) {

        event.preventDefault();

        addPleasureRow(
          "",
          ""
        );

        setStatus(
          "Pleasure Note added."
        );

      }
    );

  }

}


/* =========================================================
   DELEGATED CONTROLS
========================================================= */

function bindDelegatedControls() {

  /*
   * BUTTONS
   */
  document.addEventListener(
    "click",
    function (event) {

      const button =
        event.target.closest("button");

      if (!button) {
        return;
      }


      /*
       * Static controls.
       */
      if (
        button.id === "previewBtn" ||
        button.id === "insertBtn" ||
        button.id === "addImageBlock" ||
        button.id === "addPleasure"
      ) {

        return;

      }


      event.preventDefault();


      /*
       * MOVE UP
       */
      if (
        button.classList.contains(
          "move-up"
        )
      ) {

        const block =
          button.closest(
            ".image-block"
          );

        if (block) {

          moveImageBlock(
            block,
            -1
          );

        }

        return;
      }


      /*
       * MOVE DOWN
       */
      if (
        button.classList.contains(
          "move-down"
        )
      ) {

        const block =
          button.closest(
            ".image-block"
          );

        if (block) {

          moveImageBlock(
            block,
            1
          );

        }

        return;
      }


      /*
       * DUPLICATE
       */
      if (
        button.classList.contains(
          "duplicate"
        )
      ) {

        const block =
          button.closest(
            ".image-block"
          );

        if (block) {

          duplicateBlock(
            block
          );

        }

        return;
      }


      /*
       * REMOVE MODULE
       */
      if (
        button.classList.contains(
          "remove"
        )
      ) {

        const block =
          button.closest(
            ".image-block"
          );

        if (block) {

          block.remove();

          renumberImageBlocks();

          setStatus(
            "Image Module removed."
          );

        }

        return;
      }


      /*
       * REMOVE IMAGE
       */
      if (
        button.classList.contains(
          "remove-image"
        )
      ) {

        removeImageItem(
          button.closest(
            ".image-item"
          )
        );

        return;
      }


      /*
       * FOCUS IMAGE URL
       */
      if (
        button.classList.contains(
          "url-item"
        )
      ) {

        const item =
          button.closest(
            ".image-item"
          );

        if (item) {

          const input =
            item.querySelector(
              ".module-url"
            );

          if (input) {

            input.focus();

            setStatus(
              "Paste a direct HTTPS image URL."
            );

          }

        }

        return;
      }

    }
  );


  /*
   * LAYOUTS + FILE UPLOADS
   */
  document.addEventListener(
    "change",
    function (event) {

      const select =
        event.target.closest(
          ".layout-select"
        );


      if (select) {

        const block =
          select.closest(
            ".image-block"
          );

        if (!block) {
          return;
        }


        block.dataset.layout =
          select.value;


        syncImageInputs(
          block
        );


        const selected =
          select.options[
            select.selectedIndex
          ];


        setStatus(
          "Layout changed to " +
          selected.text +
          "."
        );


        return;
      }


      /*
       * Modular image upload.
       */
      if (
        event.target.matches(
          ".module-file-input"
        )
      ) {

        const item =
          event.target.closest(
            ".image-item"
          );

        const file =
          event.target.files &&
          event.target.files[0];


        if (
          item &&
          file
        ) {

          handleModuleUpload(
            item,
            file
          );

        }

        return;
      }

    }
  );


  /*
   * IMAGE URL INPUTS
   */
  document.addEventListener(
    "input",
    function (event) {

      /*
       * Hero 01.
       */
      if (
        event.target.id ===
        "hero1Url"
      ) {

        const preview =
          document.getElementById(
            "hero1Preview"
          );


        if (preview) {

          delete preview.dataset.fullUrl;
          delete preview.dataset.fileId;

        }


        renderHeroPreview(
          "hero1",
          event.target.value.trim()
        );


        return;
      }


      /*
       * Hero 02.
       */
      if (
        event.target.id ===
        "hero2Url"
      ) {

        const preview =
          document.getElementById(
            "hero2Preview"
          );


        if (preview) {

          delete preview.dataset.fullUrl;
          delete preview.dataset.fileId;

        }


        renderHeroPreview(
          "hero2",
          event.target.value.trim()
        );


        return;
      }


      /*
       * Modular image URL.
       */
      if (
        event.target.classList.contains(
          "module-url"
        )
      ) {

        const item =
          event.target.closest(
            ".image-item"
          );


        if (item) {

          delete item.dataset.fullUrl;
          delete item.dataset.fileId;
          delete item.dataset.driveUrl;


          renderModulePreview(
            item,
            event.target.value.trim()
          );

        }

      }

    }
  );

}


/* =========================================================
   BASIC HELPERS
========================================================= */

function value(id) {

  const element =
    document.getElementById(
      id
    );


  if (!element) {
    return "";
  }


  return element.value.trim();

}


function escapeHtml(text) {

  return String(
    text || ""
  ).replace(
    /[&<>"']/g,
    function (character) {

      return {

        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#039;"

      }[character];

    }
  );

}


function escapeAttribute(text) {

  return escapeHtml(
    text
  );

}


function paragraph(text) {

  if (!text) {
    return "";
  }


  return (

    '<p style="' +
      "font:17px/1.7 Garamond,Georgia,Times New Roman,serif;" +
      "margin:0 0 24px;" +
      "color:#151515;" +
    '">' +

      escapeHtml(
        text
      ).replace(
        /\n/g,
        "<br>"
      ) +

    "</p>"

  );

}


function setStatus(
  message
) {

  const status =
    document.getElementById(
      "status"
    );


  if (status) {

    status.textContent =
      message;

  }

}


/* =========================================================
   DRIVE URLS
========================================================= */

function buildDriveImageUrl(
  fileId
) {

  return (
    "https://drive.google.com/thumbnail?id=" +
    encodeURIComponent(
      fileId
    ) +
    "&sz=w" +
    DRIVE_IMAGE_WIDTH
  );

}


function buildDriveClickUrl(
  fileId
) {

  return (
    "https://drive.google.com/file/d/" +
    encodeURIComponent(
      fileId
    ) +
    "/view"
  );

}


/* =========================================================
   IMAGE COMPRESSION
========================================================= */

function compressImage(
  file,
  callback
) {

  if (!file) {

    callback(
      null,
      new Error(
        "No image selected."
      )
    );

    return;
  }


  const sizeMB =
    file.size /
    (1024 * 1024);


  if (
    sizeMB >
    MAX_SOURCE_IMAGE_MB
  ) {

    callback(
      null,
      new Error(
        "That image is larger than " +
        MAX_SOURCE_IMAGE_MB +
        " MB."
      )
    );

    return;
  }


  if (
    !file.type ||
    !file.type.match(
      /^image\/(jpeg|jpg|png|webp)$/
    )
  ) {

    callback(
      null,
      new Error(
        "Please choose a JPG, PNG, or WebP image."
      )
    );

    return;
  }


  const reader =
    new FileReader();


  reader.onload =
    function () {

      const image =
        new Image();


      image.onload =
        function () {

          let width =
            image.naturalWidth;


          let height =
            image.naturalHeight;


          const scale =
            Math.min(
              1,
              IMAGE_MAX_WIDTH /
                width,
              IMAGE_MAX_HEIGHT /
                height
            );


          width =
            Math.round(
              width *
              scale
            );


          height =
            Math.round(
              height *
              scale
            );


          const canvas =
            document.createElement(
              "canvas"
            );


          canvas.width =
            width;


          canvas.height =
            height;


          const context =
            canvas.getContext(
              "2d"
            );


          context.fillStyle =
            "#ffffff";


          context.fillRect(
            0,
            0,
            width,
            height
          );


          context.drawImage(
            image,
            0,
            0,
            width,
            height
          );


          canvas.toBlob(
            function (blob) {

              if (!blob) {

                callback(
                  null,
                  new Error(
                    "Image compression failed."
                  )
                );

                return;
              }


              blobToDataUrl(
                blob,
                function (
                  dataUrl
                ) {

                  callback(
                    {
                      blob:
                        blob,

                      dataUrl:
                        dataUrl,

                      width:
                        width,

                      height:
                        height
                    },

                    null
                  );

                }
              );

            },

            "image/jpeg",

            IMAGE_QUALITY
          );

        };


      image.onerror =
        function () {

          callback(
            null,
            new Error(
              "The image could not be decoded."
            )
          );

        };


      image.src =
        reader.result;

    };


  reader.onerror =
    function () {

      callback(
        null,
        new Error(
          "The image could not be read."
        )
      );

    };


  reader.readAsDataURL(
    file
  );

}


function blobToDataUrl(
  blob,
  callback
) {

  const reader =
    new FileReader();


  reader.onload =
    function () {

      callback(
        reader.result
      );

    };


  reader.onerror =
    function () {

      callback(
        null
      );

    };


  reader.readAsDataURL(
    blob
  );

}


/* =========================================================
   GOOGLE DRIVE UPLOAD
========================================================= */

async function uploadToDrive(
  dataUrl,
  originalFileName
) {

  setStatus(
    "Saving image to Google Drive…"
  );


  const payload = {

    action:
      "upload",

    fileName:
      "PD_" +
      getEditionSafeName() +
      "_" +
      Date.now() +
      "_" +
      cleanFileName(
        originalFileName
      ),

    mimeType:
      "image/jpeg",

    fileContent:
      stripDataUrlPrefix(
        dataUrl
      )

  };


  let response;


  try {

    response =
      await fetch(
        DRIVE_API_URL,
        {

          method:
            "POST",

          headers:
            {
              "Content-Type":
                "text/plain;charset=utf-8"
            },

          body:
            JSON.stringify(
              payload
            )

        }
      );

  } catch (error) {

    throw new Error(
      "Could not connect to Google Drive: " +
      error.message
    );

  }


  if (
    !response.ok
  ) {

    throw new Error(
      "Google Drive returned HTTP " +
      response.status +
      "."
    );

  }


  let result;


  try {

    result =
      await response.json();

  } catch (error) {

    throw new Error(
      "Google Drive returned an unreadable response."
    );

  }


  if (
    !result ||
    result.status !==
      "created"
  ) {

    throw new Error(
      result &&
      result.message
        ? result.message
        : "Google Drive did not create the image."
    );

  }


  /*
   * Build our own reliable URLs from fileId.
   */
  result.imageUrl =
    buildDriveImageUrl(
      result.fileId
    );


  result.driveUrl =
    buildDriveClickUrl(
      result.fileId
    );


  return result;

}


function stripDataUrlPrefix(
  dataUrl
) {

  const comma =
    dataUrl.indexOf(
      ","
    );


  if (
    comma ===
    -1
  ) {

    return dataUrl;

  }


  return dataUrl.substring(
    comma + 1
  );

}


function cleanFileName(
  name
) {

  return String(
    name ||
    "image"
  )

  .replace(
    /[\/\\:*?"<>|#%{}[\]]/g,
    "_"
  )

  .replace(
    /\s+/g,
    "_"
  )

  .replace(
    /\.[^.]+$/,
    ""
  )

  .substring(
    0,
    80
  );

}


function getEditionSafeName() {

  return (
    value("edition") ||
    "001"
  )
  .replace(
    /[^a-zA-Z0-9_-]/g,
    "_"
  );

}


/* =========================================================
   HERO IMAGE
========================================================= */

function setupHero(
  key
) {

  const fileInput =
    document.getElementById(
      key +
      "File"
    );


  const urlInput =
    document.getElementById(
      key +
      "Url"
    );


  if (fileInput) {

    fileInput.addEventListener(
      "change",
      function () {

        const file =
          fileInput.files &&
          fileInput.files[0];


        if (!file) {
          return;
        }


        handleHeroUpload(
          key,
          file
        );

      }
    );

  }


  if (urlInput) {

    urlInput.addEventListener(
      "input",
      function () {

        renderHeroPreview(
          key,
          urlInput.value.trim()
        );

      }
    );

  }

}


function handleHeroUpload(
  key,
  file
) {

  compressImage(
    file,
    async function (
      result,
      error
    ) {

      if (error) {

        setStatus(
          error.message
        );

        return;
      }


      const input =
        document.getElementById(
          key +
          "Url"
        );


      const preview =
        document.getElementById(
          key +
          "Preview"
        );


      /*
       * Show immediately.
       */
      input.value =
        result.dataUrl;


      renderHeroPreview(
        key,
        result.dataUrl
      );


      try {

        const drive =
          await uploadToDrive(
            result.dataUrl,
            file.name
          );


        /*
         * Permanent Drive image URL.
         */
        input.value =
          drive.imageUrl;


        if (preview) {

          preview.dataset.fullUrl =
            drive.driveUrl;

          preview.dataset.fileId =
            drive.fileId;

        }


        renderHeroPreview(
          key,
          drive.imageUrl
        );


        setStatus(
          key === "hero1"
            ? "Hero Image 01 saved to Drive."
            : "Hero Image 02 saved to Drive."
        );


      } catch (uploadError) {

        setStatus(
          "Preview ready. Drive upload failed: " +
          uploadError.message
        );

      }

    }
  );

}


function renderHeroPreview(
  key,
  url
) {

  const preview =
    document.getElementById(
      key +
      "Preview"
    );


  if (!preview) {
    return;
  }


  if (!url) {

    preview.className =
      "preview hero-preview empty";

    preview.textContent =
      "No hero image selected";

    return;
  }


  preview.className =
    "preview hero-preview";


  preview.innerHTML =
    "";


  const image =
    new Image();


  image.onload =
    function () {

      preview.appendChild(
        image
      );

    };


  image.onerror =
    function () {

      preview.className =
        "preview hero-preview error";

      preview.textContent =
        "Image could not be loaded.";

    };


  image.src =
    url;

}


/* =========================================================
   MODULAR IMAGE UPLOAD
========================================================= */

async function handleModuleUpload(
  item,
  file
) {

  if (!item || !file) {

    setStatus(
      "No modular image selected."
    );

    return;
  }


  setStatus(
    "Optimizing modular image…"
  );


  compressImage(
    file,
    async function (
      result,
      error
    ) {

      if (error) {

        setStatus(
          error.message
        );

        return;
      }


      const urlInput =
        item.querySelector(
          ".module-url"
        );


      if (!urlInput) {

        setStatus(
          "The modular image field could not be found."
        );

        return;
      }


      /*
       * Immediate local preview.
       */
      urlInput.value =
        result.dataUrl;


      renderModulePreview(
        item,
        result.dataUrl
      );


      try {

        const drive =
          await uploadToDrive(
            result.dataUrl,
            file.name
          );


        /*
         * Store both URLs.
         */
        item.dataset.fileId =
          drive.fileId;


        item.dataset.imageUrl =
          drive.imageUrl;


        item.dataset.fullUrl =
          drive.driveUrl;


        item.dataset.driveUrl =
          drive.driveUrl;


        /*
         * Replace temporary Base64
         * with permanent Drive thumbnail.
         */
        urlInput.value =
          drive.imageUrl;


        renderModulePreview(
          item,
          drive.imageUrl
        );


        setStatus(
          "✓ Modular image saved to Drive."
        );


      } catch (uploadError) {

        setStatus(
          "Preview ready. Drive upload failed: " +
          uploadError.message
        );


        console.error(
          uploadError
        );

      }

    }
  );

}


/* =========================================================
   IMAGE MODULES
========================================================= */

function addImageBlock() {

  const container =
    document.getElementById(
      "imageBlocks"
    );


  if (!container) {

    setStatus(
      "Image module container not found."
    );

    return;
  }


  blockCounter++;


  const block =
    document.createElement(
      "article"
    );


  block.className =
    "image-block";


  block.dataset.id =
    String(
      blockCounter
    );


  block.dataset.layout =
    "full";


  block.innerHTML =

    '<div class="block-top">' +

      '<span class="block-number">' +
        "Image Module " +
        blockCounter +
      "</span>" +

      '<select class="layout-select" aria-label="Image layout">' +

        '<option value="full">' +
          "Full Width" +
        "</option>" +

        '<option value="two">' +
          "Two Up" +
        "</option>" +

        '<option value="three">' +
          "Three Up" +
        "</option>" +

        '<option value="four">' +
          "Four Up" +
        "</option>" +

      "</select>" +

    "</div>" +

    '<div class="image-items one"></div>' +

    '<div class="block-actions">' +

      '<button type="button" class="move-up">' +
        "↑ Move Up" +
      "</button>" +

      '<button type="button" class="move-down">' +
        "↓ Move Down" +
      "</button>" +

      '<button type="button" class="duplicate">' +
        "Duplicate" +
      "</button>" +

      '<button type="button" class="remove">' +
        "Remove" +
      "</button>" +

    "</div>";


  container.appendChild(
    block
  );


  syncImageInputs(
    block
  );


  renumberImageBlocks();


  setStatus(
    "Image Module added."
  );

}


/*
 * THIS FUNCTION WAS MISSING BEFORE.
 */
function getImageBlocks() {

  const container =
    document.getElementById(
      "imageBlocks"
    );


  if (!container) {
    return [];
  }


  return Array.from(
    container.querySelectorAll(
      ":scope > .image-block"
    )
  );

}


function renumberImageBlocks() {

  const blocks =
    getImageBlocks();


  blocks.forEach(
    function (
      block,
      index
    ) {

      const label =
        block.querySelector(
          ".block-number"
        );


      if (label) {

        label.textContent =
          "Image Module " +
          (index + 1);

      }


      block.dataset.position =
        String(
          index + 1
        );

    }
  );

}


/* =========================================================
   MOVE MODULES
========================================================= */

function moveImageBlock(
  block,
  direction
) {

  const container =
    document.getElementById(
      "imageBlocks"
    );


  const blocks =
    getImageBlocks();


  if (
    !container ||
    !block
  ) {

    setStatus(
      "Unable to move image module."
    );

    return;

  }


  const index =
    blocks.indexOf(
      block
    );


  if (index === -1) {

    setStatus(
      "Unable to find image module."
    );

    return;

  }


  const targetIndex =
    index +
    direction;


  if (
    targetIndex < 0 ||
    targetIndex >=
      blocks.length
  ) {

    setStatus(
      direction < 0
        ? "Already at the top."
        : "Already at the bottom."
    );

    return;

  }


  const target =
    blocks[targetIndex];


  if (
    direction < 0
  ) {

    container.insertBefore(
      block,
      target
    );

  } else {

    container.insertBefore(
      block,
      target.nextSibling
    );

  }


  renumberImageBlocks();


  setStatus(
    direction < 0
      ? "Image Module moved up."
      : "Image Module moved down."
  );

}


/* =========================================================
   IMAGE LAYOUTS
========================================================= */

function syncImageInputs(
  block
) {

  const layout =
    block.dataset.layout ||
    "full";


  let count =
    1;


  if (
    layout ===
    "two"
  ) {

    count = 2;

  }


  if (
    layout ===
    "three"
  ) {

    count = 3;

  }


  if (
    layout ===
    "four"
  ) {

    count = 4;

  }


  const wrap =
    block.querySelector(
      ".image-items"
    );


  if (!wrap) {
    return;
  }


  wrap.className =
    "image-items " +
    (
      layout === "full"
        ? "one"
        : layout
    );


  while (
    wrap.children.length >
    count
  ) {

    wrap.lastElementChild.remove();

  }


  while (
    wrap.children.length <
    count
  ) {

    addImageItem(
      wrap,
      wrap.children.length + 1,
      {}
    );

  }


  renumberImageItems(
    wrap
  );

}


function addImageItem(
  wrap,
  number,
  preset
) {

  preset =
    preset ||
    {};


  const item =
    document.createElement(
      "div"
    );


  item.className =
    "image-item";


  item.innerHTML =

    '<label>Image ' +
      number +
    "</label>" +

    '<label class="upload-button">' +

      "Upload" +

      '<input type="file" class="module-file-input" accept="image/jpeg,image/png,image/webp">' +

    "</label>" +

    '<button type="button" class="secondary url-item">' +

      "Image URL" +

    "</button>" +

    '<input class="module-url source-url" value="' +

      escapeAttribute(
        preset.url ||
        ""
      ) +

    '" placeholder="Paste direct image URL">' +

    '<div class="preview module-preview empty">' +

      "No image selected" +

    "</div>" +

    '<input class="module-caption" value="' +

      escapeAttribute(
        preset.caption ||
        ""
      ) +

    '" placeholder="Caption">' +

    '<button type="button" class="remove-image">' +

      "Remove image" +

    "</button>";


  wrap.appendChild(
    item
  );


  /*
   * Preserve Drive click destination.
   */
  if (
    preset.clickUrl
  ) {

    item.dataset.fullUrl =
      preset.clickUrl;

  }


  if (
    preset.url
  ) {

    renderModulePreview(
      item,
      preset.url
    );

  }

}


function renumberImageItems(
  wrap
) {

  Array.from(
    wrap.children
  )
  .forEach(
    function (
      item,
      index
    ) {

      const label =
        item.querySelector(
          "label"
        );


      if (label) {

        label.textContent =
          "Image " +
          (index + 1);

      }

    }
  );

}


function removeImageItem(
  item
) {

  if (!item) {
    return;
  }


  const wrap =
    item.parentElement;


  if (!wrap) {

    item.remove();

    return;

  }


  /*
   * Leave one blank slot in a module.
   */
  if (
    wrap.children.length ===
    1
  ) {

    const url =
      item.querySelector(
        ".module-url"
      );


    const caption =
      item.querySelector(
        ".module-caption"
      );


    if (url) {
      url.value = "";
    }


    if (caption) {
      caption.value = "";
    }


    delete item.dataset.fullUrl;
    delete item.dataset.driveUrl;
    delete item.dataset.fileId;


    renderModulePreview(
      item,
      ""
    );


    setStatus(
      "Image removed."
    );


    return;

  }


  item.remove();


  renumberImageItems(
    wrap
  );


  setStatus(
    "Image removed."
  );

}


/* =========================================================
   DUPLICATE MODULE
========================================================= */

function duplicateBlock(
  block
) {

  if (!block) {
    return;
  }


  const layout =
    block.dataset.layout ||
    "full";


  const sourceItems =
    Array.from(
      block.querySelectorAll(
        ".image-item"
      )
    )
    .map(
      function (item) {

        const url =
          item.querySelector(
            ".module-url"
          );


        const caption =
          item.querySelector(
            ".module-caption"
          );


        return {

          url:
            url
              ? url.value.trim()
              : "",

          caption:
            caption
              ? caption.value.trim()
              : "",

          clickUrl:
            item.dataset.fullUrl ||
            item.dataset.driveUrl ||
            ""

        };

      }
    );


  const clone =
    document.createElement(
      "article"
    );


  blockCounter++;


  clone.className =
    "image-block";


  clone.dataset.id =
    String(
      blockCounter
    );


  clone.dataset.layout =
    layout;


  clone.innerHTML =

    '<div class="block-top">' +

      '<span class="block-number">' +
        "Image Module" +
      "</span>" +

      '<select class="layout-select" aria-label="Image layout">' +

        '<option value="full">' +
          "Full Width" +
        "</option>" +

        '<option value="two">' +
          "Two Up" +
        "</option>" +

        '<option value="three">' +
          "Three Up" +
        "</option>" +

        '<option value="four">' +
          "Four Up" +
        "</option>" +

      "</select>" +

    "</div>" +

    '<div class="image-items"></div>' +

    '<div class="block-actions">' +

      '<button type="button" class="move-up">↑ Move Up</button>' +

      '<button type="button" class="move-down">↓ Move Down</button>' +

      '<button type="button" class="duplicate">Duplicate</button>' +

      '<button type="button" class="remove">Remove</button>' +

    "</div>";


  const select =
    clone.querySelector(
      ".layout-select"
    );


  select.value =
    layout;


  const wrap =
    clone.querySelector(
      ".image-items"
    );


  const count =
    layout === "full"
      ? 1
      : layout === "two"
        ? 2
        : layout === "three"
          ? 3
          : 4;


  for (
    let i = 0;
    i < count;
    i++
  ) {

    addImageItem(
      wrap,
      i + 1,
      sourceItems[i] ||
      {}
    );

  }


  const container =
    document.getElementById(
      "imageBlocks"
    );


  container.insertBefore(
    clone,
    block.nextSibling
  );


  renumberImageBlocks();


  setStatus(
    "Image Module duplicated."
  );

}


/* =========================================================
   COLLECT IMAGE DATA
========================================================= */

function collectImageBlocks() {

  return getImageBlocks()
    .map(
      function (block) {

        return {

          layout:
            block.dataset.layout ||
            "full",

          items:
            Array.from(
              block.querySelectorAll(
                ".image-item"
              )
            )
            .map(
              function (item) {

                const url =
                  item.querySelector(
                    ".module-url"
                  );


                const caption =
                  item.querySelector(
                    ".module-caption"
                  );


                return {

                  url:
                    url
                      ? url.value.trim()
                      : "",

                  clickUrl:
                    item.dataset.fullUrl ||
                    item.dataset.driveUrl ||
                    "",

                  caption:
                    caption
                      ? caption.value.trim()
                      : ""

                };

              }
            )
            .filter(
              function (item) {

                return (
                  item.url ||
                  item.caption
                );

              }
            )

        };

      }
    )
    .filter(
      function (block) {

        return (
          block.items.length
        );

      }
    );

}


/* =========================================================
   PLEASURE NOTES
========================================================= */

function collectPleasureNotes() {

  return Array.from(
    document.querySelectorAll(
      ".pleasure-row"
    )
  )
  .map(
    function (row) {

      const label =
        row.querySelector(
          ".pleasure-label"
        );


      const note =
        row.querySelector(
          ".pleasure-value"
        );


      return {

        label:
          label
            ? label.value.trim()
            : "",

        value:
          note
            ? note.value.trim()
            : ""

      };

    }
  )
  .filter(
    function (item) {

      return (
        item.label ||
        item.value
      );

    }
  );

}


function addPleasureRow(
  labelValue,
  noteValue
) {

  const container =
    document.getElementById(
      "pleasureRows"
    );


  if (!container) {
    return;
  }


  pleasureCounter++;


  const row =
    document.createElement(
      "div"
    );


  row.className =
    "pleasure-row";


  row.innerHTML =

    '<input class="pleasure-label" value="' +
      escapeAttribute(
        labelValue ||
        ""
      ) +
    '" placeholder="Category">' +

    '<input class="pleasure-value" value="' +
      escapeAttribute(
        noteValue ||
        ""
      ) +
    '" placeholder="What has held your attention?">' +

    '<button type="button">×</button>';


  row.querySelector(
    "button"
  ).addEventListener(
    "click",
    function () {

      row.remove();

      setStatus(
        "Pleasure Note removed."
      );

    }
  );


  container.appendChild(
    row
  );

}


function buildPleasureNotesHtml(
  notes
) {

  if (!notes.length) {
    return "";
  }


  return (

    '<table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;width:100%;margin:4px 0 30px;">' +

      notes.map(
        function (note) {

          return (

            "<tr>" +

              '<td style="width:120px;vertical-align:top;padding:5px 18px 5px 0;font:10px Arial,Helvetica,sans-serif;letter-spacing:1px;text-transform:uppercase;color:#777;">' +

                escapeHtml(
                  note.label
                ) +

              "</td>" +

              '<td style="vertical-align:top;padding:5px 0;font:16px/1.5 Garamond,Georgia,Times New Roman,serif;color:#151515;">' +

                escapeHtml(
                  note.value
                ) +

              "</td>" +

            "</tr>"

          );

        }
      ).join("") +

    "</table>"

  );

}


/* =========================================================
   EMAIL IMAGE
========================================================= */

function buildEmailImage(
  imageUrl,
  clickUrl,
  caption
) {

  if (!imageUrl) {
    return "";
  }


  const destination =
    clickUrl ||
    imageUrl;


  let html =

    '<a href="' +
      escapeAttribute(
        destination
      ) +
      '" target="_blank" style="text-decoration:none;">' +

      '<img src="' +
        escapeAttribute(
          imageUrl
        ) +
        '" alt="" style="display:block;width:100%;height:auto;border:0;">' +

    "</a>";


  if (caption) {

    html +=

      '<div style="font:12px/1.45 Arial,Helvetica,sans-serif;color:#777;margin-top:7px;">' +

        escapeHtml(
          caption
        ) +

      "</div>";

  }


  return html;

}


function buildFullWidthImage(
  item
) {

  if (
    !item ||
    !item.url
  ) {
    return "";
  }


  return (

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;margin:30px 0;">' +

      "<tr>" +

        '<td style="padding:0;">' +

          buildEmailImage(
            item.url,
            item.clickUrl ||
              item.url,
            item.caption
          ) +

        "</td>" +

      "</tr>" +

    "</table>"

  );

}


function buildImageRow(
  items,
  columns
) {

  const usable =
    items.slice(
      0,
      columns
    );


  const width =
    Math.floor(
      100 /
      columns
    );


  const cells =
    usable.map(
      function (item) {

        return (

          '<td width="' +
            width +
            '%" valign="top" style="width:' +
            width +
            '%;padding:3px;vertical-align:top;">' +

            buildEmailImage(
              item.url,
              item.clickUrl ||
                item.url,
              item.caption
            ) +

          "</td>"

        );

      }
    ).join("");


  return (

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;table-layout:fixed;margin:28px 0;">' +

      "<tr>" +

        cells +

      "</tr>" +

    "</table>"

  );

}


/*
 * FOUR-UP — one single row.
 */
function buildFourImageRow(
  items
) {

  const usable =
    items.slice(
      0,
      4
    );


  const cells =
    usable.map(
      function (item) {

        return (

          '<td width="25%" valign="top" style="width:25%;padding:3px;vertical-align:top;">' +

            buildEmailImage(
              item.url,
              item.clickUrl ||
                item.url,
              item.caption
            ) +

          "</td>"

        );

      }
    ).join("");


  return (

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;table-layout:fixed;margin:28px 0;">' +

      "<tr>" +

        cells +

      "</tr>" +

    "</table>"

  );

}


function buildImageModuleHtml(
  block
) {

  if (
    !block ||
    !block.items ||
    !block.items.length
  ) {

    return "";

  }


  if (
    block.layout ===
    "full"
  ) {

    return buildFullWidthImage(
      block.items[0]
    );

  }


  if (
    block.layout ===
    "two"
  ) {

    return buildImageRow(
      block.items,
      2
    );

  }


  if (
    block.layout ===
    "three"
  ) {

    return buildImageRow(
      block.items,
      3
    );

  }


  if (
    block.layout ===
    "four"
  ) {

    return buildFourImageRow(
      block.items
    );

  }


  return "";

}


/* =========================================================
   COMPLETE NEWSLETTER
========================================================= */

function buildNewsletterHtml() {

  const edition =
    value("edition");

  const date =
    value("date");

  const title =
    value("title");

  const subtitle =
    value("subtitle");

  const reflection =
    value("reflection");

  const workText =
    value("workText");

  const studioText =
    value("studioText");

  const hero1 =
    value("hero1Url");

  const hero2 =
    value("hero2Url");

  const hero1Caption =
    value("hero1Caption");

  const hero2Caption =
    value("hero2Caption");

  const invitationTitle =
    value("inviteTitle");

  const invitationText =
    value("inviteText");

  const ctaLabel =
    value("ctaLabel") ||
    "INQUIRE";

  const ctaUrl =
    value("ctaUrl");

  const question =
    value("question");

  const notes =
    collectPleasureNotes();

  const modules =
    collectImageBlocks();


  /*
   * Hero click destinations.
   */
  const hero1Preview =
    document.getElementById(
      "hero1Preview"
    );


  const hero2Preview =
    document.getElementById(
      "hero2Preview"
    );


  const hero1Click =
    hero1Preview &&
    hero1Preview.dataset.fullUrl
      ? hero1Preview.dataset.fullUrl
      : hero1;


  const hero2Click =
    hero2Preview &&
    hero2Preview.dataset.fullUrl
      ? hero2Preview.dataset.fullUrl
      : hero2;


  /*
   * Hero HTML.
   */
  const hero1Html =
    hero1
      ? buildEmailImage(
          hero1,
          hero1Click,
          hero1Caption
        )
      : "";


  const hero2Html =
    hero2
      ? buildEmailImage(
          hero2,
          hero2Click,
          hero2Caption
        )
      : "";


  /*
   * Modular HTML.
   */
  let modulesHtml =
    "";


  modules.forEach(
    function (block) {

      modulesHtml +=
        buildImageModuleHtml(
          block
        );

    }
  );


  /*
   * Invitation.
   */
  let invitationHtml =
    "";


  if (
    invitationTitle ||
    invitationText ||
    ctaUrl
  ) {

    invitationHtml =

      '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:40px 0 11px;">' +

        "05 — AN INVITATION" +

      "</div>" +

      (
        invitationTitle
          ? '<div style="font:27px/1.15 Garamond,Georgia,Times New Roman,serif;margin:0 0 10px;">' +

              escapeHtml(
                invitationTitle
              ) +

            "</div>"
          : ""
      ) +

      paragraph(
        invitationText
      ) +

      (
        ctaUrl
          ? '<div style="padding:0 0 25px;">' +

              '<a href="' +
                escapeAttribute(
                  ctaUrl
                ) +
                '" style="display:inline-block;background:#151515;color:#fff;text-decoration:none;padding:12px 18px;font:10px Arial,Helvetica,sans-serif;letter-spacing:1.2px;">' +

                escapeHtml(
                  ctaLabel
                ) +

              "</a>" +

            "</div>"
          : ""
      );

  }


  const dateLine =
    [
      edition,
      date,
      title
    ]
      .filter(Boolean)
      .map(
        function (item) {
          return escapeHtml(item);
        }
      )
      .join(" · ");


  /*
   * COMPLETE EMAIL.
   */
  return (

    '<div style="margin:0;padding:0;background:#f4f0e8;color:#151515;">' +

      '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;background:#f4f0e8;">' +

        "<tr>" +

          '<td align="center" style="padding:28px 12px;">' +

            '<table role="presentation" width="680" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;width:100%;max-width:680px;background:#fffdf8;">' +

              /*
               * HEADER
               */

              "<tr>" +

                '<td style="padding:42px 42px 20px;">' +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:2px;">' +

                    "FLRS GLOBAL" +

                  "</div>" +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.4px;color:#777;margin-top:8px;">' +

                    "FROM THE STUDIO OF FREDDIE L. RANKIN II" +

                  "</div>" +

                  '<h1 style="font:400 50px/0.96 Garamond,Georgia,Times New Roman,serif;margin:20px 0 10px;">' +

                    "The Pleasure Dispatch" +

                  "</h1>" +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.3px;color:#777;border-bottom:1px solid #151515;padding-bottom:20px;">' +

                    dateLine +

                  "</div>" +

                  (
                    subtitle
                      ? '<div style="font:18px/1.45 Garamond,Georgia,Times New Roman,serif;margin-top:20px;">' +

                          escapeHtml(
                            subtitle
                          ) +

                        "</div>"
                      : ""
                  ) +

                "</td>" +

              "</tr>" +

              /*
               * CONTENT
               */

              "<tr>" +

                '<td style="padding:0 42px;">' +

                  /*
                   * Motif
                   */

                  '<div style="text-align:center;font:27px Garamond,Georgia,serif;margin:0 0 20px;">' +

                    "◒" +

                  "</div>" +

                  /*
                   * Hero
                   */

                  hero1Html +

                  /*
                   * Reflection
                   */

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:30px 0 11px;">' +

                    "01 — A REFLECTION" +

                  "</div>" +

                  paragraph(
                    reflection
                  ) +

                  /*
                   * Work
                   */

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +

                    "02 — THE WORK" +

                  "</div>" +

                  paragraph(
                    workText
                  ) +

                  /*
                   * Modules
                   */

                  modulesHtml +

                  /*
                   * Studio Notes
                   */

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +

                    "03 — STUDIO NOTES" +

                  "</div>" +

                  paragraph(
                    studioText
                  ) +

                  /*
                   * Hero 02
                   */

                  hero2Html +

                  /*
                   * Pleasure Notes
                   */

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +

                    "04 — PLEASURE NOTES" +

                  "</div>" +

                  '<div style="font:19px/1.4 Garamond,Georgia,Times New Roman,serif;margin:0 0 8px;">' +

                    "An offering of what has held my attention." +

                  "</div>" +

                  buildPleasureNotesHtml(
                    notes
                  ) +

                  /*
                   * Invitation
                   */

                  invitationHtml +

                  /*
                   * Question
                   */

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +

                    "06 — A QUESTION" +

                  "</div>" +

                  '<div style="font:25px/1.35 Garamond,Georgia,Times New Roman,serif;margin:0 0 36px;">' +

                    escapeHtml(
                      question
                    ) +

                  "</div>" +

                  /*
                   * Closing motif
                   */

                  '<div style="text-align:center;font:27px Garamond,Georgia,serif;margin:20px 0 32px;">' +

                    "◒" +

                  "</div>" +

                "</td>" +

              "</tr>" +

              /*
               * FOOTER
               */

              "<tr>" +

                '<td style="padding:18px 42px 34px;border-top:1px solid #151515;">' +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.1px;color:#777;">' +

                    "THE PLEASURE DISPATCH · BY FLRS GLOBAL" +

                  "</div>" +

                "</td>" +

              "</tr>" +

            "</table>" +

          "</td>" +

        "</tr>" +

      "</table>" +

    "</div>"

  );

}


/* =========================================================
   OUTLOOK BUILD
========================================================= */

function buildSubject() {

  return (

    "The Pleasure Dispatch — " +

    (
      value("edition") ||
      "No. 001"
    ) +

    ": " +

    (
      value("title") ||
      "A Note on Pleasure"
    )

  );

}


function buildInOutlook() {

  setStatus(
    "Preparing The Pleasure Dispatch…"
  );


  if (
    typeof Office ===
    "undefined"
  ) {

    setStatus(
      "Office.js is not available."
    );

    return;

  }


  if (
    !Office.context ||
    !Office.context.mailbox ||
    !Office.context.mailbox.item
  ) {

    setStatus(
      "Open The Pleasure Dispatch from a new Outlook message."
    );

    return;

  }


  const item =
    Office.context.mailbox.item;


  if (
    !item.body ||
    typeof item.body.setAsync !==
      "function"
  ) {

    setStatus(
      "Outlook did not provide access to the message body."
    );

    return;

  }


  let html;


  try {

    html =
      buildNewsletterHtml();

  } catch (error) {

    console.error(
      error
    );


    setStatus(
      "Newsletter build error: " +
      error.message
    );


    return;

  }


  const subject =
    buildSubject();


  /*
   * Drive-hosted images are regular HTTPS URLs.
   */
  item.subject.setAsync(
    subject,
    function (
      subjectResult
    ) {

      if (
        subjectResult &&
        subjectResult.status !==
          Office.AsyncResultStatus.Succeeded
      ) {

        setStatus(
          "Could not set subject: " +
          getAsyncError(
            subjectResult
          )
        );

        return;

      }


      item.body.setAsync(
        html,
        {
          coercionType:
            Office.CoercionType.Html
        },
        function (
          bodyResult
        ) {

          if (
            bodyResult &&
            bodyResult.status ===
              Office.AsyncResultStatus.Succeeded
          ) {

            setStatus(
              "✓ Dispatch built in Outlook. Review it before sending."
            );

          } else {

            setStatus(
              "Could not insert the Dispatch: " +
              getAsyncError(
                bodyResult
              )
            );

          }

        }
      );

    }
  );

}


function getAsyncError(
  result
) {

  if (
    result &&
    result.error &&
    result.error.message
  ) {

    return result.error.message;

  }


  return "Unknown Outlook error.";

}


/* =========================================================
   PREVIEW
========================================================= */

function previewNewsletter() {

  let html;


  try {

    html =
      buildNewsletterHtml();

  } catch (error) {

    setStatus(
      "Preview error: " +
      error.message
    );

    return;

  }


  const previewWindow =
    window.open(
      "",
      "_blank"
    );


  if (!previewWindow) {

    setStatus(
      "Preview was blocked by the browser."
    );

    return;

  }


  previewWindow.document.open();


  previewWindow.document.write(

    "<!doctype html>" +

    "<html>" +

      "<head>" +

        '<meta charset="utf-8">' +

        "<title>The Pleasure Dispatch Preview</title>" +

        "<style>" +

          "body{margin:0;}" +

          "img{cursor:pointer;}" +

        "</style>" +

      "</head>" +

      "<body>" +

        html +

      "</body>" +

    "</html>"

  );


  previewWindow.document.close();


  setStatus(
    "✓ Preview opened."
  );

}
