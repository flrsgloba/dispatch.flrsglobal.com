/* =========================================================
   THE PLEASURE DISPATCH
   v1.1
========================================================= */


/* =========================================================
   CONFIGURATION
========================================================= */

const DRIVE_API_URL =
  "https://script.google.com/macros/s/AKfycbyTLvYbe1O_BbzsH09UMSZdbY9_XZXga-TbSPkR3UclT3Qhlaj7gy5yhXPA_UpE6Fu7tw/exec";

const IMAGE_MAX_WIDTH = 1800;
const IMAGE_MAX_HEIGHT = 1800;

const IMAGE_QUALITY = 0.82;

const MAX_SOURCE_IMAGE_MB = 40;


/*
 * Google Drive thumbnail delivery size.
 *
 * This is intentionally slightly larger than the
 * newsletter's visible image width.
 */
const DRIVE_IMAGE_WIDTH = 1800;


/* =========================================================
   STATE
========================================================= */

let blockCounter = 0;
let pleasureCounter = 0;
let dispatchInitialized = false;


/* =========================================================
   INITIALIZATION
========================================================= */

Office.onReady(function () {

  initializeDispatch();

});


function initializeDispatch() {

  if (dispatchInitialized) {
    return;
  }

  dispatchInitialized = true;

  setupStaticControls();
  setupDelegatedControls();

  setupHero("hero1");
  setupHero("hero2");

  addPleasureRow("Coffee", "");
  addPleasureRow("Art", "");
  addPleasureRow("Object", "");

  addImageBlock();

  setStatus(
    "The Pleasure Dispatch is ready."
  );

}


/* =========================================================
   STATIC CONTROLS
========================================================= */

function setupStaticControls() {

  const previewButton =
    document.getElementById(
      "previewBtn"
    );

  const buildButton =
    document.getElementById(
      "insertBtn"
    );

  const addImageButton =
    document.getElementById(
      "addImageBlock"
    );

  const addPleasureButton =
    document.getElementById(
      "addPleasure"
    );


  if (previewButton) {

    previewButton.addEventListener(
      "click",
      function (event) {

        event.preventDefault();

        preview();

      }
    );

  }


  if (buildButton) {

    buildButton.addEventListener(
      "click",
      function (event) {

        event.preventDefault();

        buildInOutlook();

      }
    );

  }


  if (addImageButton) {

    addImageButton.addEventListener(
      "click",
      function (event) {

        event.preventDefault();

        addImageBlock();

      }
    );

  }


  if (addPleasureButton) {

    addPleasureButton.addEventListener(
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
   DYNAMIC CONTROLS
========================================================= */

function setupDelegatedControls() {

  document.addEventListener(
    "click",
    function (event) {

      const button =
        event.target.closest(
          "button"
        );


      if (!button) {
        return;
      }


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
       * IMAGE URL BUTTON
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


      /*
       * HERO URL BUTTON
       */

      if (
        button.dataset.url
      ) {

        const input =
          document.getElementById(
            button.dataset.url +
            "Url"
          );

        if (input) {

          input.focus();

          setStatus(
            "Paste a direct HTTPS image URL."
          );

        }

      }

    }
  );


  /*
   * SELECTS / FILES
   */

  document.addEventListener(
    "change",
    function (event) {

      const layout =
        event.target.closest(
          ".layout-select"
        );


      if (layout) {

        const block =
          layout.closest(
            ".image-block"
          );

        if (!block) {
          return;
        }

        block.dataset.layout =
          layout.value;

        syncImageInputs(
          block
        );


        const selected =
          layout.options[
            layout.selectedIndex
          ];


        setStatus(
          "Layout changed to " +
          selected.text +
          "."
        );


        return;

      }


      if (
        event.target.matches(
          'input[type="file"]'
        )
      ) {

        handleFileInput(
          event.target
        );

      }

    }
  );


  /*
   * INPUTS / IMAGE URLS
   */

  document.addEventListener(
    "input",
    function (event) {

      if (
        event.target.id ===
        "hero1Url"
      ) {

        renderHeroPreview(
          "hero1",
          event.target.value.trim()
        );

        return;

      }


      if (
        event.target.id ===
        "hero2Url"
      ) {

        renderHeroPreview(
          "hero2",
          event.target.value.trim()
        );

        return;

      }


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
   GENERAL HELPERS
========================================================= */

function value(id) {

  const element =
    document.getElementById(
      id
    );

  return element
    ? element.value.trim()
    : "";

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
   GOOGLE DRIVE IMAGE URLS
========================================================= */

/*
 * This is the important change.
 *
 * DO NOT use:
 *
 * drive.google.com/uc?export=view&id=...
 *
 * for embedded images.
 *
 * Use Google's thumbnail endpoint instead.
 */
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


/*
 * Full clickable Drive destination.
 */
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


  const sourceSizeMB =
    file.size /
    (1024 * 1024);


  if (
    sourceSizeMB >
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


  setStatus(
    "Optimizing " +
    file.name +
    "…"
  );


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
              "The selected image could not be decoded."
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
          "The selected image could not be read."
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
   DRIVE UPLOAD
========================================================= */

async function uploadToDrive(
  dataUrl,
  originalFileName
) {

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


  setStatus(
    "Saving image to Google Drive…"
  );


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


  if (!response.ok) {

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
   * Normalize the image URL ourselves.
   *
   * This avoids relying on the old /uc endpoint
   * returned by Apps Script.
   */
  const fileId =
    result.fileId;


  result.imageUrl =
    buildDriveImageUrl(
      fileId
    );


  result.driveUrl =
    buildDriveClickUrl(
      fileId
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
   HERO UPLOADS
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
      function (event) {

        const file =
          event.target.files &&
          event.target.files[0];


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


async function handleHeroUpload(
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


      const urlInput =
        document.getElementById(
          key +
          "Url"
        );


      /*
       * Local preview while uploading.
       */
      urlInput.value =
        result.dataUrl;


      renderHeroPreview(
        key,
        result.dataUrl
      );


      try {

        const driveResult =
          await uploadToDrive(
            result.dataUrl,
            file.name
          );


        /*
         * Store the public delivery URL.
         */
        urlInput.value =
          driveResult.imageUrl;


        /*
         * Remember full click target.
         */
        const preview =
          document.getElementById(
            key +
            "Preview"
          );


        if (preview) {

          preview.dataset.fullUrl =
            driveResult.driveUrl;

          preview.dataset.fileId =
            driveResult.fileId;

        }


        renderHeroPreview(
          key,
          driveResult.imageUrl
        );


        setStatus(
          key === "hero1"
            ? "Hero Image 01 saved to Drive."
            : "Hero Image 02 saved to Drive."
        );


      } catch (error) {

        setStatus(
          "Preview ready. Drive upload failed: " +
          error.message
        );

      }

    }
  );

}


/* =========================================================
   HERO PREVIEW
========================================================= */

function renderHeroPreview(
  key,
  url
) {

  const box =
    document.getElementById(
      key +
      "Preview"
    );


  if (!box) {
    return;
  }


  if (!url) {

    box.className =
      "preview hero-preview empty";

    box.textContent =
      "No hero image selected";

    box.removeAttribute(
      "data-full-url"
    );

    return;

  }


  box.className =
    "preview hero-preview";


  box.innerHTML =
    "";


  const image =
    new Image();


  image.onload =
    function () {

      box.appendChild(
        image
      );

    };


  image.onerror =
    function () {

      box.className =
        "preview hero-preview error";


      box.innerHTML =
        "";


      const message =
        document.createElement(
          "div"
        );


      message.textContent =
        "Image could not be loaded.";


      box.appendChild(
        message
      );

    };


  image.src =
    url;

}


/* =========================================================
   MODULAR UPLOAD
========================================================= */

function handleFileInput(
  input
) {

  /*
   * Hero files are handled by setupHero().
   */

  if (
    input.id ===
    "hero1File" ||
    input.id ===
    "hero2File"
  ) {

    return;

  }


  const item =
    input.closest(
      ".image-item"
    );


  if (!item) {
    return;
  }


  const file =
    input.files &&
    input.files[0];


  if (!file) {
    return;
  }


  handleModuleUpload(
    item,
    file
  );

}


async function handleModuleUpload(
  item,
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


      const url =
        item.querySelector(
          ".module-url"
        );


      /*
       * Local preview.
       */
      url.value =
        result.dataUrl;


      renderModulePreview(
        item,
        result.dataUrl
      );


      try {

        const driveResult =
          await uploadToDrive(
            result.dataUrl,
            file.name
          );


        /*
         * Replace local data URL with
         * Google's thumbnail delivery URL.
         */
        url.value =
          driveResult.imageUrl;


        /*
         * Keep the full Drive file URL
         * for click-through.
         */
        item.dataset.fullUrl =
          driveResult.driveUrl;


        item.dataset.fileId =
          driveResult.fileId;


        renderModulePreview(
          item,
          driveResult.imageUrl
        );


        setStatus(
          "Image saved to Google Drive."
        );


      } catch (error) {

        setStatus(
          "Preview ready. Drive upload failed: " +
          error.message
        );

      }

    }
  );

}


/* =========================================================
   MODULAR IMAGE PREVIEW
========================================================= */

function renderModulePreview(
  item,
  url
) {

  const preview =
    item.querySelector(
      ".module-preview"
    );


  if (!preview) {
    return;
  }


  if (!url) {

    preview.className =
      "preview module-preview empty";

    preview.textContent =
      "No image selected";

    return;

  }


  preview.className =
    "preview module-preview";


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
        "preview module-preview error";


      preview.textContent =
        "Image could not be loaded.";

    };


  image.src =
    url;

}


/* =========================================================
   PLEASURE NOTES
========================================================= */

function addPleasureRow(
  labelValue,
  noteValue
) {

  pleasureCounter++;


  const container =
    document.getElementById(
      "pleasureRows"
    );


  if (!container) {
    return;
  }


  const row =
    document.createElement(
      "div"
    );


  row.className =
    "pleasure-row";


  row.dataset.id =
    String(
      pleasureCounter
    );


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

    '<button type="button" aria-label="Remove pleasure note">×</button>';


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


function buildPleasureNotesHtml(
  notes
) {

  if (!notes.length) {
    return "";
  }


  return (

    '<table role="presentation" cellspacing="0" cellpadding="0" border="0" style="' +
      "border-collapse:collapse;" +
      "width:100%;" +
      "margin:4px 0 30px;" +
    '">' +

      notes.map(
        function (note) {

          return (

            "<tr>" +

              '<td style="' +
                "width:120px;" +
                "vertical-align:top;" +
                "padding:5px 18px 5px 0;" +
                "font:10px Arial,Helvetica,sans-serif;" +
                "letter-spacing:1px;" +
                "text-transform:uppercase;" +
                "color:#777;" +
              '">' +

                escapeHtml(
                  note.label
                ) +

              "</td>" +

              '<td style="' +
                "vertical-align:top;" +
                "padding:5px 0;" +
                "font:16px/1.5 Garamond,Georgia,Times New Roman,serif;" +
                "color:#151515;" +
              '">' +

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
   NEWSLETTER IMAGE DATA
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
   EMAIL IMAGE BUILDER
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
        '" alt="" style="' +
          "display:block;" +
          "width:100%;" +
          "height:auto;" +
          "border:0;" +
        '">' +

    "</a>";


  if (caption) {

    html +=

      '<div style="' +
        "font:12px/1.45 Arial,Helvetica,sans-serif;" +
        "color:#777;" +
        "margin-top:7px;" +
      '">' +

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
            '%" valign="top" style="' +
              "width:" +
              width +
              "%;" +
              "padding:3px;" +
              "vertical-align:top;" +
            '">' +

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
   NEWSLETTER HTML
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

  const inviteTitle =
    value("inviteTitle");

  const inviteText =
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
   * Hero click URLs.
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
   * Heroes.
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
   * Modular images.
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
    inviteTitle ||
    inviteText ||
    ctaUrl
  ) {

    invitationHtml =

      '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:40px 0 11px;">' +
        "05 — AN INVITATION" +
      "</div>" +

      (
        inviteTitle
          ? '<div style="font:27px/1.15 Garamond,Georgia,Times New Roman,serif;margin:0 0 10px;">' +

              escapeHtml(
                inviteTitle
              ) +

            "</div>"
          : ""
      ) +

      paragraph(
        inviteText
      ) +

      (
        ctaUrl
          ? '<div style="padding:0 0 25px;">' +

              '<a href="' +
                escapeAttribute(
                  ctaUrl
                ) +
                '" style="' +
                  "display:inline-block;" +
                  "background:#151515;" +
                  "color:#fff;" +
                  "text-decoration:none;" +
                  "padding:12px 18px;" +
                  "font:10px Arial,Helvetica,sans-serif;" +
                  "letter-spacing:1.2px;" +
                '">' +

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

          return escapeHtml(
            item
          );

        }
      )
      .join(" · ");


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
               * BODY
               */

              "<tr>" +

                '<td style="padding:0 42px;">' +

                  '<div style="text-align:center;font:27px Garamond,Georgia,serif;margin:0 0 20px;">◒</div>' +

                  hero1Html +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:30px 0 11px;">' +
                    "01 — A REFLECTION" +
                  "</div>" +

                  paragraph(
                    reflection
                  ) +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +
                    "02 — THE WORK" +
                  "</div>" +

                  paragraph(
                    workText
                  ) +

                  modulesHtml +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +
                    "03 — STUDIO NOTES" +
                  "</div>" +

                  paragraph(
                    studioText
                  ) +

                  hero2Html +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +
                    "04 — PLEASURE NOTES" +
                  "</div>" +

                  '<div style="font:19px/1.4 Garamond,Georgia,Times New Roman,serif;margin:0 0 8px;">' +
                    "An offering of what has held my attention." +
                  "</div>" +

                  buildPleasureNotesHtml(
                    notes
                  ) +

                  invitationHtml +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">' +
                    "06 — A QUESTION" +
                  "</div>" +

                  '<div style="font:25px/1.35 Garamond,Georgia,Times New Roman,serif;margin:0 0 36px;">' +
                    escapeHtml(
                      question
                    ) +
                  "</div>" +

                  '<div style="text-align:center;font:27px Garamond,Georgia,serif;margin:20px 0 32px;">◒</div>' +

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
      "Office.js is not available. Reload the add-in."
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
    typeof item.body.getAsync !==
      "function" ||
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
   * Normally Drive-hosted images are ordinary URLs,
   * so Outlook doesn't need to embed them.
   */
  item.body.getAsync(
    Office.CoercionType.Html,
    function (
      bodyResult
    ) {

      if (
        bodyResult &&
        bodyResult.status !==
          Office.AsyncResultStatus.Succeeded
      ) {

        setStatus(
          "Could not read the Outlook body: " +
          getAsyncError(
            bodyResult
          )
        );

        return;

      }


      /*
       * Fallback only for an image that remained
       * Base64 because Drive upload failed.
       */
      const localImages =
        findBase64Images(
          html
        );


      if (
        localImages.length > 0
      ) {

        addFallbackInlineImages(
          item,
          html,
          localImages,
          0,
          function (
            error,
            finalHtml
          ) {

            if (error) {

              setStatus(
                error.message
              );

              return;

            }


            writeNewsletter(
              item,
              subject,
              finalHtml
            );

          }
        );


        return;

      }


      writeNewsletter(
        item,
        subject,
        html
      );

    }
  );

}


function writeNewsletter(
  item,
  subject,
  html
) {

  setStatus(
    "Writing The Pleasure Dispatch into Outlook…"
  );


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
          "Could not set the subject: " +
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
   BASE64 FALLBACK
========================================================= */

function findBase64Images(
  html
) {

  const regex =
    /data:image\/(?:jpeg|jpg|png|webp);base64,[A-Za-z0-9+/=]+/g;


  const matches =
    html.match(
      regex
    ) || [];


  return matches.filter(
    function (
      item,
      index,
      array
    ) {

      return (
        array.indexOf(item) ===
        index
      );

    }
  );

}


function parseDataUrl(
  dataUrl
) {

  const match =
    dataUrl.match(
      /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/
    );


  if (!match) {
    return null;
  }


  return {

    mimeType:
      match[1],

    base64:
      match[2]

  };

}


function addFallbackInlineImages(
  item,
  html,
  images,
  index,
  callback
) {

  if (
    index >=
    images.length
  ) {

    callback(
      null,
      html
    );

    return;

  }


  const dataUrl =
    images[index];


  const parsed =
    parseDataUrl(
      dataUrl
    );


  if (!parsed) {

    addFallbackInlineImages(
      item,
      html,
      images,
      index + 1,
      callback
    );

    return;

  }


  const filename =
    "pleasure-dispatch-" +
    Date.now() +
    "-" +
    (index + 1) +
    ".jpg";


  item.addFileAttachmentFromBase64Async(
    parsed.base64,
    filename,
    {
      isInline:
        true
    },
    function (
      result
    ) {

      if (
        result.status ===
        Office.AsyncResultStatus.Failed
      ) {

        callback(
          new Error(
            "Outlook could not embed image " +
            (index + 1) +
            ": " +
            (
              result.error
                ? result.error.message
                : "Unknown Outlook error."
            )
          ),
          null
        );

        return;

      }


      html =
        html
          .split(
            dataUrl
          )
          .join(
            "cid:" +
            filename
          );


      addFallbackInlineImages(
        item,
        html,
        images,
        index + 1,
        callback
      );

    }
  );

}


/* =========================================================
   PREVIEW
========================================================= */

function preview() {

  setStatus(
    "Opening Dispatch preview…"
  );


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
      "Preview was blocked. Allow pop-ups for this add-in."
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

      "</head>" +

      "<body style='margin:0;'>" +

        html +

      "</body>" +

    "</html>"

  );


  previewWindow.document.close();


  setStatus(
    "✓ Preview opened."
  );

}
