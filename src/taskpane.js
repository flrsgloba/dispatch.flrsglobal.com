/* =========================================================
   THE PLEASURE DISPATCH
   v0.9

   Desktop image workflow:
   Desktop Upload
        ↓
   Resize / compress
        ↓
   Upload to Google Apps Script
        ↓
   Google Drive
        ↓
   Hosted image URL
        ↓
   Composer preview
        ↓
   Clickable image
        ↓
   Build in Outlook

   Text / externally hosted images still work normally.
========================================================= */


/* =========================================================
   CONFIGURATION
========================================================= */

/*
 * Your CURRENT Google Apps Script Web App deployment.
 */
const DRIVE_API_URL =
  "https://script.google.com/macros/s/AKfycbyTLvYbe1O_BbzsH09UMSZdbY9_XZXga-TbSPkR3UclT3Qhlaj7gy5yhXPA_UpE6Fu7tw/exec";


/*
 * Email image sizing.
 *
 * 1800px is large enough for email viewing while
 * preventing giant camera files from entering Outlook.
 */
const IMAGE_MAX_WIDTH = 1800;
const IMAGE_MAX_HEIGHT = 1800;


/*
 * JPEG compression quality.
 */
const IMAGE_QUALITY = 0.82;


/*
 * Maximum file size we will accept BEFORE
 * compression.
 */
const MAX_SOURCE_IMAGE_MB = 40;


/*
 * Maximum compressed image size we
 * prefer to send into Outlook.
 */
const MAX_EMAIL_IMAGE_MB = 5;


/* =========================================================
   GLOBAL STATE
========================================================= */

let blockCounter = 0;
let pleasureCounter = 0;


/*
 * Holds information about images uploaded
 * from the desktop.
 *
 * Example:
 *
 * {
 *   dataUrl: "...",
 *   fileName: "...",
 *   driveUrl: "...",
 *   imageUrl: "...",
 *   fileId: "..."
 * }
 */
const uploadedImages = [];


/* =========================================================
   INITIALIZATION
========================================================= */

Office.onReady(function () {

  initializeDispatch();

});


function initializeDispatch() {

  setupStaticControls();
  setupDelegatedControls();

  setupHero("hero1");
  setupHero("hero2");

  /*
   * Default Pleasure Notes.
   */
  addPleasureRow(
    "Coffee",
    ""
  );

  addPleasureRow(
    "Art",
    ""
  );

  addPleasureRow(
    "Object",
    ""
  );


  /*
   * Start with one modular image block.
   */
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


  /*
   * PREVIEW
   */

  if (previewButton) {

    previewButton.addEventListener(
      "click",
      function (event) {

        event.preventDefault();

        preview();

      }
    );

  }


  /*
   * BUILD IN OUTLOOK
   */

  if (buildButton) {

    buildButton.addEventListener(
      "click",
      function (event) {

        event.preventDefault();

        buildInOutlook();

      }
    );

  }


  /*
   * ADD IMAGE MODULE
   */

  if (addImageButton) {

    addImageButton.addEventListener(
      "click",
      function (event) {

        event.preventDefault();

        addImageBlock();

      }
    );

  }


  /*
   * ADD PLEASURE NOTE
   */

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

  /*
   * BUTTONS
   */

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


      /*
       * Static controls already have listeners.
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


        if (!block) {

          setStatus(
            "Could not find image module."
          );

          return;

        }


        moveImageBlock(
          block,
          -1
        );


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


        if (!block) {

          setStatus(
            "Could not find image module."
          );

          return;

        }


        moveImageBlock(
          block,
          1
        );


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


        if (
          block
        ) {

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
       * URL BUTTON
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
   * SELECTS + FILE INPUTS
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


      /*
       * Desktop image uploads.
       */

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
   * TEXT / IMAGE URL INPUTS
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


function setStatus(message) {

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
        "No image was selected."
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


          /*
           * Preserve aspect ratio.
           */

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


          /*
           * Fill the canvas white before converting
           * transparent PNGs to JPEG.
           */
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


              /*
               * If the image is still unusually large,
               * progressively reduce dimensions/quality.
               */
              reduceBlobIfNeeded(
                blob,
                width,
                height,
                function (finalBlob) {

                  blobToDataUrl(
                    finalBlob,
                    function (dataUrl) {

                      callback(
                        {
                          blob:
                            finalBlob,

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


function reduceBlobIfNeeded(
  blob,
  width,
  height,
  callback
) {

  const maxBytes =
    MAX_EMAIL_IMAGE_MB *
    1024 *
    1024;


  if (
    blob.size <=
    maxBytes
  ) {

    callback(
      blob
    );

    return;

  }


  /*
   * Reduce dimensions by 20%.
   */

  const image =
    new Image();


  const objectUrl =
    URL.createObjectURL(
      blob
    );


  image.onload =
    function () {

      URL.revokeObjectURL(
        objectUrl
      );


      const newWidth =
        Math.round(
          image.naturalWidth *
          0.8
        );


      const newHeight =
        Math.round(
          image.naturalHeight *
          0.8
        );


      const canvas =
        document.createElement(
          "canvas"
        );


      canvas.width =
        newWidth;


      canvas.height =
        newHeight;


      const context =
        canvas.getContext(
          "2d"
        );


      context.drawImage(
        image,
        0,
        0,
        newWidth,
        newHeight
      );


      canvas.toBlob(
        function (smallerBlob) {

          if (!smallerBlob) {

            callback(
              blob
            );

            return;

          }


          if (
            smallerBlob.size <
            blob.size
          ) {

            reduceBlobIfNeeded(
              smallerBlob,
              newWidth,
              newHeight,
              callback
            );

          } else {

            callback(
              blob
            );

          }

        },
        "image/jpeg",
        0.72
      );

    };


  image.onerror =
    function () {

      URL.revokeObjectURL(
        objectUrl
      );

      callback(
        blob
      );

    };


  image.src =
    objectUrl;

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

      setStatus(
        "Could not create image data."
      );

    };


  reader.readAsDataURL(
    blob
  );

}


/* =========================================================
   GOOGLE DRIVE UPLOAD
========================================================= */

async function uploadToGoogleDrive(
  data,
  originalFileName
) {

  /*
   * The Apps Script accepts JSON.
   *
   * We deliberately use text/plain instead of application/json
   * to avoid an unnecessary CORS preflight request.
   */

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
        data
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
      "Could not connect to Google Drive. " +
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


  return result;

}


function stripDataUrlPrefix(
  dataUrl
) {

  const separator =
    dataUrl.indexOf(
      ","
    );


  if (
    separator ===
    -1
  ) {

    return dataUrl;

  }


  return dataUrl.substring(
    separator + 1
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


/*
 * Store the Drive result so the image can use
 * the same asset later.
 */
function rememberUploadedImage(
  result,
  localDataUrl
) {

  uploadedImages.push({

    dataUrl:
      localDataUrl,

    fileId:
      result.fileId,

    fileName:
      result.fileName,

    driveUrl:
      result.driveUrl,

    imageUrl:
      result.imageUrl

  });

}


/* =========================================================
   HERO IMAGE UPLOADS
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

  setStatus(
    "Optimizing hero image…"
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
        document.getElementById(
          key +
          "Url"
        );


      /*
       * Preview immediately.
       */
      if (urlInput) {

        urlInput.value =
          result.dataUrl;

      }


      renderHeroPreview(
        key,
        result.dataUrl
      );


      /*
       * Upload to Drive.
       */
      try {

        const driveResult =
          await uploadToGoogleDrive(
            result.dataUrl,
            file.name
          );


        rememberUploadedImage(
          driveResult,
          result.dataUrl
        );


        /*
         * Replace temporary local image
         * with the persistent Drive asset.
         */
        if (urlInput) {

          urlInput.value =
            driveResult.imageUrl;

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


      } catch (uploadError) {

        /*
         * Keep local preview functioning even
         * when Drive upload isn't available.
         */
        setStatus(
          "Preview ready, but Drive upload failed: " +
          uploadError.message
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

      box.textContent =
        "This image could not be loaded.";

    };


  image.src =
    url;

}


/* =========================================================
   MODULAR DESKTOP IMAGE UPLOAD
========================================================= */

async function handleFileInput(
  input
) {

  /*
   * Hero uploads are handled separately.
   */

  if (
    input.id ===
    "hero1File" ||
    input.id ===
    "hero2File"
  ) {

    return;

  }


  const file =
    input.files &&
    input.files[0];


  if (!file) {
    return;
  }


  const item =
    input.closest(
      ".image-item"
    );


  if (!item) {
    return;
  }


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


      if (url) {

        url.value =
          result.dataUrl;

      }


      renderModulePreview(
        item,
        result.dataUrl
      );


      try {

        const driveResult =
          await uploadToGoogleDrive(
            result.dataUrl,
            file.name
          );


        rememberUploadedImage(
          driveResult,
          result.dataUrl
        );


        if (url) {

          url.value =
            driveResult.imageUrl;

        }


        renderModulePreview(
          item,
          driveResult.imageUrl
        );


        setStatus(
          "Image saved to Google Drive."
        );


      } catch (uploadError) {

        setStatus(
          "Preview ready, but Drive upload failed: " +
          uploadError.message
        );

      }

    }
  );

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

    '<table role="presentation" cellspacing="0" cellpadding="0" border="0" ' +

      'style="border-collapse:collapse;width:100%;margin:4px 0 30px;">' +

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
   IMAGE MODULES
========================================================= */

function addImageBlock() {

  blockCounter++;


  const container =
    document.getElementById(
      "imageBlocks"
    );


  if (!container) {
    return;
  }


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

  getImageBlocks()
    .forEach(
      function (
        block,
        index
      ) {

        const number =
          block.querySelector(
            ".block-number"
          );


        if (number) {

          number.textContent =
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


function moveImageBlock(
  block,
  direction
) {

  const container =
    document.getElementById(
      "imageBlocks"
    );


  if (
    !container ||
    !block
  ) {

    setStatus(
      "Unable to move image module."
    );

    return;

  }


  const blocks =
    getImageBlocks();


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

    setStatus(
      "Image Module moved up."
    );

  } else {

    container.insertBefore(
      block,
      target.nextSibling
    );

    setStatus(
      "Image Module moved down."
    );

  }


  renumberImageBlocks();

}


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
      layout ===
      "full"
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
      '<input type="file" accept="image/jpeg,image/png,image/webp">' +
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


  if (
    preset.url
  ) {

    renderModulePreview(
      item,
      preset.url
    );

  }

}


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


function renumberImageItems(
  wrap
) {

  Array.from(
    wrap.children
  ).forEach(
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
   DUPLICATE IMAGE MODULE
========================================================= */

function duplicateBlock(
  block
) {

  const layout =
    block.dataset.layout ||
    "full";


  const items =
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
              : ""

        };

      }
    );


  const newBlock =
    document.createElement(
      "article"
    );


  blockCounter++;


  newBlock.className =
    "image-block";


  newBlock.dataset.id =
    String(
      blockCounter
    );


  newBlock.dataset.layout =
    layout;


  newBlock.innerHTML =

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


  const select =
    newBlock.querySelector(
      ".layout-select"
    );


  select.value =
    layout;


  const wrap =
    newBlock.querySelector(
      ".image-items"
    );


  wrap.className =
    "image-items " +
    (
      layout ===
      "full"
        ? "one"
        : layout
    );


  items.forEach(
    function (
      item,
      index
    ) {

      addImageItem(
        wrap,
        index + 1,
        item
      );

    }
  );


  const container =
    document.getElementById(
      "imageBlocks"
    );


  container.insertBefore(
    newBlock,
    block.nextSibling
  );


  renumberImageBlocks();


  setStatus(
    "Image Module duplicated."
  );

}


/* =========================================================
   COLLECT MODULES
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
   CLICKABLE EMAIL IMAGES
========================================================= */

/*
 * Every externally hosted / Drive-hosted image becomes:
 *
 * <a href="full image">
 *    <img src="display image">
 * </a>
 *
 * Desktop images are converted to CID attachments later.
 */
function emailImage(
  imageUrl,
  displayUrl,
  caption
) {

  if (!imageUrl) {
    return "";
  }


  const source =
    displayUrl ||
    imageUrl;


  let html =

    '<a href="' +
      escapeAttribute(
        imageUrl
      ) +
      '" target="_blank" style="text-decoration:none;">' +

      '<img src="' +
        escapeAttribute(
          source
        ) +
        '" alt="" ' +
        'style="display:block;width:100%;height:auto;border:0;">' +

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


/* =========================================================
   EMAIL IMAGE MODULES
========================================================= */

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

          emailImage(
            item.url,
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

            emailImage(
              item.url,
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
 * Four Up = one horizontal row.
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

            emailImage(
              item.url,
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
   NEWSLETTER BUILDER
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
   * HERO 01
   */

  let hero1Html =
    "";


  if (hero1) {

    hero1Html =

      emailImage(
        hero1,
        hero1,
        hero1Caption
      );

  }


  /*
   * HERO 02
   */

  let hero2Html =
    "";


  if (hero2) {

    hero2Html =

      emailImage(
        hero2,
        hero2,
        hero2Caption
      );

  }


  /*
   * MODULAR IMAGE BLOCKS
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
   * INVITATION
   */

  let invitationHtml =
    "";


  if (
    invitationTitle ||
    invitationText ||
    ctaUrl
  ) {

    invitationHtml =

      '<div style="' +
        "font:10px Arial,Helvetica,sans-serif;" +
        "letter-spacing:1.5px;" +
        "color:#777;" +
        "margin:40px 0 11px;" +
      '">' +

        "05 — AN INVITATION" +

      "</div>" +

      (
        invitationTitle
          ? '<div style="' +
              "font:27px/1.15 Garamond,Georgia,Times New Roman,serif;" +
              "margin:0 0 10px;" +
            '">' +

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


  /*
   * FULL EMAIL
   */

  return (

    '<div style="' +
      "margin:0;" +
      "padding:0;" +
      "background:#f4f0e8;" +
      "color:#151515;" +
    '">' +

      '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="' +
        "border-collapse:collapse;" +
        "background:#f4f0e8;" +
      '">' +

        "<tr>" +

          '<td align="center" style="padding:28px 12px;">' +

            '<table role="presentation" width="680" cellspacing="0" cellpadding="0" border="0" style="' +
              "border-collapse:collapse;" +
              "width:100%;" +
              "max-width:680px;" +
              "background:#fffdf8;" +
            '">' +


              /*
               * HEADER
               */

              "<tr>" +

                '<td style="padding:42px 42px 20px;">' +

                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:2px;" +
                  '">' +

                    "FLRS GLOBAL" +

                  "</div>" +


                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.4px;" +
                    "color:#777;" +
                    "margin-top:8px;" +
                  '">' +

                    "FROM THE STUDIO OF FREDDIE L. RANKIN II" +

                  "</div>" +


                  '<h1 style="' +
                    "font:400 50px/0.96 Garamond,Georgia,Times New Roman,serif;" +
                    "margin:20px 0 10px;" +
                  '">' +

                    "The Pleasure Dispatch" +

                  "</h1>" +


                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.3px;" +
                    "color:#777;" +
                    "border-bottom:1px solid #151515;" +
                    "padding-bottom:20px;" +
                  '">' +

                    dateLine +

                  "</div>" +


                  (
                    subtitle
                      ? '<div style="' +
                          "font:18px/1.45 Garamond,Georgia,Times New Roman,serif;" +
                          "margin-top:20px;" +
                        '">' +

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


                  /*
                   * MOTIF
                   */

                  '<div style="' +
                    "text-align:center;" +
                    "font:27px Garamond,Georgia,serif;" +
                    "margin:0 0 20px;" +
                  '">' +

                    "◒" +

                  "</div>" +


                  /*
                   * HERO 01
                   */

                  (
                    hero1Html
                      ? '<div style="margin-bottom:4px;">' +
                          hero1Html +
                        "</div>"
                      : ""
                  ) +


                  /*
                   * REFLECTION
                   */

                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.5px;" +
                    "color:#777;" +
                    "margin:30px 0 11px;" +
                  '">' +

                    "01 — A REFLECTION" +

                  "</div>" +


                  paragraph(
                    reflection
                  ) +


                  /*
                   * THE WORK
                   */

                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.5px;" +
                    "color:#777;" +
                    "margin:38px 0 11px;" +
                  '">' +

                    "02 — THE WORK" +

                  "</div>" +


                  paragraph(
                    workText
                  ) +


                  /*
                   * MODULAR IMAGES
                   */

                  modulesHtml +


                  /*
                   * STUDIO NOTES
                   */

                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.5px;" +
                    "color:#777;" +
                    "margin:38px 0 11px;" +
                  '">' +

                    "03 — STUDIO NOTES" +

                  "</div>" +


                  paragraph(
                    studioText
                  ) +


                  /*
                   * HERO 02
                   */

                  (
                    hero2Html
                      ? '<div style="margin-bottom:4px;">' +
                          hero2Html +
                        "</div>"
                      : ""
                  ) +


                  /*
                   * PLEASURE NOTES
                   */

                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.5px;" +
                    "color:#777;" +
                    "margin:38px 0 11px;" +
                  '">' +

                    "04 — PLEASURE NOTES" +

                  "</div>" +


                  '<div style="' +
                    "font:19px/1.4 Garamond,Georgia,Times New Roman,serif;" +
                    "margin:0 0 8px;" +
                  '">' +

                    "An offering of what has held my attention." +

                  "</div>" +


                  buildPleasureNotesHtml(
                    notes
                  ) +


                  /*
                   * INVITATION
                   */

                  invitationHtml +


                  /*
                   * QUESTION
                   */

                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.5px;" +
                    "color:#777;" +
                    "margin:38px 0 11px;" +
                  '">' +

                    "06 — A QUESTION" +

                  "</div>" +


                  '<div style="' +
                    "font:25px/1.35 Garamond,Georgia,Times New Roman,serif;" +
                    "margin:0 0 36px;" +
                  '">' +

                    escapeHtml(
                      question
                    ) +

                  "</div>" +


                  /*
                   * CLOSING MOTIF
                   */

                  '<div style="' +
                    "text-align:center;" +
                    "font:27px Garamond,Georgia,serif;" +
                    "margin:20px 0 32px;" +
                  '">' +

                    "◒" +

                  "</div>" +

                "</td>" +

              "</tr>" +


              /*
               * FOOTER
               */

              "<tr>" +

                '<td style="' +
                  "padding:18px 42px 34px;" +
                  "border-top:1px solid #151515;" +
                '">' +

                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.1px;" +
                    "color:#777;" +
                  '">' +

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
   OUTLOOK INLINE IMAGE CONVERSION
========================================================= */

/*
 * Find desktop-uploaded Base64 images still present
 * in the generated HTML.
 *
 * Drive-hosted images will NOT match this and therefore
 * remain ordinary hosted images.
 */
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


function getExtension(
  mimeType
) {

  switch (
    mimeType.toLowerCase()
  ) {

    case "image/png":
      return "png";

    case "image/webp":
      return "webp";

    case "image/jpeg":
    case "image/jpg":
    default:
      return "jpg";

  }

}


/*
 * Add locally uploaded Base64 images as Outlook
 * inline attachments.
 *
 * This is the fallback for a Drive upload that wasn't
 * successful. The normal desktop workflow should already
 * have a Drive URL by the time Build is clicked.
 */
function addInlineImages(
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

    addInlineImages(
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
    "." +
    getExtension(
      parsed.mimeType
    );


  setStatus(
    "Embedding fallback image " +
    (index + 1) +
    " of " +
    images.length +
    "…"
  );


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
                : "Unknown error."
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


      addInlineImages(
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
   BUILD IN OUTLOOK
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
      "Newsletter build error:",
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
   * Microsoft recommends getting the current body
   * before working with inline Base64 images.
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
       * Normally there should be ZERO Base64 images
       * because desktop images are uploaded to Drive.
       *
       * This is retained as a fallback.
       */
      const localImages =
        findBase64Images(
          html
        );


      if (
        localImages.length ===
        0
      ) {

        writeNewsletter(
          item,
          subject,
          html
        );

        return;

      }


      setStatus(
        "Preparing fallback image embedding…"
      );


      addInlineImages(
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

    }
  );

}


/* =========================================================
   WRITE NEWSLETTER TO OUTLOOK
========================================================= */

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
   PREVIEW
========================================================= */

function preview() {

  setStatus(
    "Opening Dispatch preview…"
  );


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


  const html =
    buildNewsletterHtml();


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
