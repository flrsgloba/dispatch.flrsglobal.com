```javascript
/* =========================================================
   THE PLEASURE DISPATCH
   taskpane.js
   Outlook-Safe Production Version
========================================================= */

const DRIVE_API_URL =
  "https://script.google.com/macros/s/AKfycbyTLvYbe1O_BbzsH09UMSZdbY9_XZXga-TbSPkR3UclT3Qhlaj7gy5yhXPA_UpE6Fu7tw/exec";

const LOOP_ASSET_URL =
  "https://flrsgloba.github.io/dispatch.flrsglobal.com/assets/pleasure-loop.svg";


/* =========================================================
   DESIGN
========================================================= */

const COLORS = {
  background: "#303030",
  surface: "#595959",
  text: "#F2EEE5",
  secondary: "#C9C3B8",
  rule: "#777777",
  accent: "#D8D0C3"
};


/* =========================================================
   IMAGE SETTINGS
========================================================= */

const IMAGE_MAX_WIDTH = 1800;
const IMAGE_MAX_HEIGHT = 1800;
const IMAGE_QUALITY = 0.82;

const MAX_SOURCE_IMAGE_MB = 40;

const DRIVE_IMAGE_WIDTH = 1800;


/*
 * Outlook gets a longer timeout than before.
 */
const BUILD_TIMEOUT_MS = 90000;


/* =========================================================
   STATE
========================================================= */

let blockCounter = 0;
let pleasureCounter = 0;
let initialized = false;


/* =========================================================
   OFFICE INITIALIZATION
========================================================= */

Office.onReady(function () {

  initializeDispatch();

});


function initializeDispatch() {

  if (initialized) {
    return;
  }

  initialized = true;

  bindStaticControls();
  bindDelegatedControls();

  setupHero("hero1");
  setupHero("hero2");


  const pleasureRows =
    document.getElementById(
      "pleasureRows"
    );


  if (
    pleasureRows &&
    pleasureRows.children.length === 0
  ) {

    addPleasureRow("Coffee", "");
    addPleasureRow("Art", "");
    addPleasureRow("Object", "");

  }


  const imageBlocks =
    document.getElementById(
      "imageBlocks"
    );


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
   BASIC HELPERS
========================================================= */

function value(id) {

  const element =
    document.getElementById(id);


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

  return escapeHtml(text);

}


function paragraph(text) {

  if (!text) {
    return "";
  }


  return (

    '<p style="' +
      "margin-top:0;" +
      "margin-right:0;" +
      "margin-bottom:24px;" +
      "margin-left:0;" +
      "font-family:Georgia,Times New Roman,serif;" +
      "font-size:17px;" +
      "line-height:1.7;" +
      "color:" +
        COLORS.text +
      ";" +
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


  console.log(
    "[Pleasure Dispatch]",
    message
  );

}


/* =========================================================
   STATIC CONTROLS
========================================================= */

function bindStaticControls() {

  const previewButton =
    document.getElementById(
      "previewBtn"
    );


  const buildButton =
    document.getElementById(
      "insertBtn"
    );


  const addPleasureButton =
    document.getElementById(
      "addPleasure"
    );


  let addImageButton =
    document.getElementById(
      "addImageBlock"
    );


  if (previewButton) {

    previewButton.onclick =
      function (event) {

        event.preventDefault();
        event.stopPropagation();

        previewNewsletter();

      };

  }


  if (buildButton) {

    buildButton.onclick =
      function (event) {

        event.preventDefault();
        event.stopPropagation();

        buildInOutlook();

      };

  }


  if (addPleasureButton) {

    addPleasureButton.onclick =
      function (event) {

        event.preventDefault();
        event.stopPropagation();

        addPleasureRow("", "");

        setStatus(
          "Pleasure Note added."
        );

      };

  }


  if (!addImageButton) {

    const candidates =
      document.querySelectorAll(
        "button, a, [role='button']"
      );


    for (
      let i = 0;
      i < candidates.length;
      i++
    ) {

      const text =
        (
          candidates[i].textContent ||
          ""
        )
        .trim()
        .toLowerCase();


      if (
        text.includes(
          "add image block"
        )
      ) {

        addImageButton =
          candidates[i];

        break;

      }

    }

  }


  if (addImageButton) {

    addImageButton.onclick =
      function (event) {

        event.preventDefault();
        event.stopPropagation();

        addImageBlock();

      };

  }

}


/* =========================================================
   DYNAMIC CONTROLS
========================================================= */

function bindDelegatedControls() {

  document.addEventListener(
    "click",
    function (event) {

      const button =
        event.target.closest(
          "button, a, [role='button']"
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


      if (
        button.classList.contains(
          "move-up"
        )
      ) {

        event.preventDefault();

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


      if (
        button.classList.contains(
          "move-down"
        )
      ) {

        event.preventDefault();

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


      if (
        button.classList.contains(
          "duplicate"
        )
      ) {

        event.preventDefault();

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


      if (
        button.classList.contains(
          "remove"
        )
      ) {

        event.preventDefault();

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


      if (
        button.classList.contains(
          "remove-image"
        )
      ) {

        event.preventDefault();

        removeImageItem(
          button.closest(
            ".image-item"
          )
        );

        return;

      }


      if (
        button.classList.contains(
          "url-item"
        )
      ) {

        event.preventDefault();

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


        setStatus(
          "Layout changed to " +
          layout.options[
            layout.selectedIndex
          ].text +
          "."
        );


        return;

      }


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

      }

    }
  );


  document.addEventListener(
    "input",
    function (event) {

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
          delete item.dataset.driveUrl;
          delete item.dataset.fileId;


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
    !/^image\/(jpeg|jpg|png|webp)$/i.test(
      file.type
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
                      dataUrl:
                        dataUrl,

                      blob:
                        blob,

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
   HERO IMAGE HANDLING
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


        if (file) {

          handleHeroUpload(
            key,
            file
          );

        }

      }
    );

  }


  if (urlInput) {

    urlInput.addEventListener(
      "input",
      function () {

        const preview =
          document.getElementById(
            key +
            "Preview"
          );


        if (preview) {

          delete preview.dataset.fullUrl;
          delete preview.dataset.fileId;

        }


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
       * Local preview first.
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
            ? "✓ Hero Image 01 saved to Drive."
            : "✓ Hero Image 02 saved to Drive."
        );


      } catch (error) {

        setStatus(
          "Preview ready. Drive upload failed: " +
          error.message
        );


        console.error(
          "Hero upload error:",
          error
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

function handleModuleUpload(
  item,
  file
) {

  if (
    !item ||
    !file
  ) {

    setStatus(
      "No modular image selected."
    );

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


      const urlInput =
        item.querySelector(
          ".module-url"
        );


      if (!urlInput) {

        setStatus(
          "Could not find the modular image field."
        );

        return;

      }


      urlInput.value =
        result.dataUrl;


      renderModulePreview(
        item,
        result.dataUrl
      );


      try {

        setStatus(
          "Saving modular image to Google Drive…"
        );


        const drive =
          await uploadToDrive(
            result.dataUrl,
            file.name
          );


        item.dataset.fileId =
          drive.fileId;

        item.dataset.imageUrl =
          drive.imageUrl;

        item.dataset.fullUrl =
          drive.driveUrl;

        item.dataset.driveUrl =
          drive.driveUrl;


        urlInput.value =
          drive.imageUrl;


        renderModulePreview(
          item,
          drive.imageUrl
        );


        setStatus(
          "✓ Modular image saved to Drive."
        );


      } catch (error) {

        setStatus(
          "Preview ready. Drive upload failed: " +
          error.message
        );


        console.error(
          "Modular upload error:",
          error
        );

      }

    }
  );

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


/* =========================================================
   IMAGE MODULES
========================================================= */

function getImageBlocks() {

  const container =
    document.getElementById(
      "imageBlocks"
    );


  if (!container) {
    return [];
  }


  return Array.from(
    container.children
  ).filter(
    function (element) {

      return element.classList.contains(
        "image-block"
      );

    }
  );

}


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


  container.appendChild(
    block
  );


  syncImageInputs(
    block
  );


  renumberImageBlocks();


  setStatus(
    "✓ Image Module added."
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


  const blocks =
    getImageBlocks();


  if (
    !container ||
    !block
  ) {

    return;

  }


  const index =
    blocks.indexOf(
      block
    );


  if (index === -1) {
    return;
  }


  const targetIndex =
    index +
    direction;


  if (
    targetIndex < 0 ||
    targetIndex >= blocks.length
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
      ? "✓ Image Module moved up."
      : "✓ Image Module moved down."
  );

}


/* =========================================================
   IMAGE INPUTS
========================================================= */

function syncImageInputs(
  block
) {

  const layout =
    block.dataset.layout ||
    "full";


  let count =
    1;


  if (layout === "two") {
    count = 2;
  }

  if (layout === "three") {
    count = 3;
  }

  if (layout === "four") {
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


  if (preset.clickUrl) {

    item.dataset.fullUrl =
      preset.clickUrl;

  }


  if (preset.url) {

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
   DUPLICATE IMAGE MODULE
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


  blockCounter++;


  const clone =
    document.createElement(
      "article"
    );


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
    "✓ Image Module duplicated."
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


/* =========================================================
   PLEASURE NOTES HTML
   OUTLOOK-SAFE TABLE VERSION
========================================================= */

function buildPleasureNotesHtml(
  notes
) {

  if (
    !notes ||
    !notes.length
  ) {

    return "";

  }


  let rows =
    "";


  notes.forEach(
    function (note) {

      rows +=

        "<tr>" +

          '<td valign="top" width="120" style="' +
            "width:120px;" +
            "padding-top:5px;" +
            "padding-right:18px;" +
            "padding-bottom:5px;" +
            "font-family:Arial,Helvetica,sans-serif;" +
            "font-size:10px;" +
            "line-height:14px;" +
            "letter-spacing:1px;" +
            "text-transform:uppercase;" +
            "color:" +
              COLORS.secondary +
            ";" +
          '">' +

            escapeHtml(
              note.label
            ) +

          "</td>" +

          '<td valign="top" style="' +
            "padding-top:5px;" +
            "padding-bottom:5px;" +
            "font-family:Georgia,Times New Roman,serif;" +
            "font-size:16px;" +
            "line-height:24px;" +
            "color:" +
              COLORS.text +
            ";" +
          '">' +

            escapeHtml(
              note.value
            ) +

          "</td>" +

        "</tr>";

    }
  );


  return (

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="' +
      "border-collapse:collapse;" +
      "width:100%;" +
      "margin-bottom:30px;" +
    '">' +

      rows +

    "</table>"

  );

}


/* =========================================================
   COLLECT IMAGE MODULES
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
              function (image) {

                return (
                  image.url ||
                  image.caption
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
   OUTLOOK-SAFE IMAGE
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
      '" style="text-decoration:none;">' +

      '<img src="' +
        escapeAttribute(
          imageUrl
        ) +
        '" alt="" width="100%" style="' +
          "display:block;" +
          "width:100%;" +
          "height:auto;" +
          "border:0;" +
          "outline:none;" +
          "text-decoration:none;" +
        '">' +

    "</a>";


  if (caption) {

    html +=

      '<p style="' +
        "margin-top:7px;" +
        "margin-right:0;" +
        "margin-bottom:0;" +
        "margin-left:0;" +
        "font-family:Arial,Helvetica,sans-serif;" +
        "font-size:12px;" +
        "line-height:17px;" +
        "color:" +
          COLORS.secondary +
        ";" +
      '">' +

        escapeHtml(
          caption
        ) +

      "</p>";

  }


  return html;

}


/* =========================================================
   FULL-WIDTH IMAGE
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

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="' +
      "border-collapse:collapse;" +
      "width:100%;" +
      "margin-top:30px;" +
      "margin-bottom:30px;" +
    '">' +

      "<tr>" +

        '<td style="' +
          "padding-top:8px;" +
          "padding-right:8px;" +
          "padding-bottom:8px;" +
          "padding-left:8px;" +
          "background:" +
            COLORS.surface +
          ";" +
        '">' +

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


/* =========================================================
   TWO / THREE / FOUR UP
========================================================= */

function buildImageRow(
  items,
  columns
) {

  if (
    !items ||
    !items.length
  ) {

    return "";

  }


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


  let cells =
    "";


  usable.forEach(
    function (item) {

      cells +=

        '<td valign="top" width="' +
          width +
          '%" style="' +
            "width:" +
            width +
            "%;" +
            "padding-top:3px;" +
            "padding-right:3px;" +
            "padding-bottom:3px;" +
            "padding-left:3px;" +
            "vertical-align:top;" +
            "background:" +
              COLORS.surface +
            ";" +
          '">' +

          buildEmailImage(
            item.url,
            item.clickUrl ||
              item.url,
            item.caption
          ) +

        "</td>";

    }
  );


  return (

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="' +
      "border-collapse:collapse;" +
      "width:100%;" +
      "table-layout:fixed;" +
      "margin-top:28px;" +
      "margin-bottom:28px;" +
    '">' +

      "<tr>" +

        cells +

      "</tr>" +

    "</table>"

  );

}


function buildFourImageRow(
  items
) {

  return buildImageRow(
    items,
    4
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
   SECTION LABEL
========================================================= */

function sectionLabel(
  text
) {

  return (

    '<p style="' +
      "margin-top:30px;" +
      "margin-right:0;" +
      "margin-bottom:11px;" +
      "margin-left:0;" +
      "font-family:Arial,Helvetica,sans-serif;" +
      "font-size:10px;" +
      "line-height:14px;" +
      "letter-spacing:1.5px;" +
      "color:" +
        COLORS.secondary +
      ";" +
    '">' +

      escapeHtml(
        text
      ) +

    "</p>"

  );

}


/* =========================================================
   OUTLOOK-SAFE NEWSLETTER HTML
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


  /* =======================================================
     HERO CLICK DESTINATIONS
  ======================================================= */

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


  /* =======================================================
     HEROES
  ======================================================= */

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


  /* =======================================================
     MODULAR IMAGES
  ======================================================= */

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


  /* =======================================================
     INVITATION
  ======================================================= */

  let invitationHtml =
    "";


  if (
    inviteTitle ||
    inviteText ||
    ctaUrl
  ) {

    invitationHtml +=
      sectionLabel(
        "05 — AN INVITATION"
      );


    if (inviteTitle) {

      invitationHtml +=

        '<p style="' +
          "margin-top:0;" +
          "margin-right:0;" +
          "margin-bottom:10px;" +
          "margin-left:0;" +
          "font-family:Georgia,Times New Roman,serif;" +
          "font-size:27px;" +
          "line-height:32px;" +
          "color:" +
            COLORS.text +
          ";" +
        '">' +

          escapeHtml(
            inviteTitle
          ) +

        "</p>";

    }


    invitationHtml +=
      paragraph(
        inviteText
      );


    if (ctaUrl) {

      invitationHtml +=

        '<table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;margin-bottom:25px;">' +

          "<tr>" +

            "<td>" +

              '<a href="' +
                escapeAttribute(
                  ctaUrl
                ) +
                '" style="' +
                  "display:inline-block;" +
                  "background:" +
                    COLORS.text +
                  ";" +
                  "color:" +
                    COLORS.background +
                  ";" +
                  "text-decoration:none;" +
                  "padding-top:12px;" +
                  "padding-right:18px;" +
                  "padding-bottom:12px;" +
                  "padding-left:18px;" +
                  "font-family:Arial,Helvetica,sans-serif;" +
                  "font-size:10px;" +
                  "line-height:12px;" +
                  "letter-spacing:1.2px;" +
                '">' +

                escapeHtml(
                  ctaLabel
                ) +

              "</a>" +

            "</td>" +

          "</tr>" +

        "</table>";

    }

  }


  /* =======================================================
     DATE LINE
  ======================================================= */

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


  /* =======================================================
     NEWSLETTER
  ======================================================= */

  let html = "";


  html +=

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="' +
      "border-collapse:collapse;" +
      "width:100%;" +
      "background:" +
        COLORS.background +
      ";" +
    '">' +

      "<tr>" +

        '<td align="center" style="' +
          "padding-top:28px;" +
          "padding-right:12px;" +
          "padding-bottom:28px;" +
          "padding-left:12px;" +
        '">' +


          '<table role="presentation" width="680" cellspacing="0" cellpadding="0" border="0" style="' +
            "border-collapse:collapse;" +
            "width:100%;" +
            "max-width:680px;" +
            "background:" +
              COLORS.background +
            ";" +
          '">' +


            /* =============================================
               HEADER
            ============================================= */

            "<tr>" +

              '<td style="' +
                "padding-top:42px;" +
                "padding-right:42px;" +
                "padding-bottom:24px;" +
                "padding-left:42px;" +
                "border-bottom:1px solid " +
                  COLORS.rule +
                ";" +
              '">' +


                '<p style="' +
                  "margin-top:0;" +
                  "margin-right:0;" +
                  "margin-bottom:0;" +
                  "margin-left:0;" +
                  "font-family:Arial,Helvetica,sans-serif;" +
                  "font-size:10px;" +
                  "line-height:14px;" +
                  "letter-spacing:2px;" +
                  "color:" +
                    COLORS.text +
                  ";" +
                '">' +

                  "FLRS GLOBAL" +

                "</p>" +


                '<p style="' +
                  "margin-top:8px;" +
                  "margin-right:0;" +
                  "margin-bottom:0;" +
                  "margin-left:0;" +
                  "font-family:Arial,Helvetica,sans-serif;" +
                  "font-size:10px;" +
                  "line-height:14px;" +
                  "letter-spacing:1.4px;" +
                  "color:" +
                    COLORS.secondary +
                  ";" +
                '">' +

                  "FROM THE STUDIO OF FREDDIE L. RANKIN II" +

                "</p>" +


                '<p style="' +
                  "margin-top:20px;" +
                  "margin-right:0;" +
                  "margin-bottom:10px;" +
                  "margin-left:0;" +
                  "font-family:Georgia,Times New Roman,serif;" +
                  "font-size:50px;" +
                  "line-height:48px;" +
                  "font-weight:normal;" +
                  "color:" +
                    COLORS.text +
                  ";" +
                '">' +

                  "The Pleasure Dispatch" +

                "</p>" +


                '<p style="' +
                  "margin-top:0;" +
                  "margin-right:0;" +
                  "margin-bottom:0;" +
                  "margin-left:0;" +
                  "font-family:Arial,Helvetica,sans-serif;" +
                  "font-size:10px;" +
                  "line-height:14px;" +
                  "letter-spacing:1.3px;" +
                  "color:" +
                    COLORS.secondary +
                  ";" +
                '">' +

                  dateLine +

                "</p>" +


              "</td>" +

            "</tr>" +


            /* =============================================
               CONTENT
            ============================================= */

            "<tr>" +

              '<td style="' +
                "padding-top:26px;" +
                "padding-right:42px;" +
                "padding-bottom:42px;" +
                "padding-left:42px;" +
              '">' +


                /* LOOP */

                '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;margin-bottom:28px;">' +

                  "<tr>" +

                    '<td align="center">' +

                      '<img src="' +
                        escapeAttribute(
                          LOOP_ASSET_URL
                        ) +
                        '" alt="The Loop" width="92" style="' +
                          "display:block;" +
                          "width:92px;" +
                          "height:auto;" +
                          "border:0;" +
                        '">' +

                    "</td>" +

                  "</tr>" +

                "</table>" +


                /* SUBTITLE */

                (
                  subtitle
                    ? '<p style="' +
                        "margin-top:0;" +
                        "margin-right:0;" +
                        "margin-bottom:28px;" +
                        "margin-left:0;" +
                        "font-family:Georgia,Times New Roman,serif;" +
                        "font-size:18px;" +
                        "line-height:26px;" +
                        "color:" +
                          COLORS.secondary +
                        ";" +
                      '">' +

                        escapeHtml(
                          subtitle
                        ) +

                      "</p>"
                    : ""
                ) +


                /* HERO 01 */

                (
                  hero1Html
                    ? '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;margin-bottom:8px;">' +

                        "<tr>" +

                          "<td>" +

                            hero1Html +

                          "</td>" +

                        "</tr>" +

                      "</table>"
                    : ""
                ) +


                /* REFLECTION */

                sectionLabel(
                  "01 — A REFLECTION"
                ) +

                paragraph(
                  reflection
                ) +


                /* WORK */

                sectionLabel(
                  "02 — THE WORK"
                ) +

                paragraph(
                  workText
                ) +


                /* MODULES */

                modulesHtml +


                /* STUDIO */

                sectionLabel(
                  "03 — STUDIO NOTES"
                ) +

                paragraph(
                  studioText
                ) +


                /* HERO 02 */

                (
                  hero2Html
                    ? '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;margin-top:8px;">' +

                        "<tr>" +

                          "<td>" +

                            hero2Html +

                          "</td>" +

                        "</tr>" +

                      "</table>"
                    : ""
                ) +


                /* PLEASURE NOTES */

                sectionLabel(
                  "04 — PLEASURE NOTES"
                ) +


                '<p style="' +
                  "margin-top:0;" +
                  "margin-right:0;" +
                  "margin-bottom:8px;" +
                  "margin-left:0;" +
                  "font-family:Georgia,Times New Roman,serif;" +
                  "font-size:19px;" +
                  "line-height:27px;" +
                  "color:" +
                    COLORS.text +
                  ";" +
                '">' +

                  "An offering of what has held my attention." +

                "</p>" +


                buildPleasureNotesHtml(
                  notes
                ) +


                /* INVITATION */

                invitationHtml +


                /* QUESTION */

                sectionLabel(
                  "06 — A QUESTION"
                ) +


                '<p style="' +
                  "margin-top:0;" +
                  "margin-right:0;" +
                  "margin-bottom:36px;" +
                  "margin-left:0;" +
                  "font-family:Georgia,Times New Roman,serif;" +
                  "font-size:25px;" +
                  "line-height:34px;" +
                  "color:" +
                    COLORS.text +
                  ";" +
                '">' +

                  escapeHtml(
                    question
                  ) +

                "</p>" +


                /* CLOSING LOOP */

                '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;margin-bottom:18px;">' +

                  "<tr>" +

                    '<td align="center">' +

                      '<img src="' +
                        escapeAttribute(
                          LOOP_ASSET_URL
                        ) +
                        '" alt="The Loop" width="66" style="' +
                          "display:block;" +
                          "width:66px;" +
                          "height:auto;" +
                          "border:0;" +
                        '">' +

                    "</td>" +

                  "</tr>" +

                "</table>" +


                '<p align="center" style="' +
                  "margin-top:0;" +
                  "margin-right:0;" +
                  "margin-bottom:10px;" +
                  "margin-left:0;" +
                  "font-family:Georgia,Times New Roman,serif;" +
                  "font-size:15px;" +
                  "line-height:21px;" +
                  "font-style:italic;" +
                  "color:" +
                    COLORS.secondary +
                  ";" +
                '">' +

                  "Pleasure is the desire to return." +

                "</p>" +


              "</td>" +

            "</tr>" +


            /* =============================================
               FOOTER
            ============================================= */

            "<tr>" +

              '<td style="' +
                "padding-top:18px;" +
                "padding-right:42px;" +
                "padding-bottom:34px;" +
                "padding-left:42px;" +
                "border-top:1px solid " +
                  COLORS.rule +
                ";" +
              '">' +

                '<p style="' +
                  "margin:0;" +
                  "font-family:Arial,Helvetica,sans-serif;" +
                  "font-size:10px;" +
                  "line-height:14px;" +
                  "letter-spacing:1.1px;" +
                  "color:" +
                    COLORS.secondary +
                  ";" +
                '">' +

                  "THE PLEASURE DISPATCH · BY FLRS GLOBAL" +

                "</p>" +

              "</td>" +

            "</tr>" +


          "</table>" +

        "</td>" +

      "</tr>" +

    "</table>";


  return html;

}


/* =========================================================
   BUILD PROGRESS
========================================================= */

function updateBuildProgress(
  current,
  label
) {

  setStatus(
    current +
    " / 4  " +
    label
  );

}


/* =========================================================
   OUTLOOK ERROR HANDLING
========================================================= */

function getAsyncError(
  result
) {

  if (
    result &&
    result.error
  ) {

    if (
      result.error.message
    ) {

      return result.error.message;

    }


    if (
      result.error.name
    ) {

      return result.error.name;

    }

  }


  return "Unknown Outlook error.";

}


/* =========================================================
   HTML DIAGNOSTICS
========================================================= */

function inspectNewsletterHtml(
  html
) {

  const diagnostics = {

    length:
      html.length,

    hasDiv:
      /<div\b/i.test(html),

    hasH1:
      /<h1\b/i.test(html),

    hasStyleTag:
      /<style\b/i.test(html),

    hasScript:
      /<script\b/i.test(html),

    hasDataUrl:
      /data:image/i.test(html),

    imageCount:
      (
        html.match(
          /<img\b/gi
        ) ||
        []
      ).length,

    tableCount:
      (
        html.match(
          /<table\b/gi
        ) ||
        []
      ).length

  };


  console.log(
    "Newsletter diagnostics:",
    diagnostics
  );


  return diagnostics;

}


/* =========================================================
   BUILD IN OUTLOOK
========================================================= */

function buildInOutlook() {

  updateBuildProgress(
    0,
    "Starting…"
  );


  if (
    typeof Office ===
    "undefined"
  ) {

    setStatus(
      "Build failed — Office.js is unavailable."
    );

    return;

  }


  if (
    !Office.context ||
    !Office.context.mailbox ||
    !Office.context.mailbox.item
  ) {

    setStatus(
      "Build failed — open The Pleasure Dispatch from a new Outlook message."
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
      "Build failed — Outlook body access is unavailable."
    );

    return;

  }


  /* =======================================================
     STEP 1
  ======================================================= */

  updateBuildProgress(
    1,
    "Generating newsletter…"
  );


  let html;


  try {

    html =
      buildNewsletterHtml();

  } catch (error) {

    console.error(
      "Newsletter generation error:",
      error
    );


    setStatus(
      "Build failed generating newsletter: " +
      error.message
    );


    return;

  }


  if (
    !html ||
    typeof html !==
      "string"
  ) {

    setStatus(
      "Build failed — newsletter HTML is empty."
    );

    return;

  }


  inspectNewsletterHtml(
    html
  );


  console.log(
    "Newsletter HTML length:",
    html.length
  );


  /* =======================================================
     STEP 2
  ======================================================= */

  updateBuildProgress(
    2,
    "Checking image sources…"
  );


  const modules =
    collectImageBlocks();


  let modularCount =
    0;


  modules.forEach(
    function (block) {

      modularCount +=
        block.items.length;

    }
  );


  let heroCount =
    0;


  if (
    value("hero1Url")
  ) {

    heroCount++;

  }


  if (
    value("hero2Url")
  ) {

    heroCount++;

  }


  const totalImages =
    modularCount +
    heroCount;


  updateBuildProgress(
    2,
    totalImages +
    " image" +
    (
      totalImages ===
      1
        ? ""
        : "s"
    ) +
    " ready…"
  );


  console.log(
    "Total images:",
    totalImages
  );


  /* =======================================================
     STEP 3
     READ BODY
  ======================================================= */

  updateBuildProgress(
    3,
    "Preparing Outlook message…"
  );


  let finished =
    false;


  const timeout =
    setTimeout(
      function () {

        if (finished) {
          return;
        }


        finished =
          true;


        setStatus(
          "Build timed out — Outlook did not respond after 90 seconds."
        );


        console.error(
          "Outlook operation timed out."
        );

      },
      BUILD_TIMEOUT_MS
    );


  item.body.getAsync(
    Office.CoercionType.Html,
    function (
      bodyResult
    ) {

      if (finished) {
        return;
      }


      if (
        !bodyResult ||
        bodyResult.status !==
          Office.AsyncResultStatus.Succeeded
      ) {

        finished =
          true;


        clearTimeout(
          timeout
        );


        setStatus(
          "Build failed reading Outlook body: " +
          getAsyncError(
            bodyResult
          )
        );


        return;

      }


      console.log(
        "Current Outlook body length:",
        (
          bodyResult.value ||
          ""
        ).length
      );


      /* ===================================================
         SUBJECT
      =================================================== */

      updateBuildProgress(
        3,
        "Setting subject…"
      );


      const subject =
        buildSubject();


      if (
        !item.subject ||
        typeof item.subject.setAsync !==
          "function"
      ) {

        finished =
          true;

        clearTimeout(
          timeout
        );


        setStatus(
          "Build failed — Outlook subject access is unavailable."
        );


        return;

      }


      item.subject.setAsync(
        subject,
        function (
          subjectResult
        ) {

          if (finished) {
            return;
          }


          if (
            !subjectResult ||
            subjectResult.status !==
              Office.AsyncResultStatus.Succeeded
          ) {

            finished =
              true;


            clearTimeout(
              timeout
            );


            setStatus(
              "Build failed setting subject: " +
              getAsyncError(
                subjectResult
              )
            );


            return;

          }


          console.log(
            "Subject set successfully."
          );


          /* ===============================================
             BODY
          =============================================== */

          updateBuildProgress(
            3,
            "Writing newsletter body…"
          );


          item.body.setAsync(
            html,
            {
              coercionType:
                Office.CoercionType.Html
            },
            function (
              bodySetResult
            ) {

              if (finished) {
                return;
              }


              finished =
                true;


              clearTimeout(
                timeout
              );


              if (
                bodySetResult &&
                bodySetResult.status ===
                  Office.AsyncResultStatus.Succeeded
              ) {

                updateBuildProgress(
                  4,
                  "✓ Complete — Dispatch built in Outlook."
                );


                console.log(
                  "Newsletter body inserted successfully."
                );


                return;

              }


              const errorMessage =
                getAsyncError(
                  bodySetResult
                );


              setStatus(
                "Build failed writing newsletter: " +
                errorMessage
              );


              console.error(
                "body.setAsync result:",
                bodySetResult
              );


              /*
               * Second diagnostic:
               * attempt an Outlook-safe emergency
               * body if the full newsletter fails.
               */
              console.warn(
                "Full newsletter rejected by Outlook."
              );


              console.warn(
                "Newsletter HTML length:",
                html.length
              );


              console.warn(
                "Trying emergency fallback body..."
              );


              const fallback =
                buildEmergencyFallback();


              item.body.setAsync(
                fallback,
                {
                  coercionType:
                    Office.CoercionType.Html
                },
                function (
                  fallbackResult
                ) {

                  if (
                    fallbackResult &&
                    fallbackResult.status ===
                      Office.AsyncResultStatus.Succeeded
                  ) {

                    setStatus(
                      "Outlook rejected the full newsletter. A diagnostic fallback was inserted instead. Check the console."
                    );


                    console.log(
                      "Emergency fallback inserted successfully."
                    );


                  } else {

                    setStatus(
                      "Outlook rejected both the newsletter and the fallback body. Check the console."
                    );


                    console.error(
                      "Fallback result:",
                      fallbackResult
                    );

                  }

                }
              );

            }
          );

        }
      );

    }
  );

}


/* =========================================================
   EMERGENCY FALLBACK
========================================================= */

function buildEmergencyFallback() {

  return (

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">' +

      "<tr>" +

        '<td style="padding:30px;font-family:Arial,Helvetica,sans-serif;">' +

          "<strong>The Pleasure Dispatch</strong>" +

          "<br><br>" +

          "The newsletter HTML was rejected by Outlook." +

          "<br>" +

          "The add-in is functioning, but Outlook's body renderer rejected the full HTML payload." +

        "</td>" +

      "</tr>" +

    "</table>"

  );

}


/* =========================================================
   SUBJECT
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


/* =========================================================
   PREVIEW
========================================================= */

function previewNewsletter() {

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

          "body{" +
            "margin:0;" +
            "background:" +
              COLORS.background +
            ";" +
          "}" +

          "img{" +
            "cursor:pointer;" +
          "}" +

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
```
