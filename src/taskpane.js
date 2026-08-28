/* =========================================================
   THE PLEASURE DISPATCH
   taskpane.js
   Stable Production Version
   Outlook-safe body insertion
========================================================= */

const DRIVE_API_URL =
  "https://script.google.com/macros/s/AKfycbyTLvYBe1O_BbzsH09UMSZdbY9_XZXga-TbSPkR3UclT3Qhlaj7gy5yhXPA_UpE6Fu7tw/exec";

const LOOP_ASSET_URL =
  "https://flrsgloba.github.io/dispatch.flrsglobal.com/assets/pleasure-loop.svg";

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


/* =========================================================
   OFFICE INITIALIZATION
========================================================= */

Office.onReady(function () {

  console.log(
    "[Pleasure Dispatch] Office.js ready."
  );

  initializeDispatch();

});


function initializeDispatch() {

  if (initialized) {
    return;
  }

  initialized = true;

  console.log(
    "[Pleasure Dispatch] Initializing UI..."
  );

  bindStaticControls();
  bindDelegatedControls();

  setupHero("hero1");
  setupHero("hero2");

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
   BASIC HELPERS
========================================================= */

function value(id) {

  const element =
    document.getElementById(id);

  if (!element) {
    return "";
  }

  return String(
    element.value || ""
  ).trim();

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
      "font-size:17px;" +
      "line-height:1.7;" +
      "font-family:Georgia,Times New Roman,serif;" +
      "margin:0 0 24px;" +
      "color:" +
        COLORS.text +
      ";" +
    '">' +

      escapeHtml(text).replace(
        /\n/g,
        "<br>"
      ) +

    "</p>"
  );

}


function setStatus(message) {

  const status =
    document.getElementById("status");

  if (status) {
    status.textContent = message;
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

  console.log(
    "[Pleasure Dispatch] Binding static controls."
  );


  const previewButton =
    document.getElementById("previewBtn");

  const buildButton =
    document.getElementById("insertBtn");

  const addPleasureButton =
    document.getElementById("addPleasure");

  let addImageButton =
    document.getElementById("addImageBlock");


  if (previewButton) {

    previewButton.onclick =
      function (event) {

        event.preventDefault();
        event.stopPropagation();

        console.log(
          "Preview clicked."
        );

        previewNewsletter();

      };

  } else {

    console.error(
      "previewBtn not found."
    );

  }


  if (buildButton) {

    buildButton.onclick =
      function (event) {

        event.preventDefault();
        event.stopPropagation();

        console.log(
          "Build clicked."
        );

        buildInOutlook();

      };

  } else {

    console.error(
      "insertBtn not found."
    );

  }


  if (addPleasureButton) {

    addPleasureButton.onclick =
      function (event) {

        event.preventDefault();
        event.stopPropagation();

        addPleasureRow(
          "",
          ""
        );

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

  } else {

    console.error(
      "Add Image Block control not found."
    );

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


      /* MOVE UP */

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


      /* MOVE DOWN */

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


      /* DUPLICATE */

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


      /* REMOVE MODULE */

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


      /* REMOVE IMAGE */

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


      /* URL BUTTON */

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


  /* CHANGE */

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


  /* INPUT */

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

function buildDriveImageUrl(fileId) {

  return (
    "https://drive.google.com/thumbnail?id=" +
    encodeURIComponent(fileId) +
    "&sz=w" +
    DRIVE_IMAGE_WIDTH
  );

}


function buildDriveClickUrl(fileId) {

  return (
    "https://drive.google.com/file/d/" +
    encodeURIComponent(fileId) +
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
      /^image\/(jpeg|jpg|png|webp)$/i
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
                function (dataUrl) {

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


  console.log(
    "Drive upload payload size:",
    payload.fileContent.length
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
    dataUrl.indexOf(",");


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
   HERO HANDLING
========================================================= */

function setupHero(key) {

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


      if (!input) {

        setStatus(
          "Hero URL field not found."
        );

        return;

      }


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

        console.error(
          "Hero upload failed:",
          error
        );


        setStatus(
          "Preview ready. Drive upload failed: " +
          error.message
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

        console.error(
          "Modular upload error:",
          error
        );


        setStatus(
          "Preview ready. Drive upload failed: " +
          error.message
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
   DUPLICATE
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


  if (container) {

    container.insertBefore(
      clone,
      block.nextSibling
    );

  }


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


  let html =
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;margin:4px 0 30px;">';


  notes.forEach(
    function (note) {

      html +=

        "<tr>" +

          '<td width="120" valign="top" style="' +
            "width:120px;" +
            "padding:5px 18px 5px 0;" +
            "vertical-align:top;" +
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
            "padding:5px 0;" +
            "vertical-align:top;" +
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


  html +=
    "</table>";


  return html;

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
          block.items.length > 0
        );

      }
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
        '">' +

    "</a>";


  if (caption) {

    html +=

      '<div style="' +
        "font-family:Arial,Helvetica,sans-serif;" +
        "font-size:12px;" +
        "line-height:17px;" +
        "color:" +
          COLORS.secondary +
        ";" +
        "padding-top:7px;" +
      '">' +

        escapeHtml(
          caption
        ) +

      "</div>";

  }


  return html;

}


/* =========================================================
   FULL WIDTH IMAGE
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

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;margin:30px 0;background:' +
      COLORS.surface +
    ';">' +

      "<tr>" +

        '<td style="padding:8px;">' +

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
   IMAGE ROW
========================================================= */

function buildImageRow(
  items,
  columns
) {

  const usable =
    items.slice(
      0,
      columns
    );


  if (!usable.length) {
    return "";
  }


  const width =
    Math.floor(
      100 /
      columns
    );


  let html =

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;margin:28px 0;">' +

      "<tr>";


  usable.forEach(
    function (item) {

      html +=

        '<td width="' +
          width +
          '%" valign="top" style="' +
            "width:" +
            width +
            "%;" +
            "padding:3px;" +
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


  html +=

      "</tr>" +

    "</table>";


  return html;

}


function buildFourImageRow(
  items
) {

  return buildImageRow(
    items,
    4
  );

}


/* =========================================================
   IMAGE MODULE HTML
========================================================= */

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
   OUTLOOK-SAFE NEWSLETTER HTML
========================================================= */

/*
 * IMPORTANT:
 *
 * This function intentionally uses conservative HTML.
 *
 * Outlook's body.setAsync() is much less forgiving than a
 * browser. We therefore avoid:
 *
 * - CSS classes
 * - style blocks
 * - flexbox
 * - grid
 * - SVG inside the newsletter
 * - JavaScript
 * - <div> as structural containers where tables work
 * - unsupported CSS
 *
 * All important styling is inline.
 */

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


  let invitationHtml =
    "";


  if (
    inviteTitle ||
    inviteText ||
    ctaUrl
  ) {

    invitationHtml =

      '<tr>' +

        '<td style="padding:30px 0 10px;">' +

          '<div style="' +
            "font-family:Arial,Helvetica,sans-serif;" +
            "font-size:10px;" +
            "line-height:14px;" +
            "letter-spacing:1.5px;" +
            "color:" +
              COLORS.secondary +
            ";" +
          '">' +

            "05 — AN INVITATION" +

          "</div>" +

        "</td>" +

      "</tr>";


    if (inviteTitle) {

      invitationHtml +=

        "<tr>" +

          '<td style="padding:0 0 10px;">' +

            '<div style="' +
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

            "</div>" +

          "</td>" +

        "</tr>";

    }


    if (inviteText) {

      invitationHtml +=

        "<tr>" +

          '<td style="padding:0 0 10px;">' +

            paragraph(
              inviteText
            ) +

          "</td>" +

        "</tr>";

    }


    if (ctaUrl) {

      invitationHtml +=

        "<tr>" +

          '<td style="padding:0 0 25px;">' +

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
                "padding:12px 18px;" +
                "font-family:Arial,Helvetica,sans-serif;" +
                "font-size:10px;" +
                "line-height:14px;" +
                "letter-spacing:1.2px;" +
              '">' +

              escapeHtml(
                ctaLabel
              ) +

            "</a>" +

          "</td>" +

        "</tr>";

    }

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


  let html = "";


  /*
   * OUTER TABLE
   */

  html +=

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;background:' +
      COLORS.background +
    ';">' +

      "<tr>" +

        '<td align="center" style="padding:20px 10px;">' +

          '<table cellpadding="0" cellspacing="0" border="0" width="680" style="border-collapse:collapse;width:100%;max-width:680px;background:' +
            COLORS.background +
          ';">';


  /*
   * HEADER
   */

  html +=

    "<tr>" +

      '<td style="padding:38px 32px 24px;border-bottom:1px solid ' +
        COLORS.rule +
      ';">' +

        '<div style="' +
          "font-family:Arial,Helvetica,sans-serif;" +
          "font-size:10px;" +
          "line-height:14px;" +
          "letter-spacing:2px;" +
          "color:" +
            COLORS.text +
          ";" +
        '">' +

          "FLRS GLOBAL" +

        "</div>" +

        '<div style="' +
          "font-family:Arial,Helvetica,sans-serif;" +
          "font-size:10px;" +
          "line-height:14px;" +
          "letter-spacing:1.4px;" +
          "color:" +
            COLORS.secondary +
          ";" +
          "padding-top:8px;" +
        '">' +

          "FROM THE STUDIO OF FREDDIE L. RANKIN II" +

        "</div>" +

        '<div style="' +
          "font-family:Georgia,Times New Roman,serif;" +
          "font-size:46px;" +
          "line-height:48px;" +
          "font-weight:normal;" +
          "color:" +
            COLORS.text +
          ";" +
          "padding-top:20px;" +
        '">' +

          "The Pleasure Dispatch" +

        "</div>" +

        '<div style="' +
          "font-family:Arial,Helvetica,sans-serif;" +
          "font-size:10px;" +
          "line-height:14px;" +
          "letter-spacing:1.3px;" +
          "color:" +
            COLORS.secondary +
          ";" +
          "padding-top:10px;" +
        '">' +

          dateLine +

        "</div>" +

      "</td>" +

    "</tr>";


  /*
   * CONTENT TABLE
   */

  html +=

    "<tr>" +

      '<td style="padding:26px 32px 40px;">' +

        '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">';


  /*
   * SUBTITLE
   */

  if (subtitle) {

    html +=

      "<tr>" +

        '<td style="padding:0 0 28px;">' +

          '<div style="' +
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

          "</div>" +

        "</td>" +

      "</tr>";

  }


  /*
   * HERO 01
   */

  if (hero1) {

    html +=

      "<tr>" +

        '<td style="padding:0 0 8px;">' +

          buildEmailImage(
            hero1,
            hero1Click,
            hero1Caption
          ) +

        "</td>" +

      "</tr>";

  }


  /*
   * REFLECTION
   */

  html +=

    "<tr>" +

      '<td style="padding:30px 0 11px;">' +

        '<div style="' +
          "font-family:Arial,Helvetica,sans-serif;" +
          "font-size:10px;" +
          "line-height:14px;" +
          "letter-spacing:1.5px;" +
          "color:" +
            COLORS.secondary +
          ";" +
        '">' +

          "01 — A REFLECTION" +

        "</div>" +

      "</td>" +

    "</tr>" +

    "<tr>" +

      '<td>' +

        paragraph(
          reflection
        ) +

      "</td>" +

    "</tr>";


  /*
   * THE WORK
   */

  html +=

    "<tr>" +

      '<td style="padding:14px 0 11px;">' +

        '<div style="' +
          "font-family:Arial,Helvetica,sans-serif;" +
          "font-size:10px;" +
          "line-height:14px;" +
          "letter-spacing:1.5px;" +
          "color:" +
            COLORS.secondary +
          ";" +
        '">' +

          "02 — THE WORK" +

        "</div>" +

      "</td>" +

    "</tr>" +

    "<tr>" +

      '<td>' +

        paragraph(
          workText
        ) +

      "</td>" +

    "</tr>";


  /*
   * MODULAR IMAGES
   */

  if (modulesHtml) {

    html +=

      "<tr>" +

        '<td>' +

          modulesHtml +

        "</td>" +

      "</tr>";

  }


  /*
   * STUDIO NOTES
   */

  html +=

    "<tr>" +

      '<td style="padding:14px 0 11px;">' +

        '<div style="' +
          "font-family:Arial,Helvetica,sans-serif;" +
          "font-size:10px;" +
          "line-height:14px;" +
          "letter-spacing:1.5px;" +
          "color:" +
            COLORS.secondary +
          ";" +
        '">' +

          "03 — STUDIO NOTES" +

        "</div>" +

      "</td>" +

    "</tr>" +

    "<tr>" +

      '<td>' +

        paragraph(
          studioText
        ) +

      "</td>" +

    "</tr>";


  /*
   * HERO 02
   */

  if (hero2) {

    html +=

      "<tr>" +

        '<td style="padding:8px 0 0;">' +

          buildEmailImage(
            hero2,
            hero2Click,
            hero2Caption
          ) +

        "</td>" +

      "</tr>";

  }


  /*
   * PLEASURE NOTES
   */

  html +=

    "<tr>" +

      '<td style="padding:38px 0 11px;">' +

        '<div style="' +
          "font-family:Arial,Helvetica,sans-serif;" +
          "font-size:10px;" +
          "line-height:14px;" +
          "letter-spacing:1.5px;" +
          "color:" +
            COLORS.secondary +
          ";" +
        '">' +

          "04 — PLEASURE NOTES" +

        "</div>" +

      "</td>" +

    "</tr>" +

    "<tr>" +

      '<td style="padding:0 0 8px;">' +

        '<div style="' +
          "font-family:Georgia,Times New Roman,serif;" +
          "font-size:19px;" +
          "line-height:27px;" +
          "color:" +
            COLORS.text +
          ";" +
        '">' +

          "An offering of what has held my attention." +

        "</div>" +

      "</td>" +

    "</tr>" +

    "<tr>" +

      "<td>" +

        buildPleasureNotesHtml(
          notes
        ) +

      "</td>" +

    "</tr>";


  /*
   * INVITATION
   */

  html +=
    invitationHtml;


  /*
   * QUESTION
   */

  html +=

    "<tr>" +

      '<td style="padding:38px 0 11px;">' +

        '<div style="' +
          "font-family:Arial,Helvetica,sans-serif;" +
          "font-size:10px;" +
          "line-height:14px;" +
          "letter-spacing:1.5px;" +
          "color:" +
            COLORS.secondary +
          ";" +
        '">' +

          "06 — A QUESTION" +

        "</div>" +

      "</td>" +

    "</tr>" +

    "<tr>" +

      '<td style="padding:0 0 36px;">' +

        '<div style="' +
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

        "</div>" +

      "</td>" +

    "</tr>";


  /*
   * CLOSING
   *
   * SVG intentionally removed from Outlook build.
   */

  html +=

    "<tr>" +

      '<td align="center" style="padding:30px 0 18px;">' +

        '<div style="' +
          "font-family:Georgia,Times New Roman,serif;" +
          "font-size:24px;" +
          "line-height:30px;" +
          "color:" +
            COLORS.secondary +
          ";" +
        '">' +

          "◌" +

        "</div>" +

      "</td>" +

    "</tr>" +

    "<tr>" +

      '<td align="center" style="padding:0 0 10px;">' +

        '<div style="' +
          "font-family:Georgia,Times New Roman,serif;" +
          "font-size:15px;" +
          "line-height:21px;" +
          "font-style:italic;" +
          "color:" +
            COLORS.secondary +
          ";" +
        '">' +

          "Pleasure is the desire to return." +

        "</div>" +

      "</td>" +

    "</tr>";


  html +=

          "</table>" +

        "</td>" +

      "</tr>";


  /*
   * FOOTER
   */

  html +=

    "<tr>" +

      '<td style="padding:18px 32px 34px;border-top:1px solid ' +
        COLORS.rule +
      ';">' +

        '<div style="' +
          "font-family:Arial,Helvetica,sans-serif;" +
          "font-size:10px;" +
          "line-height:14px;" +
          "letter-spacing:1.1px;" +
          "color:" +
            COLORS.secondary +
          ";" +
        '">' +

          "THE PLEASURE DISPATCH · BY FLRS GLOBAL" +

        "</div>" +

      "</td>" +

    "</tr>";


  html +=

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
   OUTLOOK ERROR
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
   HTML VALIDATION
========================================================= */

function validateNewsletterHtml(
  html
) {

  if (
    !html ||
    typeof html !== "string"
  ) {

    return {
      valid: false,
      message: "Newsletter HTML is empty."
    };

  }


  const checks = [

    {
      name: "script",
      pattern: /<script\b/i,
      message: "HTML contains a script tag."
    },

    {
      name: "svg",
      pattern: /<svg\b/i,
      message: "HTML contains SVG markup."
    },

    {
      name: "style",
      pattern: /<style\b/i,
      message: "HTML contains a style block."
    },

    {
      name: "javascript",
      pattern: /javascript:/i,
      message: "HTML contains JavaScript."
    },

    {
      name: "data-image",
      pattern: /data:image\//i,
      message: "HTML contains an embedded data image."
    }

  ];


  for (
    let i = 0;
    i < checks.length;
    i++
  ) {

    if (
      checks[i].pattern.test(
        html
      )
    ) {

      return {
        valid: false,
        message:
          checks[i].message
      };

    }

  }


  return {
    valid: true,
    message: "HTML passed Outlook safety checks."
  };

}


/* =========================================================
   BUILD IN OUTLOOK
========================================================= */

function buildInOutlook() {

  console.log(
    "===================================="
  );

  console.log(
    "THE PLEASURE DISPATCH BUILD"
  );

  console.log(
    "===================================="
  );


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
      "Newsletter build error:",
      error
    );


    setStatus(
      "Build failed generating newsletter: " +
      error.message
    );

    return;

  }


  console.log(
    "Newsletter HTML length:",
    html.length
  );


  /* =======================================================
     VALIDATE HTML
  ======================================================= */

  const validation =
    validateNewsletterHtml(
      html
    );


  console.log(
    "HTML validation:",
    validation
  );


  if (
    !validation.valid
  ) {

    setStatus(
      "Build stopped — " +
      validation.message
    );

    return;

  }


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


  console.log(
    "Hero images:",
    heroCount
  );

  console.log(
    "Modular images:",
    modularCount
  );

  console.log(
    "Total images:",
    totalImages
  );


  updateBuildProgress(
    2,
    totalImages +
    " image" +
    (
      totalImages === 1
        ? ""
        : "s"
    ) +
    " ready…"
  );


  /* =======================================================
     STEP 3
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
          "Build timed out — Outlook did not respond after 60 seconds."
        );


        console.error(
          "Outlook operation timed out."
        );

      },
      BUILD_TIMEOUT_MS
    );


  function finish() {

    if (finished) {
      return false;
    }

    finished = true;

    clearTimeout(
      timeout
    );

    return true;

  }


  /* =======================================================
     SUBJECT
  ======================================================= */

  const subject =
    buildSubject();


  console.log(
    "Subject:",
    subject
  );


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

        finish();


        setStatus(
          "Build failed setting subject: " +
          getAsyncError(
            subjectResult
          )
        );


        console.error(
          "subject.setAsync result:",
          subjectResult
        );


        return;

      }


      console.log(
        "Subject set successfully."
      );


      /* ===================================================
         WRITE BODY
      =================================================== */

      updateBuildProgress(
        3,
        "Writing newsletter body…"
      );


      console.log(
        "Calling body.setAsync..."
      );


      try {

        item.body.setAsync(
          html,
          {
            coercionType:
              Office.CoercionType.Html
          },
          function (
            bodyResult
          ) {

            if (finished) {
              return;
            }


            if (
              bodyResult &&
              bodyResult.status ===
                Office.AsyncResultStatus.Succeeded
            ) {

              finish();


              updateBuildProgress(
                4,
                "✓ Complete — Dispatch built in Outlook."
              );


              console.log(
                "body.setAsync succeeded."
              );


              return;

            }


            finish();


            const errorMessage =
              getAsyncError(
                bodyResult
              );


            setStatus(
              "Build failed writing newsletter: " +
              errorMessage
            );


            console.error(
              "body.setAsync result:",
              bodyResult
            );


            console.error(
              "Rejected HTML length:",
              html.length
            );


            console.error(
              "First 2000 characters:",
              html.substring(
                0,
                2000
              )
            );

          }
        );

      } catch (error) {

        finish();


        setStatus(
          "Build failed writing newsletter: " +
          error.message
        );


        console.error(
          "body.setAsync threw exception:",
          error
        );

      }

    }
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
   MINIMAL OUTLOOK BODY TEST
========================================================= */

function testOutlookBody() {

  console.log(
    "Running minimal Outlook body test..."
  );


  if (
    typeof Office ===
    "undefined"
  ) {

    setStatus(
      "Office.js unavailable."
    );

    return;

  }


  if (
    !Office.context ||
    !Office.context.mailbox ||
    !Office.context.mailbox.item
  ) {

    setStatus(
      "Open a new Outlook message first."
    );

    return;

  }


  const item =
    Office.context.mailbox.item;


  const testHtml =

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">' +

      "<tr>" +

        '<td style="padding:30px;font-family:Arial,sans-serif;font-size:20px;color:#151515;">' +

          "The Pleasure Dispatch body test." +

        "</td>" +

      "</tr>" +

    "</table>";


  try {

    item.body.setAsync(
      testHtml,
      {
        coercionType:
          Office.CoercionType.Html
      },
      function (
        result
      ) {

        console.log(
          "Minimal body test result:",
          result
        );


        if (
          result &&
          result.status ===
            Office.AsyncResultStatus.Succeeded
        ) {

          setStatus(
            "✓ Outlook body test succeeded."
          );

        } else {

          setStatus(
            "Outlook body test failed: " +
            getAsyncError(
              result
            )
          );

        }

      }
    );

  } catch (error) {

    console.error(
      "Minimal body test exception:",
      error
    );


    setStatus(
      "Outlook body test failed: " +
      error.message
    );

  }

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

    console.error(
      "Preview error:",
      error
    );


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
