/* =========================================================
   THE PLEASURE DISPATCH
   taskpane.js
   Consolidated Production Version
========================================================= */

const DRIVE_API_URL =
  "https://script.google.com/macros/s/AKfycbyTLvYbe1O_BbzsH09UMSZdbY9_XZXga-TbSPkR3UclT3Qhlaj7gy5yhXPA_UpE6Fu7tw/exec";

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
const BUILD_TIMEOUT_MS = 15000;

let blockCounter = 0;
let pleasureCounter = 0;
let initialized = false;


/* =========================================================
   OFFICE
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

  return element
    ? element.value.trim()
    : "";
}


function escapeHtml(text) {

  return String(text || "").replace(
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
    '<p style="font:17px/1.75 Garamond,Georgia,Times New Roman,serif;' +
    'margin:0 0 24px;' +
    'color:' + COLORS.text + ';">' +
    escapeHtml(text).replace(/\n/g, "<br>") +
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


  /*
   * Find Add Image Block by ID first.
   * If missing, search by visible text.
   */

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


    console.log(
      "Add Image Block bound."
    );

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

  /*
   * BUTTONS
   */

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


      /*
       * Move Up.
       */

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


      /*
       * Move Down.
       */

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


      /*
       * Duplicate.
       */

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


      /*
       * Remove module.
       */

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


      /*
       * Remove image.
       */

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


      /*
       * Image URL button.
       */

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

      }

    }
  );


  /*
   * SELECTS + FILES
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


        setStatus(
          "Layout changed to " +
          layout.options[
            layout.selectedIndex
          ].text +
          "."
        );


        return;

      }


      /*
       * Modular upload.
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

      }

    }
  );


  /*
   * URL INPUTS
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
    encodeURIComponent(fileId) +
    "&sz=w" +
    DRIVE_IMAGE_WIDTH
  );

}


function buildDriveClickUrl(
  fileId
) {

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
              IMAGE_MAX_WIDTH / width,
              IMAGE_MAX_HEIGHT / height
            );


          width =
            Math.round(
              width * scale
            );

          height =
            Math.round(
              height * scale
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


  return comma === -1
    ? dataUrl
    : dataUrl.substring(
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
   HERO IMAGES
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
   MODULAR IMAGES
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
          "Could not find the modular image field."
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
         * Drive data.
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
         * Replace local source.
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


      } catch (error) {

        setStatus(
          "Preview ready. Drive upload failed: " +
          error.message
        );


        console.error(
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


    console.error(
      "#imageBlocks was not found."
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


  const count =
    layout === "full"
      ? 1
      : layout === "two"
        ? 2
        : layout === "three"
          ? 3
          : 4;


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
                "color:" +
                  COLORS.secondary +
                ";" +
              '">' +

                escapeHtml(
                  note.label
                ) +

              "</td>" +

              '<td style="' +
                "vertical-align:top;" +
                "padding:5px 0;" +
                "font:16px/1.5 Garamond,Georgia,Times New Roman,serif;" +
                "color:" +
                  COLORS.text +
                ";" +
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
        "color:" +
          COLORS.secondary +
        ";" +
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

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="' +
      "border-collapse:collapse;" +
      "margin:30px 0;" +
      "background:" +
        COLORS.surface +
      ";" +
    '">' +

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
   * Hero click targets.
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
    inviteTitle ||
    inviteText ||
    ctaUrl
  ) {

    invitationHtml =

      '<div style="' +
        "font:10px Arial,Helvetica,sans-serif;" +
        "letter-spacing:1.5px;" +
        "color:" +
          COLORS.secondary +
        ";" +
        "margin:40px 0 11px;" +
      '">' +

        "05 — AN INVITATION" +

      "</div>" +

      (
        inviteTitle
          ? '<div style="' +
              "font:27px/1.15 Garamond,Georgia,Times New Roman,serif;" +
              "color:" +
                COLORS.text +
              ";" +
              "margin:0 0 10px;" +
            '">' +

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
                  "background:" +
                    COLORS.text +
                  ";" +
                  "color:" +
                    COLORS.background +
                  ";" +
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
   * Complete newsletter.
   */

  return (

    '<div style="' +
      "margin:0;" +
      "padding:0;" +
      "background:" +
        COLORS.background +
      ";" +
      "color:" +
        COLORS.text +
      ";" +
    '">' +


      '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="' +
        "border-collapse:collapse;" +
        "background:" +
          COLORS.background +
        ";" +
      '">' +


        "<tr>" +


          '<td align="center" style="padding:28px 12px;">' +


            '<table role="presentation" width="680" cellspacing="0" cellpadding="0" border="0" style="' +
              "border-collapse:collapse;" +
              "width:100%;" +
              "max-width:680px;" +
              "background:" +
                COLORS.background +
              ";" +
            '">' +


              /*
               * HEADER
               */

              "<tr>" +

                '<td style="' +
                  "padding:42px 42px 24px;" +
                  "border-bottom:1px solid " +
                    COLORS.rule +
                  ";" +
                '">' +

                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:2px;" +
                    "color:" +
                      COLORS.text +
                    ";" +
                  '">' +

                    "FLRS GLOBAL" +

                  "</div>" +


                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.4px;" +
                    "color:" +
                      COLORS.secondary +
                    ";" +
                    "margin-top:8px;" +
                  '">' +

                    "FROM THE STUDIO OF FREDDIE L. RANKIN II" +

                  "</div>" +


                  '<h1 style="' +
                    "font:400 50px/0.96 Garamond,Georgia,Times New Roman,serif;" +
                    "color:" +
                      COLORS.text +
                    ";" +
                    "margin:20px 0 10px;" +
                  '">' +

                    "The Pleasure Dispatch" +

                  "</h1>" +


                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.3px;" +
                    "color:" +
                      COLORS.secondary +
                    ";" +
                  '">' +

                    dateLine +

                  "</div>" +

                "</td>" +

              "</tr>" +


              /*
               * CONTENT
               */

              "<tr>" +

                '<td style="padding:26px 42px 42px;">' +


                  /*
                   * THE LOOP
                   */

                  '<div style="text-align:center;margin:0 0 28px;">' +

                    '<img src="' +
                      LOOP_ASSET_URL +
                      '" alt="The Loop" width="92" style="' +
                        "display:inline-block;" +
                        "width:92px;" +
                        "height:auto;" +
                        "border:0;" +
                      '">' +

                  "</div>" +


                  /*
                   * SUBTITLE
                   */

                  (
                    subtitle
                      ? '<div style="' +
                          "font:18px/1.45 Garamond,Georgia,Times New Roman,serif;" +
                          "color:" +
                            COLORS.secondary +
                          ";" +
                          "margin:0 0 28px;" +
                        '">' +

                          escapeHtml(
                            subtitle
                          ) +

                        "</div>"
                      : ""
                  ) +


                  /*
                   * HERO 01
                   */

                  (
                    hero1Html
                      ? '<div style="margin-bottom:8px;">' +
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
                    "color:" +
                      COLORS.secondary +
                    ";" +
                    "margin:30px 0 11px;" +
                  '">' +

                    "01 — A REFLECTION" +

                  "</div>" +

                  paragraph(
                    reflection
                  ) +


                  /*
                   * WORK
                   */

                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.5px;" +
                    "color:" +
                      COLORS.secondary +
                    ";" +
                    "margin:38px 0 11px;" +
                  '">' +

                    "02 — THE WORK" +

                  "</div>" +

                  paragraph(
                    workText
                  ) +


                  /*
                   * MODULAR IMAGE BLOCKS
                   */

                  modulesHtml +


                  /*
                   * STUDIO NOTES
                   */

                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.5px;" +
                    "color:" +
                      COLORS.secondary +
                    ";" +
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
                      ? '<div style="margin:8px 0 0;">' +
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
                    "color:" +
                      COLORS.secondary +
                    ";" +
                    "margin:38px 0 11px;" +
                  '">' +

                    "04 — PLEASURE NOTES" +

                  "</div>" +


                  '<div style="' +
                    "font:19px/1.4 Garamond,Georgia,Times New Roman,serif;" +
                    "color:" +
                      COLORS.text +
                    ";" +
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
                    "color:" +
                      COLORS.secondary +
                    ";" +
                    "margin:38px 0 11px;" +
                  '">' +

                    "06 — A QUESTION" +

                  "</div>" +


                  '<div style="' +
                    "font:25px/1.35 Garamond,Georgia,Times New Roman,serif;" +
                    "color:" +
                      COLORS.text +
                    ";" +
                    "margin:0 0 36px;" +
                  '">' +

                    escapeHtml(
                      question
                    ) +

                  "</div>" +


                  /*
                   * CLOSING LOOP
                   */

                  '<div style="' +
                    "text-align:center;" +
                    "margin:30px 0 18px;" +
                  '">' +

                    '<img src="' +
                      LOOP_ASSET_URL +
                      '" alt="The Loop" width="66" style="' +
                        "display:inline-block;" +
                        "width:66px;" +
                        "height:auto;" +
                        "border:0;" +
                      '">' +

                  "</div>" +


                  '<div style="' +
                    "text-align:center;" +
                    "font:15px/1.4 Garamond,Georgia,Times New Roman,serif;" +
                    "font-style:italic;" +
                    "color:" +
                      COLORS.secondary +
                    ";" +
                    "margin:0 0 10px;" +
                  '">' +

                    "Pleasure is the desire to return." +

                  "</div>" +


                "</td>" +

              "</tr>" +


              /*
               * FOOTER
               */

              "<tr>" +

                '<td style="' +
                  "padding:18px 42px 34px;" +
                  "border-top:1px solid " +
                    COLORS.rule +
                  ";" +
                '">' +

                  '<div style="' +
                    "font:10px Arial,Helvetica,sans-serif;" +
                    "letter-spacing:1.1px;" +
                    "color:" +
                      COLORS.secondary +
                    ";" +
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
   BUILD SUBJECT
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
    typeof item.body.setAsync !==
      "function"
  ) {

    setStatus(
      "Build failed — Outlook body access is unavailable."
    );

    return;

  }


  /*
   * STEP 1
   */

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
      "Build failed — " +
      error.message
    );


    return;

  }


  /*
   * STEP 2
   */

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


  if (
    totalImages > 0
  ) {

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

  }


  /*
   * SUBJECT
   */

  const subject =
    buildSubject();


  /*
   * STEP 3
   */

  updateBuildProgress(
    3,
    "Writing newsletter to Outlook…"
  );


  let completed =
    false;


  const timeout =
    setTimeout(
      function () {

        if (
          completed
        ) {

          return;

        }


        completed =
          true;


        setStatus(
          "Build timed out — Outlook did not finish writing the newsletter."
        );


      },
      BUILD_TIMEOUT_MS
    );


  /*
   * SUBJECT
   */

  item.subject.setAsync(
    subject,
    function (
      subjectResult
    ) {

      if (
        completed
      ) {

        return;

      }


      if (
        subjectResult &&
        subjectResult.status !==
          Office.AsyncResultStatus.Succeeded
      ) {

        completed =
          true;


        clearTimeout(
          timeout
        );


        setStatus(
          "Build failed while writing subject: " +
          getAsyncError(
            subjectResult
          )
        );


        return;

      }


      /*
       * BODY
       */

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
            completed
          ) {

            return;

          }


          completed =
            true;


          clearTimeout(
            timeout
          );


          if (
            bodyResult &&
            bodyResult.status ===
              Office.AsyncResultStatus.Succeeded
          ) {

            updateBuildProgress(
              4,
              "✓ Complete — Dispatch built in Outlook."
            );


            return;

          }


          setStatus(
            "Build failed while writing newsletter: " +
            getAsyncError(
              bodyResult
            )
          );

        }
      );

    }
  );

}


/* =========================================================
   ERROR
========================================================= */

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
