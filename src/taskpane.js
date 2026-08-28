/* =========================================================
   THE PLEASURE DISPATCH
   v0.8 — Image Compression + Reliable Outlook Build
========================================================= */

let blockCounter = 0;
let pleasureCounter = 0;

const IMAGE_MAX_WIDTH = 1800;
const IMAGE_MAX_HEIGHT = 1800;
const IMAGE_QUALITY = 0.82;
const MAX_INLINE_IMAGE_BYTES = 8 * 1024 * 1024;


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
    document.getElementById("previewBtn");

  const buildButton =
    document.getElementById("insertBtn");

  const addImageButton =
    document.getElementById("addImageBlock");

  const addPleasureButton =
    document.getElementById("addPleasure");


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

        addPleasureRow("", "");

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

function setupDelegatedControls() {

  document.addEventListener(
    "click",
    function (event) {

      const button =
        event.target.closest("button");

      if (!button) {
        return;
      }


      /*
       * Static controls already have handlers.
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
        button.classList.contains("move-up")
      ) {

        const block =
          button.closest(".image-block");

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
        button.classList.contains("move-down")
      ) {

        const block =
          button.closest(".image-block");

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
        button.classList.contains("duplicate")
      ) {

        const block =
          button.closest(".image-block");

        if (!block) {
          setStatus(
            "Could not find image module."
          );
          return;
        }

        duplicateBlock(block);

        return;
      }


      /*
       * REMOVE MODULE
       */
      if (
        button.classList.contains("remove")
      ) {

        const block =
          button.closest(".image-block");

        if (!block) {
          return;
        }

        block.remove();

        renumberImageBlocks();

        setStatus(
          "Image Module removed."
        );

        return;
      }


      /*
       * REMOVE IMAGE
       */
      if (
        button.classList.contains("remove-image")
      ) {

        removeImageItem(
          button.closest(".image-item")
        );

        return;
      }


      /*
       * IMAGE URL BUTTON
       */
      if (
        button.classList.contains("url-item")
      ) {

        const item =
          button.closest(".image-item");

        if (item) {

          const input =
            item.querySelector(".module-url");

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
            button.dataset.url + "Url"
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
   * SELECT + FILE INPUTS
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
    document.getElementById(id);

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
  return escapeHtml(text);
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
    document.getElementById("status");

  if (status) {
    status.textContent =
      message;
  }

}


/* =========================================================
   DESKTOP IMAGE COMPRESSION
========================================================= */

function compressImage(
  file,
  callback
) {

  if (
    !file ||
    !file.type ||
    !file.type.match(
      /^image\/(jpeg|png|webp)$/
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
    "Preparing image…"
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
           * Scale down large camera images.
           */
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


          context.drawImage(
            image,
            0,
            0,
            width,
            height
          );


          /*
           * JPEG is much more efficient for
           * photographic material.
           */
          const outputType =
            "image/jpeg";


          canvas.toBlob(
            function (blob) {

              if (!blob) {

                callback(
                  null,
                  new Error(
                    "The image could not be compressed."
                  )
                );

                return;
              }


              /*
               * Keep quality conservative.
               * If still oversized, progressively
               * reduce quality.
               */
              compressBlobToLimit(
                blob,
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
            outputType,
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


function compressBlobToLimit(
  blob,
  callback
) {

  /*
   * Most normal photographs should
   * already be below this threshold.
   */
  if (
    blob.size <=
    MAX_INLINE_IMAGE_BYTES
  ) {

    callback(
      blob
    );

    return;
  }


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


      const canvas =
        document.createElement(
          "canvas"
        );


      canvas.width =
        Math.round(
          image.naturalWidth *
          0.75
        );

      canvas.height =
        Math.round(
          image.naturalHeight *
          0.75
        );


      const context =
        canvas.getContext(
          "2d"
        );


      context.drawImage(
        image,
        0,
        0,
        canvas.width,
        canvas.height
      );


      canvas.toBlob(
        function (smallerBlob) {

          if (
            !smallerBlob
          ) {

            callback(
              blob
            );

            return;
          }


          if (
            smallerBlob.size <
            blob.size
          ) {

            compressBlobToLimit(
              smallerBlob,
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


  reader.readAsDataURL(
    blob
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
      key + "File"
    );

  const urlInput =
    document.getElementById(
      key + "Url"
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


        compressImage(
          file,
          function (result, error) {

            if (error) {

              setStatus(
                error.message
              );

              return;
            }


            urlInput.value =
              result.dataUrl;


            renderHeroPreview(
              key,
              result.dataUrl
            );


            setStatus(
              key === "hero1"
                ? "Hero Image 01 optimized and ready."
                : "Hero Image 02 optimized and ready."
            );

          }
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


function renderHeroPreview(
  key,
  url
) {

  const box =
    document.getElementById(
      key + "Preview"
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
   MODULAR IMAGE UPLOAD
========================================================= */

function handleFileInput(
  input
) {

  const file =
    input.files &&
    input.files[0];


  if (!file) {
    return;
  }


  /*
   * Hero 01
   */
  if (
    input.id ===
    "hero1File"
  ) {

    compressImage(
      file,
      function (result, error) {

        if (error) {

          setStatus(
            error.message
          );

          return;
        }


        document.getElementById(
          "hero1Url"
        ).value =
          result.dataUrl;


        renderHeroPreview(
          "hero1",
          result.dataUrl
        );


        setStatus(
          "Hero Image 01 optimized and ready."
        );

      }
    );

    return;
  }


  /*
   * Hero 02
   */
  if (
    input.id ===
    "hero2File"
  ) {

    compressImage(
      file,
      function (result, error) {

        if (error) {

          setStatus(
            error.message
          );

          return;
        }


        document.getElementById(
          "hero2Url"
        ).value =
          result.dataUrl;


        renderHeroPreview(
          "hero2",
          result.dataUrl
        );


        setStatus(
          "Hero Image 02 optimized and ready."
        );

      }
    );

    return;
  }


  /*
   * Modular image
   */
  const item =
    input.closest(
      ".image-item"
    );


  if (!item) {
    return;
  }


  compressImage(
    file,
    function (result, error) {

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


      setStatus(
        "Image optimized and loaded into module."
      );

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
        labelValue || ""
      ) +
      '" placeholder="Category">' +

    '<input class="pleasure-value" value="' +
      escapeAttribute(
        noteValue || ""
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

            '<td style="width:120px;vertical-align:top;padding:5px 18px 5px 0;' +
            'font:10px Arial,Helvetica,sans-serif;letter-spacing:1px;' +
            'text-transform:uppercase;color:#777;">' +

              escapeHtml(
                note.label
              ) +

            "</td>" +

            '<td style="vertical-align:top;padding:5px 0;' +
            'font:16px/1.5 Garamond,Georgia,Times New Roman,serif;color:#151515;">' +

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


  if (
    index === -1
  ) {

    setStatus(
      "Unable to find image module."
    );

    return;

  }


  const targetIndex =
    index + direction;


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


  if (direction < 0) {

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
    preset || {};


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


  if (
    !wrap
  ) {

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
              : ""

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


  wrap.className =
    "image-items " +
    (
      layout ===
      "full"
        ? "one"
        : layout
    );


  const count =
    layout ===
    "full"
      ? 1
      : layout ===
        "two"
          ? 2
          : layout ===
            "three"
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
   EMAIL IMAGE BUILDERS
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

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" ' +
    'style="border-collapse:collapse;margin:30px 0;">' +

      "<tr>" +

        '<td style="padding:0;">' +

          '<img src="' +
          escapeAttribute(
            item.url
          ) +
          '" alt="" ' +
          'style="display:block;width:100%;height:auto;border:0;">' +

          (
            item.caption
              ? '<div style="font:12px/1.45 Arial,Helvetica,sans-serif;color:#777;margin-top:7px;">' +
                escapeHtml(
                  item.caption
                ) +
                "</div>"
              : ""
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

            (
              item.url
                ? '<img src="' +
                  escapeAttribute(
                    item.url
                  ) +
                  '" alt="" style="display:block;width:100%;height:auto;border:0;">'
                : ""
            ) +

            (
              item.caption
                ? '<div style="font:11px/1.4 Arial,Helvetica,sans-serif;color:#777;padding-top:6px;">' +
                  escapeHtml(
                    item.caption
                  ) +
                  "</div>"
                : ""
            ) +

          "</td>"

        );

      }
    ).join("");


  return (

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" ' +
    'style="border-collapse:collapse;table-layout:fixed;margin:28px 0;">' +

      "<tr>" +

        cells +

      "</tr>" +

    "</table>"

  );

}


/*
 * Four Up ALWAYS means:
 *
 * [ IMG 1 ] [ IMG 2 ] [ IMG 3 ] [ IMG 4 ]
 *
 * one single horizontal row.
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

            (
              item.url
                ? '<img src="' +
                  escapeAttribute(
                    item.url
                  ) +
                  '" alt="" style="display:block;width:100%;height:auto;border:0;">'
                : ""
            ) +

            (
              item.caption
                ? '<div style="font:10px/1.4 Arial,Helvetica,sans-serif;color:#777;padding-top:6px;">' +
                  escapeHtml(
                    item.caption
                  ) +
                  "</div>"
                : ""
            ) +

          "</td>"

        );

      }
    ).join("");


  return (

    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" ' +
    'style="border-collapse:collapse;table-layout:fixed;margin:28px 0;">' +

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
   CONVERT DESKTOP DATA URLS TO INLINE CID IMAGES
========================================================= */

function findInlineDataUrls(
  html
) {

  const regex =
    /data:image\/(?:jpeg|jpg|png|webp);base64,[A-Za-z0-9+/=]+/g;


  const matches =
    html.match(
      regex
    ) || [];


  /*
   * Preserve order and remove duplicates.
   */
  return matches.filter(
    function (
      url,
      index,
      array
    ) {

      return (
        array.indexOf(url) ===
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


function getImageExtension(
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
 * Outlook's documented inline attachment
 * pattern uses the filename as the CID:
 *
 * <img src="cid:sample.png">
 */
function addInlineImages(
  item,
  html,
  dataUrls,
  index,
  completed
) {

  if (
    index >=
    dataUrls.length
  ) {

    completed(
      null,
      html
    );

    return;

  }


  const dataUrl =
    dataUrls[index];


  const parsed =
    parseDataUrl(
      dataUrl
    );


  if (!parsed) {

    addInlineImages(
      item,
      html,
      dataUrls,
      index + 1,
      completed
    );

    return;

  }


  const extension =
    getImageExtension(
      parsed.mimeType
    );


  const attachmentName =
    "pleasure-dispatch-image-" +
    (index + 1) +
    "-" +
    Date.now() +
    "." +
    extension;


  setStatus(
    "Embedding image " +
    (index + 1) +
    " of " +
    dataUrls.length +
    " in Outlook…"
  );


  item.addFileAttachmentFromBase64Async(
    parsed.base64,
    attachmentName,
    {
      isInline: true
    },
    function (result) {

      if (
        result.status ===
        Office.AsyncResultStatus.Failed
      ) {

        completed(
          new Error(
            "Outlook could not embed image " +
            (index + 1) +
            ": " +
            (
              result.error
                ? result.error.message
                : "Unknown attachment error."
            )
          ),
          null
        );

        return;

      }


      /*
       * Replace the Base64 data URL with
       * the inline attachment CID.
       */
      html =
        html.split(
          dataUrl
        ).join(
          "cid:" +
          attachmentName
        );


      addInlineImages(
        item,
        html,
        dataUrls,
        index + 1,
        completed
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


  /*
   * Confirm Office.js.
   */
  if (
    typeof Office ===
    "undefined"
  ) {

    setStatus(
      "Office.js is not available. Reload the add-in."
    );

    return;

  }


  /*
   * Confirm compose context.
   */
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


  /*
   * Build the HTML.
   */
  let html;

  try {

    html =
      buildNewsletterHtml();

  } catch (error) {

    setStatus(
      "Newsletter build error: " +
      error.message
    );

    return;

  }


  const subject =
    buildSubject();


  /*
   * IMPORTANT:
   *
   * Microsoft recommends getting the current body
   * before inserting Base64 inline images.
   */
  item.body.getAsync(
    Office.CoercionType.Html,
    function (bodyResult) {

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
       * Find any desktop-uploaded data URLs.
       */
      const dataUrls =
        findInlineDataUrls(
          html
        );


      /*
       * No local images.
       */
      if (
        dataUrls.length ===
        0
      ) {

        writeNewsletter(
          item,
          subject,
          html
        );

        return;

      }


      /*
       * Local images need to be added
       * as inline Outlook attachments.
       */
      setStatus(
        "Preparing " +
        dataUrls.length +
        " uploaded image" +
        (
          dataUrls.length ===
          1
            ? ""
            : "s"
        ) +
        "…"
      );


      addInlineImages(
        item,
        html,
        dataUrls,
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
   WRITE NEWSLETTER
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
    function (subjectResult) {

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
        function (bodyResult) {

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
