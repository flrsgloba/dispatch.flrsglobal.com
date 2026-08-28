/* =========================================================
   THE PLEASURE DISPATCH v0.7
   Consolidated controls
========================================================= */

let blockCounter = 0;
let pleasureCounter = 0;


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
   * Starting Pleasure Notes
   */
  addPleasureRow("Coffee", "");
  addPleasureRow("Art", "");
  addPleasureRow("Object", "");

  /*
   * Start with one image module.
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
       * These are already handled
       * by the static control listeners.
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


        if (!block) {

          setStatus(
            "Could not find image module."
          );

          return;

        }


        duplicateBlock(
          block
        );

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
       * REMOVE INDIVIDUAL IMAGE
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
       * HERO IMAGE URL BUTTONS
       */

      if (
        button.dataset.url
      ) {

        const key =
          button.dataset.url;


        const input =
          document.getElementById(
            key + "Url"
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
   * SELECT / FILE INPUTS
   */

  document.addEventListener(
    "change",
    function (event) {


      /*
       * IMAGE LAYOUT
       */

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
       * FILE UPLOAD
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
   * TEXT / URL INPUTS
   */

  document.addEventListener(
    "input",
    function (event) {


      /*
       * HERO 01
       */

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


      /*
       * HERO 02
       */

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


      /*
       * MODULE IMAGE
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
   HELPERS
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
   HERO IMAGES
========================================================= */

function setupHero(key) {

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


        handleLocalImage(
          file,
          function (dataUrl) {

            if (urlInput) {

              urlInput.value =
                dataUrl;

            }


            renderHeroPreview(
              key,
              dataUrl
            );


            setStatus(
              key === "hero1"
                ? "Hero Image 01 loaded."
                : "Hero Image 02 loaded."
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


function handleLocalImage(
  file,
  callback
) {

  if (
    !file.type ||
    !file.type.match(
      /^image\/(jpeg|png|webp)$/
    )
  ) {

    setStatus(
      "Please choose a JPG, PNG, or WebP image."
    );

    return;

  }


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
        "The image could not be read."
      );

    };


  reader.readAsDataURL(
    file
  );

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
        "This image could not be loaded. Please use a direct HTTPS image URL.";

    };


  image.src =
    url;

}


/* =========================================================
   FILE INPUTS
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
   * HERO 01
   */

  if (
    input.id ===
    "hero1File"
  ) {

    handleLocalImage(
      file,
      function (dataUrl) {

        const url =
          document.getElementById(
            "hero1Url"
          );


        if (url) {

          url.value =
            dataUrl;

        }


        renderHeroPreview(
          "hero1",
          dataUrl
        );


        setStatus(
          "Hero Image 01 loaded."
        );

      }
    );

    return;

  }


  /*
   * HERO 02
   */

  if (
    input.id ===
    "hero2File"
  ) {

    handleLocalImage(
      file,
      function (dataUrl) {

        const url =
          document.getElementById(
            "hero2Url"
          );


        if (url) {

          url.value =
            dataUrl;

        }


        renderHeroPreview(
          "hero2",
          dataUrl
        );


        setStatus(
          "Hero Image 02 loaded."
        );

      }
    );

    return;

  }


  /*
   * MODULAR IMAGE
   */

  const item =
    input.closest(
      ".image-item"
    );


  if (!item) {
    return;
  }


  handleLocalImage(
    file,
    function (dataUrl) {

      const url =
        item.querySelector(
          ".module-url"
        );


      if (url) {

        url.value =
          dataUrl;

      }


      renderModulePreview(
        item,
        dataUrl
      );


      setStatus(
        "Image loaded into module."
      );

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

  const blocks =
    getImageBlocks();


  blocks.forEach(
    function (block, index) {

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


/* =========================================================
   MOVE
========================================================= */

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
    index + direction;


  /*
   * Already at edge.
   */

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


/* =========================================================
   LAYOUT
========================================================= */

function syncImageInputs(
  block
) {

  const layout =
    block.dataset.layout ||
    "full";


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


/* =========================================================
   IMAGE ITEMS
========================================================= */

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
        preset.url || ""
      ) +
      '" placeholder="Paste direct image URL">' +

    '<div class="preview module-preview empty">' +
      "No image selected" +
    "</div>" +

    '<input class="module-caption" value="' +
      escapeAttribute(
        preset.caption || ""
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
        "Image could not be loaded. Please use a direct HTTPS image URL.";

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
    function (item, index) {

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
   * Don't allow an image module
   * to become completely empty.
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
    ).map(
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


  const count =
    layout === "full"
      ? 1
      : layout === "two"
        ? 2
        : layout === "three"
          ? 3
          : 4;


  wrap.className =
    "image-items " +
    (
      layout === "full"
        ? "one"
        : layout
    );


  for (
    let i = 0;
    i < count;
    i++
  ) {

    addImageItem(
      wrap,
      i + 1,
      sourceItems[i] || {}
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
            ).map(
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
            ).filter(
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
   EMAIL IMAGE BLOCKS
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
      100 / columns
    );


  const cells =
    usable.map(
      function (item) {

        return (

          '<td width="' +
          width +
          '%" valign="top" style="padding:3px;vertical-align:top;">' +

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

    'style="border-collapse:collapse;margin:28px 0;table-layout:fixed;">' +

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


  const hero1 =
    value("hero1Url");

  const hero2 =
    value("hero2Url");


  const hero1Caption =
    value("hero1Caption");

  const hero2Caption =
    value("hero2Caption");


  const notes =
    collectPleasureNotes();


  const blocks =
    collectImageBlocks();


  const hero1Html =
    hero1
      ? '<img src="' +
        escapeAttribute(
          hero1
        ) +
        '" alt="" style="display:block;width:100%;height:auto;border:0;margin:0 0 10px;">' +

        (
          hero1Caption
            ? '<div style="font:12px/1.45 Arial,Helvetica,sans-serif;color:#777;margin:0 0 32px;">' +
              escapeHtml(
                hero1Caption
              ) +
              "</div>"
            : ""
        )
      : "";


  const hero2Html =
    hero2
      ? '<img src="' +
        escapeAttribute(
          hero2
        ) +
        '" alt="" style="display:block;width:100%;height:auto;border:0;margin:0 0 10px;">' +

        (
          hero2Caption
            ? '<div style="font:12px/1.45 Arial,Helvetica,sans-serif;color:#777;margin:0 0 32px;">' +
              escapeHtml(
                hero2Caption
              ) +
              "</div>"
            : ""
        )
      : "";


  let invitationHtml =
    "";


  if (
    value("inviteTitle") ||
    value("inviteText") ||
    value("ctaUrl")
  ) {

    invitationHtml =

      '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:40px 0 11px;">' +

        "05 — AN INVITATION" +

      "</div>" +

      (
        value("inviteTitle")
          ? '<div style="font:27px/1.15 Garamond,Georgia,Times New Roman,serif;margin:0 0 10px;">' +
            escapeHtml(
              value("inviteTitle")
            ) +
            "</div>"
          : ""
      ) +

      paragraph(
        value("inviteText")
      ) +

      (
        value("ctaUrl")
          ? '<div style="padding:0 0 25px;">' +

            '<a href="' +
            escapeAttribute(
              value("ctaUrl")
            ) +
            '" style="display:inline-block;background:#151515;color:#fff;text-decoration:none;padding:12px 18px;font:10px Arial,Helvetica,sans-serif;letter-spacing:1.2px;">' +

            escapeHtml(
              value("ctaLabel") ||
              "INQUIRE"
            ) +

            "</a>" +

          "</div>"
          : ""
      );

  }


  const dateLine =
    [edition, date, title]
      .filter(Boolean)
      .map(escapeHtml)
      .join(" · ");


  return (

    '<div style="margin:0;padding:0;background:#f4f0e8;color:#151515;">' +

      '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;background:#f4f0e8;">' +

        "<tr>" +

          '<td align="center" style="padding:28px 12px;">' +

            '<table role="presentation" width="680" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse;width:100%;max-width:680px;background:#fffdf8;">' +

              "<tr>" +

                '<td style="padding:42px 42px 20px;">' +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:2px;">FLRS GLOBAL</div>' +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.4px;color:#777;margin-top:8px;">FROM THE STUDIO OF FREDDIE L. RANKIN II</div>' +

                  '<h1 style="font:400 50px/0.96 Garamond,Georgia,Times New Roman,serif;margin:20px 0 10px;">The Pleasure Dispatch</h1>' +

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

              "<tr>" +

                '<td style="padding:0 42px;">' +

                  '<div style="text-align:center;font:27px Garamond,Georgia,serif;margin:0 0 20px;">◒</div>' +

                  hero1Html +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:30px 0 11px;">01 — A REFLECTION</div>' +

                  paragraph(
                    value("reflection")
                  ) +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">02 — THE WORK</div>' +

                  paragraph(
                    value("workText")
                  ) +

                  blocks
                    .map(function (block) {
                      return buildImageModuleHtml(
                        block
                      );
                    })
                    .join("") +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">03 — STUDIO NOTES</div>' +

                  paragraph(
                    value("studioText")
                  ) +

                  hero2Html +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">04 — PLEASURE NOTES</div>' +

                  '<div style="font:19px/1.4 Garamond,Georgia,Times New Roman,serif;margin:0 0 8px;">An offering of what has held my attention.</div>' +

                  buildPleasureNotesHtml(
                    notes
                  ) +

                  invitationHtml +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.5px;color:#777;margin:38px 0 11px;">06 — A QUESTION</div>' +

                  '<div style="font:25px/1.35 Garamond,Georgia,Times New Roman,serif;margin:0 0 36px;">' +

                    escapeHtml(
                      value("question")
                    ) +

                  "</div>" +

                  '<div style="text-align:center;font:27px Garamond,Georgia,serif;margin:20px 0 32px;">◒</div>' +

                "</td>" +

              "</tr>" +

              '<tr>' +

                '<td style="padding:18px 42px 34px;border-top:1px solid #151515;">' +

                  '<div style="font:10px Arial,Helvetica,sans-serif;letter-spacing:1.1px;color:#777;">THE PLEASURE DISPATCH · BY FLRS GLOBAL</div>' +

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
   OUTLOOK
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
    "Building The Pleasure Dispatch in Outlook…"
  );


  /*
   * Make sure Office.js is actually
   * available.
   */

  if (
    typeof Office ===
    "undefined"
  ) {

    setStatus(
      "Office.js is not available. Reload the Outlook add-in."
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


  const subject =
    buildSubject();


  const html =
    buildNewsletterHtml();


  /*
   * Set subject.
   */

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


      /*
       * Set HTML body.
       */

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
