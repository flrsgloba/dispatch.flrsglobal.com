```javascript
/* =========================================================
   THE PLEASURE DISPATCH
   publish.js

   OUTLOOK PUBLISHING BRIDGE

   Outlook blocks direct fetch() calls to Google Apps Script
   because of its Content Security Policy.

   Therefore:

       Outlook Task Pane
            ↓
       Office Dialog
            ↓
       Google Apps Script Publisher
            ↓
       GitHub + Google Drive
========================================================= */

(function () {

  "use strict";


  /* =======================================================
     CONFIGURATION
  ======================================================= */

  const PUBLISHER_URL =
    "https://script.google.com/macros/s/AKfycbzavxknADmXnvAhRqcf9areGCRpfAJIZ62v84kqb_hpfgfAWIUbngcCH4B8M9TpkuA-uw/exec";


  const SECRET_STORAGE_KEY =
    "pd_publish_secret";


  let publishDialog =
    null;


  let publishPayload =
    null;


  let publishResolve =
    null;


  let publishReject =
    null;


  /* =======================================================
     STATUS
  ======================================================= */

  function status(message) {

    if (
      typeof window.setStatus ===
      "function"
    ) {

      window.setStatus(
        message
      );

    }


    console.log(
      "[Pleasure Dispatch]",
      message
    );

  }


  /* =======================================================
     INPUT
  ======================================================= */

  function inputValue(id) {

    const element =
      document.getElementById(
        id
      );


    if (!element) {
      return "";
    }


    return String(
      element.value || ""
    ).trim();

  }


  /* =======================================================
     PUBLISH SECRET
  ======================================================= */

  function getPublishSecret() {

    const stored =
      window.localStorage.getItem(
        SECRET_STORAGE_KEY
      );


    if (stored) {

      return Promise.resolve(
        stored.trim()
      );

    }


    return new Promise(
      function (resolve) {

        const existing =
          document.getElementById(
            "pdPublishKeyDialog"
          );


        if (existing) {
          existing.remove();
        }


        const overlay =
          document.createElement(
            "div"
          );


        overlay.id =
          "pdPublishKeyDialog";


        overlay.style.cssText =
          "position:fixed;" +
          "inset:0;" +
          "z-index:999999;" +
          "display:flex;" +
          "align-items:center;" +
          "justify-content:center;" +
          "padding:20px;" +
          "background:rgba(0,0,0,.55);" +
          "box-sizing:border-box;";


        const panel =
          document.createElement(
            "div"
          );


        panel.style.cssText =
          "width:100%;" +
          "max-width:360px;" +
          "box-sizing:border-box;" +
          "padding:24px;" +
          "background:#303030;" +
          "color:#F2EEE5;" +
          "border:1px solid #777;" +
          "box-shadow:0 12px 40px rgba(0,0,0,.35);" +
          "font-family:Arial,sans-serif;";


        const title =
          document.createElement(
            "div"
          );


        title.textContent =
          "Publish Dispatch";


        title.style.cssText =
          "font-size:16px;" +
          "font-weight:600;" +
          "margin-bottom:8px;";


        const description =
          document.createElement(
            "div"
          );


        description.textContent =
          "Enter your publishing key to send this Dispatch to the public archive.";


        description.style.cssText =
          "font-size:13px;" +
          "line-height:1.5;" +
          "color:#C9C3B8;" +
          "margin-bottom:16px;";


        const input =
          document.createElement(
            "input"
          );


        input.type =
          "password";


        input.autocomplete =
          "off";


        input.spellcheck =
          false;


        input.style.cssText =
          "display:block;" +
          "width:100%;" +
          "box-sizing:border-box;" +
          "height:40px;" +
          "padding:8px 10px;" +
          "background:#595959;" +
          "color:#F2EEE5;" +
          "border:1px solid #777;" +
          "outline:none;" +
          "font:14px Arial,sans-serif;";


        const error =
          document.createElement(
            "div"
          );


        error.style.cssText =
          "display:none;" +
          "color:#F2EEE5;" +
          "font-size:12px;" +
          "margin-top:8px;";


        const actions =
          document.createElement(
            "div"
          );


        actions.style.cssText =
          "display:flex;" +
          "justify-content:flex-end;" +
          "gap:8px;" +
          "margin-top:18px;";


        const cancel =
          document.createElement(
            "button"
          );


        cancel.type =
          "button";


        cancel.textContent =
          "Cancel";


        cancel.style.cssText =
          "height:36px;" +
          "padding:0 14px;" +
          "background:transparent;" +
          "color:#C9C3B8;" +
          "border:1px solid #777;" +
          "cursor:pointer;";


        const submit =
          document.createElement(
            "button"
          );


        submit.type =
          "button";


        submit.textContent =
          "Publish";


        submit.style.cssText =
          "height:36px;" +
          "padding:0 14px;" +
          "background:#D8D0C3;" +
          "color:#303030;" +
          "border:1px solid #D8D0C3;" +
          "cursor:pointer;" +
          "font-weight:600;";


        function close(value) {

          document.removeEventListener(
            "keydown",
            onKeyDown,
            true
          );


          overlay.remove();


          resolve(
            value || ""
          );

        }


        function submitKey() {

          const secret =
            String(
              input.value || ""
            ).trim();


          if (!secret) {

            error.textContent =
              "Please enter your publishing key.";


            error.style.display =
              "block";


            input.focus();


            return;

          }


          window.localStorage.setItem(
            SECRET_STORAGE_KEY,
            secret
          );


          close(
            secret
          );

        }


        function onKeyDown(event) {

          if (
            event.key ===
            "Escape"
          ) {

            event.preventDefault();


            close(
              ""
            );


          } else if (
            event.key ===
            "Enter"
          ) {

            event.preventDefault();


            submitKey();

          }

        }


        cancel.addEventListener(
          "click",
          function () {

            close(
              ""
            );

          }
        );


        submit.addEventListener(
          "click",
          submitKey
        );


        document.addEventListener(
          "keydown",
          onKeyDown,
          true
        );


        actions.appendChild(
          cancel
        );


        actions.appendChild(
          submit
        );


        panel.appendChild(
          title
        );


        panel.appendChild(
          description
        );


        panel.appendChild(
          input
        );


        panel.appendChild(
          error
        );


        panel.appendChild(
          actions
        );


        overlay.appendChild(
          panel
        );


        document.body.appendChild(
          overlay
        );


        input.focus();

      }
    );

  }


  /* =======================================================
     EDITION
  ======================================================= */

  function normalizeEdition() {

    const raw =
      inputValue(
        "edition"
      );


    if (!raw) {

      throw new Error(
        "Publish stopped — please enter an edition number."
      );

    }


    const match =
      raw.match(
        /(\d+(?:\.\d+)?)/
      );


    if (!match) {

      throw new Error(
        "Publish stopped — edition must contain a number, such as No. 001."
      );

    }


    const number =
      match[1];


    const numeric =
      Number(
        number
      );


    if (
      !isFinite(
        numeric
      ) ||
      numeric < 0
    ) {

      throw new Error(
        "Publish stopped — invalid edition number."
      );

    }


    return {

      number:
        number,

      label:
        "No. " +
        number

    };

  }


  /* =======================================================
     IMAGE TEST
  ======================================================= */

  function hasMeaningfulImage(element) {

    const images =
      element.querySelectorAll
        ? element.querySelectorAll(
            "img"
          )
        : [];


    for (
      let i = 0;
      i < images.length;
      i++
    ) {

      const src =
        (
          images[i]
            .getAttribute(
              "src"
            ) ||
          ""
        ).trim();


      if (
        src &&
        src !== "#"
      ) {

        return true;

      }

    }


    return false;

  }


  /* =======================================================
     TEXT TEST
  ======================================================= */

  function hasMeaningfulText(element) {

    const clone =
      element.cloneNode(
        true
      );


    clone
      .querySelectorAll(
        "img,svg,style,script,noscript"
      )
      .forEach(
        function (node) {

          node.remove();

        }
      );


    return (
      clone.textContent ||
      ""
    )
      .replace(
        /\u00a0/g,
        " "
      )
      .replace(
        /\s+/g,
        " "
      )
      .trim()
      .length > 0;

  }


  /* =======================================================
     REMOVE SECTION
  ======================================================= */

  function removeSectionByHeading(
    doc,
    headingText
  ) {

    const elements =
      Array.from(
        doc.body.querySelectorAll(
          "div,td"
        )
      );


    for (
      let i = 0;
      i < elements.length;
      i++
    ) {

      const element =
        elements[i];


      if (
        (
          element.textContent ||
          ""
        ).trim() !==
        headingText
      ) {

        continue;

      }


      if (
        !element.parentNode
      ) {

        continue;

      }


      const parent =
        element.parentNode;


      const children =
        Array.from(
          parent.children
        );


      const index =
        children.indexOf(
          element
        );


      if (
        index < 0
      ) {

        continue;

      }


      element.remove();


      if (
        children[index + 1] &&
        !hasMeaningfulText(
          children[index + 1]
        ) &&
        !hasMeaningfulImage(
          children[index + 1]
        )
      ) {

        children[index + 1].remove();

      }


      return;

    }

  }


  /* =======================================================
     CLEAN EDITORIAL SECTIONS
  ======================================================= */

  function removeEmptyEditorialSections(html) {

    if (
      !html ||
      typeof DOMParser ===
        "undefined"
    ) {

      return html;

    }


    const parser =
      new DOMParser();


    const doc =
      parser.parseFromString(
        String(html),
        "text/html"
      );


    if (
      !inputValue(
        "reflection"
      )
    ) {

      removeSectionByHeading(
        doc,
        "01 — A REFLECTION"
      );

    }


    const imageBlocks =
      typeof window.collectImageBlocks ===
        "function"
        ? window.collectImageBlocks()
        : [];


    if (
      !inputValue(
        "workText"
      ) &&
      !imageBlocks.length
    ) {

      removeSectionByHeading(
        doc,
        "02 — THE WORK"
      );

    }


    if (
      !inputValue(
        "studioText"
      )
    ) {

      removeSectionByHeading(
        doc,
        "03 — STUDIO NOTES"
      );

    }


    const notes =
      typeof window.collectPleasureNotes ===
        "function"
        ? window.collectPleasureNotes()
        : [];


    if (
      !notes.length
    ) {

      removeSectionByHeading(
        doc,
        "04 — PLEASURE NOTES"
      );


      const intro =
        Array.from(
          doc.body.querySelectorAll(
            "div"
          )
        ).find(
          function (element) {

            return (
              (
                element.textContent ||
                ""
              ).trim() ===
              "An offering of what has held my attention."
            );

          }
        );


      if (intro) {
        intro.remove();
      }

    }


    if (
      !inputValue(
        "inviteTitle"
      ) &&
      !inputValue(
        "inviteText"
      ) &&
      !inputValue(
        "ctaUrl"
      )
    ) {

      removeSectionByHeading(
        doc,
        "05 — AN INVITATION"
      );

    }


    return doc.body.innerHTML.trim();

  }


  /* =======================================================
     CLEAN PUBLISH HTML
  ======================================================= */

  function cleanPublishHtml(html) {

    if (
      !html ||
      typeof DOMParser ===
        "undefined"
    ) {

      return html;

    }


    const parser =
      new DOMParser();


    const doc =
      parser.parseFromString(
        String(html),
        "text/html"
      );


    doc
      .querySelectorAll(
        "img"
      )
      .forEach(
        function (image) {

          const src =
            (
              image.getAttribute(
                "src"
              ) ||
              ""
            ).trim();


          if (
            !src ||
            src === "#"
          ) {

            image.remove();

          }

        }
      );


    doc
      .querySelectorAll(
        "p"
      )
      .forEach(
        function (paragraph) {

          if (
            !hasMeaningfulText(
              paragraph
            ) &&
            !hasMeaningfulImage(
              paragraph
            )
          ) {

            paragraph.remove();

          }

        }
      );


    return removeEmptyEditorialSections(
      doc.body.innerHTML.trim()
    );

  }


  /* =======================================================
     BUILD PAYLOAD
  ======================================================= */

  function collectPayload(secret) {

    if (
      typeof window.buildNewsletterHtml !==
      "function"
    ) {

      throw new Error(
        "The newsletter builder is not available. Reload the Dispatch add-in and try again."
      );

    }


    const edition =
      normalizeEdition();


    const html =
      cleanPublishHtml(
        window.buildNewsletterHtml()
      );


    if (
      typeof window.validateNewsletterHtml ===
      "function"
    ) {

      const problem =
        window.validateNewsletterHtml(
          html
        );


      if (problem) {

        throw new Error(
          problem
        );

      }

    }


    const title =
      inputValue(
        "title"
      );


    const date =
      inputValue(
        "date"
      ) ||
      new Date()
        .toISOString()
        .slice(
          0,
          10
        );


    let pleasureText =
      "";


    if (
      typeof window.collectPleasureNotes ===
      "function"
    ) {

      pleasureText =
        window
          .collectPleasureNotes()
          .map(
            function (note) {

              return (
                (note.label || "") +
                " " +
                (note.value || "")
              );

            }
          )
          .join(
            " "
          );

    }


    const subject =
      typeof window.buildSubject ===
        "function"

        ? window.buildSubject()

        : "The Pleasure Dispatch — " +
          edition.label +
          ": " +
          (
            title ||
            "A Note on Pleasure"
          );


    return {

      action:
        "publish",

      secret:
        secret,

      edition:
        edition.number,

      editionLabel:
        edition.label,

      date:
        date,

      title:
        title ||
        "The Pleasure Dispatch",

      subtitle:
        inputValue(
          "subtitle"
        ),

      subject:
        subject,

      reflection:
        inputValue(
          "reflection"
        ),

      html:
        html,

      searchText:
        [
          edition.number,
          edition.label,
          date,
          title,
          inputValue("subtitle"),
          inputValue("reflection"),
          inputValue("workText"),
          inputValue("studioText"),
          inputValue("inviteTitle"),
          inputValue("inviteText"),
          inputValue("question"),
          pleasureText
        ]
          .filter(
            Boolean
          )
          .join(
            " "
          )

    };

  }


  /* =======================================================
     OFFICE DIALOG
  ======================================================= */

  function openPublisherDialog(payload) {

    return new Promise(
      function (resolve, reject) {

        publishPayload =
          payload;


        publishResolve =
          resolve;


        publishReject =
          reject;


        if (
          !window.Office ||
          !Office.context ||
          !Office.context.ui ||
          typeof Office.context.ui.displayDialogAsync !==
            "function"
        ) {

          reject(
            new Error(
              "Office Dialog API is unavailable in this Outlook context."
            )
          );


          return;

        }


        status(
          "Opening secure publishing window…"
        );


        const dialogUrl =
          PUBLISHER_URL +
          "?action=publisherDialog";


        console.log(
          "[Pleasure Dispatch] Opening publisher:",
          dialogUrl
        );


        Office.context.ui.displayDialogAsync(
          dialogUrl,

          {
            height:
              55,

            width:
              40,

            displayInIframe:
              false

          },

          function (result) {

            if (
              result.status !==
              Office.AsyncResultStatus.Succeeded
            ) {

              const message =
                result.error &&
                result.error.message
                  ? result.error.message
                  : "Could not open the publishing window.";


              publishDialog =
                null;


              publishPayload =
                null;


              publishResolve =
                null;


              publishReject =
                null;


              reject(
                new Error(
                  message
                )
              );


              return;

            }


            publishDialog =
              result.value;


            console.log(
              "[Pleasure Dispatch] Publisher dialog opened."
            );


            publishDialog.addEventHandler(
              Office.EventType.DialogMessageReceived,

              function (arg) {

                let message =
                  arg &&
                  arg.message;


                console.log(
                  "[Pleasure Dispatch] Dialog message:",
                  message
                );


                try {

                  message =
                    JSON.parse(
                      message
                    );

                } catch (
                  error
                ) {

                  return;

                }


                if (
                  message.type ===
                  "publisherReady"
                ) {

                  status(
                    "Sending Dispatch to publisher…"
                  );


                  publishDialog.messageChild(
                    JSON.stringify({

                      type:
                        "publishPayload",

                      payload:
                        publishPayload

                    })
                  );


                  return;

                }


                if (
                  message.type ===
                  "publishResult"
                ) {

                  const success =
                    !!message.success;


                  const data =
                    message.data ||
                    message;


                  if (
                    publishDialog
                  ) {

                    try {
                      publishDialog.close();
                    } catch (
                      closeError
                    ) {}
                  }


                  publishDialog =
                    null;


                  publishPayload =
                    null;


                  if (
                    success
                  ) {

                    if (
                      publishResolve
                    ) {

                      publishResolve(
                        data
                      );

                    }

                  } else {

                    if (
                      publishReject
                    ) {

                      publishReject(
                        new Error(
                          message.error ||
                          "Publishing failed."
                        )
                      );

                    }

                  }


                  publishResolve =
                    null;


                  publishReject =
                    null;

                }

              }
            );


            publishDialog.addEventHandler(
              Office.EventType.DialogEventReceived,

              function (event) {

                console.log(
                  "[Pleasure Dispatch] Dialog event:",
                  event
                );


                if (
                  event &&
                  event.error ===
                  12006
                ) {

                  publishDialog =
                    null;


                  publishPayload =
                    null;


                  if (
                    publishReject
                  ) {

                    publishReject(
                      new Error(
                        "Publishing window was closed."
                      )
                    );

                  }


                  publishResolve =
                    null;


                  publishReject =
                    null;

                }

              }
            );

          }
        );

      }
    );

  }


  /* =======================================================
     PUBLISH
  ======================================================= */

  async function publishDispatch() {

    console.log(
      "[Pleasure Dispatch] Publish button clicked."
    );


    status(
      "Publish button clicked."
    );


    const button =
      document.getElementById(
        "publishBtn"
      );


    if (button) {

      button.disabled =
        true;

    }


    try {

      status(
        "Preparing Dispatch…"
      );


      const secret =
        await getPublishSecret();


      if (!secret) {

        status(
          "Publish cancelled."
        );


        return;

      }


      status(
        "Building Dispatch payload…"
      );


      const payload =
        collectPayload(
          secret
        );


      console.log(
        "[Pleasure Dispatch] Payload prepared:",
        {

          edition:
            payload.edition,

          editionLabel:
            payload.editionLabel,

          title:
            payload.title,

          htmlLength:
            payload.html.length

        }
      );


      status(
        "Connecting to publisher…"
      );


      const result =
        await openPublisherDialog(
          payload
        );


      console.log(
        "[Pleasure Dispatch] Publisher result:",
        result
      );


      status(
        "✓ Dispatch published successfully."
      );


    } catch (
      error
    ) {

      console.error(
        "[Pleasure Dispatch] Publish failed:",
        error
      );


      status(
        "Publish failed: " +
        (
          error &&
          error.message
            ? error.message
            : String(error)
        )
      );


    } finally {

      if (button) {

        button.disabled =
          false;

      }

    }

  }


  /* =======================================================
     BUTTON BINDING
  ======================================================= */

  function bindPublishButton() {

    const button =
      document.getElementById(
        "publishBtn"
      );


    if (!button) {

      console.log(
        "[Pleasure Dispatch] publishBtn not found yet."
      );


      return false;

    }


    if (
      button.dataset.publishBound ===
      "true"
    ) {

      return true;

    }


    button.dataset.publishBound =
      "true";


    button.type =
      "button";


    button.addEventListener(
      "click",

      function (event) {

        event.preventDefault();

        event.stopPropagation();

        publishDispatch();

      },

      false
    );


    console.log(
      "[Pleasure Dispatch] Publish button bound."
    );


    return true;

  }


  /* =======================================================
     WAIT FOR BUTTON
  ======================================================= */

  function waitForPublishButton() {

    if (
      bindPublishButton()
    ) {

      return;

    }


    setTimeout(
      waitForPublishButton,
      250
    );

  }


  /* =======================================================
     OFFICE READY
  ======================================================= */

  function initializePublishBridge() {

    console.log(
      "[Pleasure Dispatch] Publish bridge initializing."
    );


    bindPublishButton();


    waitForPublishButton();

  }


  if (
    window.Office &&
    typeof Office.onReady ===
      "function"
  ) {

    Office.onReady(
      initializePublishBridge
    );

  } else {

    window.addEventListener(
      "load",
      initializePublishBridge
    );

  }


  /* =======================================================
     PUBLIC API
  ======================================================= */

  window.publishDispatch =
    publishDispatch;


  /* =======================================================
     NORMALIZE LOOP IMAGE
  ======================================================= */

  (function wrapNewsletterBuilder() {

    function install() {

      if (
        typeof window.buildNewsletterHtml !==
        "function"
      ) {

        return false;

      }


      if (
        window.__pdPublishBuilderWrapped
      ) {

        return true;

      }


      const original =
        window.buildNewsletterHtml;


      window.buildNewsletterHtml =
        function () {

          let html =
            original.apply(
              this,
              arguments
            );


          if (
            !html ||
            typeof html !==
            "string"
          ) {

            return html;

          }


          return html.replace(

            /https:\/\/flrsgloba\.github\.io\/dispatch\.flrsglobal\.com\/assets\/pleasure-loop\.(png|svg)/gi,

            "https://dispatch.flrsglobal.com/assets/pleasure-loop.png"

          );

        };


      window.__pdPublishBuilderWrapped =
        true;


      console.log(
        "[Pleasure Dispatch] Newsletter builder wrapped."
      );


      return true;

    }


    if (!install()) {

      setTimeout(
        install,
        250
      );


      setTimeout(
        install,
        1000
      );


      setTimeout(
        install,
        2000
      );

    }

  })();

})();
```
