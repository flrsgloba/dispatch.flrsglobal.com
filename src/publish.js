```javascript
/* =========================================================
   THE PLEASURE DISPATCH
   publish.js

   Outlook publishing bridge

   IMPORTANT:
   Outlook's task pane CSP blocks direct fetch() requests
   to Google Apps Script.

   Publishing therefore uses the Office Dialog API:

   Outlook task pane
        ↓
   Office Dialog
        ↓
   Publisher endpoint
        ↓
   GitHub / Drive
========================================================= */

(function () {
  "use strict";

  const PUBLISHER_URL =
    "https://script.google.com/macros/s/AKfycbzavxknADmXnvAhRqcf9areGCRpfAJIZ62v84kqb_hpfgfAWIUbngcCH4B8M9TpkuA-uw/exec";

  const SECRET_STORAGE_KEY =
    "pd_publish_secret";

  let publishDialog = null;
  let publishPayload = null;


  /* =======================================================
     PUBLISH KEY
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

        overlay.setAttribute(
          "role",
          "dialog"
        );

        overlay.setAttribute(
          "aria-modal",
          "true"
        );

        overlay.style.cssText =
          "position:fixed;inset:0;z-index:999999;" +
          "display:flex;align-items:center;" +
          "justify-content:center;padding:20px;" +
          "background:rgba(0,0,0,.55);" +
          "box-sizing:border-box";


        const panel =
          document.createElement(
            "div"
          );

        panel.style.cssText =
          "width:100%;max-width:360px;" +
          "box-sizing:border-box;padding:24px;" +
          "background:#303030;color:#F2EEE5;" +
          "border:1px solid #777;" +
          "box-shadow:0 12px 40px rgba(0,0,0,.35);" +
          "font-family:Arial,sans-serif";


        const title =
          document.createElement(
            "div"
          );

        title.textContent =
          "Publish Dispatch";

        title.style.cssText =
          "font-size:16px;font-weight:600;" +
          "margin-bottom:8px";


        const description =
          document.createElement(
            "div"
          );

        description.textContent =
          "Enter your publishing key to send this Dispatch to the public archive.";

        description.style.cssText =
          "font-size:13px;line-height:1.5;" +
          "color:#C9C3B8;margin-bottom:16px";


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
          "display:block;width:100%;" +
          "box-sizing:border-box;height:40px;" +
          "padding:8px 10px;background:#595959;" +
          "color:#F2EEE5;border:1px solid #777;" +
          "outline:none;font:14px Arial,sans-serif";


        const error =
          document.createElement(
            "div"
          );

        error.style.cssText =
          "display:none;color:#F2EEE5;" +
          "font-size:12px;margin-top:8px";


        const actions =
          document.createElement(
            "div"
          );

        actions.style.cssText =
          "display:flex;justify-content:flex-end;" +
          "gap:8px;margin-top:18px";


        const cancel =
          document.createElement(
            "button"
          );

        cancel.type =
          "button";

        cancel.textContent =
          "Cancel";

        cancel.style.cssText =
          "height:36px;padding:0 14px;" +
          "background:transparent;color:#C9C3B8;" +
          "border:1px solid #777;cursor:pointer";


        const submit =
          document.createElement(
            "button"
          );

        submit.type =
          "button";

        submit.textContent =
          "Publish";

        submit.style.cssText =
          "height:36px;padding:0 14px;" +
          "background:#D8D0C3;color:#303030;" +
          "border:1px solid #D8D0C3;" +
          "cursor:pointer;font-weight:600";


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


        function onKeyDown(
          event
        ) {

          if (
            event.key ===
            "Escape"
          ) {

            event.preventDefault();

            close("");

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
            close("");
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
     STATUS
  ======================================================= */

  function status(
    message
  ) {

    if (
      typeof setStatus ===
      "function"
    ) {

      setStatus(
        message
      );

    } else {

      console.log(
        "[Pleasure Dispatch]",
        message
      );

    }

  }


  /* =======================================================
     INPUT
  ======================================================= */

  function inputValue(
    id
  ) {

    const element =
      document.getElementById(
        id
      );


    return element
      ? String(
          element.value || ""
        ).trim()
      : "";

  }


  /* =======================================================
     IMAGE TEST
  ======================================================= */

  function hasMeaningfulImage(
    element
  ) {

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

  function hasMeaningfulText(
    element
  ) {

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
     REMOVE EMPTY EDITORIAL SECTIONS
  ======================================================= */

  function removeEmptyEditorialSections(
    html
  ) {

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


    if (
      !inputValue(
        "workText"
      ) &&
      typeof collectImageBlocks ===
        "function" &&
      !collectImageBlocks().length
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
      typeof collectPleasureNotes ===
        "function"
        ? collectPleasureNotes()
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
     CLEAN HTML
  ======================================================= */

  function cleanPublishHtml(
    html
  ) {

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
     BUILD PAYLOAD
  ======================================================= */

  function collectPayload(
    secret
  ) {

    if (
      typeof buildNewsletterHtml !==
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
        buildNewsletterHtml()
      );


    if (
      typeof validateNewsletterHtml ===
      "function"
    ) {

      const problem =
        validateNewsletterHtml(
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
      typeof collectPleasureNotes ===
      "function"
    ) {

      pleasureText =
        collectPleasureNotes()
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
      typeof buildSubject ===
        "function"

        ? buildSubject()

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

      /*
       * Publisher .gs currently expects
       * "secret" rather than "publishKey".
       */
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
     OPEN PUBLISH DIALOG
  ======================================================= */

  function openPublisherDialog(
    payload
  ) {

    return new Promise(
      function (resolve, reject) {

        if (
          !window.Office ||
          !Office.context ||
          !Office.context.ui ||
          typeof Office.context.ui.displayDialogAsync !==
            "function"
        ) {

          reject(
            new Error(
              "Outlook does not support the publishing dialog API in this context."
            )
          );

          return;

        }


        publishPayload =
          payload;


        status(
          "Opening secure publishing window…"
        );


        Office.context.ui.displayDialogAsync(
          PUBLISHER_URL +
            "?action=publisherDialog",

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

              reject(
                new Error(
                  "Could not open the publishing window."
                )
              );

              return;

            }


            publishDialog =
              result.value;


            publishDialog.addEventHandler(
              Office.EventType.DialogMessageReceived,

              function (
                arg
              ) {

                let message =
                  arg &&
                  arg.message;


                try {

                  message =
                    JSON.parse(
                      message
                    );

                } catch (
                  parseError
                ) {

                  /*
                   * Ignore non-JSON messages.
                   */

                  return;

                }


                if (
                  message.type ===
                  "publisherReady"
                ) {

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

                  if (
                    publishDialog
                  ) {

                    publishDialog.close();

                  }


                  publishDialog =
                    null;

                  publishPayload =
                    null;


                  if (
                    message.success
                  ) {

                    resolve(
                      message.data ||
                      message
                    );

                  } else {

                    reject(
                      new Error(
                        message.error ||
                        "Publishing failed."
                      )
                    );

                  }

                }

              }
            );


            publishDialog.addEventHandler(
              Office.EventType.DialogEventReceived,

              function (
                event
              ) {

                if (
                  event &&
                  event.error ===
                  12006
                ) {

                  publishDialog =
                    null;

                  publishPayload =
                    null;

                  reject(
                    new Error(
                      "Publishing window was closed."
                    )
                  );

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


      const payload =
        collectPayload(
          secret
        );


      console.log(
        "[Pleasure Dispatch] Payload prepared:",
        {
          action:
            payload.action,

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

      status(
        "Publish failed: " +
        (
          error &&
          error.message
            ? error.message
            : String(error)
        )
      );


      console.error(
        "[Pleasure Dispatch] Publish failed:",
        error
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


    if (
      !button ||
      button.dataset.publishBound ===
        "true"
    ) {

      return;

    }


    button.dataset.publishBound =
      "true";


    button.addEventListener(
      "click",
      function (event) {

        event.preventDefault();

        event.stopPropagation();

        publishDispatch();

      }
    );

  }


  /* =======================================================
     LOAD
  ======================================================= */

  window.addEventListener(
    "load",
    bindPublishButton
  );


  setTimeout(
    bindPublishButton,
    250
  );

  setTimeout(
    bindPublishButton,
    1000
  );

  setTimeout(
    bindPublishButton,
    2000
  );


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
