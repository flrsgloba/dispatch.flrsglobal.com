/* =========================================================
   THE PLEASURE DISPATCH
   publish.js

   OUTLOOK PUBLISHING BRIDGE

   FLOW:

       Outlook Task Pane
            ↓
       Same-domain publisher.html
            ↓
       Google Apps Script Publisher
            ↓
       GitHub + Google Drive
            ↓
       Result back to Outlook

   IMPORTANT:
   The first Office Dialog URL MUST be on the same
   full domain as the Outlook add-in.
========================================================= */

(function () {

  "use strict";


  /* =======================================================
     CONFIGURATION
  ======================================================= */

  const ADDIN_ORIGIN =
    "https://dispatch.flrsglobal.com";


  const PUBLISHER_URL =
    "https://script.google.com/macros/s/AKfycbzavxknADmXnvAhRqcf9areGCRpfAJIZ62v84kqb_hpfgfAWIUbngcCH4B8M9TpkuA-uw/exec";


  const DIALOG_URL =
    ADDIN_ORIGIN +
    "/src/publisher.html";


  const SECRET_STORAGE_KEY =
    "pd_publish_secret";


  let publishDialog =
    null;


  let publishPayload =
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


    if (!stored) {

      throw new Error(
        "Publishing key is not configured."
      );

    }


    return stored.trim();

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
     REMOVE EMPTY EDITORIAL SECTIONS
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
     CLEAN HTML
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
     OPEN PUBLISHER DIALOG
  ======================================================= */

  function openPublisherDialog(payload) {

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
              "Office Dialog API is unavailable in this Outlook context."
            )
          );

          return;

        }


        if (
          publishDialog
        ) {

          try {

            publishDialog.close();

          } catch (
            error
          ) {}

          publishDialog =
            null;

        }


        publishPayload =
          payload;


        status(
          "Connecting to publisher…"
        );


        console.log(
          "[Pleasure Dispatch] Opening same-domain publisher:",
          DIALOG_URL
        );


        Office.context.ui.displayDialogAsync(

          DIALOG_URL,

          {

            height:
              10,

            width:
              10,

            displayInIframe:
              true

          },

          function (result) {

            if (
              result.status !==
              Office.AsyncResultStatus.Succeeded
            ) {

              const error =
                result.error;


              publishDialog =
                null;


              publishPayload =
                null;


              reject(
                new Error(

                  error &&
                  error.message

                    ? error.message

                    : "Could not open the publisher."

                )
              );


              return;

            }


            publishDialog =
              result.value;


            console.log(
              "[Pleasure Dispatch] Same-domain publisher opened."
            );


            publishDialog.addEventHandler(

              Office.EventType.DialogMessageReceived,

              function (arg) {

                let message =
                  arg &&
                  arg.message;


                console.log(
                  "[Pleasure Dispatch] Publisher message:",
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

                    }),

                    {

                      targetOrigin:
                        ADDIN_ORIGIN

                    }

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
                      error
                    ) {}

                  }


                  publishDialog =
                    null;


                  publishPayload =
                    null;


                  if (
                    success
                  ) {

                    resolve(
                      data
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

              function (event) {

                console.log(
                  "[Pleasure Dispatch] Publisher dialog event:",
                  event
                );


                if (
                  event &&
                  event.error
                ) {

                  const errorCode =
                    event.error;


                  publishDialog =
                    null;


                  publishPayload =
                    null;


                  if (
                    errorCode ===
                    12006
                  ) {

                    reject(
                      new Error(
                        "Publisher dialog was closed."
                      )
                    );

                  } else if (
                    errorCode ===
                    12009
                  ) {

                    reject(
                      new Error(
                        "Outlook blocked the publisher dialog."
                      )
                    );

                  } else if (
                    errorCode ===
                    12011
                  ) {

                    reject(
                      new Error(
                        "Your Outlook/browser is blocking the publisher window."
                      )
                    );

                  } else {

                    reject(
                      new Error(
                        "Publisher dialog error: " +
                        errorCode
                      )
                    );

                  }

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
        getPublishSecret();


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


    if (
      !install()
    ) {

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
