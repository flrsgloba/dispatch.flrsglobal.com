/* =========================================================
   THE PLEASURE DISPATCH
   publish.js

   NET-NEW PUBLISH BRIDGE

   Architecture:

   Outlook taskpane
        ↓
   Office Dialog
        ↓
   Publisher Apps Script deployment
        ↓
   GitHub / Dispatch archive

   IMPORTANT:
   - This file owns publishing only.
   - drive-bridge.js owns the current Drive deployment.
   - taskpane.js owns content construction and Outlook insertion.
   - The publisher deployment URL is intentionally separate from
     the current Drive API URL.
========================================================= */

(function () {
  "use strict";

  const PUBLISHER_URL =
    "https://script.google.com/macros/s/AKfycbzavxknADmXnvAhRqcf9areGCRpfAJIZ62v84kqb_hpfgfAWIUbngcCH4B8M9TpkuA-uw/exec";

  const SECRET_STORAGE_KEY = "pd_publish_secret";
  const DIALOG_TIMEOUT_MS = 30000;
  const PING_INTERVAL_MS = 500;

  let publisherDialog = null;
  let pendingPayload = null;
  let pendingResolve = null;
  let pendingReject = null;
  let timeoutTimer = null;
  let pingTimer = null;
  let resultReceived = false;

  function setPublishStatus(message) {
    if (typeof window.setStatus === "function") {
      window.setStatus(message);
    }
    console.log("[Pleasure Dispatch]", message);
  }

  function inputValue(id) {
    const element = document.getElementById(id);
    return element ? String(element.value || "").trim() : "";
  }

  function getPublishSecret() {
    let secret = "";

    try {
      secret = window.localStorage.getItem(SECRET_STORAGE_KEY) || "";
    } catch (error) {
      console.warn("[Pleasure Dispatch] Local storage is unavailable.");
    }

    secret = secret.trim();

    if (secret) return secret;

    secret = window.prompt(
      "Enter the Pleasure Dispatch publishing key:"
    );

    if (!secret || !secret.trim()) {
      throw new Error("Publishing key is required.");
    }

    secret = secret.trim();

    try {
      window.localStorage.setItem(
        SECRET_STORAGE_KEY,
        secret
      );
    } catch (error) {
      console.warn(
        "[Pleasure Dispatch] Could not save publishing key locally."
      );
    }

    return secret;
  }

  function clearStoredPublishSecret() {
    try {
      window.localStorage.removeItem(SECRET_STORAGE_KEY);
    } catch (error) {
      /* Ignore local-storage failures. */
    }
  }

  function normalizeEdition() {
    const raw = inputValue("edition");
    const match = raw.match(/(\d+(?:\.\d+)?)/);

    if (!match) {
      throw new Error(
        "Publish stopped — enter an edition number, such as No. 001."
      );
    }

    const number = match[1];

    if (!isFinite(Number(number)) || Number(number) < 0) {
      throw new Error("Publish stopped — invalid edition number.");
    }

    return {
      number: number,
      label: "No. " + number
    };
  }

  function hasMeaningfulText(element) {
    if (!element) return false;

    const clone = element.cloneNode(true);

    clone
      .querySelectorAll("img,svg,style,script,noscript")
      .forEach(function (node) {
        node.remove();
      });

    return String(clone.textContent || "")
      .replace(/\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .length > 0;
  }

  function hasMeaningfulImage(element) {
    if (!element || !element.querySelectorAll) return false;

    const images = element.querySelectorAll("img");

    for (let i = 0; i < images.length; i++) {
      const src = String(
        images[i].getAttribute("src") || ""
      ).trim();

      if (src && src !== "#") return true;
    }

    return false;
  }

  function removeSectionByHeading(doc, headingText) {
    const elements = Array.from(
      doc.body.querySelectorAll("div,td")
    );

    for (let i = 0; i < elements.length; i++) {
      const element = elements[i];

      if (
        String(element.textContent || "").trim() !==
        headingText
      ) {
        continue;
      }

      const parent = element.parentNode;
      if (!parent) return;

      const children = Array.from(parent.children);
      const index = children.indexOf(element);

      element.remove();

      const next = children[index + 1];

      if (
        next &&
        !hasMeaningfulText(next) &&
        !hasMeaningfulImage(next)
      ) {
        next.remove();
      }

      return;
    }
  }

  function removeEmptyEditorialSections(html) {
    if (!html || typeof DOMParser === "undefined") {
      return html;
    }

    const doc = new DOMParser().parseFromString(
      String(html),
      "text/html"
    );

    if (!inputValue("reflection")) {
      removeSectionByHeading(doc, "01 — A REFLECTION");
    }

    const imageBlocks =
      typeof window.collectImageBlocks === "function"
        ? window.collectImageBlocks()
        : [];

    if (!inputValue("workText") && !imageBlocks.length) {
      removeSectionByHeading(doc, "02 — THE WORK");
    }

    if (!inputValue("studioText")) {
      removeSectionByHeading(doc, "03 — STUDIO NOTES");
    }

    const pleasureNotes =
      typeof window.collectPleasureNotes === "function"
        ? window.collectPleasureNotes()
        : [];

    if (!pleasureNotes.length) {
      removeSectionByHeading(doc, "04 — PLEASURE NOTES");

      Array.from(doc.body.querySelectorAll("div")).forEach(
        function (element) {
          if (
            String(element.textContent || "").trim() ===
            "An offering of what has held my attention."
          ) {
            element.remove();
          }
        }
      );
    }

    if (
      !inputValue("inviteTitle") &&
      !inputValue("inviteText") &&
      !inputValue("ctaUrl")
    ) {
      removeSectionByHeading(doc, "05 — AN INVITATION");
    }

    return doc.body.innerHTML.trim();
  }

  function cleanPublishHtml(html) {
    if (!html || typeof DOMParser === "undefined") {
      return html;
    }

    const doc = new DOMParser().parseFromString(
      String(html),
      "text/html"
    );

    doc.querySelectorAll("img").forEach(function (image) {
      const src = String(
        image.getAttribute("src") || ""
      ).trim();

      if (!src || src === "#") image.remove();
    });

    doc.querySelectorAll("p").forEach(function (paragraph) {
      if (
        !hasMeaningfulText(paragraph) &&
        !hasMeaningfulImage(paragraph)
      ) {
        paragraph.remove();
      }
    });

    return removeEmptyEditorialSections(
      doc.body.innerHTML.trim()
    );
  }

  function collectPleasureSearchText() {
    if (typeof window.collectPleasureNotes !== "function") {
      return "";
    }

    return window
      .collectPleasureNotes()
      .map(function (note) {
        return [
          String(note.label || ""),
          String(note.value || "")
        ].join(" ");
      })
      .join(" ");
  }

  function collectPayload(secret) {
    if (typeof window.buildNewsletterHtml !== "function") {
      throw new Error(
        "The newsletter builder is not available. Reload the Dispatch add-in and try again."
      );
    }

    const edition = normalizeEdition();
    const rawHtml = window.buildNewsletterHtml();
    const html = cleanPublishHtml(rawHtml);

    if (typeof window.validateNewsletterHtml === "function") {
      const problem = window.validateNewsletterHtml(html);
      if (problem) throw new Error(problem);
    }

    if (!html) {
      throw new Error(
        "Publish stopped — the Dispatch contains no publishable content."
      );
    }

    const title = inputValue("title");
    const date =
      inputValue("date") ||
      new Date().toISOString().slice(0, 10);

    const pleasureText = collectPleasureSearchText();

    const subject =
      typeof window.buildSubject === "function"
        ? window.buildSubject()
        : "The Pleasure Dispatch — " +
          edition.label +
          ": " +
          (title || "A Note on Pleasure");

    return {
      action: "publish",
      secret: secret,
      edition: edition.number,
      editionLabel: edition.label,
      date: date,
      title: title || "The Pleasure Dispatch",
      subtitle: inputValue("subtitle"),
      subject: subject,
      reflection: inputValue("reflection"),
      html: html,
      searchText: [
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
        .filter(Boolean)
        .join(" ")
    };
  }

  function clearTimers() {
    if (pingTimer) {
      clearInterval(pingTimer);
      pingTimer = null;
    }

    if (timeoutTimer) {
      clearTimeout(timeoutTimer);
      timeoutTimer = null;
    }
  }

  function clearPending() {
    clearTimers();
    pendingPayload = null;
    pendingResolve = null;
    pendingReject = null;
  }

  function closeDialog() {
    if (!publisherDialog) return;

    try {
      publisherDialog.close();
    } catch (error) {
      console.warn(
        "[Pleasure Dispatch] Could not close publisher dialog:",
        error
      );
    }

    publisherDialog = null;
  }

  function failPublish(message) {
    const reject = pendingReject;

    clearPending();
    closeDialog();

    if (reject) reject(new Error(message));
  }

  function handleDialogMessage(arg) {
    const raw = arg && arg.message;
    if (!raw) return;

    let message;

    try {
      message = JSON.parse(raw);
    } catch (error) {
      console.warn(
        "[Pleasure Dispatch] Ignoring non-JSON publisher message."
      );
      return;
    }

    console.log(
      "[Pleasure Dispatch] Publisher message:",
      message
    );

    if (message.type === "publisherReady") {
      if (!publisherDialog || !pendingPayload) return;

      if (pingTimer) {
        clearInterval(pingTimer);
        pingTimer = null;
      }

      setPublishStatus("Sending Dispatch to publisher…");

      try {
        publisherDialog.messageChild(
          JSON.stringify({
            type: "publishPayload",
            payload: pendingPayload
          })
        );
      } catch (error) {
        failPublish(
          "Could not send the Dispatch to the publisher: " +
            error.message
        );
      }

      return;
    }

    if (message.type === "publishResult") {
      resultReceived = true;

      const resolve = pendingResolve;
      const reject = pendingReject;
      const data = message.data || message;

      closeDialog();
      clearPending();

      if (message.success) {
        if (resolve) resolve(data);
      } else if (reject) {
        reject(
          new Error(
            message.error || "Publishing failed."
          )
        );
      }
    }
  }

  function sendPing() {
    if (!publisherDialog || !pendingPayload) return;

    try {
      publisherDialog.messageChild(
        JSON.stringify({ type: "publisherPing" })
      );
    } catch (error) {
      console.warn(
        "[Pleasure Dispatch] Publisher ping failed:",
        error
      );
    }
  }

  function openPublisherDialog(payload) {
    return new Promise(function (resolve, reject) {
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

      closeDialog();
      clearPending();

      pendingPayload = payload;
      pendingResolve = resolve;
      pendingReject = reject;
      resultReceived = false;

      setPublishStatus("Opening secure publisher…");

      const dialogUrl =
        PUBLISHER_URL + "?action=publisherDialog";

      console.log(
        "[Pleasure Dispatch] Opening publisher:",
        dialogUrl
      );

      Office.context.ui.displayDialogAsync(
        dialogUrl,
        {
          height: 25,
          width: 30,
          displayInIframe: false
        },
        function (result) {
          if (
            result.status !==
            Office.AsyncResultStatus.Succeeded
          ) {
            failPublish(
              result.error && result.error.message
                ? result.error.message
                : "Could not open the publisher."
            );
            return;
          }

          publisherDialog = result.value;

          publisherDialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            handleDialogMessage
          );

          publisherDialog.addEventHandler(
            Office.EventType.DialogEventReceived,
            function (event) {
              console.log(
                "[Pleasure Dispatch] Publisher dialog event:",
                event
              );

              if (resultReceived) return;

              if (event && event.error === 12006) {
                failPublish(
                  "Publisher dialog closed before publishing completed."
                );
              }
            }
          );

          sendPing();
          pingTimer = setInterval(sendPing, PING_INTERVAL_MS);

          timeoutTimer = setTimeout(function () {
            if (
              !resultReceived &&
              publisherDialog &&
              pendingPayload
            ) {
              failPublish(
                "Publisher did not respond within 30 seconds. Check the publisher Apps Script deployment and try again."
              );
            }
          }, DIALOG_TIMEOUT_MS);
        }
      );
    });
  }

  async function publishDispatch() {
    const button = document.getElementById("publishBtn");

    if (button) button.disabled = true;

    try {
      setPublishStatus("Preparing Dispatch…");

      const secret = getPublishSecret();
      const payload = collectPayload(secret);

      console.log(
        "[Pleasure Dispatch] Publish payload prepared:",
        {
          edition: payload.edition,
          title: payload.title,
          htmlLength: payload.html.length
        }
      );

      setPublishStatus("Connecting to publisher…");

      const result = await openPublisherDialog(payload);

      console.log(
        "[Pleasure Dispatch] Publisher result:",
        result
      );

      setPublishStatus(
        "✓ Dispatch published successfully."
      );

      return result;
    } catch (error) {
      console.error(
        "[Pleasure Dispatch] Publish failed:",
        error
      );

      if (
        error &&
        /publish.*key|secret|unauthor/i.test(
          String(error.message || error)
        )
      ) {
        clearStoredPublishSecret();
      }

      setPublishStatus(
        "Publish failed: " +
          (
            error && error.message
              ? error.message
              : String(error)
          )
      );

      throw error;
    } finally {
      if (button) button.disabled = false;
    }
  }

  function bindPublishButton() {
    const button = document.getElementById("publishBtn");

    if (!button) return false;

    if (button.dataset.publishBound === "true") {
      return true;
    }

    button.dataset.publishBound = "true";
    button.type = "button";

    button.addEventListener("click", function (event) {
      event.preventDefault();
      event.stopPropagation();

      publishDispatch().catch(function () {
        /* Status is already shown to the user. */
      });
    });

    console.log(
      "[Pleasure Dispatch] Publish button bound."
    );

    return true;
  }

  window.publishDispatch = publishDispatch;

  window.setPublishSecret = function (secret) {
    const value = String(secret || "").trim();

    if (!value) {
      throw new Error(
        "Publishing key cannot be empty."
      );
    }

    window.localStorage.setItem(
      SECRET_STORAGE_KEY,
      value
    );

    return true;
  };

  window.clearPublishSecret =
    clearStoredPublishSecret;

  function initializePublisher() {
    bindPublishButton();
  }

  if (document.readyState === "loading") {
    document.addEventListener(
      "DOMContentLoaded",
      initializePublisher
    );
  } else {
    initializePublisher();
  }

  if (
    window.Office &&
    typeof Office.onReady === "function"
  ) {
    Office.onReady(function () {
      bindPublishButton();
    });
  }

})();
