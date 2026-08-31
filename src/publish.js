/* =========================================================
   THE PLEASURE DISPATCH
   publish.js

   Minimal Outlook publishing bridge.

   Outlook taskpane
        ↓
   Office Dialog
        ↓
   Google Apps Script publisherDialog
        ↓
   Google Apps Script publishFromDialog()
        ↓
   GitHub

   The dialog is served by Google Apps Script. The actual publish
   operation is performed server-side with google.script.run so the
   browser never POSTs to the Apps Script web-app URL.
========================================================= */

(function () {
  "use strict";

  const PUBLISHER_URL =
    "https://script.google.com/macros/s/AKfycbzavxknADmXnvAhRqcf9areGCRpfAJIZ62v84kqb_hpfgfAWIUbngcCH4B8M9TpkuA-uw/exec";

  const SECRET_STORAGE_KEY = "pd_publish_secret";
  const DIALOG_TIMEOUT_MS = 30000;

  let dialog = null;
  let pendingPayload = null;
  let pendingResolve = null;
  let pendingReject = null;
  let resultReceived = false;
  let pingTimer = null;
  let timeoutTimer = null;

  function status(message) {
    if (typeof window.setStatus === "function") window.setStatus(message);
    console.log("[Pleasure Dispatch]", message);
  }

  function inputValue(id) {
    const element = document.getElementById(id);
    return element ? String(element.value || "").trim() : "";
  }

  function getPublishSecret() {
    const secret = window.localStorage.getItem(SECRET_STORAGE_KEY);
    if (!secret || !secret.trim()) {
      throw new Error("Publishing key is not configured.");
    }
    return secret.trim();
  }

  function normalizeEdition() {
    const raw = inputValue("edition");
    const match = raw.match(/(\d+(?:\.\d+)?)/);
    if (!match) {
      throw new Error("Publish stopped — enter an edition number, such as No. 001.");
    }
    const number = match[1];
    if (!isFinite(Number(number)) || Number(number) < 0) {
      throw new Error("Publish stopped — invalid edition number.");
    }
    return { number: number, label: "No. " + number };
  }

  function hasMeaningfulImage(element) {
    if (!element || !element.querySelectorAll) return false;
    const images = element.querySelectorAll("img");
    for (let i = 0; i < images.length; i++) {
      const src = String(images[i].getAttribute("src") || "").trim();
      if (src && src !== "#") return true;
    }
    return false;
  }

  function hasMeaningfulText(element) {
    if (!element) return false;
    const clone = element.cloneNode(true);
    clone.querySelectorAll("img,svg,style,script,noscript").forEach(function (node) {
      node.remove();
    });
    return String(clone.textContent || "")
      .replace(/\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim().length > 0;
  }

  function removeSectionByHeading(doc, headingText) {
    const elements = Array.from(doc.body.querySelectorAll("div,td"));
    for (let i = 0; i < elements.length; i++) {
      const element = elements[i];
      if (String(element.textContent || "").trim() !== headingText) continue;
      const parent = element.parentNode;
      if (!parent) continue;
      const children = Array.from(parent.children);
      const index = children.indexOf(element);
      element.remove();
      const next = children[index + 1];
      if (next && !hasMeaningfulText(next) && !hasMeaningfulImage(next)) next.remove();
      return;
    }
  }

  function removeEmptyEditorialSections(html) {
    if (!html || typeof DOMParser === "undefined") return html;
    const doc = new DOMParser().parseFromString(String(html), "text/html");

    if (!inputValue("reflection")) {
      removeSectionByHeading(doc, "01 — A REFLECTION");
    }

    const imageBlocks = typeof window.collectImageBlocks === "function"
      ? window.collectImageBlocks()
      : [];

    if (!inputValue("workText") && !imageBlocks.length) {
      removeSectionByHeading(doc, "02 — THE WORK");
    }

    if (!inputValue("studioText")) {
      removeSectionByHeading(doc, "03 — STUDIO NOTES");
    }

    const notes = typeof window.collectPleasureNotes === "function"
      ? window.collectPleasureNotes()
      : [];

    if (!notes.length) {
      removeSectionByHeading(doc, "04 — PLEASURE NOTES");
      Array.from(doc.body.querySelectorAll("div")).forEach(function (element) {
        if (String(element.textContent || "").trim() === "An offering of what has held my attention.") {
          element.remove();
        }
      });
    }

    if (!inputValue("inviteTitle") && !inputValue("inviteText") && !inputValue("ctaUrl")) {
      removeSectionByHeading(doc, "05 — AN INVITATION");
    }

    return doc.body.innerHTML.trim();
  }

  function cleanPublishHtml(html) {
    if (!html || typeof DOMParser === "undefined") return html;
    const doc = new DOMParser().parseFromString(String(html), "text/html");

    doc.querySelectorAll("img").forEach(function (image) {
      const src = String(image.getAttribute("src") || "").trim();
      if (!src || src === "#") image.remove();
    });

    doc.querySelectorAll("p").forEach(function (paragraph) {
      if (!hasMeaningfulText(paragraph) && !hasMeaningfulImage(paragraph)) {
        paragraph.remove();
      }
    });

    return removeEmptyEditorialSections(doc.body.innerHTML.trim());
  }

  function collectPayload(secret) {
    if (typeof window.buildNewsletterHtml !== "function") {
      throw new Error("The newsletter builder is not available. Reload the Dispatch add-in and try again.");
    }

    const edition = normalizeEdition();
    const html = cleanPublishHtml(window.buildNewsletterHtml());

    if (typeof window.validateNewsletterHtml === "function") {
      const problem = window.validateNewsletterHtml(html);
      if (problem) throw new Error(problem);
    }

    if (!html) {
      throw new Error("Publish stopped — the Dispatch contains no publishable content.");
    }

    const title = inputValue("title");
    const date = inputValue("date") || new Date().toISOString().slice(0, 10);

    let pleasureText = "";
    if (typeof window.collectPleasureNotes === "function") {
      pleasureText = window.collectPleasureNotes().map(function (note) {
        return String(note.label || "") + " " + String(note.value || "");
      }).join(" ");
    }

    const subject = typeof window.buildSubject === "function"
      ? window.buildSubject()
      : "The Pleasure Dispatch — " + edition.label + ": " + (title || "A Note on Pleasure");

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
      ].filter(Boolean).join(" ")
    };
  }

  function clearPending() {
    if (pingTimer) {
      clearInterval(pingTimer);
      pingTimer = null;
    }
    if (timeoutTimer) {
      clearTimeout(timeoutTimer);
      timeoutTimer = null;
    }
    pendingPayload = null;
    pendingResolve = null;
    pendingReject = null;
  }

  function fail(message) {
    const reject = pendingReject;
    clearPending();
    if (dialog) {
      try { dialog.close(); } catch (error) {}
    }
    dialog = null;
    if (reject) reject(new Error(message));
  }

  function handleDialogMessage(arg) {
    const raw = arg && arg.message;
    if (!raw) return;

    let message;
    try {
      message = JSON.parse(raw);
    } catch (error) {
      console.warn("[Pleasure Dispatch] Ignoring non-JSON dialog message.");
      return;
    }

    console.log("[Pleasure Dispatch] Publisher message:", message);

    if (message.type === "publisherReady") {
      if (!dialog || !pendingPayload) return;

      if (pingTimer) {
        clearInterval(pingTimer);
        pingTimer = null;
      }

      status("Sending Dispatch to publisher…");

      try {
        dialog.messageChild(JSON.stringify({
          type: "publishPayload",
          payload: pendingPayload
        }));
      } catch (error) {
        fail("Could not send the Dispatch to the publisher: " + error.message);
      }
      return;
    }

    if (message.type === "publishResult") {
      resultReceived = true;

      const resolve = pendingResolve;
      const reject = pendingReject;
      const data = message.data || message;

      if (dialog) {
        try { dialog.close(); } catch (error) {}
      }

      dialog = null;
      clearPending();

      if (message.success) {
        if (resolve) resolve(data);
      } else if (reject) {
        reject(new Error(message.error || "Publishing failed."));
      }
    }
  }

  function sendPing() {
    if (!dialog || !pendingPayload) return;
    try {
      dialog.messageChild(JSON.stringify({ type: "publisherPing" }));
    } catch (error) {
      console.warn("[Pleasure Dispatch] Publisher ping failed:", error);
    }
  }

  function openPublisherDialog(payload) {
    return new Promise(function (resolve, reject) {
      if (!window.Office || !Office.context || !Office.context.ui ||
          typeof Office.context.ui.displayDialogAsync !== "function") {
        reject(new Error("Office Dialog API is unavailable in this Outlook context."));
        return;
      }

      if (dialog) {
        try { dialog.close(); } catch (error) {}
        dialog = null;
      }

      pendingPayload = payload;
      pendingResolve = resolve;
      pendingReject = reject;
      resultReceived = false;

      status("Opening secure publisher…");
      console.log("[Pleasure Dispatch] Opening Apps Script publisher dialog:", PUBLISHER_URL);

      const dialogUrl = PUBLISHER_URL + "?action=publisherDialog";

      Office.context.ui.displayDialogAsync(dialogUrl, {
        height: 25,
        width: 30,
        displayInIframe: false
      }, function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          fail(result.error && result.error.message
            ? result.error.message
            : "Could not open the publisher.");
          return;
        }

        dialog = result.value;
        console.log("[Pleasure Dispatch] Publisher dialog opened.");

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          handleDialogMessage
        );

        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          function (event) {
            console.log("[Pleasure Dispatch] Publisher dialog event:", event);
            if (resultReceived) return;
            if (event && event.error === 12006) {
              fail("Publisher dialog closed before publishing completed.");
            }
          }
        );

        sendPing();
        pingTimer = setInterval(sendPing, 500);

        timeoutTimer = setTimeout(function () {
          if (!resultReceived && dialog && pendingPayload) {
            fail("Publisher did not respond within 30 seconds. Check the Apps Script deployment and try again.");
          }
        }, DIALOG_TIMEOUT_MS);
      });
    });
  }

  async function publishDispatch() {
    const button = document.getElementById("publishBtn");
    if (button) button.disabled = true;

    try {
      status("Preparing Dispatch…");
      const secret = getPublishSecret();
      const payload = collectPayload(secret);

      console.log("[Pleasure Dispatch] Payload prepared:", {
        edition: payload.edition,
        title: payload.title,
        htmlLength: payload.html.length
      });

      status("Connecting to publisher…");
      const result = await openPublisherDialog(payload);

      console.log("[Pleasure Dispatch] Publisher result:", result);
      status("✓ Dispatch published successfully.");
    } catch (error) {
      console.error("[Pleasure Dispatch] Publish failed:", error);
      status("Publish failed: " + (error && error.message ? error.message : String(error)));
    } finally {
      if (button) button.disabled = false;
    }
  }

  function bindPublishButton() {
    const button = document.getElementById("publishBtn");
    if (!button || button.dataset.publishBound === "true") return !!button;

    button.dataset.publishBound = "true";
    button.type = "button";

    button.addEventListener("click", function (event) {
      event.preventDefault();
      event.stopPropagation();
      publishDispatch();
    }, false);

    console.log("[Pleasure Dispatch] Publish button bound.");
    return true;
  }

  function initializePublishBridge() {
    if (bindPublishButton()) return;
    setTimeout(initializePublishBridge, 250);
  }

  if (window.Office && typeof Office.onReady === "function") {
    Office.onReady(initializePublishBridge);
  } else {
    window.addEventListener("load", initializePublishBridge);
  }

  window.publishDispatch = publishDispatch;
})();
