/* =========================================================
   THE PLEASURE DISPATCH
   publish.js
   Production Send → Publish bridge

   IMPORTANT:
   Publishing uses the same Apps Script deployment as the
   Drive bridge. The compatibility patch intercepts fetch()
   requests, so publishing uses XMLHttpRequest here to avoid
   being rerouted to the retired publisher deployment.
========================================================= */
(function () {
  "use strict";

  const PUBLISH_API_URL =
    "https://script.google.com/macros/s/AKfycbzavxknADmXnvAhRqcf9areGCRpfAJIZ62v84kqb_hpfgfAWIUbngcCH4B8M9TpkuA-uw/exec";

  const SECRET_STORAGE_KEY = "pd_publish_secret";

  function getPublishSecret() {
    const stored = window.localStorage.getItem(SECRET_STORAGE_KEY);
    if (stored) return Promise.resolve(stored.trim());

    return new Promise(function (resolve) {
      const existing = document.getElementById("pdPublishKeyDialog");
      if (existing) existing.remove();

      const overlay = document.createElement("div");
      overlay.id = "pdPublishKeyDialog";
      overlay.setAttribute("role", "dialog");
      overlay.setAttribute("aria-modal", "true");
      overlay.style.cssText = "position:fixed;inset:0;z-index:999999;display:flex;align-items:center;justify-content:center;padding:20px;background:rgba(0,0,0,.55);box-sizing:border-box";

      const panel = document.createElement("div");
      panel.style.cssText = "width:100%;max-width:360px;box-sizing:border-box;padding:24px;background:#303030;color:#F2EEE5;border:1px solid #777;box-shadow:0 12px 40px rgba(0,0,0,.35);font-family:Arial,sans-serif";

      const title = document.createElement("div");
      title.textContent = "Publish Dispatch";
      title.style.cssText = "font-size:16px;font-weight:600;margin-bottom:8px;";

      const description = document.createElement("div");
      description.textContent = "Enter your publishing key to send this Dispatch to the public archive.";
      description.style.cssText = "font-size:13px;line-height:1.5;color:#C9C3B8;margin-bottom:16px;";

      const input = document.createElement("input");
      input.type = "password";
      input.autocomplete = "off";
      input.spellcheck = false;
      input.style.cssText = "display:block;width:100%;box-sizing:border-box;height:40px;padding:8px 10px;background:#595959;color:#F2EEE5;border:1px solid #777;outline:none;font:14px Arial,sans-serif";

      const error = document.createElement("div");
      error.style.cssText = "display:none;color:#F2EEE5;font-size:12px;margin-top:8px;";

      const actions = document.createElement("div");
      actions.style.cssText = "display:flex;justify-content:flex-end;gap:8px;margin-top:18px;";

      const cancel = document.createElement("button");
      cancel.type = "button";
      cancel.textContent = "Cancel";
      cancel.style.cssText = "height:36px;padding:0 14px;background:transparent;color:#C9C3B8;border:1px solid #777;cursor:pointer";

      const submit = document.createElement("button");
      submit.type = "button";
      submit.textContent = "Publish";
      submit.style.cssText = "height:36px;padding:0 14px;background:#D8D0C3;color:#303030;border:1px solid #D8D0C3;cursor:pointer;font-weight:600";

      function close(value) {
        document.removeEventListener("keydown", onKeyDown, true);
        overlay.remove();
        resolve(value || "");
      }

      function submitKey() {
        const secret = String(input.value || "").trim();
        if (!secret) {
          error.textContent = "Please enter your publishing key.";
          error.style.display = "block";
          input.focus();
          return;
        }
        window.localStorage.setItem(SECRET_STORAGE_KEY, secret);
        close(secret);
      }

      function onKeyDown(event) {
        if (event.key === "Escape") {
          event.preventDefault();
          close("");
        } else if (event.key === "Enter") {
          event.preventDefault();
          submitKey();
        }
      }

      cancel.addEventListener("click", function () { close(""); });
      submit.addEventListener("click", submitKey);
      document.addEventListener("keydown", onKeyDown, true);

      actions.appendChild(cancel);
      actions.appendChild(submit);
      panel.appendChild(title);
      panel.appendChild(description);
      panel.appendChild(input);
      panel.appendChild(error);
      panel.appendChild(actions);
      overlay.appendChild(panel);
      document.body.appendChild(overlay);
      input.focus();
    });
  }

  function clearStoredPublishSecret() {
    window.localStorage.removeItem(SECRET_STORAGE_KEY);
  }

  function status(message) {
    if (typeof setStatus === "function") setStatus(message);
    else console.log("[Pleasure Dispatch]", message);
  }

  function hasMeaningfulImage(element) {
    const images = element.querySelectorAll ? element.querySelectorAll("img") : [];
    for (let i = 0; i < images.length; i++) {
      const src = (images[i].getAttribute("src") || "").trim();
      if (src && src !== "#") return true;
    }
    return false;
  }

  function hasMeaningfulText(element) {
    const clone = element.cloneNode(true);
    clone.querySelectorAll("img,svg,style,script,noscript").forEach(function (node) { node.remove(); });
    return (clone.textContent || "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim().length > 0;
  }

  function looksLikeOptionalModule(element) {
    const className = String(element.className || "").toLowerCase();
    const id = String(element.id || "").toLowerCase();
    return /module|section|block|figure|image|text|reflection|studio|invite|question|pleasure|hero/.test((className + " " + id).trim());
  }

  function cleanPublishHtml(html) {
    if (!html || typeof DOMParser === "undefined") return html;
    const parser = new DOMParser();
    const document = parser.parseFromString(html, "text/html");

    document.querySelectorAll("img").forEach(function (image) {
      const src = (image.getAttribute("src") || "").trim();
      if (!src || src === "#") image.remove();
    });

    Array.from(document.body.querySelectorAll("div,section,article,figure,td")).reverse().forEach(function (element) {
      if (!element.parentNode || !looksLikeOptionalModule(element)) return;
      if (!hasMeaningfulText(element) && !hasMeaningfulImage(element)) element.remove();
    });

    document.querySelectorAll("p").forEach(function (paragraph) {
      if (!hasMeaningfulText(paragraph) && !hasMeaningfulImage(paragraph)) paragraph.remove();
    });

    return document.body.innerHTML.trim();
  }

  function collectPayload(secret) {
    if (typeof buildNewsletterHtml !== "function") {
      throw new Error("The newsletter builder is not available. Reload the Dispatch add-in and try again.");
    }

    const rawHtml = buildNewsletterHtml();
    const html = cleanPublishHtml(rawHtml);

    if (typeof validateNewsletterHtml === "function") {
      const problem = validateNewsletterHtml(html);
      if (problem) throw new Error(problem);
    }

    const edition = typeof singleLineText === "function" ? singleLineText(value("edition")) : value("edition");
    const safeEdition = edition || "No. 001";
    const numberMatch = safeEdition.match(/(\d+(?:\.\d+)?)/);
    const number = numberMatch ? numberMatch[1] : safeEdition;
    const title = typeof singleLineText === "function" ? singleLineText(value("title")) : value("title");
    const date = value("date") || new Date().toISOString().slice(0, 10);

    let pleasureText = "";
    if (typeof collectPleasureNotes === "function") {
      pleasureText = collectPleasureNotes().map(function (note) {
        return (note.label || "") + " " + (note.value || "");
      }).join(" ");
    }

    let subject = "";
    if (typeof buildSubject === "function") subject = buildSubject();

    return {
      action: "publish",
      publishKey: secret,
      edition: number,
      editionLabel: safeEdition,
      date: date,
      title: title || "The Pleasure Dispatch",
      subtitle: value("subtitle"),
      subject: subject,
      html: html,
      searchText: [number, safeEdition, date, title, value("subtitle"), value("reflection"), value("workText"), value("studioText"), value("inviteTitle"), value("inviteText"), value("question"), pleasureText].join(" ")
    };
  }

  function postPublish(payload) {
    return new Promise(function (resolve, reject) {
      const xhr = new XMLHttpRequest();
      xhr.open("POST", PUBLISH_API_URL, true);
      xhr.setRequestHeader("Content-Type", "text/plain;charset=utf-8");
      xhr.timeout = 30000;

      xhr.onload = function () {
        const responseText = xhr.responseText || "";
        let result = null;

        try {
          result = responseText ? JSON.parse(responseText) : null;
        } catch (error) {
          reject(new Error("Publish service returned an unreadable response (HTTP " + xhr.status + ")."));
          return;
        }

        console.log("[Pleasure Dispatch] Publish HTTP status:", xhr.status);
        console.log("[Pleasure Dispatch] Publish response:", result);

        if (xhr.status < 200 || xhr.status >= 300) {
          reject(new Error((result && result.message) || "Publish service returned HTTP " + xhr.status + "."));
          return;
        }

        resolve(result);
      };

      xhr.onerror = function () {
        reject(new Error("Could not connect to the Dispatch publishing bridge."));
      };

      xhr.ontimeout = function () {
        reject(new Error("The Dispatch publishing bridge timed out."));
      };

      xhr.send(JSON.stringify(payload));
    });
  }

  async function publishDispatch() {
    const button = document.getElementById("publishBtn");
    if (button) button.disabled = true;

    try {
      status("Publishing Dispatch…");
      const secret = await getPublishSecret();
      if (!secret) {
        status("Publish cancelled.");
        return;
      }

      const payload = collectPayload(secret);
      const result = await postPublish(payload);

      if (!result || result.status !== "published") {
        if (result && result.code === "INVALID_PUBLISH_KEY") {
          clearStoredPublishSecret();
          throw new Error("Publishing key rejected. Click Publish again and enter the current key.");
        }
        throw new Error(result && result.message ? result.message : "Publish service did not publish the Dispatch.");
      }

      status("✓ Published " + payload.editionLabel + " to dispatch.flrsglobal.com.");
      console.log("[Pleasure Dispatch] Published:", result);
    } catch (error) {
      status("Publish failed: " + error.message);
      console.error("[Pleasure Dispatch] Publish failed:", error);
    } finally {
      if (button) button.disabled = false;
    }
  }

  function bindPublishButton() {
    const button = document.getElementById("publishBtn");
    if (!button || button.dataset.bound === "true") return;
    button.dataset.bound = "true";
    button.addEventListener("click", function (event) {
      event.preventDefault();
      event.stopPropagation();
      publishDispatch();
    });
  }

  window.addEventListener("load", bindPublishButton);
  setTimeout(bindPublishButton, 250);
  setTimeout(bindPublishButton, 1000);
  setTimeout(bindPublishButton, 2000);
  window.publishDispatch = publishDispatch;
})();

/* =========================================================
   OPTIONAL MODULE CLEANUP
========================================================= */
(function () {
  if (typeof window === "undefined") return;
  const originalBuildNewsletterHtml = window.buildNewsletterHtml;
  if (typeof originalBuildNewsletterHtml !== "function") return;

  function fieldHasText(id) {
    const element = document.getElementById(id);
    return !!(element && String(element.value || "").trim());
  }

  function hasPleasureNotes() {
    if (typeof collectPleasureNotes === "function") {
      return collectPleasureNotes().some(function (note) {
        return String(note.label || "").trim() || String(note.value || "").trim();
      });
    }
    const rows = document.querySelectorAll("#pleasureRows > *");
    for (let i = 0; i < rows.length; i++) {
      const inputs = rows[i].querySelectorAll ? rows[i].querySelectorAll("input,textarea") : [];
      for (let j = 0; j < inputs.length; j++) {
        if (String(inputs[j].value || "").trim()) return true;
      }
    }
    return false;
  }

  function removeHeading(html, label) {
    const escapedLabel = label.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    return html.replace(new RegExp("<div[^>]*>\\s*" + escapedLabel + "\\s*</div>", "gi"), "");
  }

  window.buildNewsletterHtml = function () {
    let html = originalBuildNewsletterHtml.apply(this, arguments);
    if (!fieldHasText("reflection")) html = removeHeading(html, "01 — A REFLECTION");
    if (!fieldHasText("workText")) html = removeHeading(html, "02 — THE WORK");
    if (!fieldHasText("studioText")) html = removeHeading(html, "03 — STUDIO NOTES");
    if (!hasPleasureNotes()) {
      html = removeHeading(html, "04 — PLEASURE NOTES");
      html = html.replace(/<div[^>]*>\s*An offering of what has held my attention\.\s*<\/div>/gi, "");
    }
    if (!fieldHasText("question")) html = removeHeading(html, "06 — A QUESTION");
    return html;
  };
})();
