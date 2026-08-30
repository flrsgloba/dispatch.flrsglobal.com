/* =========================================================
   THE PLEASURE DISPATCH
   publish.js
   Production Send → Publish bridge

   Notes:
   - Uses the dedicated publishing Google Apps Script endpoint.
   - The GitHub token remains server-side in Apps Script.
   - The publishing secret is requested at publish time and stored
     only in this browser's localStorage; it is never committed here.
   - Existing Build in Outlook behavior remains untouched.
========================================================= */
(function () {
  const PUBLISH_API_URL =
    "https://script.google.com/macros/s/AKfycbzavxknADmXnvAhRqcf9areGCRpfAJIZ62v84kqb_hpfgfAWIUbngcCH4B8M9TpkuA-uw/exec";

  const SECRET_STORAGE_KEY = "pd_publish_secret";

  function getPublishSecret() {
    let secret = window.localStorage.getItem(SECRET_STORAGE_KEY);

    if (!secret) {
      secret = window.prompt(
        "Enter your Dispatch publishing key:"
      );

      if (secret) {
        window.localStorage.setItem(
          SECRET_STORAGE_KEY,
          secret.trim()
        );
      }
    }

    return (secret || "").trim();
  }

  function clearStoredPublishSecret() {
    window.localStorage.removeItem(SECRET_STORAGE_KEY);
  }

  function status(message) {
    if (typeof setStatus === "function") {
      setStatus(message);
    } else {
      console.log("[Pleasure Dispatch]", message);
    }
  }

  function collectPayload(secret) {
    if (typeof buildNewsletterHtml !== "function") {
      throw new Error(
        "The newsletter builder is not available. Reload the Dispatch add-in and try again."
      );
    }

    const html = buildNewsletterHtml();

    if (typeof validateNewsletterHtml === "function") {
      const problem = validateNewsletterHtml(html);

      if (problem) {
        throw new Error(problem);
      }
    }

    const edition =
      typeof singleLineText === "function"
        ? singleLineText(value("edition"))
        : value("edition");

    const safeEdition = edition || "No. 001";
    const numberMatch = safeEdition.match(/(\d+(?:\.\d+)?)/);
    const number = numberMatch
      ? numberMatch[1]
      : safeEdition;

    const title =
      typeof singleLineText === "function"
        ? singleLineText(value("title"))
        : value("title");

    const date =
      value("date") ||
      new Date().toISOString().slice(0, 10);

    let pleasureText = "";

    if (typeof collectPleasureNotes === "function") {
      pleasureText = collectPleasureNotes()
        .map(function (note) {
          return (
            (note.label || "") +
            " " +
            (note.value || "")
          );
        })
        .join(" ");
    }

    let subject = "";

    if (typeof buildSubject === "function") {
      subject = buildSubject();
    }

    return {
      action: "publish",
      secret: secret,
      edition: number,
      editionLabel: safeEdition,
      date: date,
      title:
        title ||
        "The Pleasure Dispatch",
      subtitle: value("subtitle"),
      subject: subject,
      html: html,
      searchText: [
        number,
        safeEdition,
        date,
        title,
        value("subtitle"),
        value("reflection"),
        value("workText"),
        value("studioText"),
        value("inviteTitle"),
        value("inviteText"),
        value("question"),
        pleasureText
      ].join(" ")
    };
  }

  async function publishDispatch() {
    const button =
      document.getElementById("publishBtn");

    if (button) {
      button.disabled = true;
    }

    try {
      status("Publishing Dispatch…");

      const secret = getPublishSecret();

      if (!secret) {
        throw new Error(
          "A publishing key is required."
        );
      }

      const payload = collectPayload(secret);

      const response = await fetch(
        PUBLISH_API_URL,
        {
          method: "POST",
          headers: {
            "Content-Type":
              "text/plain;charset=utf-8"
          },
          body: JSON.stringify(payload)
        }
      );

      if (!response.ok) {
        throw new Error(
          "Publish service returned HTTP " +
          response.status + "."
        );
      }

      const result =
        await response.json();

      if (
        !result ||
        result.status !== "published"
      ) {
        if (
          result &&
          result.message === "Unauthorized."
        ) {
          clearStoredPublishSecret();
          throw new Error(
            "Publishing key rejected. Please publish again and enter the current key."
          );
        }

        throw new Error(
          result && result.message
            ? result.message
            : "Publish service did not publish the Dispatch."
        );
      }

      status(
        "✓ Published " +
        payload.editionLabel +
        " to dispatch.flrsglobal.com."
      );

      console.log(
        "[Pleasure Dispatch] Published:",
        result
      );
    } catch (error) {
      status(
        "Publish failed: " +
        error.message
      );

      console.error(
        "[Pleasure Dispatch] Publish failed:",
        error
      );
    } finally {
      if (button) {
        button.disabled = false;
      }
    }
  }

  function bindPublishButton() {
    const button =
      document.getElementById("publishBtn");

    if (
      !button ||
      button.dataset.bound === "true"
    ) {
      return;
    }

    button.dataset.bound = "true";

    button.addEventListener(
      "click",
      function (event) {
        event.preventDefault();
        event.stopPropagation();
        publishDispatch();
      }
    );
  }

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
})();
