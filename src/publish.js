/* The Pleasure Dispatch — publish bridge */
(function () {
  const PUBLISH_API_URL =
    "https://script.google.com/macros/s/AKfycbyTLvYbe1O_BbzsH09UMSZdbY9_XZXga-TbSPkR3UclT3Qhlaj7gy5yhXPA_UpE6Fu7tw/exec";

  function getPublishSecret() {
    let secret = window.localStorage.getItem("pd_publish_secret");
    if (!secret) {
      secret = window.prompt("Enter your Dispatch publishing key:");
      if (secret) window.localStorage.setItem("pd_publish_secret", secret);
    }
    return secret || "";
  }

  function status(message) {
    if (typeof setStatus === "function") setStatus(message);
    else console.log("[Pleasure Dispatch]", message);
  }

  function collectPayload() {
    const html = buildNewsletterHtml();
    const problem = validateNewsletterHtml(html);
    if (problem) throw new Error(problem);

    const edition = singleLineText(value("edition")) || "No. 001";
    const numberMatch = edition.match(/(\d+(?:\.\d+)?)/);
    const number = numberMatch ? numberMatch[1] : edition;
    const title = singleLineText(value("title")) || "The Pleasure Dispatch";
    const date = value("date") || new Date().toISOString().slice(0, 10);

    return {
      action: "publish",
      secret: getPublishSecret(),
      edition: number,
      editionLabel: edition,
      date: date,
      title: title,
      subtitle: value("subtitle"),
      subject: buildSubject(),
      html: html,
      searchText: [
        number, edition, date, title, value("subtitle"),
        value("reflection"), value("workText"), value("studioText"),
        value("inviteTitle"), value("inviteText"), value("question"),
        collectPleasureNotes().map(function (n) {
          return n.label + " " + n.value;
        }).join(" ")
      ].join(" ")
    };
  }

  async function publishDispatch() {
    const button = document.getElementById("publishBtn");
    if (button) button.disabled = true;

    try {
      status("Publishing Dispatch…");
      const payload = collectPayload();
      if (!payload.secret) throw new Error("A publishing key is required.");

      const response = await fetch(PUBLISH_API_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(payload)
      });

      if (!response.ok) {
        throw new Error("Publish service returned HTTP " + response.status + ".");
      }

      const result = await response.json();
      if (!result || result.status !== "published") {
        throw new Error(result && result.message
          ? result.message
          : "Publish service did not publish the Dispatch.");
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
      publishDispatch();
    });
  }

  window.addEventListener("load", bindPublishButton);
  setTimeout(bindPublishButton, 250);
  setTimeout(bindPublishButton, 1000);
  setTimeout(bindPublishButton, 2000);
  window.publishDispatch = publishDispatch;
})();
