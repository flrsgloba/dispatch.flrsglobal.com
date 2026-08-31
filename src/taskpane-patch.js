/* =========================================================
   THE PLEASURE DISPATCH
   taskpane-patch.js
   Production compatibility patch

   Fixes:
   - Routes Drive requests to the current Drive Apps Script.
   - Routes publishing requests to the current publisher Apps Script.
   - Uses dispatch.flrsglobal.com for the public lightbox.
   - Uses dispatch.flrsglobal.com for The Loop image.
   - Removes inline CSS from newsletter HTML before Outlook insertion.
   - Uses Outlook Body.setAsync with HostConfig when available.
   - Keeps the existing editor, image modules, Drive storage,
     progress counter, contacts, and publisher intact.
========================================================= */

const PD_DRIVE_API_URL =
  "https://script.google.com/macros/s/AKfycbzavxknADmXnvAhRqcf9areGCRpfAJIZ62v84kqb_hpfgfAWIUbngcCH4B8M9TpkuA-uw/exec";

const PD_PUBLISH_API_URL =
  "https://script.google.com/macros/s/AKfycbyTLvYbe1O_BbzsH09UMSZdbY9_XZXga-TbSPkR3UclT3Qhlaj7gy5yhXPA_UpE6Fu7tw/exec";

const PD_LEGACY_API_URL =
  "https://script.google.com/macros/s/AKfycbzavxknADmXnvAhRqcf9areGCRpfAJIZ62v84kqb_hpfgfAWIUbngcCH4B8M9TpkuA-uw/exec";

const PD_PUBLIC_SITE =
  "https://dispatch.flrsglobal.com";

const PD_LIGHTBOX_BASE_URL =
  PD_PUBLIC_SITE +
  "/assets/lightbox.html";

const PD_LOOP_ASSET_URL =
  PD_PUBLIC_SITE +
  "/assets/pleasure-loop.png";

/* =========================================================
   API ROUTING

   The current contacts.js and publish.js still point at the
   previous bridge URL. Route those requests here without
   changing their working logic.
========================================================= */

(function installApiRouter() {

  if (
    window.__pdApiRouterInstalled
  ) {
    return;
  }

  const nativeFetch =
    window.fetch.bind(window);

  window.fetch = function (input, init) {

    let url =
      typeof input === "string"
        ? input
        : input && input.url
          ? input.url
          : "";

    if (
      url.indexOf(PD_LEGACY_API_URL) === 0
    ) {

      let action = "";

      try {
        const parsed =
          new URL(url);
        action =
          String(
            parsed.searchParams.get("action") || ""
          ).toLowerCase();
      } catch (error) {
        action = "";
      }

      let target =
        PD_DRIVE_API_URL;

      if (
        init &&
        String(init.method || "GET").toUpperCase() === "POST"
      ) {

        try {
          const body =
            typeof init.body === "string"
              ? JSON.parse(init.body)
              : null;

          if (
            body &&
            body.action === "publish"
          ) {
            target =
              PD_PUBLISH_API_URL;
          }
        } catch (error) {
          /* Leave the request routed to Drive if the body is not JSON. */
        }
      }

      url =
        target +
        (
          url.indexOf("?") >= 0
            ? url.substring(url.indexOf("?"))
            : ""
        );

      if (
        typeof input === "string"
      ) {
        input = url;
      } else if (input) {
        input = new Request(
          url,
          input
        );
      }
    }

    return nativeFetch(
      input,
      init
    );
  };

  window.__pdApiRouterInstalled =
    true;

})();

/* =========================================================
   PUBLIC URL BUILDERS
========================================================= */

function buildLightboxUrl(fileId, caption) {

  let url =
    PD_LIGHTBOX_BASE_URL +
    "?id=" +
    encodeURIComponent(fileId);

  if (caption) {
    url +=
      "&caption=" +
      encodeURIComponent(caption);
  }

  return url;
}

/* =========================================================
   DRIVE UPLOAD OVERRIDE
========================================================= */

async function uploadToDriveOnce(
  dataUrl,
  originalFileName
) {

  const payload = {
    action: "upload",
    publishKey:
      typeof PUBLISH_KEY !== "undefined"
        ? PUBLISH_KEY
        : "",
    fileName:
      "PD_" +
      getEditionSafeName() +
      "_" +
      Date.now() +
      "_" +
      cleanFileName(originalFileName),
    mimeType: "image/jpeg",
    fileContent:
      stripDataUrlPrefix(dataUrl)
  };

  let response;

  try {

    response =
      await fetch(
        PD_DRIVE_API_URL,
        {
          method: "POST",
          headers: {
            "Content-Type":
              "text/plain;charset=utf-8"
          },
          body:
            JSON.stringify(payload)
        }
      );

  } catch (error) {

    throw new Error(
      "Could not connect to Google Drive: " +
      error.message
    );

  }

  if (!response.ok) {

    throw new Error(
      "Google Drive returned HTTP " +
      response.status +
      "."
    );

  }

  let result;

  try {
    result =
      await response.json();
  } catch (error) {
    throw new Error(
      "Google Drive returned an unreadable response."
    );
  }

  if (
    !result ||
    result.status !== "created"
  ) {
    throw new Error(
      result && result.message
        ? result.message
        : "Google Drive did not create the image."
    );
  }

  result.imageUrl =
    buildDriveImageUrl(
      result.fileId
    );

  result.driveUrl =
    buildDriveFileUrl(
      result.fileId
    );

  result.lightboxUrl =
    buildLightboxUrl(
      result.fileId
    );

  return result;
}

/* =========================================================
   OUTLOOK HTML NORMALIZER

   Outlook Body.setAsync does not support inline CSS.
   The existing newsletter builder uses inline styles heavily,
   so convert each unique style attribute into a CSS class.
========================================================= */

function normalizeNewsletterHtmlForOutlook(html) {

  if (
    !html ||
    typeof DOMParser === "undefined"
  ) {
    return html;
  }

  const parser =
    new DOMParser();

  const document =
    parser.parseFromString(
      String(html),
      "text/html"
    );

  const styleMap =
    new Map();

  let styleCounter = 0;

  Array.from(
    document.querySelectorAll("[style]")
  ).forEach(function (element) {

    const style =
      String(
        element.getAttribute("style") || ""
      ).trim();

    if (!style) {
      element.removeAttribute("style");
      return;
    }

    let className =
      styleMap.get(style);

    if (!className) {
      styleCounter++;
      className =
        "pd-style-" +
        styleCounter;
      styleMap.set(
        style,
        className
      );
    }

    element.removeAttribute("style");
    element.classList.add(
      className
    );
  });

  document
    .querySelectorAll("svg")
    .forEach(function (svg) {
      svg.remove();
    });

  document
    .querySelectorAll("a")
    .forEach(function (anchor) {
      if (!anchor.id) {
        anchor.id = "LPNoLP";
      }
    });

  const css =
    Array.from(styleMap.entries())
      .map(function (entry) {
        return (
          "." +
          entry[1] +
          "{" +
          entry[0] +
          "}"
        );
      })
      .join("");

  return (
    "<style type=\"text/css\">" +
    css +
    "</style>" +
    document.body.innerHTML
  );
}

/* =========================================================
   NEWSLETTER BUILDER WRAPPER
========================================================= */

(function wrapNewsletterBuilder() {

  function install() {

    if (
      typeof window.buildNewsletterHtml !==
      "function"
    ) {
      return false;
    }

    if (
      window.__pdProductionBuilderWrapped
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
          html &&
          typeof html === "string"
        ) {

          html =
            html.replace(
              /https:\/\/flrsgloba\.github\.io\/dispatch\.flrsglobal\.com\/assets\/pleasure-loop\.png/gi,
              PD_LOOP_ASSET_URL
            );

          html =
            html.replace(
              /https:\/\/flrsgloba\.github\.io\/dispatch\.flrsglobal\.com\/assets\/pleasure-loop\.svg/gi,
              PD_LOOP_ASSET_URL
            );
        }

        return html;
      };

    window.__pdProductionBuilderWrapped =
      true;

    return true;
  }

  if (!install()) {
    setTimeout(install, 250);
    setTimeout(install, 1000);
    setTimeout(install, 2000);
  }

})();

/* =========================================================
   OUTLOOK BUILD OVERRIDE
========================================================= */

function buildInOutlook() {

  if (
    typeof Office === "undefined"
  ) {
    setStatus(
      "Build failed — Office.js is unavailable."
    );
    return;
  }

  if (
    !Office.context ||
    !Office.context.mailbox ||
    !Office.context.mailbox.item
  ) {
    setStatus(
      "Build failed — open The Pleasure Dispatch from a new Outlook message."
    );
    return;
  }

  const item =
    Office.context.mailbox.item;

  if (
    !item.body ||
    typeof item.body.getAsync !== "function" ||
    typeof item.body.setAsync !== "function"
  ) {
    setStatus(
      "Build failed — Outlook body access is unavailable."
    );
    return;
  }

  if (
    typeof findEmbeddedDataImage ===
    "function"
  ) {
    const problem =
      findEmbeddedDataImage();

    if (problem) {
      setStatus(
        "Build stopped — " +
        problem
      );
      return;
    }
  }

  updateBuildProgress(
    1,
    "Generating newsletter…"
  );

  let html;

  try {
    html =
      buildNewsletterHtml();
  } catch (error) {
    console.error(
      "Newsletter build error:",
      error
    );
    setStatus(
      "Build failed — " +
      error.message
    );
    return;
  }

  const validationProblem =
    typeof validateNewsletterHtml ===
    "function"
      ? validateNewsletterHtml(html)
      : "";

  if (validationProblem) {
    setStatus(
      "Build stopped — " +
      validationProblem
    );
    return;
  }

  const outlookHtml =
    normalizeNewsletterHtmlForOutlook(
      html
    );

  console.log(
    "Newsletter HTML length before Outlook normalization:",
    html.length
  );

  console.log(
    "Newsletter HTML length after Outlook normalization:",
    outlookHtml.length
  );

  updateBuildProgress(
    2,
    "Checking image sources…"
  );

  const modules =
    typeof collectImageBlocks ===
    "function"
      ? collectImageBlocks()
      : [];

  let modularCount = 0;

  modules.forEach(function (block) {
    modularCount +=
      block.items.length;
  });

  let heroCount = 0;

  if (value("hero1Url")) heroCount++;
  if (value("hero2Url")) heroCount++;

  const totalImages =
    modularCount +
    heroCount;

  updateBuildProgress(
    2,
    totalImages +
    " image" +
    (totalImages === 1 ? "" : "s") +
    " ready…"
  );

  updateBuildProgress(
    3,
    "Preparing Outlook message…"
  );

  let finished = false;

  const timeout =
    setTimeout(function () {
      if (finished) return;
      finished = true;
      setStatus(
        "Build timed out — Outlook did not respond after 60 seconds."
      );
    }, BUILD_TIMEOUT_MS);

  item.body.getAsync(
    Office.CoercionType.Html,
    function (bodyResult) {

      if (finished) return;

      if (
        !bodyResult ||
        bodyResult.status !==
        Office.AsyncResultStatus.Succeeded
      ) {
        finished = true;
        clearTimeout(timeout);
        setStatus(
          "Build failed reading Outlook body: " +
          (
            bodyResult &&
            bodyResult.error &&
            bodyResult.error.message
              ? bodyResult.error.message
              : "Unknown Outlook error."
          )
        );
        return;
      }

      updateBuildProgress(
        3,
        "Setting subject…"
      );

      const subject =
        buildSubject();

      function writeBody() {

        if (finished) return;

        updateBuildProgress(
          3,
          "Writing newsletter body…"
        );

        const options = {
          coercionType:
            Office.CoercionType.Html
        };

        if (
          Office.MailboxEnums &&
          Office.MailboxEnums.BodyMode &&
          Office.MailboxEnums.BodyMode.HostConfig
        ) {
          options.bodyMode =
            Office.MailboxEnums.BodyMode.HostConfig;
        }

        item.body.setAsync(
          outlookHtml,
          options,
          function (result) {

            if (finished) return;

            finished = true;
            clearTimeout(timeout);

            if (
              result &&
              result.status ===
              Office.AsyncResultStatus.Succeeded
            ) {
              updateBuildProgress(
                4,
                "✓ Complete — Dispatch built in Outlook."
              );
              console.log(
                "Newsletter body inserted successfully."
              );
              return;
            }

            const message =
              result &&
              result.error &&
              result.error.message
                ? result.error.message
                : "Outlook could not write the HTML body.";

            setStatus(
              "Build failed writing newsletter: " +
              message
            );

            console.error(
              "body.setAsync result:",
              result
            );
          }
        );
      }

      if (
        !item.subject ||
        typeof item.subject.setAsync !==
        "function"
      ) {
        writeBody();
        return;
      }

      item.subject.setAsync(
        subject,
        function (subjectResult) {

          if (finished) return;

          if (
            !subjectResult ||
            subjectResult.status !==
            Office.AsyncResultStatus.Succeeded
          ) {
            finished = true;
            clearTimeout(timeout);
            setStatus(
              "Build failed setting subject: " +
              (
                subjectResult &&
                subjectResult.error &&
                subjectResult.error.message
                  ? subjectResult.error.message
                  : "Unknown Outlook error."
              )
            );
            return;
          }

          writeBody();
        }
      );
    }
  );
}

setStatus(
  "The Pleasure Dispatch production patch is loaded."
);
