/* =========================================================
   THE PLEASURE DISPATCH
   drive-bridge.js

   Compatibility bridge for the Google Drive Apps Script endpoint.

   The Drive project was redeployed while the Outlook taskpane was
   intentionally kept on the last known-good split-file build. Rather
   than touching the large taskpane.js production file, this bridge
   rewrites only known legacy Drive API URLs to the current deployment.

   Publisher traffic is NOT changed here. publish.js owns publishing.
========================================================= */

(function () {
  "use strict";

  const CURRENT_DRIVE_API_URL =
    "https://script.google.com/macros/s/AKfycbxY93Vr1Zuij1sIKM7X0sgmyT5ipFnufnYGUrw6DqSAQL8QQYM6juVkRszGf-QdRKMEWQ/exec";

  const LEGACY_DRIVE_API_URLS = [
    "https://script.google.com/macros/s/AKfycbzavxknADmXnvAhRqcf9areGCRpfAJIZ62v84kqb_hpfgfAWIUbngcCH4B8M9TpkuA-uw/exec",
    "https://script.google.com/macros/s/AKfycbyTLvYbe1O_BbzsH09UMSZdbY9_XZXga-TbSPkR3UclT3Qhlaj7gy5yhXPA_UpE6Fu7tw/exec"
  ];

  const originalFetch = window.fetch.bind(window);

  function rewriteUrl(input) {
    if (typeof input === "string") {
      return LEGACY_DRIVE_API_URLS.indexOf(input) !== -1
        ? CURRENT_DRIVE_API_URL
        : input;
    }

    if (input instanceof URL) {
      return LEGACY_DRIVE_API_URLS.indexOf(input.href) !== -1
        ? CURRENT_DRIVE_API_URL
        : input.href;
    }

    if (input && typeof input.url === "string") {
      return LEGACY_DRIVE_API_URLS.indexOf(input.url) !== -1
        ? CURRENT_DRIVE_API_URL
        : input.url;
    }

    return input;
  }

  window.fetch = function (input, init) {
    const rewritten = rewriteUrl(input);

    if (rewritten !== input) {
      console.log(
        "[Pleasure Dispatch] Redirecting legacy Drive API to current deployment."
      );

      if (input && typeof input.url === "string" && !(input instanceof URL)) {
        return originalFetch(new Request(rewritten, input), init);
      }

      return originalFetch(rewritten, init);
    }

    return originalFetch(input, init);
  };

  window.PLEASURE_DISPATCH_DRIVE_API_URL = CURRENT_DRIVE_API_URL;
})();
