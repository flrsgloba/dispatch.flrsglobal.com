/* =========================================================
   THE PLEASURE DISPATCH
   editorial-layout-patch.js

   Keeps the proven taskpane builder intact and normalizes the
   finished newsletter HTML to the editorial structure used by
   the Pleasure Dispatch template.

   Current rules:
   - Hero 01 gets a deliberate pause before 01 — A REFLECTION.
   - Hero 02 belongs to 03 — STUDIO NOTES.
   - Hero 02 appears before the Studio Notes copy.
   - Studio Notes copy follows the image.
   - Pleasure Notes retain the Outlook table treatment and gain
     the same subtle editorial indent used in the edition pages.

   This is intentionally a wrapper around buildNewsletterHtml()
   rather than a rewrite of taskpane.js.
========================================================= */

(function () {
  "use strict";

  const originalBuildNewsletterHtml =
    window.buildNewsletterHtml;

  if (typeof originalBuildNewsletterHtml !== "function") {
    console.error(
      "[Pleasure Dispatch] Editorial layout patch could not find buildNewsletterHtml()."
    );
    return;
  }

  function findExactText(root, text) {
    const wanted = String(text || "").trim();

    if (!wanted) return null;

    const elements = root.querySelectorAll(
      "div,p,td,h1,h2,h3,span"
    );

    for (let i = 0; i < elements.length; i++) {
      if (
        String(elements[i].textContent || "").trim() ===
        wanted
      ) {
        return elements[i];
      }
    }

    return null;
  }

  function findDirectChildAncestor(element, parent) {
    let current = element;

    while (
      current &&
      current.parentElement &&
      current.parentElement !== parent
    ) {
      current = current.parentElement;
    }

    return current && current.parentElement === parent
      ? current
      : null;
  }

  function findImageBlockBetween(root, startElement, endElement) {
    const images = Array.from(
      root.querySelectorAll("img")
    ).filter(function (image) {
      const src = String(
        image.getAttribute("src") || ""
      ).trim();

      return (
        src &&
        !/pleasure-loop\.png(?:\?|$)/i.test(src)
      );
    });

    for (let i = 0; i < images.length; i++) {
      const image = images[i];

      if (
        startElement &&
        !startElement.contains(image) &&
        !isAfter(image, startElement)
      ) {
        continue;
      }

      if (
        endElement &&
        !endElement.contains(image) &&
        !isBefore(image, endElement)
      ) {
        continue;
      }

      return image;
    }

    return null;
  }

  function isAfter(node, reference) {
    if (!node || !reference) return false;

    const position = reference.compareDocumentPosition(node);

    return !!(
      position &
      Node.DOCUMENT_POSITION_FOLLOWING
    );
  }

  function isBefore(node, reference) {
    if (!node || !reference) return false;

    const position = reference.compareDocumentPosition(node);

    return !!(
      position &
      Node.DOCUMENT_POSITION_PRECEDING
    );
  }

  function applyLayout(html) {
    if (
      !html ||
      typeof DOMParser === "undefined"
    ) {
      return html;
    }

    const doc = new DOMParser().parseFromString(
      String(html),
      "text/html"
    );

    const reflectionHeading = findExactText(
      doc,
      "01 — A REFLECTION"
    );

    const studioHeading = findExactText(
      doc,
      "03 — STUDIO NOTES"
    );

    const pleasureHeading = findExactText(
      doc,
      "04 — PLEASURE NOTES"
    );

    if (!reflectionHeading || !studioHeading) {
      return html;
    }

    const contentCell =
      studioHeading.closest("td") ||
      reflectionHeading.closest("td");

    if (!contentCell) {
      return html;
    }

    /* -------------------------------------------------------
       HERO 01 → REFLECTION
    ------------------------------------------------------- */

    const hero1Image = findImageBlockBetween(
      contentCell,
      null,
      reflectionHeading
    );

    if (hero1Image) {
      const hero1Block = findDirectChildAncestor(
        hero1Image,
        contentCell
      );

      if (hero1Block) {
        hero1Block.style.marginBottom = "0";
        reflectionHeading.style.marginTop = "42px";
      }
    }

    /* -------------------------------------------------------
       HERO 02 → STUDIO NOTES

       The original builder places Hero 02 after the Studio
       Notes paragraph. Move its direct content block so the
       editorial sequence becomes:

       03 — STUDIO NOTES
       Hero 02
       Studio Notes copy
    ------------------------------------------------------- */

    const hero2Image = findImageBlockBetween(
      contentCell,
      studioHeading,
      pleasureHeading
    );

    const studioParagraph = Array.from(
      contentCell.querySelectorAll("p")
    ).find(function (paragraph) {
      return (
        isAfter(paragraph, studioHeading) &&
        (!pleasureHeading ||
          isBefore(paragraph, pleasureHeading))
      );
    });

    if (hero2Image && studioParagraph) {
      const hero2Block = findDirectChildAncestor(
        hero2Image,
        contentCell
      );

      if (hero2Block) {
        hero2Block.style.margin = "20px 0";

        contentCell.insertBefore(
          hero2Block,
          studioParagraph
        );
      }
    }

    /* -------------------------------------------------------
       PLEASURE NOTES

       Preserve the Outlook table structure but give it the
       same subtle editorial indentation used on the site.
    ------------------------------------------------------- */

    if (pleasureHeading) {
      const tables = Array.from(
        contentCell.querySelectorAll(
          'table[role="presentation"]'
        )
      );

      tables.forEach(function (table) {
        if (
          isAfter(table, pleasureHeading) &&
          (!findExactText(contentCell, "05 — AN INVITATION") ||
            isBefore(
              table,
              findExactText(
                contentCell,
                "05 — AN INVITATION"
              )
            ))
        ) {
          table.style.width =
            "calc(100% - 40px)";
          table.style.marginLeft = "40px";
          table.style.marginRight = "0";
        }
      });
    }

    return doc.body.innerHTML.trim();
  }

  window.buildNewsletterHtml = function () {
    const html = originalBuildNewsletterHtml.apply(
      this,
      arguments
    );

    try {
      return applyLayout(html);
    } catch (error) {
      console.error(
        "[Pleasure Dispatch] Editorial layout normalization failed; preserving original HTML:",
        error
      );
      return html;
    }
  };

  console.log(
    "[Pleasure Dispatch] Editorial layout patch loaded."
  );
})();
