/* =========================================================
   THE PLEASURE DISPATCH
   contacts.js
   Google Drive CSV contact picker + Outlook BCC
========================================================= */
(function () {
  "use strict";

  const CONTACTS_API_URL =
    "https://script.google.com/macros/s/AKfycbzavxknADmXnvAhRqcf9areGCRpfAJIZ62v84kqb_hpfgfAWIUbngcCH4B8M9TpkuA-uw/exec";

  const UNSUBSCRIBE_URL = "https://dispatch.flrsglobal.com/unsubscribe/";
  const state = { files: [], contacts: [], selected: new Set(), loadedFileId: "", loadedFileName: "" };

  function esc(value) { return String(value || "").replace(/[&<>\"']/g, function (c) { return { "&":"&amp;", "<":"&lt;", ">":"&gt;", '\"':"&quot;", "'":"&#039;" }[c]; }); }
  function setStatus(message) { if (typeof window.setStatus === "function") window.setStatus(message); else console.log("[Pleasure Dispatch Contacts]", message); }

  function normalizeContact(contact, index) {
    const email = String(contact.email || "").trim().toLowerCase();
    const name = String(contact.name || "").trim();
    const status = String(contact.status || "active").trim().toLowerCase();
    return { id:String(contact.id || email || "row-" + index), name:name || email, email:email, status:status, active:contact.active !== false && status !== "unsubscribed" && status !== "inactive" && status !== "opted_out" && status !== "opted-out" };
  }

  async function apiGet(params) {
    const url = CONTACTS_API_URL + "?" + new URLSearchParams(params).toString();
    const response = await fetch(url, { method:"GET", cache:"no-store" });
    if (!response.ok) throw new Error("Contacts service returned HTTP " + response.status + ".");
    const result = await response.json();
    if (result && result.status === "error") throw new Error(result.message || "Contacts service error.");
    return result;
  }

  function injectUi() {
    if (document.getElementById("pdContactsSection")) return;
    const section = document.createElement("section");
    section.id = "pdContactsSection"; section.className = "editor-section pd-contacts";
    section.innerHTML = '<div class="section-head"><span>RECIPIENTS</span><strong>Contact List</strong></div><div class="pd-contact-toolbar"><select id="pdContactFile" aria-label="Select contact CSV"><option value="">Select a CSV from Google Drive…</option></select><button id="pdRefreshContacts" class="secondary" type="button">Refresh</button></div><div class="pd-contact-meta" id="pdContactMeta">Choose the CSV you want to use for this Dispatch.</div><div class="pd-contact-actions"><label class="pd-check-all"><input id="pdSelectAll" type="checkbox"> Select all active</label><input id="pdContactSearch" class="pd-contact-search" type="search" placeholder="Search contacts…" autocomplete="off"></div><div id="pdContactList" class="pd-contact-list"><div class="pd-contact-empty">No contact list selected.</div></div><div class="pd-contact-footer"><strong id="pdSelectedCount">0 selected</strong><span>Recipients will be added to Outlook BCC.</span></div>';
    const actions = document.querySelector(".actions"), sections = document.querySelectorAll(".editor-section"), target = sections.length ? sections[0] : actions;
    if (target && target.parentNode) target.parentNode.insertBefore(section, target);
    injectStyles();
    document.getElementById("pdContactFile").addEventListener("change", function () { if (this.value) loadContacts(this.value); });
    document.getElementById("pdRefreshContacts").addEventListener("click", loadContactFiles);
    document.getElementById("pdSelectAll").addEventListener("change", function () { filteredContacts().forEach(function (contact) { if (contact.active && contact.email) this.checked ? state.selected.add(contact.email) : state.selected.delete(contact.email); }, this); renderContacts(); });
    document.getElementById("pdContactSearch").addEventListener("input", renderContacts);
  }

  function injectStyles() {
    if (document.getElementById("pdContactsStyles")) return;
    const style = document.createElement("style"); style.id = "pdContactsStyles";
    style.textContent = `.pd-contact-toolbar{display:grid;grid-template-columns:1fr auto;gap:10px;margin-bottom:10px}.pd-contact-toolbar select,.pd-contact-search{width:100%;min-height:40px;background:#595959;color:#F2EEE5;border:1px solid #777;padding:8px 10px;box-sizing:border-box;outline:none}.pd-contact-meta{font:10px/1.5 Arial,Helvetica,sans-serif;letter-spacing:.08em;text-transform:uppercase;color:#C9C3B8;margin:8px 0 14px}.pd-contact-actions{display:flex;align-items:center;gap:12px;margin-bottom:10px}.pd-check-all{font:11px Arial,Helvetica,sans-serif;color:#C9C3B8;white-space:nowrap}.pd-contact-search{max-width:240px;margin-left:auto}.pd-contact-list{border-top:1px solid #777;border-bottom:1px solid #777;max-height:260px;overflow:auto}.pd-contact-row{display:grid;grid-template-columns:24px 1fr auto;gap:10px;align-items:center;padding:9px 0;border-bottom:1px solid rgba(119,119,119,.45)}.pd-contact-row:last-child{border-bottom:0}.pd-contact-row input{margin:0}.pd-contact-name{font:14px/1.2 Arial,Helvetica,sans-serif;color:#F2EEE5;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.pd-contact-email{font:11px/1.3 Arial,Helvetica,sans-serif;color:#C9C3B8;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.pd-contact-status{font:9px Arial,Helvetica,sans-serif;letter-spacing:.08em;text-transform:uppercase;color:#777}.pd-contact-row.inactive{opacity:.45}.pd-contact-empty{padding:18px 0;color:#C9C3B8;font:13px Arial,Helvetica,sans-serif}.pd-contact-footer{display:flex;justify-content:space-between;gap:12px;margin-top:12px;font:10px Arial,Helvetica,sans-serif;color:#C9C3B8;letter-spacing:.04em}.pd-contact-footer strong{color:#F2EEE5}@media(max-width:600px){.pd-contact-toolbar{grid-template-columns:1fr}.pd-contact-actions{align-items:stretch;flex-direction:column}.pd-contact-search{max-width:none;margin-left:0}.pd-contact-footer{flex-direction:column}}`;
    document.head.appendChild(style);
  }

  async function loadContactFiles() {
    const select = document.getElementById("pdContactFile"); if (!select) return;
    select.disabled = true; setStatus("Loading contact lists from Google Drive…");
    try {
      const result = await apiGet({ action:"contactFiles" }); state.files = Array.isArray(result.files) ? result.files : [];
      select.innerHTML = '<option value="">Select a CSV from Google Drive…</option>';
      state.files.forEach(function (file) { const option = document.createElement("option"); option.value = file.fileId; option.textContent = file.fileName + (file.updatedAt ? " · " + new Date(file.updatedAt).toLocaleDateString() : ""); select.appendChild(option); });
      if (state.loadedFileId) select.value = state.loadedFileId;
      setStatus(state.files.length ? "Contact lists ready." : "No CSV files found in the Dispatch Drive folder.");
    } catch (error) { setStatus("Could not load contact lists: " + error.message); } finally { select.disabled = false; }
  }

  async function loadContacts(fileId) {
    state.selected.clear(); state.contacts = []; state.loadedFileId = fileId;
    const file = state.files.find(function (item) { return item.fileId === fileId; }); state.loadedFileName = file ? file.fileName : "Contact CSV"; renderContacts(); setStatus("Loading contacts from " + state.loadedFileName + "…");
    try {
      const result = await apiGet({ action:"contacts", fileId:fileId }); state.contacts = (Array.isArray(result.contacts) ? result.contacts : []).map(normalizeContact).filter(function (contact) { return !!contact.email; });
      const active = state.contacts.filter(function (contact) { return contact.active; }).length;
      document.getElementById("pdContactMeta").textContent = state.loadedFileName + " · " + active + " active · " + (state.contacts.length-active) + " excluded"; renderContacts(); setStatus("Contacts loaded: " + active + " active.");
    } catch (error) { document.getElementById("pdContactMeta").textContent = "Could not read this CSV."; setStatus("Could not load contacts: " + error.message); }
  }

  function filteredContacts() { const search = String((document.getElementById("pdContactSearch") || {}).value || "").trim().toLowerCase(); if (!search) return state.contacts; return state.contacts.filter(function (contact) { return (contact.name + " " + contact.email).toLowerCase().indexOf(search) !== -1; }); }
  function updateContactCounts() { const count = document.getElementById("pdSelectedCount"); if (count) count.textContent = state.selected.size + " selected"; }
  function selectedEmails() { return Array.from(state.selected).filter(function (email) { return state.contacts.some(function (contact) { return contact.email === email && contact.active; }); }); }

  function renderContacts() {
    const list = document.getElementById("pdContactList"), count = document.getElementById("pdSelectedCount"), selectAll = document.getElementById("pdSelectAll"); if (!list || !count) return;
    const contacts = filteredContacts(); list.innerHTML = "";
    if (!contacts.length) list.innerHTML = '<div class="pd-contact-empty">No contacts match the current search.</div>';
    else contacts.forEach(function (contact) {
      const row = document.createElement("label"); row.className = "pd-contact-row" + (contact.active ? "" : " inactive");
      const checkbox = document.createElement("input"); checkbox.type = "checkbox"; checkbox.disabled = !contact.active; checkbox.checked = state.selected.has(contact.email); checkbox.addEventListener("change", function () { if (checkbox.checked) state.selected.add(contact.email); else state.selected.delete(contact.email); updateContactCounts(); });
      const info = document.createElement("span"); info.innerHTML = '<span class="pd-contact-name">' + esc(contact.name) + '</span><br><span class="pd-contact-email">' + esc(contact.email) + '</span>';
      const status = document.createElement("span"); status.className = "pd-contact-status"; status.textContent = contact.active ? "active" : "excluded";
      row.appendChild(checkbox); row.appendChild(info); row.appendChild(status); list.appendChild(row);
    });
    updateContactCounts();
    if (selectAll) { const activeVisible = contacts.filter(function (c) { return c.active && c.email; }); selectAll.checked = activeVisible.length > 0 && activeVisible.every(function (c) { return state.selected.has(c.email); }); selectAll.indeterminate = activeVisible.some(function (c) { return state.selected.has(c.email); }) && !selectAll.checked; }
  }

  function addUnsubscribeFooter(html) {
    return String(html || "") + '<div style="border-top:1px solid #777;margin-top:36px;padding:18px 42px 28px;text-align:center;font:11px/1.5 Arial,Helvetica,sans-serif;color:#C9C3B8;">You are receiving The Pleasure Dispatch from FLRSGLOBAL. <a href="' + UNSUBSCRIBE_URL + '" style="color:#F2EEE5;text-decoration:underline;">Unsubscribe</a></div>';
  }

  function setRecipientsAndBuild() {
    const emails = selectedEmails(); if (!emails.length) { setStatus("Choose at least one active recipient before building the Dispatch."); return; }
    const item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
    if (!item || !item.bcc) { setStatus("Recipient access is unavailable in this Outlook compose window."); return; }
    setStatus("Adding " + emails.length + " recipients to BCC…");
    item.bcc.setAsync(emails, function (result) {
      if (!result || result.status !== Office.AsyncResultStatus.Succeeded) { setStatus("Recipient setup failed: " + ((result && result.error && result.error.message) || "Outlook could not set the BCC recipients.")); return; }
      setStatus("✓ " + emails.length + " recipients added to BCC. Building Dispatch…");
      if (typeof window.__pdOriginalBuildInOutlook === "function") window.__pdOriginalBuildInOutlook();
    });
  }

  function wrapBuild() {
    if (window.__pdContactsWrapped || typeof window.buildInOutlook !== "function") return;
    window.__pdOriginalBuildInOutlook = window.buildInOutlook;
    window.buildInOutlook = function () { if (selectedEmails().length) { setRecipientsAndBuild(); return; } window.__pdOriginalBuildInOutlook.apply(this, arguments); };
    window.__pdContactsWrapped = true;
  }

  function addUnsubscribeToNewsletterBuilder() {
    if (window.__pdUnsubscribeWrapped || typeof window.buildNewsletterHtml !== "function") return;
    const original = window.buildNewsletterHtml;
    window.buildNewsletterHtml = function () { return addUnsubscribeFooter(original.apply(this, arguments)); };
    window.__pdUnsubscribeWrapped = true;
  }

  function init() { injectUi(); loadContactFiles(); wrapBuild(); addUnsubscribeToNewsletterBuilder(); setTimeout(wrapBuild,250); setTimeout(wrapBuild,1000); setTimeout(addUnsubscribeToNewsletterBuilder,250); setTimeout(addUnsubscribeToNewsletterBuilder,1000); }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
  window.PleasureDispatchContacts = { refresh:loadContactFiles, getSelectedEmails:selectedEmails, load:loadContacts };
})();
