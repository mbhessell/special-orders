/* Special Orders: MS ABC special purchase orders, saved to the Sales SharePoint lists after Microsoft 365 sign-in. */
(function () {
  "use strict";
  const CFG = {
    clientId: "a2b31038-4fac-4e8b-8516-5b4c931f5a16",
    tenantId: "0c07cc35-4442-46a2-a2b1-a250d653b92f",
    site: "febdistributing365.sharepoint.com:/sites/PristineSpecialOrders",
    lists: { accounts: "baed8fad-89d9-4462-93df-caf317c1d463", products: "5f5b99a8-46f7-46dc-936f-77a9f3dfae54", orders: "d1705b46-6f32-4e97-b834-0babea675f4d" },
    scopes: ["Sites.ReadWrite.All"],
    tax: 0.07
  };
  const $ = s => document.querySelector(s);
  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const money = n => "$" + (Math.round(n * 100) / 100).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const r2 = n => Math.round(n * 100) / 100;
  const key = s => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");

  /* ---------- sign-in + Graph ---------- */
  let msalApp = null, account = null, siteId = null;
  async function token() {
    const req = { scopes: CFG.scopes, account };
    try { return (await msalApp.acquireTokenSilent(req)).accessToken; }
    catch (e) { await msalApp.acquireTokenRedirect(req); throw new Error("redirecting"); }
  }
  async function graph(path, opts) {
    const t = await token();
    opts = opts || {};
    let url = path.startsWith("https://") ? path : "https://graph.microsoft.com/v1.0" + path, out = [];
    const headers = { Authorization: "Bearer " + t, Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" };
    if (opts.body) headers["Content-Type"] = "application/json";
    while (url) {
      const r = await fetch(url, { method: opts.method || "GET", headers, body: opts.body ? JSON.stringify(opts.body) : undefined });
      if (r.status === 401 || r.status === 403) { const e = new Error("noaccess"); e.status = r.status; throw e; }
      if (!r.ok) { let m = ""; try { m = (await r.json()).error.message; } catch (x) {} throw new Error("Graph " + r.status + (m ? ": " + m : "")); }
      if (r.status === 204) return null;
      const j = await r.json();
      if (!j.value || opts.method) return j;
      out = out.concat(j.value); url = j["@odata.nextLink"] || null;
    }
    return out;
  }
  const L = id => "/sites/" + siteId + "/lists/" + id + "/items";

  /* ---------- data ---------- */
  let ACCTS = [], PRODS = [];
  async function loadData() {
    status("Loading accounts and products…");
    siteId = (await graph("/sites/" + CFG.site + "?$select=id")).id;
    const [a, p] = await Promise.all([
      graph(L(CFG.lists.accounts) + "?$top=999&$filter=fields/Status eq 'Active'&$expand=fields($select=Title,PermitHolder,PermitNumber,LocationName,Address,City,County,Route,Phone)"),
      graph(L(CFG.lists.products) + "?$top=999&$expand=fields($select=Title,ProductCode,Size,CasePrice,Supplier,Active)")
    ]);
    ACCTS = a.map(i => { const f = i.fields; return { id: i.id, dba: f.Title || "", holder: f.PermitHolder || "", permit: f.PermitNumber || "", loc: f.LocationName || "", addr: f.Address || "", city: f.City || "", county: f.County || "", route: f.Route || "", phone: f.Phone || "",
      k: [key(f.Title), key(f.LocationName), key(f.PermitHolder)], kp: key(f.PermitNumber), kc: key(f.City) }; });
    PRODS = p.map(i => i.fields).filter(f => f.Active !== false).map(f => ({ code: String(f.ProductCode || ""), name: f.Title || "", size: f.Size || "", price: Number(f.CasePrice) || 0, supplier: f.Supplier || "" }))
      .sort((x, y) => x.name.localeCompare(y.name));
  }

  /* ---------- typeahead ---------- */
  function typeahead(input, dd, search, render, pick) {
    let items = [], hl = -1;
    const show = () => {
      items = search(input.value);
      hl = items.length ? 0 : -1;
      dd.innerHTML = items.length ? items.map((it, i) => '<div data-i="' + i + '"' + (i === hl ? ' class="hl"' : "") + ">" + render(it) + "</div>").join("") : (input.value.trim() ? '<div class="muted">No matches</div>' : "");
      dd.hidden = !dd.innerHTML;
    };
    const choose = i => { if (items[i]) { dd.hidden = true; pick(items[i]); } };
    input.addEventListener("input", show);
    input.addEventListener("focus", () => { if (input.value.trim() || input.dataset.showAll) show(); });
    input.addEventListener("keydown", e => {
      if (dd.hidden) return;
      if (e.key === "ArrowDown" || e.key === "ArrowUp") { e.preventDefault(); hl = Math.max(0, Math.min(items.length - 1, hl + (e.key === "ArrowDown" ? 1 : -1))); [...dd.children].forEach((c, i) => c.classList.toggle("hl", i === hl)); dd.children[hl] && dd.children[hl].scrollIntoView({ block: "nearest" }); }
      else if (e.key === "Enter") { e.preventDefault(); choose(hl); }
      else if (e.key === "Escape") dd.hidden = true;
    });
    dd.addEventListener("mousedown", e => { const d = e.target.closest("[data-i]"); if (d) { e.preventDefault(); choose(+d.dataset.i); } });
    input.addEventListener("blur", () => setTimeout(() => dd.hidden = true, 120));
  }

  function searchAccounts(q) {
    const k = key(q); if (!k) return [];
    const starts = [], contains = [];
    for (const a of ACCTS) {
      if (a.kp.startsWith(k) || a.kp.replace(/^0+/, "").startsWith(k.replace(/^0+/, "")) && /^\d+$/.test(k) || a.k.some(x => x.startsWith(k))) starts.push(a);
      else if (k.length >= 3 && (a.k.some(x => x.includes(k)) || a.kc.startsWith(k))) contains.push(a);
      if (starts.length >= 25) break;
    }
    return starts.concat(contains).slice(0, 25);
  }
  function searchProducts(q) {
    const k = key(q);
    if (!k) return PRODS.slice();
    return PRODS.filter(p => p.code.startsWith(k) || key(p.name).includes(k) || key(p.supplier).startsWith(k));
  }

  /* ---------- order state ---------- */
  let ACCT = null, LINES = [];
  function pickAccount(a) {
    ACCT = a;
    $("#acctQ").value = a.dba || a.holder;
    $("#acct").className = "acct";
    const cell = (k, v) => '<div><div class="k">' + k + '</div><div class="v">' + esc(v || "–") + "</div></div>";
    $("#acct").innerHTML = cell("DBA name", a.dba) + cell("Permit holder", a.holder) + cell("Permit no.", a.permit) + cell("Address", a.addr) + cell("County", a.county) + cell("Route", a.route);
    $("#acctExtra").hidden = false;
    $("#phone").value = a.phone || "";
    validate();
  }
  function lineRow(i) {
    const ln = LINES[i];
    return '<tr data-i="' + i + '"><td class="p"><div class="search"><input class="pq" type="search" autocomplete="off" placeholder="Code or brand" value="' + esc(ln.p ? ln.p.code + " – " + ln.p.name : "") + '"><div class="dd" hidden></div></div></td>' +
      '<td class="q num"><input class="qty" type="number" min="1" step="1" inputmode="numeric" value="' + (ln.cases || "") + '" aria-label="Cases"></td>' +
      '<td class="pr num">' + (ln.p ? money(ln.p.price) : "–") + '</td><td class="t num">' + (ln.p && ln.cases ? money(ln.p.price * ln.cases) : "–") + "</td>" +
      '<td class="xx"><button class="x" title="Remove">×</button></td></tr>';
  }
  function renderLines() {
    $("#lines").innerHTML = LINES.map((_, i) => lineRow(i)).join("");
    [...$("#lines").children].forEach(tr => {
      const i = +tr.dataset.i, inp = tr.querySelector(".pq");
      inp.dataset.showAll = "1";
      typeahead(inp, tr.querySelector(".dd"), searchProducts,
        p => esc(p.code) + " – " + esc(p.name) + "<small>" + esc(p.size) + " · " + money(p.price) + " per case" + (p.supplier ? " · " + esc(p.supplier) : "") + "</small>",
        p => { LINES[i].p = p; if (!LINES[i].cases) LINES[i].cases = 1; renderLines(); const q = $("#lines").children[i].querySelector(".qty"); q.focus(); q.select(); });
      tr.querySelector(".qty").addEventListener("input", e => { LINES[i].cases = Math.max(0, parseInt(e.target.value, 10) || 0); tr.querySelector(".t").textContent = LINES[i].p && LINES[i].cases ? money(LINES[i].p.price * LINES[i].cases) : "–"; totals(); });
      tr.querySelector(".x").addEventListener("click", () => { LINES.splice(i, 1); if (!LINES.length) LINES.push({}); renderLines(); });
    });
    totals();
  }
  function totals() {
    const sub = r2(LINES.reduce((s, l) => s + (l.p && l.cases ? r2(l.p.price * l.cases) : 0), 0)), tax = r2(sub * CFG.tax);
    $("#sub").textContent = money(sub); $("#tax").textContent = money(tax); $("#tot").textContent = money(sub + tax);
    validate();
  }
  function validate() { $("#submit").disabled = !(ACCT && LINES.some(l => l.p && l.cases > 0)); }
  function resetForm() {
    ACCT = null; LINES = [{}];
    $("#acctQ").value = ""; $("#acct").className = "acct empty"; $("#acct").textContent = "No account selected."; $("#acctExtra").hidden = true; $("#phone").value = ""; $("#auth").value = "";
    renderLines();
  }

  async function submit() {
    const btn = $("#submit"); btn.disabled = true; btn.textContent = "Saving…";
    const lines = LINES.filter(l => l.p && l.cases > 0);
    const now = new Date(), p2 = n => String(n).padStart(2, "0");
    const who = account.name || account.username || "";
    const initials = (who.match(/\b[A-Za-z]/g) || []).join("").toUpperCase().slice(0, 3) || "X";
    const orderId = "SO-" + String(now.getFullYear()).slice(2) + p2(now.getMonth() + 1) + p2(now.getDate()) + "-" + p2(now.getHours()) + p2(now.getMinutes()) + p2(now.getSeconds()) + "-" + initials;
    const orderDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 12).toISOString();
    const phone = $("#phone").value.trim(), auth = $("#auth").value.trim();
    try {
      for (const l of lines) {
        await graph(L(CFG.lists.orders), { method: "POST", body: { fields: {
          Title: orderId, OrderDate: orderDate, EnteredBy: who, PermitHolder: ACCT.holder, DBAName: ACCT.dba, PermitNumber: ACCT.permit,
          Address: ACCT.addr, City: ACCT.city, County: ACCT.county, Route: ACCT.route, Phone: phone, AuthCode: auth,
          ProductCode: l.p.code, BrandName: l.p.name, Size: l.p.size, Cases: l.cases, CasePrice: l.p.price, LineTotal: r2(l.p.price * l.cases), Sent: false } } });
      }
      if (phone && phone !== ACCT.phone) {
        try { await graph(L(CFG.lists.accounts) + "/" + ACCT.id + "/fields", { method: "PATCH", body: { Phone: phone } }); ACCT.phone = phone; } catch (e) { /* phone save is a nice-to-have */ }
      }
      const name = ACCT.dba || ACCT.holder;
      resetForm();
      status("Order " + esc(orderId) + " saved for " + esc(name) + " (" + lines.length + " product" + (lines.length > 1 ? "s" : "") + "). It goes out in Friday's email.", false, true);
      window.scrollTo({ top: 0, behavior: "smooth" });
    } catch (e) {
      if (e.message === "redirecting") return;
      status("The order didn't save (" + esc(e.message) + "). Nothing was lost on this screen. Try Submit again.", true);
    } finally { btn.textContent = "Submit order"; validate(); }
  }

  /* ---------- this week's (unsent) orders ---------- */
  async function loadWeek() {
    $("#week").innerHTML = '<p class="muted">Loading…</p>';
    try {
      const rows = (await graph(L(CFG.lists.orders) + "?$top=999&$filter=fields/Sent eq 0&$expand=fields($select=Title,OrderDate,EnteredBy,DBAName,PermitHolder,PermitNumber,ProductCode,BrandName,Cases,CasePrice,LineTotal)"))
        .map(i => Object.assign({ id: i.id }, i.fields));
      if (!rows.length) { $("#week").innerHTML = '<p class="muted">No orders waiting. Everything has been sent.</p>'; return; }
      const by = {};
      rows.forEach(r => (by[r.PermitNumber] = by[r.PermitNumber] || []).push(r));
      let grand = 0;
      $("#week").innerHTML = Object.values(by).sort((a, b) => String(a[0].DBAName).localeCompare(String(b[0].DBAName))).map(g => {
        const sub = r2(g.reduce((s, r) => s + (Number(r.LineTotal) || 0), 0)); grand += sub;
        return '<div class="grp"><b>' + esc(g[0].DBAName || g[0].PermitHolder) + '</b> <span class="muted">· Permit ' + esc(g[0].PermitNumber) + '</span><table>' +
          g.map(r => "<tr><td>" + esc(r.ProductCode) + " " + esc(r.BrandName) + '<br><small class="muted">' + esc(r.Title) + " · " + esc(r.EnteredBy) + '</small></td><td class="num">' + r.Cases + ' cs</td><td class="num">' + money(Number(r.LineTotal) || 0) +
            '</td><td class="num"><button class="x" data-del="' + r.id + '" title="Remove this line">×</button></td></tr>').join("") +
          '<tr><td colspan="2" class="muted">Sub-total + 7% tax</td><td class="num"><b>' + money(sub + r2(sub * CFG.tax)) + "</b></td><td></td></tr></table></div>";
      }).join("") + '<p class="muted" style="margin:10px 0 0">' + rows.length + " lines · " + Object.keys(by).length + " accounts · " + money(grand) + " before tax</p>";
      $("#week").querySelectorAll("[data-del]").forEach(b => b.onclick = async () => {
        if (!confirm("Remove this product line from the order?")) return;
        b.disabled = true;
        try { await graph(L(CFG.lists.orders) + "/" + b.dataset.del, { method: "DELETE" }); loadWeek(); }
        catch (e) { if (e.message !== "redirecting") alert("Couldn't remove it: " + e.message); b.disabled = false; }
      });
    } catch (e) {
      if (e.message === "redirecting") return;
      $("#week").innerHTML = '<p class="status err">Couldn\'t load orders (' + esc(e.message) + ").</p>";
    }
  }

  /* ---------- wiring ---------- */
  function status(msg, isErr, isOk) { $("#status").hidden = !msg; $("#status").className = "status" + (isErr ? " err" : isOk ? " ok" : ""); $("#status").innerHTML = msg || ""; }
  function tab(which) {
    $("#tNew").setAttribute("aria-pressed", which === "new"); $("#tWeek").setAttribute("aria-pressed", which === "week");
    $("#vNew").hidden = which !== "new"; $("#vWeek").hidden = which !== "week";
    if (which === "week") loadWeek();
  }
  function wire() {
    typeahead($("#acctQ"), $("#acctDD"), searchAccounts,
      a => "<b>" + esc(a.dba || a.holder) + "</b>" + (a.loc && key(a.loc) !== key(a.dba) ? " <span class=\"muted\">(" + esc(a.loc) + ")</span>" : "") + "<small>" + esc(a.holder) + " · Permit " + esc(a.permit) + " · " + esc(a.addr || a.city) + "</small>",
      a => { pickAccount(a); status(""); const first = $("#lines .pq"); if (first && !LINES[0].p) first.focus(); });
    $("#addLine").onclick = () => { LINES.push({}); renderLines(); const all = $("#lines").querySelectorAll(".pq"); all[all.length - 1].focus(); };
    $("#submit").onclick = submit;
    $("#tNew").onclick = () => tab("new"); $("#tWeek").onclick = () => tab("week");
    $("#reloadWeek").onclick = loadWeek;
    $("#signout").onclick = () => msalApp.logoutRedirect({ account });
    resetForm();
  }

  async function boot() {
    if (!window.msal) return status("The Microsoft sign-in library didn't load. Check your connection and reload.", true);
    const redirectUri = location.origin + location.pathname.replace(/index\.html$/, "");
    msalApp = new msal.PublicClientApplication({ auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri }, cache: { cacheLocation: "localStorage" } });
    const res = await msalApp.handleRedirectPromise().catch(e => { status("Sign-in didn't finish: " + esc(e.errorMessage || e.message), true); return null; });
    account = (res && res.account) || msalApp.getAllAccounts()[0] || null;
    if (!account) { $("#signin").hidden = false; $("#signbtn").onclick = () => msalApp.loginRedirect({ scopes: CFG.scopes }); return; }
    msalApp.setActiveAccount(account); $("#signin").hidden = true;
    try {
      await loadData(); wire();
      $("#who").textContent = account.name || account.username; $("#main").hidden = false; status("");
    } catch (e) {
      if (e.message === "redirecting") return;
      if (e.message === "noaccess") return status("You're signed in, but you don't have access to the Pristine Special Orders site. Ask Matthew Hessell to add you.", true);
      status("Couldn't load accounts and products (" + esc(e.message) + "). Reload the page in a minute.", true);
    }
  }
  boot();
})();
