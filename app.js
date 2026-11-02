/* Special Orders: MS ABC special purchase orders, saved to the Pristine Special Orders SharePoint lists after Microsoft 365 sign-in.
   Phones must be registered once with an 8-digit activation code (Phones tab, owners only) before signing in. */
(function () {
  "use strict";
  const CFG = {
    clientId: "a2b31038-4fac-4e8b-8516-5b4c931f5a16",
    tenantId: "0c07cc35-4442-46a2-a2b1-a250d653b92f",
    site: "febdistributing365.sharepoint.com:/sites/PristineSpecialOrders",
    lists: { accounts: "baed8fad-89d9-4462-93df-caf317c1d463", products: "5f5b99a8-46f7-46dc-936f-77a9f3dfae54", orders: "d1705b46-6f32-4e97-b834-0babea675f4d", devices: "7bf3a290-3936-4ee5-aee7-e0b3e64c5f35" },
    deviceUrl: "https://default0c07cc35444246a2a2b1a250d653b9.2f.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/11/workflows/fe8756f0a67144308fa3597e77ab55fd/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=1i9awuqKe9z7Y-iVOZ1ojP54_B6-bgcwkLth7JWtx2E",
    people: [["Matthew Hessell", "mbhessell@febdistributing.com"], ["Rod McCormick", "rmccormick@pristinebev.com"], ["Monica Britt", "mbritt@febdistributing.com"], ["Jeffrey Bertucci", "jbertucci@febdistributing.com"], ["Brian Drennan", "brian@capitalcitybev.com"], ["Jackson Drennan", "jack@capitalcitybev.com"], ["Ryland Bertucci", "RylandB@febdistributing.com"]],
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
      graph(L(CFG.lists.accounts) + "?$top=999&$filter=fields/Status eq 'Active'&$expand=fields($select=Title,PermitHolder,PermitNumber,RetailerId,LocationName,Address,City,County,Route,Phone)"),
      graph(L(CFG.lists.products) + "?$top=999&$expand=fields($select=Title,ProductCode,Size,CasePrice,Supplier,Active)")
    ]);
    ACCTS = a.map(i => { const f = i.fields; return { id: i.id, dba: f.Title || "", holder: f.PermitHolder || "", permit: f.PermitNumber || "", retailer: f.RetailerId || "", loc: f.LocationName || "", addr: f.Address || "", city: f.City || "", county: f.County || "", route: f.Route || "", phone: f.Phone || "",
      k: [key(f.Title), key(f.LocationName), key(f.PermitHolder)], kp: key(f.PermitNumber), kr: key(f.RetailerId), kc: key(f.City) }; });
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
      if (a.kp.startsWith(k) || (k.length >= 4 && a.kr.startsWith(k)) || a.kp.replace(/^0+/, "").startsWith(k.replace(/^0+/, "")) && /^\d+$/.test(k) || a.k.some(x => x.startsWith(k))) starts.push(a);
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
    $("#acct").innerHTML = cell("DBA name", a.dba) + cell("Permit holder", a.holder) + cell("Permit no.", a.permit) + cell("Retailer ID", a.retailer) + cell("Address", a.addr) + cell("County", a.county) + cell("Route", a.route);
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
          Title: orderId, OrderDate: orderDate, EnteredBy: who, PermitHolder: ACCT.holder, DBAName: ACCT.dba, PermitNumber: ACCT.permit, RetailerId: ACCT.retailer,
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
      const rows = (await graph(L(CFG.lists.orders) + "?$top=999&$filter=fields/Sent eq 0&$expand=fields($select=Title,OrderDate,EnteredBy,DBAName,PermitHolder,PermitNumber,RetailerId,ProductCode,BrandName,Cases,CasePrice,LineTotal)"))
        .map(i => Object.assign({ id: i.id }, i.fields));
      if (!rows.length) { $("#week").innerHTML = '<p class="muted">No orders waiting. Everything has been sent.</p>'; return; }
      const by = {};
      rows.forEach(r => (by[r.PermitNumber] = by[r.PermitNumber] || []).push(r));
      let grand = 0;
      $("#week").innerHTML = Object.values(by).sort((a, b) => String(a[0].DBAName).localeCompare(String(b[0].DBAName))).map(g => {
        const sub = r2(g.reduce((s, r) => s + (Number(r.LineTotal) || 0), 0)); grand += sub;
        return '<div class="grp"><b>' + esc(g[0].DBAName || g[0].PermitHolder) + '</b> <span class="muted">· Permit ' + esc(g[0].PermitNumber) + (g[0].RetailerId ? ' · Retailer ID ' + esc(g[0].RetailerId) : '') + '</span><table>' +
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

  /* ---------- phone registration (one-time 8-digit activation code) ---------- */
  const DEV_KEY = "so_device";
  const isPhone = () => /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1) ||
    (window.matchMedia && matchMedia("(display-mode: standalone)").matches) || navigator.standalone === true;
  const devGet = () => { try { return JSON.parse(localStorage.getItem(DEV_KEY) || "null"); } catch (e) { return null; } };
  const devSet = d => { try { d ? localStorage.setItem(DEV_KEY, JSON.stringify(d)) : localStorage.removeItem(DEV_KEY); } catch (e) {} };
  const randKey = n => { const a = new Uint8Array(n); crypto.getRandomValues(a); return btoa(String.fromCharCode.apply(null, a)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, ""); };
  const devName = () => /iPhone/.test(navigator.userAgent) ? "iPhone" : /iPad|Macintosh/.test(navigator.userAgent) ? "iPad" : /Android/.test(navigator.userAgent) ? "Android" : "Phone";
  async function devPost(body) {
    const r = await fetch(CFG.deviceUrl, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    let j = {}; try { j = await r.json(); } catch (e) {}
    return j;
  }
  function showOnly(id) { ["register", "mismatch", "signin", "main"].forEach(x => $("#" + x).hidden = x !== id); }
  function showRegister(msg) {
    showOnly("register"); status("");
    const inp = $("#regCode"), err = $("#regErr");
    err.hidden = !msg; err.textContent = msg || "";
    inp.oninput = () => { const d = inp.value.replace(/\D/g, "").slice(0, 8); inp.value = d.length > 4 ? d.slice(0, 4) + " " + d.slice(4) : d; };
    $("#regBtn").onclick = async () => {
      const code = inp.value.replace(/\D/g, "");
      if (code.length !== 8) { err.hidden = false; err.textContent = "Enter all 8 digits."; return; }
      const btn = $("#regBtn"); btn.disabled = true; btn.textContent = "Checking…"; err.hidden = true;
      const dev = { deviceId: "D-" + randKey(9), deviceKey: randKey(32) };
      try {
        const r = await devPost({ action: "register", code, deviceId: dev.deviceId, deviceKey: dev.deviceKey, deviceName: devName() });
        if (r.ok) { devSet(Object.assign(dev, { email: r.email || "", name: r.name || "" })); location.reload(); return; }
        err.hidden = false; err.textContent = "That code is wrong, already used, or expired. Ask Matthew Hessell for a new one.";
      } catch (e) { err.hidden = false; err.textContent = "Couldn't reach the server. Check your connection and try again."; }
      btn.disabled = false; btn.textContent = "Activate this phone";
    };
  }
  // true = OK to continue to Microsoft sign-in; false = a registration screen is showing
  async function phoneGate() {
    if (!isPhone()) return true;
    const dev = devGet();
    if (!dev) { showRegister(); return false; }
    try {
      const r = await devPost({ action: "check", deviceId: dev.deviceId, deviceKey: dev.deviceKey });
      if (r.ok) { if (r.email) { dev.email = r.email; dev.name = r.name || dev.name; devSet(dev); } return true; }
      if (r.error === "device") { devSet(null); showRegister("This phone is no longer registered. Ask Matthew Hessell for a new activation code."); return false; }
      throw new Error("check");
    } catch (e) {
      showOnly(""); status('Couldn\'t confirm this phone\'s registration. Check your connection. <button class="linkbtn" onclick="location.reload()">Try again</button>', true);
      return false;
    }
  }

  /* ---------- Phones tab (owners only: the Devices list is owner-only) ---------- */
  let canAdmin = false;
  async function checkAdmin() {
    try { await graph("/sites/" + siteId + "/lists/" + CFG.lists.devices + "?$select=id"); canAdmin = true; } catch (e) { canAdmin = false; }
    $("#tPhones").hidden = !canAdmin;
  }
  const fmtDT = v => v ? new Date(v).toLocaleString("en-US", { month: "numeric", day: "numeric", hour: "numeric", minute: "2-digit" }) : "–";
  async function loadPhones() {
    $("#phones").innerHTML = '<p class="muted">Loading…</p>';
    try {
      const rows = (await graph(L(CFG.lists.devices) + "?$top=999&$expand=fields($select=Title,Email,PersonName,DeviceStatus,DeviceName,CodeExpires,RegisteredAt,LastSeen,IssuedBy)"))
        .map(i => Object.assign({ id: i.id }, i.fields));
      const now = Date.now();
      rows.forEach(r => { if (r.DeviceStatus === "Pending" && r.CodeExpires && new Date(r.CodeExpires).getTime() < now) r.DeviceStatus = "Expired"; });
      const order = { Active: 0, Pending: 1, Expired: 2, Revoked: 3 };
      rows.sort((a, b) => (order[a.DeviceStatus] - order[b.DeviceStatus]) || String(a.PersonName).localeCompare(String(b.PersonName)));
      if (!rows.length) { $("#phones").innerHTML = '<p class="muted">No phones yet. Issue a code above.</p>'; return; }
      $("#phones").innerHTML = '<table class="tbl"><thead><tr><th>Person</th><th>Status</th><th>Phone</th><th>Registered</th><th>Last used</th><th></th></tr></thead><tbody>' +
        rows.map(r => "<tr><td><b>" + esc(r.PersonName || r.Title) + '</b><br><small class="muted">' + esc(r.Email) + '</small></td><td><span class="pill ' + esc(r.DeviceStatus) + '">' + esc(r.DeviceStatus) +
          (r.DeviceStatus === "Pending" ? "</span><br><small class=\"muted\">until " + fmtDT(r.CodeExpires) + "</small>" : "</span>") + "</td><td>" + esc(r.DeviceName || "–") + "</td><td>" + fmtDT(r.RegisteredAt) + "</td><td>" + fmtDT(r.LastSeen) +
          '</td><td class="num">' + (r.DeviceStatus === "Active" || r.DeviceStatus === "Pending" ? '<button class="linkbtn" data-revoke="' + r.id + '">' + (r.DeviceStatus === "Active" ? "Revoke" : "Cancel") + "</button>" : "") + "</td></tr>").join("") + "</tbody></table>";
      $("#phones").querySelectorAll("[data-revoke]").forEach(b => b.onclick = async () => {
        if (!confirm("Remove access for this phone? The person will need a new code to use it again.")) return;
        b.disabled = true;
        try { await graph(L(CFG.lists.devices) + "/" + b.dataset.revoke + "/fields", { method: "PATCH", body: { DeviceStatus: "Revoked", Code: "revoked" } }); loadPhones(); }
        catch (e) { if (e.message !== "redirecting") alert("Couldn't revoke it: " + e.message); b.disabled = false; }
      });
    } catch (e) { if (e.message !== "redirecting") $("#phones").innerHTML = '<p class="status err">Couldn\'t load phones (' + esc(e.message) + ").</p>"; }
  }
  async function issueCode() {
    const name = $("#pName").value.trim(), email = $("#pEmail").value.trim().toLowerCase();
    if (!name || !/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/i.test(email)) { alert("Enter the person's name and Microsoft 365 email."); return; }
    const btn = $("#issueBtn"); btn.disabled = true; btn.textContent = "Generating…";
    try {
      const a = new Uint32Array(1); crypto.getRandomValues(a); const code = String(a[0] % 100000000).padStart(8, "0");
      const expires = new Date(Date.now() + 72 * 3600 * 1000);
      await graph(L(CFG.lists.devices), { method: "POST", body: { fields: { Title: name, PersonName: name, Email: email, Code: code, CodeExpires: expires.toISOString(), DeviceStatus: "Pending", IssuedBy: account.name || account.username } } });
      const pretty = code.slice(0, 4) + " " + code.slice(4);
      const body = "Hi " + name.split(" ")[0] + ",\n\nHere's how to set up the Special Orders app on your phone:\n\n1. On your phone, open https://so.febdistributing.com\n2. Add it to your home screen (iPhone: Share > Add to Home Screen; Android: menu > Install app), then open it from there.\n3. Enter this activation code: " + pretty + "\n4. Sign in with your Microsoft 365 account (" + email + ").\n\nThe code works once and expires " + expires.toLocaleString("en-US", { weekday: "short", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }) + ".\n";
      $("#issued").hidden = false;
      $("#issued").innerHTML = "Code for <b>" + esc(name) + '</b>:<div class="bigcode">' + pretty + '</div><small>Expires ' + esc(expires.toLocaleString()) + '. Only works for ' + esc(email) + '.</small><br>' +
        '<a class="linkbtn" href="mailto:' + encodeURIComponent(email) + "?subject=" + encodeURIComponent("Your Special Orders activation code") + "&body=" + encodeURIComponent(body) + '">Email it to ' + esc(name.split(" ")[0]) + "</a> · " +
        '<button class="linkbtn" id="copyCode">Copy code</button>';
      $("#copyCode").onclick = () => { try { navigator.clipboard.writeText(code); $("#copyCode").textContent = "Copied"; } catch (e) {} };
      $("#pName").value = ""; $("#pEmail").value = "";
      loadPhones();
    } catch (e) { if (e.message !== "redirecting") alert("Couldn't create the code: " + e.message); }
    btn.disabled = false; btn.textContent = "Generate code";
  }

  /* ---------- wiring ---------- */
  function status(msg, isErr, isOk) { $("#status").hidden = !msg; $("#status").className = "status" + (isErr ? " err" : isOk ? " ok" : ""); $("#status").innerHTML = msg || ""; }
  function tab(which) {
    $("#tNew").setAttribute("aria-pressed", which === "new"); $("#tWeek").setAttribute("aria-pressed", which === "week"); $("#tPhones").setAttribute("aria-pressed", which === "phones");
    $("#vNew").hidden = which !== "new"; $("#vWeek").hidden = which !== "week"; $("#vPhones").hidden = which !== "phones";
    if (which === "week") loadWeek();
    if (which === "phones") loadPhones();
  }
  function wire() {
    typeahead($("#acctQ"), $("#acctDD"), searchAccounts,
      a => "<b>" + esc(a.dba || a.holder) + "</b>" + (a.loc && key(a.loc) !== key(a.dba) ? " <span class=\"muted\">(" + esc(a.loc) + ")</span>" : "") + "<small>" + esc(a.holder) + " · Permit " + esc(a.permit) + (a.retailer ? " · " + esc(a.retailer) : "") + " · " + esc(a.addr || a.city) + "</small>",
      a => { pickAccount(a); status(""); const first = $("#lines .pq"); if (first && !LINES[0].p) first.focus(); });
    $("#addLine").onclick = () => { LINES.push({}); renderLines(); const all = $("#lines").querySelectorAll(".pq"); all[all.length - 1].focus(); };
    $("#submit").onclick = submit;
    $("#tNew").onclick = () => tab("new"); $("#tWeek").onclick = () => tab("week"); $("#tPhones").onclick = () => tab("phones");
    $("#issueBtn").onclick = issueCode; $("#reloadPhones").onclick = loadPhones;
    $("#people").innerHTML = CFG.people.map(p => '<option value="' + esc(p[0]) + '">').join("");
    $("#pName").onchange = () => { const m = CFG.people.find(p => p[0].toLowerCase() === $("#pName").value.trim().toLowerCase()); if (m && !$("#pEmail").value) $("#pEmail").value = m[1]; };
    $("#reloadWeek").onclick = loadWeek;
    $("#signout").onclick = () => msalApp.logoutRedirect({ account });
    resetForm();
  }

  async function boot() {
    if ("serviceWorker" in navigator) { navigator.serviceWorker.register("sw.js").catch(() => {}); }
    if (!window.msal) return status("The Microsoft sign-in library didn't load. Check your connection and reload.", true);
    if (!(await phoneGate())) return;
    const redirectUri = location.origin + location.pathname.replace(/index\.html$/, "");
    msalApp = new msal.PublicClientApplication({ auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri }, cache: { cacheLocation: "localStorage" } });
    const res = await msalApp.handleRedirectPromise().catch(e => { status("Sign-in didn't finish: " + esc(e.errorMessage || e.message), true); return null; });
    account = (res && res.account) || msalApp.getAllAccounts()[0] || null;
    const dev = isPhone() ? devGet() : null;
    if (!account) {
      showOnly("signin"); if (dev && dev.name) $("#signin p").textContent = "Phone registered to " + dev.name + ". Sign in with your Microsoft 365 account" + (dev.email ? " (" + dev.email + ")" : "") + ".";
      $("#signbtn").onclick = () => msalApp.loginRedirect({ scopes: CFG.scopes, loginHint: dev && dev.email || undefined }); return;
    }
    msalApp.setActiveAccount(account); $("#signin").hidden = true;
    if (dev && dev.email && String(account.username || "").toLowerCase() !== dev.email.toLowerCase()) {
      showOnly("mismatch"); $("#mismatchMsg").textContent = "This phone is registered to " + (dev.name || dev.email) + " (" + dev.email + "), but you're signed in as " + account.username + ".";
      $("#mismatchOut").onclick = () => msalApp.logoutRedirect({ account }); return;
    }
    try {
      await loadData(); wire(); checkAdmin();
      $("#who").textContent = account.name || account.username; showOnly("main"); status("");
    } catch (e) {
      if (e.message === "redirecting") return;
      if (e.message === "noaccess") return status("You're signed in, but you don't have access to the Pristine Special Orders site. Ask Matthew Hessell to add you.", true);
      status("Couldn't load accounts and products (" + esc(e.message) + "). Reload the page in a minute.", true);
    }
  }
  boot();
})();
