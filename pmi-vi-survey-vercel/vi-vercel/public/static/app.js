(() => {
  'use strict';

  /* ------------------------------------------------------------ helpers */
  const $ = (s, el = document) => el.querySelector(s);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const clone = (o) => JSON.parse(JSON.stringify(o));

  const slug = location.pathname.split('/')[2];
  const qs = new URLSearchParams(location.search);
  let KEY = qs.get('k') || sessionStorage.getItem('k:' + slug) || '';
  if (KEY) sessionStorage.setItem('k:' + slug, KEY);

  const view = $('#view');

  const S = {
    data: null,
    me: localStorage.getItem('vi-name') || '',
    ws: null,
    ref: null,
    cur: null,
    rev: 0,
    missing: [],
    dirty: false,
    saving: false,
    ver: 0,
    timer: null,
    conflict: null,
    remote: null,
    ownerTimers: {},
  };

  async function api(path, body, extra) {
    const [route, ref] = path.split('/');
    const params = new URLSearchParams(Object.assign({ route, slug }, ref ? { ref } : {}, extra || {}));
    const r = await fetch(`/api/index?${params}`, {
      method: body ? 'POST' : 'GET',
      headers: { 'Content-Type': 'application/json', 'X-Key': KEY },
      body: body ? JSON.stringify(body) : undefined,
    });
    let data = null;
    try { data = await r.json(); } catch (e) { /* empty body */ }
    return { status: r.status, ok: r.ok, data };
  }

  function ago(iso) {
    if (!iso) return '';
    const s = (Date.now() - new Date(iso).getTime()) / 1000;
    if (s < 45) return 'just now';
    if (s < 3600) return Math.round(s / 60) + ' min ago';
    if (s < 86400) return Math.round(s / 3600) + ' h ago';
    return new Date(iso).toLocaleDateString();
  }
  const first = (n) => String(n || '').split(/\s+/)[0];

  function toast(msg) {
    const t = document.createElement('div');
    t.className = 'toast';
    t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(() => t.remove(), 2600);
  }

  /* -------------------------------------------------------- data access */
  const wsList = () => S.data.structure;
  const findWs = (ref) => wsList().find((w) => w.ref === ref);
  const findTp = (ref) => { for (const w of wsList()) { const t = w.touchpoints.find((x) => x.ref === ref); if (t) return t; } return null; };
  const tpRec = (ref) => S.data.tp[ref];
  const tpState = (ref) => { const r = tpRec(ref); return !r ? '' : r.missing.length ? 'part' : 'ok'; };
  const wsDone = (w) => w.touchpoints.filter((t) => tpState(t.ref) === 'ok').length;
  const owner = (ref) => S.data.owners[ref] || { value: '', comment: '' };

  const blank = () => ({
    exists: '', mark: '', assets: [{ name: '', identity: '', identity_other: '' }],
    qty_mode: '', qty_num: '', qty_unit: '', qty_band: '',
    audience: '', visibility: '', effort: '', cost_min: '', cost_max: '', cost_unknown: false,
    support_tags: [], support_text: '', constraint_tags: [], constraint_text: '',
    owner_exception: '', comment: '',
  });

  function loadCurrent() {
    const rec = S.ref ? tpRec(S.ref) : null;
    S.cur = Object.assign(blank(), rec ? clone(rec.v) : {});
    if (!S.cur.assets || !S.cur.assets.length) S.cur.assets = blank().assets;
    S.rev = rec ? rec.rev : 0;
    S.missing = rec ? rec.missing : ['Whether the touchpoint exists'];
    S.dirty = false;
    S.conflict = null;
    S.remote = null;
  }

  /* ------------------------------------------------------------ routing */
  function parseHash() {
    const h = location.hash.replace(/^#\/?/, '').split('/');
    const w = h[0] && findWs(h[0]) ? h[0] : null;
    S.ws = w;
    if (w) {
      const ws = findWs(w);
      const wanted = h[1] && ws.touchpoints.find((t) => t.ref === h[1]) ? h[1] : null;
      S.ref = wanted || (ws.touchpoints.find((t) => tpState(t.ref) !== 'ok') || ws.touchpoints[0]).ref;
    } else {
      S.ref = null;
    }
    loadCurrent();
    render();
    setTimeout(poll, 400);
  }

  async function go(hash) {
    await flush();
    if (location.hash === hash) parseHash(); else location.hash = hash;
    window.scrollTo({ top: 0 });
  }
  window.addEventListener('hashchange', async () => { if (S.data) { await flush(); parseHash(); } });

  /* ------------------------------------------------------------- saving */
  function queueSave() {
    S.dirty = true;
    S.ver++;
    const sb = $('#savebar'); if (sb) sb.textContent = 'Saving…';
    clearTimeout(S.timer);
    S.timer = setTimeout(saveNow, 600);
  }

  async function saveNow() {
    clearTimeout(S.timer);
    if (!S.dirty || S.saving || S.conflict) return;
    S.saving = true;
    const ref = S.ref, ver = S.ver, payload = clone(S.cur);
    let r;
    try { r = await api('tp/' + ref, { v: payload, base_rev: S.rev, by: S.me }); }
    catch (e) { S.saving = false; const sb = $('#savebar'); if (sb) sb.textContent = 'Connection interrupted, retrying…'; S.timer = setTimeout(saveNow, 3000); return; }
    S.saving = false;
    if (r.status === 409) { S.conflict = r.data.current; renderBanner(); return; }
    if (!r.ok) { const sb = $('#savebar'); if (sb) sb.textContent = 'Save pending'; return; }
    S.data.tp[ref] = { v: payload, rev: r.data.rev, by: S.me, at: r.data.at, missing: r.data.missing };
    S.data.progress = r.data.progress;
    S.data.ready = r.data.ready;
    if (ref === S.ref) { S.rev = r.data.rev; S.missing = r.data.missing; }
    if (ver === S.ver) S.dirty = false; else saveNow();
    updateStatus(); updateRail();
  }

  async function flush() {
    clearTimeout(S.timer);
    for (let i = 0; i < 25 && (S.dirty || S.saving) && !S.conflict; i++) {
      if (!S.saving) await saveNow(); else await sleep(120);
    }
  }

  function saveOwner(ref) {
    clearTimeout(S.ownerTimers[ref]);
    S.ownerTimers[ref] = setTimeout(async () => {
      const o = owner(ref);
      const r = await api('owner/' + ref, { value: o.value, comment: o.comment, by: S.me });
      if (r.ok) { S.data.progress = r.data.progress; S.data.ready = r.data.ready; updateOverall(); }
    }, 600);
  }

  /* ------------------------------------------------------------ polling */
  async function poll() {
    if (document.hidden || !S.data) return;
    const r = await api('state', null, { since: S.sver || 0, by: S.me, ref: S.ref || '' });
    if (!r.ok) return;
    const focused = view.contains(document.activeElement) && /INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName);
    const presChanged = JSON.stringify(r.data.presence || []) !== JSON.stringify(S.data.presence || []);
    S.data.presence = r.data.presence;
    if (r.data.unchanged) {
      if (presChanged && !focused) { if (S.ws) updateRail(); else renderHome(); }
      return;
    }
    S.sver = r.data.ver;
    const prevTp = S.data.tp;
    S.data.tp = r.data.tp; S.data.progress = r.data.progress; S.data.ready = r.data.ready;
    S.data.submissions = r.data.submissions;
    if (S.ref && S.dirty && prevTp[S.ref]) S.data.tp[S.ref] = prevTp[S.ref];
    for (const k of Object.keys(r.data.owners)) {
      if (!(S.ownerTimers[k] && focused)) S.data.owners[k] = r.data.owners[k];
    }
    if (!S.ws) {
      if (!focused) renderHome();
    } else {
      updateRail();
      const rec = tpRec(S.ref);
      if (rec && rec.rev > S.rev && !S.conflict) {
        if (!S.dirty && !focused) { S.cur = Object.assign(blank(), clone(rec.v)); S.rev = rec.rev; S.missing = rec.missing; renderForm(); }
        else if (rec.by !== S.me) { S.remote = rec; renderBanner(); }
      }
      updateStatus();
    }
  }

  /* ------------------------------------------------------------- render */
  function render() {
    $('#mname').textContent = S.data.market.name;
    $('#mename').textContent = S.me;
    if (S.ws) renderWs(); else renderHome();
  }

  function otherEditors(refMatch) {
    return (S.data.presence || []).filter((p) => p.by !== S.me && refMatch(p.ref));
  }

  /* ---------- home ---------- */
  function renderHome() {
    const p = S.data.progress;
    const pct = Math.round((p.done / p.total) * 100);
    const subs = S.data.submissions || [];
    const lastSub = subs[subs.length - 1];

    const cards = wsList().map((w) => {
      const done = wsDone(w), tot = w.touchpoints.length, pc = Math.round((done / tot) * 100);
      const cls = done === tot ? 'done' : done > 0 ? 'part' : '';
      const act = otherEditors((r) => r && r[0] === w.ref);
      const tpChips = w.touchpoints.map((t) => `<button type="button" class="tpchip ${tpState(t.ref)}" data-go="#/${w.ref}/${t.ref}" title="${esc(t.title)}">${esc(t.ref)}</button>`).join('');
      return `<article class="ws ${cls}">
        <div class="ws-head"><div class="badge">${done === tot ? '✓' : esc(w.ref)}</div>
          <div style="flex:1"><h3>${esc(w.title)}</h3><div class="count">${done} of ${tot} touchpoints answered${act.length ? ` · <span class="who">${esc(act.map((a) => first(a.by)).join(', '))} active</span>` : ''}</div></div>
          <div class="pct">${pc}%</div></div>
        <div class="bar"><i style="width:${pc}%"></i></div>
        <div><span class="field-label">Questions (tap one to answer it)</span><div class="tpchips">${tpChips}</div></div>
        <div><label class="field-label" for="own-${w.ref}">Responsible team / contact (workstream default)</label>
          <input type="text" id="own-${w.ref}" data-owner="${w.ref}" data-k="value" placeholder="e.g., Internal Communications" value="${esc(owner(w.ref).value)}"></div>
        <div class="ws-actions">
          <button class="btn" data-go="#/${w.ref}" type="button">${done === tot ? 'Review workstream' : done ? 'Continue workstream' : 'Open workstream'}</button>
          <button class="btn ghost" data-copylink="${w.ref}" type="button" title="Copies a direct link so a colleague can pick up this workstream">Copy link</button>
        </div></article>`;
    }).join('');

    const open = [];
    wsList().forEach((w) => {
      const n = w.touchpoints.length - wsDone(w);
      if (n) open.push(`${w.ref} · ${n} touchpoint${n > 1 ? 's' : ''}`);
      if (!owner(w.ref).value.trim()) open.push(`${w.ref} · owner`);
    });

    view.innerHTML = `
      <section class="overall">
        <h2>Overall progress</h2>
        <p>Quantifying where PMI corporate communications exist today, which visual identities are in use, and what each change would involve. Six workstreams cover ${p.total} touchpoints. Answers save automatically and are shared live with colleagues working on the same market link, so workstreams can be distributed across teams.</p>
        <div class="bar"><i style="width:${pct}%"></i></div>
        <div class="bar-meta"><span>${p.done} of ${p.total} touchpoints answered</span><span>${p.owners_done} of ${p.owners_total} workstream owners confirmed</span></div>
      </section>
      <div class="grid">${cards}</div>
      <section class="submit-panel" id="submit-panel">${submitPanel(open, lastSub)}</section>`;
  }

  function submitPanel(open, lastSub) {
    const can = !S.data.require_complete || S.data.ready;
    const note = lastSub ? `<div class="submitted-note">Last submitted ${esc(ago(lastSub.at))}${lastSub.by ? ' by ' + esc(lastSub.by) : ''}. Later changes can be submitted again.</div>` : '';
    return `<div>
      <h2>${open.length ? 'Submit when the inventory feels complete' : 'Everything is answered and ready to submit'}</h2>
      <p>Colleagues can answer the touchpoints they know best; answered touchpoints turn green. Submitting sends the whole inventory to the project team as one Excel workbook.${open.length ? ' Still open:' : ''}</p>
      ${open.length ? `<ul class="remaining">${open.map((o) => `<li>${esc(o)}</li>`).join('')}</ul>` : ''}
      ${note}</div>
      <button class="btn" id="open-submit" type="button" ${can ? '' : 'disabled'}>${lastSub ? 'Submit again' : 'Submit inventory'}</button>`;
  }

  function updateOverall() {
    if (!S.ws) { const f = view.contains(document.activeElement) && /INPUT|TEXTAREA/.test(document.activeElement.tagName); if (!f) renderHome(); }
  }

  /* ---------- workstream view ---------- */
  function renderWs() {
    const w = findWs(S.ws);
    view.innerHTML = `<div class="layout">
      <aside class="rail">
        <button class="btn ghost small" data-go="#/" type="button" style="margin:2px 4px 10px">← Overview</button>
        <h3>${esc(w.ref)} · ${esc(w.title)}</h3>
        <div id="rail-meta" style="margin:0 8px 10px"></div>
        <div id="rail-items"></div>
        <div style="padding:12px 8px 6px"><label class="field-label" for="own-ws">Responsible team / contact</label>
          <input type="text" id="own-ws" data-owner="${w.ref}" data-k="value" placeholder="Workstream default" value="${esc(owner(w.ref).value)}">
          <label class="field-label" for="own-ws-c" style="margin-top:10px">Workstream-wide comment</label>
          <textarea id="own-ws-c" data-owner="${w.ref}" data-k="comment" placeholder="Optional" style="min-height:60px">${esc(owner(w.ref).comment)}</textarea></div>
      </aside>
      <section class="card" id="card"></section></div>`;
    updateRail();
    renderForm();
  }

  function updateRail() {
    const box = $('#rail-items'); if (!box) return;
    const w = findWs(S.ws);
    const done = wsDone(w), pc = Math.round((done / w.touchpoints.length) * 100);
    $('#rail-meta').innerHTML = `<div class="bar"><i style="width:${pc}%"></i></div><div class="bar-meta"><span>${done} of ${w.touchpoints.length} answered</span><span>${pc}%</span></div>`;
    box.innerHTML = w.touchpoints.map((t) => {
      const others = otherEditors((r) => r === t.ref);
      return `<button class="rail-item ${t.ref === S.ref ? 'active' : ''} ${tpState(t.ref) === 'ok' ? 'ok' : ''}" data-go="#/${w.ref}/${t.ref}" type="button">
        <span class="state ${tpState(t.ref)}"></span><span class="ref">${esc(t.ref)}</span><span class="t">${esc(t.title)}</span>
        ${others.length ? `<span class="who">${esc(first(others[0].by))}</span>` : ''}</button>`;
    }).join('');
  }

  const seg = (field, opts, val) => `<div class="seg" role="radiogroup">${opts.map(([v, l]) =>
    `<button type="button" role="radio" aria-checked="${val === v}" class="${val === v ? 'on' : ''}" data-seg="${field}" data-val="${esc(v)}">${esc(l)}</button>`).join('')}</div>`;
  const chips = (field, opts, sel) => `<div class="chips">${opts.map((o) =>
    `<button type="button" class="pill ${sel.includes(o) ? 'on' : ''}" data-chip="${field}" data-val="${esc(o)}">${esc(o)}</button>`).join('')}</div>`;

  function renderForm() {
    const card = $('#card'); if (!card) return;
    const w = findWs(S.ws), t = findTp(S.ref), v = S.cur;
    const idx = w.touchpoints.findIndex((x) => x.ref === S.ref);
    const prev = w.touchpoints[idx - 1], next = w.touchpoints[idx + 1];
    const O = S.data.options;
    const isYes = v.exists === 'yes';

    const assetRows = v.assets.map((a, i) => `
      <div class="asset-row">
        <input type="text" data-asset="${i}" data-k="name" placeholder="Asset, channel or location" value="${esc(a.name)}" aria-label="Asset ${i + 1}">
        <select data-asset="${i}" data-k="identity" aria-label="Visual identity used for asset ${i + 1}">
          <option value="">Visual identity used…</option>
          ${O.identities.map((o) => `<option ${a.identity === o ? 'selected' : ''}>${esc(o)}</option>`).join('')}
        </select>
        <button type="button" class="icon-btn" data-rm-asset="${i}" title="Remove asset" aria-label="Remove asset ${i + 1}">✕</button>
        ${a.identity === 'Other' ? `<input class="other" type="text" data-asset="${i}" data-k="identity_other" placeholder="Describe the visual identity used" value="${esc(a.identity_other)}">` : ''}
      </div>`).join('');

    const qty = `
      ${seg('qty_mode', [['exact', 'Exact count'], ['band', 'Band'], ['na', 'Not applicable']], v.qty_mode)}
      ${v.qty_mode === 'exact' ? `<div class="row2" style="margin-top:10px">
        <input type="number" min="0" step="any" data-f="qty_num" placeholder="Number" value="${esc(v.qty_num)}" aria-label="Quantity">
        <input type="text" data-f="qty_unit" placeholder="Unit (sites, accounts, screens, rooms, vehicles…)" value="${esc(v.qty_unit)}" aria-label="Unit"></div>` : ''}
      ${v.qty_mode === 'band' ? `<div style="margin-top:10px">${seg('qty_band', O.bands.map((b) => [b, b]), v.qty_band)}</div>` : ''}`;

    const cost = ['some effort', 'meaningful investment'].includes(v.effort) ? `
      <div class="q"><div class="ql">Estimated one-off cost range</div>
        <p class="hint">Local currency only; no conversion to EUR is needed. A rough range is helpful at this stage.</p>
        <div class="row3"><div class="cur">${esc(S.data.market.currency)}</div>
          <input type="number" min="0" step="any" data-f="cost_min" placeholder="From" value="${esc(v.cost_min)}" aria-label="Cost from">
          <input type="number" min="0" step="any" data-f="cost_max" placeholder="To" value="${esc(v.cost_max)}" aria-label="Cost to"></div>
        <label class="inline-check"><input type="checkbox" data-f="cost_unknown" ${v.cost_unknown ? 'checked' : ''}> Not yet known</label></div>` : '';

    const details = isYes ? `<div class="followups">
      <div class="q"><div class="ql">Does any corporate visual identity mark appear on this touchpoint today?</div>${seg('mark', [['yes', 'Yes'], ['no', 'No']], v.mark)}</div>

      <div class="q"><div class="ql">What specifically exists, and which visual identity does each asset use?</div>
        <p class="hint">One row per distinct asset, channel or location, for example “One Swiss intranet” or “Swiss IT Portal”. Assets without a mark can be included with “No mark”.</p>
        ${assetRows}
        <button type="button" class="btn ghost small" data-add-asset>+ Add another asset</button></div>

      <div class="q"><div class="ql">Quantity, if relevant</div>
        <p class="hint">Active or reusable items count; old file versions do not. A practical unit keeps entries comparable across markets.</p>${qty}</div>

      <div class="q"><div class="ql">Audience</div>${seg('audience', [['internal', 'Internal'], ['external', 'External'], ['both', 'Both']], v.audience)}
        <div class="defs"><span>Internal = employees · External = customers, partners, authorities, media or public · Both = used by both groups</span></div></div>

      <div class="q"><div class="ql">Visibility</div>${seg('visibility', [['low', 'Low'], ['medium', 'Medium'], ['high', 'High']], v.visibility)}
        <div class="defs"><span>Low = limited exposure · Medium = broad market or limited external exposure · High = frequent external exposure or broad employee exposure</span></div></div>

      <div class="q"><div class="ql">Effort of changing this touchpoint</div>
        <p class="hint">A directional view is sufficient. “Unknown” fits where the effort cannot yet be assessed.</p>
        ${seg('effort', [['easy / no-cost change', 'Easy / no-cost change'], ['some effort', 'Some effort'], ['meaningful investment', 'Meaningful investment'], ['unknown', 'Unknown']], v.effort)}</div>
      ${cost}

      <div class="q"><div class="ql">Support that would help to make the change</div>
        <p class="hint">Topics can be selected and a short detail added. “None needed” can be written where no support is expected.</p>
        ${chips('support_tags', O.support, v.support_tags)}
        <textarea data-f="support_text" placeholder="Optional detail (e.g., agency help for the portal banner)">${esc(v.support_text)}</textarea></div>

      <div class="q"><div class="ql">Rollout constraints and dependencies</div>
        <p class="hint">Budget timing, contracts, release windows, refurbishment cycles, approvals or procurement shape the sequencing. “None identified” can be written where nothing applies.</p>
        ${chips('constraint_tags', O.constraints, v.constraint_tags)}
        <textarea data-f="constraint_text" placeholder="Optional detail (e.g., next IT release window in Q2)">${esc(v.constraint_text)}</textarea></div>

      <div class="q"><details ${v.owner_exception ? 'open' : ''}><summary class="ql" style="cursor:pointer">Different owner for this touchpoint?</summary>
        <p class="hint" style="margin-top:6px">Only relevant where the owner differs from the workstream default${owner(w.ref).value ? ` (currently: ${esc(owner(w.ref).value)})` : ''}.</p>
        <input type="text" data-f="owner_exception" placeholder="Team, function or named contact" value="${esc(v.owner_exception)}"></details></div>
    </div>` : '';

    const existsNote = v.exists === 'no'
      ? '<p class="hint" style="margin-top:10px">No further questions apply to this touchpoint.</p>'
      : v.exists === 'unknown'
        ? '<p class="hint" style="margin-top:10px">“Unknown” is a valid answer. A short note under Comments helps the follow-up.</p>' : '';

    card.innerHTML = `
      <div class="crumb">Workstream ${esc(w.ref)} · ${esc(w.title)} · touchpoint ${idx + 1} of ${w.touchpoints.length}</div>
      <div class="card-head"><h2><span class="ref">${esc(t.ref)}</span>${esc(t.title)}</h2><span class="status" id="status"></span></div>
      <div class="savebar" id="savebar"></div>
      <div id="banner"></div>

      <div class="q"><div class="ql">Does this touchpoint exist for PMI corporate market communications?</div>
        <p class="hint">Product-only communications are outside the scope. This question is separate from whether a visual identity mark is present.</p>
        ${seg('exists', [['yes', 'Yes'], ['no', 'No'], ['unknown', 'Unknown']], v.exists)}${existsNote}</div>
      ${details}

      ${v.exists ? `<div class="divider"></div><div class="q"><label for="comment">Comments <span style="font-weight:400;color:var(--muted)">(optional)</span></label>
        <p class="hint">Context, assumptions or open questions that help others to interpret the answers.</p>
        <textarea id="comment" data-f="comment">${esc(v.comment)}</textarea></div>` : ''}

      <div class="nav">
        ${prev ? `<button class="btn ghost" data-go="#/${w.ref}/${prev.ref}" type="button">← ${esc(prev.ref)}</button>` : '<span></span>'}
        ${next ? `<button class="btn" data-go="#/${w.ref}/${next.ref}" type="button">Next: ${esc(next.ref)} →</button>`
          : `<button class="btn" data-go="#/" type="button">Back to overview</button>`}
      </div>`;
    updateStatus();
    renderBanner();
  }

  function updateStatus() {
    const st = $('#status'), sb = $('#savebar'); if (!st) return;
    const m = S.missing;
    st.className = 'status' + (m.length ? '' : ' ok');
    st.textContent = m.length ? 'Still open: ' + m.slice(0, 3).join(', ') + (m.length > 3 ? ` +${m.length - 3}` : '') : 'Answered ✓';
    const rec = tpRec(S.ref);
    if (sb && !S.dirty && !S.saving) sb.textContent = rec ? `Saved · ${rec.by || 'contributor'} · ${ago(rec.at)}` : 'Progress saves automatically';
  }

  function renderBanner() {
    const b = $('#banner'); if (!b) return;
    if (S.conflict) {
      b.innerHTML = `<div class="banner"><span><strong>${esc(S.conflict.by || 'A colleague')}</strong> updated this touchpoint in the meantime.</span>
        <button class="btn small" data-conflict="theirs" type="button">Load their version</button>
        <button class="btn ghost small" data-conflict="mine" type="button">Keep this version</button></div>`;
    } else if (S.remote) {
      b.innerHTML = `<div class="banner"><span><strong>${esc(S.remote.by || 'A colleague')}</strong> added an update to this touchpoint.</span>
        <button class="btn small" data-conflict="remote" type="button">Load update</button></div>`;
    } else b.innerHTML = '';
  }

  /* ----------------------------------------------------------- modals */
  function modal(html, onMount) {
    const o = document.createElement('div');
    o.className = 'overlay';
    o.innerHTML = `<div class="modal" role="dialog" aria-modal="true">${html}</div>`;
    document.body.appendChild(o);
    onMount(o, () => o.remove());
    return o;
  }

  function askName(force) {
    modal(`<h2>Welcome</h2><p>A name or initials allows entries to be attributed and shown to colleagues working in parallel.</p>
      <input type="text" id="nm" placeholder="Name or initials" value="${esc(S.me)}" maxlength="60">
      <div class="actions">${force ? '<button class="btn ghost" id="cx" type="button">Cancel</button>' : ''}<button class="btn" id="ok" type="button">Continue</button></div>`,
      (o, close) => {
        const inp = $('#nm', o); inp.focus();
        const done = () => { const val = inp.value.trim(); if (!val) return; S.me = val; localStorage.setItem('vi-name', val); $('#mename').textContent = val; close(); poll(); };
        $('#ok', o).onclick = done; inp.onkeydown = (e) => { if (e.key === 'Enter') done(); };
        const cx = $('#cx', o); if (cx) cx.onclick = close;
      });
  }

  function openItems() {
    const tps = [], owners = [];
    wsList().forEach((w) => {
      w.touchpoints.forEach((x) => { if (tpState(x.ref) !== 'ok') tps.push(x.ref); });
      if (!owner(w.ref).value.trim()) owners.push(w.ref);
    });
    return { tps, owners };
  }

  function askSubmit() {
    const o = openItems();
    const blocked = S.data.require_complete && (o.tps.length || o.owners.length);
    const warn = (o.tps.length || o.owners.length) ? `<div class="banner" style="margin:0 0 14px"><span><strong>Not everything is answered yet.</strong>
      ${o.tps.length ? `Open touchpoints (${o.tps.length}): ${esc(o.tps.join(', '))}.` : ''}
      ${o.owners.length ? ` Workstream owner not set for: ${esc(o.owners.join(', '))}.` : ''}
      Open items appear empty in the workbook.</span></div>` : '';
    modal(`<h2>Submit the entire inventory?</h2>
      <p>All answers for ${esc(S.data.market.name)} are sent to the project team as an Excel workbook. Before sending, please make sure that all questions have been answered, or that “Unknown” has been selected where the information is not available.</p>
      ${warn}
      <label class="field-label" for="sb">Submitted by</label><input type="text" id="sb" value="${esc(S.me)}">
      <label class="field-label" for="sn" style="margin-top:12px">Note for the project team (optional)</label><textarea id="sn"></textarea>
      <div class="actions"><button class="btn ghost" id="cx" type="button">Not yet, keep editing</button><button class="btn" id="go" type="button" ${blocked ? 'disabled' : ''}>Yes, submit</button></div>`,
      (m, close) => {
        $('#cx', m).onclick = close;
        $('#go', m).onclick = async () => {
          const btn = $('#go', m); btn.disabled = true; btn.textContent = 'Submitting…';
          await flush();
          const r = await api('submit', { by: $('#sb', m).value.trim() || S.me, note: $('#sn', m).value.trim() });
          close();
          if (r.ok) {
            toast(r.data.mail === 'sent' ? 'Submitted. The workbook has been sent to the project team by email.' : 'Submitted and stored. Email delivery is not active yet; the project team can download the workbook from the admin page.');
            S.sver = 0; await poll(); if (!S.ws) renderHome();
          } else toast('Submission could not be completed. Open touchpoints are listed on the overview.');
        };
      });
  }

  /* ----------------------------------------------------------- events */
  view.addEventListener('click', async (e) => {
    const el = e.target.closest('button, [data-go]'); if (!el) return;
    if (el.dataset.go) return go(el.dataset.go);
    if (el.id === 'open-submit') return askSubmit();
    if (el.dataset.copylink) {
      const url = `${location.origin}/m/${slug}?k=${encodeURIComponent(KEY)}#/${el.dataset.copylink}`;
      try { await navigator.clipboard.writeText(url); toast('Link copied'); } catch (err) { prompt('Link', url); }
      return;
    }
    if (el.dataset.seg) { S.cur[el.dataset.seg] = el.dataset.val; renderForm(); return queueSave(); }
    if (el.dataset.chip) {
      const arr = S.cur[el.dataset.chip], i = arr.indexOf(el.dataset.val);
      if (i >= 0) arr.splice(i, 1); else arr.push(el.dataset.val);
      renderForm(); return queueSave();
    }
    if (el.hasAttribute('data-add-asset')) {
      S.cur.assets.push({ name: '', identity: '', identity_other: '' }); renderForm();
      const rows = view.querySelectorAll('.asset-row input[data-k=name]'); rows[rows.length - 1].focus(); return queueSave();
    }
    if (el.dataset.rmAsset !== undefined) {
      S.cur.assets.splice(+el.dataset.rmAsset, 1);
      if (!S.cur.assets.length) S.cur.assets = blank().assets;
      renderForm(); return queueSave();
    }
    if (el.dataset.conflict) {
      const kind = el.dataset.conflict;
      if (kind === 'mine') { S.rev = S.conflict.rev; S.conflict = null; renderBanner(); queueSave(); }
      else {
        const src = kind === 'theirs' ? S.conflict : S.remote;
        S.cur = Object.assign(blank(), clone(src.v)); S.rev = src.rev;
        S.missing = (tpRec(S.ref) && tpRec(S.ref).missing) || S.missing;
        S.conflict = null; S.remote = null; S.dirty = false; renderForm();
        poll();
      }
    }
  });

  view.addEventListener('input', (e) => {
    const t = e.target;
    if (t.dataset.owner) {
      const o = S.data.owners[t.dataset.owner] || (S.data.owners[t.dataset.owner] = { value: '', comment: '' });
      o[t.dataset.k] = t.value; return saveOwner(t.dataset.owner);
    }
    if (t.dataset.asset !== undefined && t.tagName === 'INPUT') { S.cur.assets[+t.dataset.asset][t.dataset.k] = t.value; return queueSave(); }
    if (t.dataset.f && t.type !== 'checkbox') { S.cur[t.dataset.f] = t.value; queueSave(); }
  });

  view.addEventListener('change', (e) => {
    const t = e.target;
    if (t.dataset.asset !== undefined && t.tagName === 'SELECT') { S.cur.assets[+t.dataset.asset].identity = t.value; renderForm(); return queueSave(); }
    if (t.dataset.f === 'cost_unknown') { S.cur.cost_unknown = t.checked; queueSave(); }
  });

  $('#me').onclick = () => askName(true);
  $('#dl').onclick = async () => {
    await flush();
    const r = await fetch(`/api/index?route=export&slug=${slug}`, { headers: { 'X-Key': KEY } });
    if (!r.ok) return toast('Download not available');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(await r.blob()); a.download = `${slug}_draft.xlsx`; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  };
  window.addEventListener('beforeunload', (e) => { if (S.dirty) { e.preventDefault(); e.returnValue = ''; } });
  document.addEventListener('visibilitychange', () => { if (document.hidden) flush(); });

  /* -------------------------------------------------------------- boot */
  (async function init() {
    const r = await api('state', null, { by: S.me });
    if (!r.ok) {
      view.innerHTML = `<div class="center"><h2>This link could not be opened</h2><p>The complete market link, including the access key, is needed. The project team can provide it.</p></div>`;
      $('#mname').textContent = 'Market inventory';
      return;
    }
    S.data = r.data;
    S.sver = r.data.ver;
    if (!S.me) askName(false);
    parseHash();
    setInterval(poll, 15000);
  })();
})();
