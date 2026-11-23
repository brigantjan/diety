// =====================================================================
// SUPTel – Diéty a cestovné príkazy – frontend
// Nastavenia (zamestnanci, sadzby, časy) sú v Supabase (tabuľka diety_nastavenia).
// Výpočet beží priamo v prehliadači: Python (Pyodide) + openpyxl, súbory
// dochádzky sa nikam neodosielajú.
// =====================================================================

const el = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const TYPY = ['dochádza', 'týždňovkár', 'týždňovkár nonstop', 'bez CP'];
const TYP_POPIS = {'dochádza':'dochádza denne domov','týždňovkár':'týždňovkár','týždňovkár nonstop':'týždňovkár nonstop','bez CP':'bez cestovného príkazu (kancelária)'};
const CZ_MODY = ['jedna viacdňová cesta', 'samostatné jednodňové cesty'];
const fmtEur = v => (Number(v) || 0).toLocaleString('sk-SK', {minimumFractionDigits:2, maximumFractionDigits:2});
const fmtNum = v => (Number(v) || 0).toLocaleString('sk-SK', {maximumFractionDigits:2});

const state = { user:null, data:null, version:null, py:null, api:null, pyPromise:null,
                dochName:'', analyza:null, vysledok:null, urls:[] };

// ---------------------------------------------------------------------
// Úložisko nastavení
// ---------------------------------------------------------------------
const LOCAL_MODE = !SUPABASE_URL;            // len na testovanie bez Supabase
const sb = LOCAL_MODE ? null : window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const Store = {
  async load() {
    if (LOCAL_MODE) {
      let raw = localStorage.getItem('diety_nastavenia');
      if (!raw) raw = await (await fetch('seed.json')).text();
      return { data: JSON.parse(raw), version: 'local' };
    }
    const { data, error } = await sb.from('diety_nastavenia').select('data,updated_at').eq('id', 'main').maybeSingle();
    if (error) throw error;
    if (!data) throw new Error('NO_ACCESS');
    return { data: data.data, version: data.updated_at };
  },
  async save(d) {
    if (LOCAL_MODE) { localStorage.setItem('diety_nastavenia', JSON.stringify(d)); return 'local'; }
    const { data, error } = await sb.from('diety_nastavenia').update({ data: d })
      .eq('id', 'main').eq('updated_at', state.version).select('updated_at').maybeSingle();
    if (error) throw error;
    if (!data) throw new Error('Nastavenia medzitým zmenil niekto iný. Obnov stránku (F5) a zopakuj zmenu.');
    return data.updated_at;
  }
};

// Archív mesiacov (súčty po ľuďoch) – tabuľka diety_archiv
const Archiv = {
  async nacitaj() {
    if (LOCAL_MODE) return JSON.parse(localStorage.getItem('diety_archiv') || '[]');
    const { data, error } = await sb.from('diety_archiv').select('rok,mesiac,data,updated_at').order('rok').order('mesiac');
    if (error) throw error;
    return data || [];
  },
  async uloz(rok, mesiac, data) {
    if (LOCAL_MODE) {
      const all = JSON.parse(localStorage.getItem('diety_archiv') || '[]').filter(x => !(x.rok === rok && x.mesiac === mesiac));
      all.push({ rok, mesiac, data, updated_at: new Date().toISOString() });
      localStorage.setItem('diety_archiv', JSON.stringify(all));
      return;
    }
    const { error } = await sb.from('diety_archiv').upsert({ rok, mesiac, data }, { onConflict: 'rok,mesiac' });
    if (error) throw error;
  }
};

async function ulozDoArchivu(res) {
  const rows = await Archiv.nacitaj();
  const stary = rows.find(x => x.rok === res.rok && x.mesiac === res.mesiac);
  const data = structuredClone(stary ? stary.data : {});
  const cas = new Date().toISOString();
  const z = { sr: {}, cz: {} };
  for (const x of res.prehlad) {
    z.sr[x[0]] = { spolu: x[7], dni: x[3] };
    if (x[8]) z.cz[x[0]] = { spolu: x[9], dni: x[8] };
  }
  const prepisane = [];
  for (const druh of ['sr', 'cz']) {
    if (data[druh] && data[druh].zdroj === 'finálny súbor') continue;   // finálne (upravené) súbory majú prednosť
    data[druh] = { zdroj: 'výpočet', cas, ludia: z[druh] };
    prepisane.push(druh);
  }
  if (prepisane.length) await Archiv.uloz(res.rok, res.mesiac, data);
  return prepisane.length === 2 ? 'uložené do archívu'
       : prepisane.length ? 'uložené do archívu (druhý príkaz už má v archíve finálny súbor)'
       : 'archív sa nemenil – tento mesiac už má nahraté finálne súbory';
}

async function ulozNastavenia(newData) {
  const v = await Store.save(newData);
  state.data = newData;
  state.version = v;
}

// ---------------------------------------------------------------------
// Výpočtový modul (Pyodide)
// ---------------------------------------------------------------------
function engineStatus(txt, cls) {
  const s = el('engine-status');
  s.querySelector('.dot').className = 'dot ' + (cls || '');
  s.querySelector('span:last-child').textContent = txt;
}

function loadScript(src) {
  return new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = src; s.onload = res; s.onerror = () => rej(new Error('Nepodarilo sa načítať ' + src));
    document.head.appendChild(s);
  });
}

function startPy() {
  if (state.pyPromise) return state.pyPromise;
  state.pyPromise = (async () => {
    engineStatus('načítavam…');
    await loadScript(PYODIDE_URL + 'pyodide.js');
    const py = await loadPyodide({ indexURL: PYODIDE_URL });
    py.setStdout({ batched: s => console.log('[py]', s) });
    const bin = async p => new Uint8Array(await (await fetch(p, {cache:'no-cache'})).arrayBuffer());
    const txt = async p => (await fetch(p, {cache:'no-cache'})).text();
    for (const w of ['et_xmlfile-2.0.0-py3-none-any.whl', 'openpyxl-3.1.5-py2.py3-none-any.whl'])
      py.unpackArchive(await bin('wheels/' + w), 'wheel');
    py.FS.mkdirTree('/home/pyodide/data/sablony');
    py.FS.writeFile('/home/pyodide/engine.py', await txt('py/engine.py'));
    py.FS.writeFile('/home/pyodide/web_api.py', await txt('py/web_api.py'));
    for (const f of ['sablona_SR.xlsx', 'sablona_CZ.xlsx'])
      py.FS.writeFile('/home/pyodide/data/sablony/' + f, await bin('sablony/' + f));
    py.runPython("import sys; sys.path.insert(0, '/home/pyodide'); import web_api");
    state.py = py;
    state.api = py.pyimport('web_api');
    engineStatus('pripravený', 'ok');
    return py;
  })().catch(e => { engineStatus('chyba', 'err'); state.pyPromise = null; throw e; });
  return state.pyPromise;
}

function pyChyba(e) {
  const m = String(e && e.message || e);
  const lines = m.trim().split('\n').filter(Boolean);
  const last = lines[lines.length - 1] || m;
  return last.replace(/^\w+Error: /, '');
}

// ---------------------------------------------------------------------
// Prihlásenie
// ---------------------------------------------------------------------
async function poPrihlaseni(user) {
  state.user = user;
  try {
    const r = await Store.load();
    state.data = r.data; state.version = r.version;
  } catch (e) {
    const msg = e.message === 'NO_ACCESS'
      ? 'Tento účet nemá prístup k Diétam. Požiadaj správcu o pridelenie prístupu.'
      : 'Nastavenia sa nepodarilo načítať: ' + e.message;
    if (!LOCAL_MODE) await sb.auth.signOut();
    el('login-error').textContent = msg;
    el('login-error').classList.remove('hidden');
    return;
  }
  el('user-name').textContent = user ? user.email : 'lokálny režim';
  el('login-view').classList.add('hidden');
  el('app-view').classList.remove('hidden');
  prepni('spracovat');
  startPy().catch(e => sprava('Výpočtový modul sa nenačítal: ' + e.message + ' – skontroluj internet a obnov stránku (F5).', 'error'));
}

el('login-form').addEventListener('submit', async e => {
  e.preventDefault();
  el('login-btn').disabled = true;
  el('login-error').classList.add('hidden');
  const { data, error } = await sb.auth.signInWithPassword({
    email: el('login-email').value.trim(), password: el('login-password').value });
  el('login-btn').disabled = false;
  if (error) {
    el('login-error').textContent = 'Prihlásenie zlyhalo: ' + error.message;
    el('login-error').classList.remove('hidden');
    return;
  }
  poPrihlaseni(data.user);
});

el('logout-btn').addEventListener('click', async () => {
  if (!LOCAL_MODE) await sb.auth.signOut();
  location.reload();
});

// ---------------------------------------------------------------------
// Navigácia a správy
// ---------------------------------------------------------------------
function sprava(text, typ = 'ok') {
  el('msg').innerHTML = text ? `<div class="banner ${typ}">${esc(text)}</div>` : '';
  if (text) window.scrollTo({ top: 0, behavior: 'smooth' });
}

function prepni(view) {
  document.querySelectorAll('header nav button').forEach(b => b.classList.toggle('on', b.dataset.view === view));
  for (const v of ['spracovat', 'zamestnanci', 'sadzby', 'archiv']) el('v-' + v).classList.toggle('hidden', v !== view);
  sprava('');
  if (view === 'spracovat') (state.vysledok ? renderVysledok() : renderUpload());
  if (view === 'zamestnanci') renderZamestnanci();
  if (view === 'sadzby') renderSadzby();
  if (view === 'archiv') renderArchiv();
}
document.querySelectorAll('header nav button').forEach(b => b.addEventListener('click', () => prepni(b.dataset.view)));

// ---------------------------------------------------------------------
// 1) Nahratie dochádzky
// ---------------------------------------------------------------------
function renderUpload() {
  state.vysledok = null;
  el('v-spracovat').innerHTML = `
    <h2>Spracovať mesiac</h2>
    <p class="sub">Nahraj dochádzku – cestovné príkazy SR a ČR sa vytvoria automaticky.</p>
    <div class="card">
      <label class="drop" id="drop">
        <i class="ti ti-cloud-upload"></i>
        <b>Pretiahni sem súbor s dochádzkou</b>
        <span class="muted">alebo klikni a vyber ho (.xlsx)</span>
        <input type="file" id="subor" accept=".xlsx,.xlsm" hidden>
      </label>
      <div id="busy" class="busy hidden"><div class="spin"></div><span id="busy-txt">Spracúvam…</span></div>
      <div class="steps">
        <div class="step"><span>1</span>Nahráš dochádzku (aj rozpracovanú).</div>
        <div class="step"><span>2</span>Ak je v nej nový človek, hneď sa ťa opýtam na jeho údaje.</div>
        <div class="step"><span>3</span>Stiahneš SR príkaz, ČR príkaz a kontrolný prehľad.</div>
      </div>
      <p class="muted" style="margin-top:10px"><i class="ti ti-lock"></i> Dochádzka sa spracuje priamo v tvojom počítači, nikam sa neodosiela.</p>
    </div>`;
  const inp = el('subor'), drop = el('drop');
  inp.addEventListener('change', () => inp.files[0] && spracuj(inp.files[0]));
  ['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => e.dataTransfer.files[0] && spracuj(e.dataTransfer.files[0]));
}

function busy(txt) {
  const b = el('busy');
  if (!b) return;
  b.classList.toggle('hidden', !txt);
  if (txt) el('busy-txt').textContent = txt;
}

async function spracuj(file) {
  sprava('');
  if (!/\.xls[xm]$/i.test(file.name)) { sprava('Súbor musí byť Excel (.xlsx).', 'error'); return; }
  try {
    busy(state.api ? 'Čítam dochádzku…' : 'Pripravujem výpočtový modul (prvýkrát to trvá dlhšie)…');
    const buf = new Uint8Array(await file.arrayBuffer());
    await startPy();
    busy('Čítam dochádzku…');
    state.py.FS.writeFile('/tmp/dochadzka.xlsx', buf);
    state.dochName = file.name;
    state.analyza = JSON.parse(state.api.analyzuj('/tmp/dochadzka.xlsx', JSON.stringify(state.data)));
    const chyba = state.analyza.ludia.filter(p => p.stav !== 'ok');
    if (chyba.length) { renderDoplnit(chyba); return; }
    await generuj();
  } catch (e) {
    console.error(e);
    busy('');
    sprava('Dochádzku sa nepodarilo spracovať: ' + pyChyba(e), 'error');
  }
}

// ---------------------------------------------------------------------
// 2) Noví / neúplní zamestnanci
// ---------------------------------------------------------------------
function unikatnyHarok(meno, zoznam) {
  const used = new Set(zoznam.map(z => z.harok).filter(Boolean));
  const parts = meno.split(/\s+/);
  let t = parts[0] || meno;
  if (used.has(t) && parts[1]) t = `${parts[0]} ${parts[1][0]}.`;
  let base = t, n = 2;
  while (used.has(t)) t = `${base} ${n++}`;
  return t.slice(0, 31);
}

function renderDoplnit(chyba) {
  const zam = Object.fromEntries(state.data.zamestnanci.map(z => [z.meno, z]));
  const karty = chyba.map((p, i) => {
    const z = zam[p.meno] || {};
    return `
    <div class="osoba" data-i="${i}" data-meno="${esc(p.meno)}">
      <h4>${esc(p.meno)} <span class="badge n">${p.stav === 'novy' ? 'nový' : 'chýbajú údaje'}</span></h4>
      <div class="typy">${TYPY.map(t => `<label><input type="radio" name="typ_${i}" value="${esc(t)}" ${z.typ === t ? 'checked' : ''}> ${esc(TYP_POPIS[t])}</label>`).join('')}</div>
      <div class="detail">
        <div class="g4">
          <div><label>Bydlisko</label><input class="byd" value="${esc(z.bydlisko)}" placeholder="napr. Rajec"></div>
          <div><label>Miesto rokovania</label><input class="mie" value="${esc(z.miesto || 'Žilina')}"></div>
          <div><label>Účel cesty</label><input class="uce" value="${esc(z.ucel || 'montážne práce')}"></div>
          <div><label>Súvislé dni SC v ČR</label><select class="czm">${CZ_MODY.map(m => `<option ${z.cz_mode === m ? 'selected' : ''}>${m}</option>`).join('')}</select></div>
        </div>
        <div class="g2">
          <div><label>Text pri odchode</label><input class="to" value="${esc(z.t_odch)}"><div class="hint">dochádza: „Rajec-Žilina“ · týždňovkár: skratka bydliska, napr. „L.Revúce“</div></div>
          <div><label>Text pri príchode</label><input class="tp" value="${esc(z.t_pri)}"><div class="hint">vyplní sa samo podľa bydliska – môžeš prepísať</div></div>
        </div>
      </div>
    </div>`;
  }).join('');
  el('v-spracovat').innerHTML = `
    <h2>Dochádzka ${esc(state.analyza.mesiac_txt)}</h2>
    <p class="sub">Pred výpočtom potrebujem údaje o ${chyba.length === 1 ? 'novom zamestnancovi' : 'nových zamestnancoch'}. Uložia sa natrvalo – nabudúce sa už pýtať nebudem.</p>
    <form id="f-doplnit">${karty}
      <div style="display:flex;gap:10px;align-items:center">
        <button type="submit"><i class="ti ti-calculator"></i> Uložiť a vypočítať diéty</button>
        <button type="button" class="secondary" id="zrusit">Zrušiť</button>
      </div>
    </form>
    <div id="busy" class="busy hidden"><div class="spin"></div><span id="busy-txt"></span></div>
    <p class="muted" style="margin-top:14px">Ostatní ľudia v dochádzke (${state.analyza.ludia.length - chyba.length}) sú v poriadku.</p>`;

  document.querySelectorAll('.osoba').forEach(box => {
    const q = c => box.querySelector(c);
    const touched = { to: !!q('.to').value, tp: !!q('.tp').value };
    q('.to').addEventListener('input', () => touched.to = true);
    q('.tp').addEventListener('input', () => touched.tp = true);
    const typ = () => (box.querySelector('input[type=radio]:checked') || {}).value || '';
    const fill = () => {
      const t = typ(), b = q('.byd').value.trim(), m = q('.mie').value.trim() || 'Žilina';
      q('.detail').classList.toggle('hidden', t === 'bez CP');
      if (!b) return;
      if (t === 'dochádza') { if (!touched.to) q('.to').value = `${b}-${m}`; if (!touched.tp) q('.tp').value = `${m}-${b}`; }
      else if (t.startsWith('týždňovkár')) { if (!touched.to) q('.to').value = b; if (!touched.tp) q('.tp').value = b; }
    };
    box.querySelectorAll('input[type=radio]').forEach(r => r.addEventListener('change', () => { touched.to = touched.tp = false; fill(); }));
    q('.byd').addEventListener('input', fill);
    q('.mie').addEventListener('input', fill);
    fill();
  });
  el('zrusit').addEventListener('click', renderUpload);
  el('f-doplnit').addEventListener('submit', async e => {
    e.preventDefault();
    const nova = structuredClone(state.data);
    const chyby = [];
    document.querySelectorAll('.osoba').forEach(box => {
      const q = c => box.querySelector(c).value.trim();
      const meno = box.dataset.meno;
      const typ = (box.querySelector('input[type=radio]:checked') || {}).value || '';
      if (!typ) { chyby.push(`${meno}: vyber typ`); return; }
      if (typ !== 'bez CP' && !q('.byd')) { chyby.push(`${meno}: doplň bydlisko`); return; }
      let z = nova.zamestnanci.find(x => x.meno === meno);
      if (!z) { z = { meno, harok: unikatnyHarok(meno, nova.zamestnanci), poznamka: '' }; nova.zamestnanci.push(z); }
      Object.assign(z, { typ, bydlisko: q('.byd'), miesto: q('.mie') || 'Žilina', ucel: q('.uce') || 'montážne práce',
                         cz_mode: q('.czm'), t_odch: q('.to'), t_pri: q('.tp') });
      if (!z.harok && typ !== 'bez CP') z.harok = unikatnyHarok(meno, nova.zamestnanci);
    });
    if (chyby.length) { sprava(chyby.join(' · '), 'error'); return; }
    try {
      busy('Ukladám a počítam…');
      await ulozNastavenia(nova);
      await generuj();
    } catch (err) {
      busy('');
      sprava(pyChyba(err), 'error');
    }
  });
}

// ---------------------------------------------------------------------
// 3) Výpočet a výsledok
// ---------------------------------------------------------------------
async function generuj() {
  busy('Počítam diéty a pripravujem súbory…');
  await new Promise(r => setTimeout(r, 30));          // nech sa stihne prekresliť hláška
  const res = JSON.parse(state.api.generuj('/tmp/dochadzka.xlsx', JSON.stringify(state.data)));
  state.urls.forEach(u => URL.revokeObjectURL(u));
  state.urls = [];
  const xlsx = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  res.subory = res.files.map(n => {
    const url = URL.createObjectURL(new Blob([state.py.FS.readFile(res.out_dir + '/' + n)], { type: xlsx }));
    state.urls.push(url);
    return { nazov: n, url };
  });
  const zipPath = state.api.zabal(res.out_dir, `Diety_${String(res.mesiac).padStart(2, '0')}_${res.rok}.zip`);
  const zipUrl = URL.createObjectURL(new Blob([state.py.FS.readFile(zipPath)], { type: 'application/zip' }));
  state.urls.push(zipUrl);
  res.zip = { nazov: zipPath.split('/').pop(), url: zipUrl };
  res.cas = new Date().toLocaleString('sk-SK');
  try { res.archiv = await ulozDoArchivu(res); }
  catch (e) { console.error(e); res.archiv = 'do archívu sa nepodarilo uložiť: ' + (e.message || e); res.archivChyba = true; }
  state.vysledok = res;
  renderVysledok();
}

function renderVysledok() {
  const r = state.vysledok;
  const sr = r.prehlad.reduce((a, x) => a + x[7], 0), cz = r.prehlad.reduce((a, x) => a + x[9], 0);
  const sum = i => r.prehlad.reduce((a, x) => a + (x[i] || 0), 0);
  const nazov = n => n.startsWith('SR') ? 'SR príkaz' : n.startsWith('CR') ? 'ČR príkaz' : 'Kontrolný prehľad';
  el('v-spracovat').innerHTML = `
    <h2>Diéty – ${esc(r.mesiac_txt)}</h2>
    <p class="sub">Spracované ${esc(r.cas)} zo súboru ${esc(state.dochName)} · <span class="${r.archivChyba ? 'err-txt' : ''}"><i class="ti ti-archive"></i> ${esc(r.archiv || '')}</span></p>
    <div class="kpi">
      <div><span class="muted">Ľudí s cestovným príkazom</span><b>${r.prehlad.length}</b></div>
      <div><span class="muted">Diéty SR spolu</span><b>${fmtEur(sr)} €</b></div>
      <div><span class="muted">Diéty ČR spolu</span><b>${fmtNum(cz)} Kč</b></div>
    </div>
    <div class="card">
      <div class="dl">
        ${r.subory.map(s => `<a href="${s.url}" download="${esc(s.nazov)}"><button type="button" class="${s.nazov.startsWith('Kontrola') ? 'secondary' : ''}"><i class="ti ti-download"></i> ${nazov(s.nazov)}</button></a>`).join('')}
        <a href="${r.zip.url}" download="${esc(r.zip.nazov)}"><button type="button" class="secondary"><i class="ti ti-file-zip"></i> Všetko (ZIP)</button></a>
      </div>
      <p class="muted" style="margin-top:10px">Jednodňové cesty do ČR majú predvyplnené časy 7:00–16:30 – prepíš ich podľa GPS, suma v Kč sa v Exceli prepočíta sama.</p>
    </div>
    ${renderKontrola(r.kontrola || [])}
    ${r.warnings.length ? `<div class="banner warn"><b>Upozornenia (${r.warnings.length})</b><ul style="margin:6px 0 0 18px">${r.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul></div>` : ''}
    <div class="card tw">
      <h3>Prehľad po ľuďoch</h3>
      <table>
        <tr><th>Meno</th><th>Typ</th><th class="n">Dni SR</th>${r.sadzby.map(s => `<th class="n">${fmtEur(s)} €</th>`).join('')}<th class="n">SR spolu €</th><th class="n">Dni ČR</th><th class="n">ČR spolu Kč</th></tr>
        ${r.prehlad.map(x => `<tr><td>${esc(x[0])}</td><td><span class="badge ${x[1].startsWith('týž') ? 't' : ''}">${esc(x[1])}</span></td>
          <td class="n">${x[3] || ''}</td><td class="n">${x[4] || ''}</td><td class="n">${x[5] || ''}</td><td class="n">${x[6] || ''}</td>
          <td class="n">${x[7] ? fmtEur(x[7]) : ''}</td><td class="n">${x[8] || ''}</td><td class="n">${x[9] ? fmtNum(x[9]) : ''}</td></tr>`).join('')}
        <tr class="sum"><td>Spolu</td><td></td><td class="n">${sum(3)}</td><td class="n">${sum(4)}</td><td class="n">${sum(5)}</td><td class="n">${sum(6)}</td>
          <td class="n">${fmtEur(sr)}</td><td class="n">${sum(8)}</td><td class="n">${fmtNum(cz)}</td></tr>
      </table>
    </div>
    <div style="display:flex;gap:10px">
      <button class="secondary" id="prepocet"><i class="ti ti-refresh"></i> Prepočítať znova (po zmene zamestnancov alebo sadzieb)</button>
      <button class="secondary" id="iny"><i class="ti ti-upload"></i> Spracovať iný súbor</button>
    </div>
    <div id="busy" class="busy hidden"><div class="spin"></div><span id="busy-txt"></span></div>`;
  el('iny').addEventListener('click', renderUpload);
  el('prepocet').addEventListener('click', async () => {
    try {
      const an = JSON.parse(state.api.analyzuj('/tmp/dochadzka.xlsx', JSON.stringify(state.data)));
      const chyba = an.ludia.filter(p => p.stav !== 'ok');
      if (chyba.length) { state.analyza = an; renderDoplnit(chyba); return; }
      await generuj();
    } catch (e) { busy(''); sprava(pyChyba(e), 'error'); }
  });
}

function renderKontrola(k) {
  if (!k.length) return `<div class="banner ok"><i class="ti ti-circle-check"></i> Kontrola dochádzky: bez zistení.</div>`;
  const chyby = k.filter(x => x.druh === 'chyba').length;
  const skupiny = [];
  for (const x of k) {
    const key = x.meno || 'Celá dochádzka';
    let g = skupiny.find(s => s.meno === key);
    if (!g) { g = { meno: key, polozky: [] }; skupiny.push(g); }
    g.polozky.push(x);
  }
  return `<details class="card kontrola" ${chyby ? 'open' : ''}>
    <summary><b>Kontrola dochádzky</b> – ${chyby ? `<span class="err-txt">${chyby} na preverenie</span>` : ''}${chyby && k.length > chyby ? ', ' : ''}${k.length > chyby ? `${k.length - chyby} informácií` : ''}
      <span class="muted">(výpočet prebehol – ak niečo opravíš v dochádzke, nahraj ju znova)</span></summary>
    <div class="kgrid">${skupiny.map(g => `<div><div class="kmeno">${esc(g.meno)}</div><ul>${g.polozky.map(x =>
      `<li class="${x.druh}"><i class="ti ${x.druh === 'chyba' ? 'ti-alert-triangle' : 'ti-info-circle'}"></i> ${esc(x.text)}</li>`).join('')}</ul></div>`).join('')}</div>
  </details>`;
}

// ---------------------------------------------------------------------
// Archív
// ---------------------------------------------------------------------
const MES = ['jan', 'feb', 'mar', 'apr', 'máj', 'jún', 'júl', 'aug', 'sep', 'okt', 'nov', 'dec'];
const archivStav = { rok: null, druh: 'sr' };

async function renderArchiv() {
  const box = el('v-archiv');
  box.innerHTML = `<h2>Archív</h2><p class="sub">Načítavam…</p>`;
  let rows;
  try { rows = await Archiv.nacitaj(); }
  catch (e) { box.innerHTML = `<h2>Archív</h2><div class="banner error">Archív sa nepodarilo načítať: ${esc(e.message || e)}</div>`; return; }
  const roky = [...new Set(rows.map(x => x.rok))].sort((a, b) => b - a);
  if (!archivStav.rok || !roky.includes(archivStav.rok)) archivStav.rok = roky[0] || new Date().getFullYear();
  const rok = archivStav.rok, druh = archivStav.druh;
  const mesiace = Array(12).fill(null);
  rows.filter(x => x.rok === rok).forEach(x => mesiace[x.mesiac - 1] = x.data[druh] || null);
  const mena = [...new Set(mesiace.flatMap(m => m ? Object.keys(m.ludia) : []))].sort((a, b) => a.localeCompare(b, 'sk'));
  const jed = druh === 'sr' ? '€' : 'Kč', f = druh === 'sr' ? fmtEur : fmtNum;
  const sumM = mesiace.map(m => m ? Object.values(m.ludia).reduce((a, x) => a + x.spolu, 0) : null);
  const tabulka = mena.length ? `
    <div class="card tw"><table class="arch">
      <tr><th>Meno</th>${MES.map((m, i) => `<th class="n">${m}${mesiace[i] ? `<span class="zdroj ${mesiace[i].zdroj === 'finálny súbor' ? 'fin' : ''}" title="${esc(mesiace[i].zdroj)}">${mesiace[i].zdroj === 'finálny súbor' ? 'F' : 'V'}</span>` : ''}</th>`).join('')}<th class="n">Spolu ${jed}</th></tr>
      ${mena.map(n => {
        const v = mesiace.map(m => m && m.ludia[n] ? m.ludia[n].spolu : null);
        return `<tr><td>${esc(n)}</td>${v.map(x => `<td class="n">${x === null ? '' : f(x)}</td>`).join('')}<td class="n"><b>${f(v.reduce((a, x) => a + (x || 0), 0))}</b></td></tr>`;
      }).join('')}
      <tr class="sum"><td>Spolu</td>${sumM.map(x => `<td class="n">${x === null ? '' : f(x)}</td>`).join('')}<td class="n">${f(sumM.reduce((a, x) => a + (x || 0), 0))}</td></tr>
    </table>
    <p class="muted" style="margin-top:8px"><span class="zdroj">V</span> = podľa výpočtu z dochádzky &nbsp; <span class="zdroj fin">F</span> = podľa finálneho súboru (aj s tvojimi ručnými úpravami, napr. GPS časy)</p></div>`
    : `<div class="card muted">Za rok ${rok} zatiaľ nie je v archíve nič. Mesiac sa uloží automaticky pri každom spracovaní dochádzky.</div>`;
  box.innerHTML = `
    <h2>Archív</h2>
    <p class="sub">Súčty diét po ľuďoch za jednotlivé mesiace.</p>
    <div style="display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap">
      <select id="a-rok" style="width:auto">${(roky.length ? roky : [rok]).map(r => `<option ${r === rok ? 'selected' : ''}>${r}</option>`).join('')}</select>
      <div class="seg"><button class="${druh === 'sr' ? 'on' : ''}" data-d="sr">Diéty SR (€)</button><button class="${druh === 'cz' ? 'on' : ''}" data-d="cz">Diéty ČR (Kč)</button></div>
    </div>
    ${tabulka}
    <div class="card">
      <h3>Aktualizovať z finálnych súborov</h3>
      <p class="muted" style="margin-bottom:10px">Keď v stiahnutom SR alebo ČR príkaze niečo upravíš (napr. GPS časy) a uložíš ho v Exceli, nahraj ho sem – archív sa prepíše skutočnými sumami. Môžeš nahrať oba súbory naraz.</p>
      <label class="drop" id="a-drop" style="padding:22px"><i class="ti ti-file-upload" style="font-size:30px"></i><b style="font-size:14px">Pretiahni sem SR / ČR príkaz</b><input type="file" id="a-subor" accept=".xlsx" multiple hidden></label>
      <div id="busy" class="busy hidden"><div class="spin"></div><span id="busy-txt"></span></div>
    </div>`;
  el('a-rok').addEventListener('change', e => { archivStav.rok = +e.target.value; renderArchiv(); });
  box.querySelectorAll('.seg button').forEach(b => b.addEventListener('click', () => { archivStav.druh = b.dataset.d; renderArchiv(); }));
  const inp = el('a-subor'), drop = el('a-drop');
  inp.addEventListener('change', () => inp.files.length && nahrajFinalne([...inp.files]));
  ['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => e.dataTransfer.files.length && nahrajFinalne([...e.dataTransfer.files]));
}

async function nahrajFinalne(files) {
  sprava('');
  try {
    busy('Pripravujem výpočtový modul…');
    await startPy();
    const rows = await Archiv.nacitaj();
    const hotove = [];
    for (const file of files) {
      busy('Čítam ' + file.name + '…');
      state.py.FS.writeFile('/tmp/finalny.xlsx', new Uint8Array(await file.arrayBuffer()));
      let r;
      try { r = JSON.parse(state.api.nacitaj_suhrn('/tmp/finalny.xlsx')); }
      catch (e) { throw new Error(file.name + ': ' + pyChyba(e)); }
      const druh = r.druh === 'CZ' ? 'cz' : 'sr';
      let row = rows.find(x => x.rok === r.rok && x.mesiac === r.mesiac);
      if (!row) { row = { rok: r.rok, mesiac: r.mesiac, data: {} }; rows.push(row); }
      row.data[druh] = { zdroj: 'finálny súbor', subor: file.name, cas: new Date().toISOString(),
                         ludia: Object.fromEntries(r.ludia.map(x => [x.meno, { spolu: x.spolu, dni: x.dni }])) };
      await Archiv.uloz(r.rok, r.mesiac, row.data);
      hotove.push(`${druh === 'sr' ? 'SR' : 'ČR'} ${MES[r.mesiac - 1]} ${r.rok}`);
      archivStav.rok = r.rok; archivStav.druh = druh;
    }
    await renderArchiv();
    sprava('Archív aktualizovaný z finálnych súborov: ' + hotove.join(', ') + '.');
  } catch (e) {
    busy('');
    sprava(e.message || String(e), 'error');
  }
}

// ---------------------------------------------------------------------
// Zamestnanci
// ---------------------------------------------------------------------
function riadokZam(z, i) {
  const sel = (cls, opts, val, empty) => `<select class="${cls}">${empty && !opts.includes(val) ? '<option value="" selected>– vyber –</option>' : ''}${opts.map(o => `<option ${o === val ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
  const inp = (cls, val, w) => `<input class="${cls}" value="${esc(val)}" style="min-width:${w}px">`;
  return `<tr data-poz="${esc(z.poznamka || '')}">
    <td>${inp('meno', z.meno, 170)}</td><td>${sel('typ', TYPY, z.typ, true)}</td><td>${inp('harok', z.harok, 90)}</td>
    <td>${inp('bydlisko', z.bydlisko, 130)}</td><td>${inp('t_odch', z.t_odch, 150)}</td><td>${inp('t_pri', z.t_pri, 150)}</td>
    <td>${inp('miesto', z.miesto, 80)}</td><td>${inp('ucel', z.ucel, 120)}</td><td>${sel('cz_mode', CZ_MODY, z.cz_mode || CZ_MODY[0])}</td>
    <td><input type="checkbox" class="zmazat" title="Odstrániť" style="width:auto"></td></tr>`;
}

function renderZamestnanci() {
  el('v-zamestnanci').innerHTML = `
    <h2>Zamestnanci</h2>
    <p class="sub">Nastavenie pre každého človeka z dochádzky. Meno musí byť presne ako v dochádzke.</p>
    <div class="card tw">
      <table id="t-zam">
        <tr><th>Meno (ako v dochádzke)</th><th>Typ</th><th>Hárok</th><th>Bydlisko</th><th>Text pri odchode</th><th>Text pri príchode</th><th>Miesto</th><th>Účel cesty</th><th>Súvislé dni SC v ČR</th><th title="Odstrániť">✕</th></tr>
        ${state.data.zamestnanci.map(riadokZam).join('')}
      </table>
      <button class="secondary sm" id="pridat-zam" style="margin-top:10px"><i class="ti ti-plus"></i> Pridať riadok</button>
    </div>
    <div class="card legend">
      <div><b>dochádza</b> – denne domov: 7:00–15:30 + nadčas</div>
      <div><b>týždňovkár</b> – 5:00–23:01, posledný deň pred voľnom 5:00–19:00</div>
      <div><b>týždňovkár nonstop</b> – každý deň 0:00–23:59</div>
      <div><b>bez CP</b> – nerobí sa cestovný príkaz (kancelária)</div>
      <div><b>jedna viacdňová cesta</b> – dni SC za sebou = zostáva v ČR cez noc</div>
      <div><b>samostatné jednodňové cesty</b> – každý deň SC tam a späť</div>
    </div>
    <div class="savebar"><button id="ulozit-zam"><i class="ti ti-device-floppy"></i> Uložiť zmeny</button><span class="muted" id="zam-info"></span></div>`;
  const tab = el('t-zam');
  tab.addEventListener('change', e => { if (e.target.classList.contains('zmazat')) e.target.closest('tr').classList.toggle('del', e.target.checked); });
  el('pridat-zam').addEventListener('click', () => {
    tab.insertAdjacentHTML('beforeend', riadokZam({ miesto: 'Žilina', ucel: 'montážne práce', cz_mode: CZ_MODY[0] }));
    tab.querySelector('tr:last-child input').focus();
  });
  el('ulozit-zam').addEventListener('click', async () => {
    const zoznam = [], chyby = [], mena = new Set();
    tab.querySelectorAll('tr').forEach((tr, i) => {
      if (i === 0 || tr.querySelector('.zmazat').checked) return;
      const v = c => tr.querySelector('.' + c).value.trim();
      const z = { meno: v('meno'), typ: v('typ'), harok: v('harok'), bydlisko: v('bydlisko'), t_odch: v('t_odch'), t_pri: v('t_pri'),
                  miesto: v('miesto'), ucel: v('ucel'), cz_mode: v('cz_mode'), poznamka: tr.dataset.poz || '' };
      if (!z.meno) return;
      if (mena.has(z.meno)) chyby.push(`${z.meno} je v zozname dvakrát`);
      mena.add(z.meno);
      if (!TYPY.includes(z.typ)) chyby.push(`${z.meno}: vyber typ`);
      if (!z.harok && z.typ !== 'bez CP') z.harok = unikatnyHarok(z.meno, zoznam.concat(state.data.zamestnanci));
      zoznam.push(z);
    });
    if (chyby.length) { sprava(chyby.join(' · '), 'error'); return; }
    try {
      await ulozNastavenia({ ...state.data, zamestnanci: zoznam });
      renderZamestnanci();
      sprava('Zoznam zamestnancov uložený.');
    } catch (e) { sprava(e.message, 'error'); }
  });
}

// ---------------------------------------------------------------------
// Sadzby a časy
// ---------------------------------------------------------------------
function riadokSadzba(s) {
  const n = (cls, v, dec) => `<td><input class="${cls}" value="${v === undefined || v === '' ? '' : (dec ? fmtEur(v) : fmtNum(v))}" style="text-align:right"></td>`;
  return `<tr><td><input type="date" class="od" value="${esc(s.od || '')}"></td>${n('a', s.a, 1)}${n('b', s.b, 1)}${n('c', s.c, 1)}${n('cz', s.cz ?? 600, 0)}
    <td><input class="zdroj" value="${esc(s.zdroj || '')}"></td><td><input type="checkbox" class="zmazat" style="width:auto"></td></tr>`;
}

function renderSadzby() {
  const d = state.data;
  el('v-sadzby').innerHTML = `
    <h2>Sadzby a časy</h2>
    <p class="sub">Pre každý deň sa použije sadzba platná k tomu dňu. Pri zmene zákonných súm pridaj nový riadok.</p>
    <div class="card">
      <h3>Sadzby stravného</h3>
      <table id="t-sad">
        <tr><th>Platné od</th><th class="n">SR 5–12 h (€)</th><th class="n">SR nad 12–18 h (€)</th><th class="n">SR nad 18 h (€)</th><th class="n">ČR plná sadzba (Kč)</th><th>Zdroj</th><th>✕</th></tr>
        ${[...d.sadzby].sort((a, b) => a.od.localeCompare(b.od)).map(riadokSadzba).join('')}
      </table>
      <button class="secondary sm" id="pridat-sad" style="margin-top:10px"><i class="ti ti-plus"></i> Nová sadzba</button>
      <p class="muted" style="margin-top:10px">ČR: do 6 h = 25 %, nad 6 do 12 h = 50 %, nad 12 h = 100 %. Viacdňová cesta = každý deň 100 %.</p>
    </div>
    <div class="card">
      <h3>Štandardné časy a údaje</h3>
      <table id="t-par">
        ${d.parametre.map((p, i) => `<tr><td style="width:36%">${esc(p.kluc)}</td>
          <td style="width:200px"><input data-i="${i}" value="${esc(p.hodnota)}" ${p.cas ? 'placeholder="hh:mm" style="width:90px"' : ''}></td>
          <td class="muted">${esc(p.popis || '')}</td></tr>`).join('')}
      </table>
    </div>
    <div class="savebar"><button id="ulozit-sad"><i class="ti ti-device-floppy"></i> Uložiť</button></div>`;
  el('pridat-sad').addEventListener('click', () => el('t-sad').insertAdjacentHTML('beforeend', riadokSadzba({})));
  el('ulozit-sad').addEventListener('click', async () => {
    const chyby = [], sadzby = [];
    const num = v => parseFloat(String(v).replace(/\s| /g, '').replace(',', '.'));
    el('t-sad').querySelectorAll('tr').forEach((tr, i) => {
      if (i === 0 || tr.querySelector('.zmazat').checked) return;
      const g = c => tr.querySelector('.' + c).value.trim();
      if (!g('od')) return;
      const s = { od: g('od'), a: num(g('a')), b: num(g('b')), c: num(g('c')), cz: num(g('cz')), zdroj: g('zdroj') };
      if ([s.a, s.b, s.c, s.cz].some(x => !isFinite(x) || x <= 0)) chyby.push(`Sadzba od ${s.od}: skontroluj sumy`);
      sadzby.push(s);
    });
    if (!sadzby.length) chyby.push('Musí zostať aspoň jedna sadzba.');
    const parametre = structuredClone(d.parametre);
    el('t-par').querySelectorAll('input').forEach(inp => {
      const p = parametre[+inp.dataset.i];
      const v = inp.value.trim();
      if (p.cas) {
        const m = v.match(/^(\d{1,2})[:.](\d{2})$/);
        if (!m || +m[1] > 23 || +m[2] > 59) { chyby.push(`${p.kluc}: „${v}“ nie je čas (napr. 07:00)`); return; }
        p.hodnota = `${m[1].padStart(2, '0')}:${m[2]}`;
      } else p.hodnota = v;
    });
    if (chyby.length) { sprava(chyby.join(' · '), 'error'); return; }
    try {
      await ulozNastavenia({ ...d, sadzby, parametre });
      renderSadzby();
      sprava('Sadzby a časy uložené.');
    } catch (e) { sprava(e.message, 'error'); }
  });
}

// ---------------------------------------------------------------------
// Štart
// ---------------------------------------------------------------------
(async () => {
  if (LOCAL_MODE) { poPrihlaseni(null); return; }
  const { data } = await sb.auth.getSession();
  if (data.session) poPrihlaseni(data.session.user);
})();
