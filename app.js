const $ = id => document.getElementById(id);
const load = (k, d) => { try { return JSON.parse(localStorage.getItem(k)) ?? d; } catch { return d; } };
const save = (k, v) => localStorage.setItem(k, JSON.stringify(v));

let entries = load('entries', []);
let running = load('running', null); // start timestamp (ms) or null
let rate = Number(load('rate', 11.75)); // pay per hour
let pending = null; // {start,end} awaiting a note

const pad = n => String(n).padStart(2, '0');
const dayKey = ms => { const d = new Date(ms); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const hm = mins => `${Math.floor(mins / 60)}h ${pad(mins % 60)}m`;
const clock = ms => { const s = Math.floor(ms / 1000); return `${Math.floor(s / 3600)}:${pad(Math.floor(s / 60) % 60)}:${pad(s % 60)}`; };
const timeOf = ms => new Date(ms).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });

// ---- stopwatch ----
function render() {
  $('startStop').textContent = running ? 'Stop' : 'Start';
  $('startStop').classList.toggle('running', !!running);
  tick();
  renderList();
}
function tick() { $('timer').textContent = clock(running ? Date.now() - running : 0); }
setInterval(tick, 500);

$('startStop').onclick = () => {
  if (!running) {
    running = Date.now();
    save('running', running);
  } else {
    pending = { start: running, end: Date.now() };
    running = null;
    save('running', null);
    askNote(`${clock(pending.end - pending.start)} worked`);
  }
  render();
};

function askNote(info) {
  $('dlgTime').textContent = info;
  $('note').value = '';
  $('dlg').showModal();
  $('note').focus();
}
$('form').onsubmit = () => {
  const note = $('note').value.trim();
  if (pending && note) {
    entries.push({ id: Date.now(), start: pending.start, end: pending.end, note });
    save('entries', entries);
  }
  pending = null;
  $('month').value = $('month').value || dayKey(Date.now()).slice(0, 7);
  render();
};
$('discard').onclick = () => {
  if (confirm('Discard this time?')) { pending = null; $('dlg').close(); render(); }
};
$('dlg').addEventListener('cancel', e => e.preventDefault()); // must save or discard

// ---- month data ----
function monthDays(ym) {
  const days = {};
  for (const e of entries) {
    const k = dayKey(e.start);
    if (!k.startsWith(ym)) continue;
    (days[k] ||= { ms: 0, notes: [], items: [] });
    days[k].ms += e.end - e.start;
    days[k].items.push(e);
    if (!days[k].notes.includes(e.note)) days[k].notes.push(e.note);
  }
  return Object.keys(days).sort().map(k => ({ date: k, notes: days[k].notes, mins: Math.round(days[k].ms / 60000), items: days[k].items }));
}

function renderList() {
  const ym = $('month').value;
  const days = monthDays(ym).reverse();
  const total = days.reduce((s, d) => s + d.mins, 0);
  $('summary').textContent = days.length ? `${hm(total)} · $${(total / 60 * rate).toFixed(2)}` : '';
  const list = $('list');
  list.innerHTML = '';
  if (!days.length) { list.innerHTML = '<p style="color:var(--muted)">No work logged this month.</p>'; return; }
  for (const d of days) {
    const h = document.createElement('div');
    h.className = 'day';
    h.textContent = `${d.date} — ${hm(d.mins)}`;
    list.append(h);
    for (const e of d.items.sort((a, b) => b.start - a.start)) {
      const row = document.createElement('div');
      row.className = 'entry';
      const text = document.createElement('div');
      const name = document.createElement('span');
      name.textContent = e.note;
      const sub = document.createElement('small');
      sub.textContent = `${timeOf(e.start)}–${timeOf(e.end)} · ${hm(Math.round((e.end - e.start) / 60000))}`;
      text.append(name, sub);
      const btns = document.createElement('div');
      btns.className = 'row';
      const edit = document.createElement('button');
      edit.className = 'small'; edit.textContent = 'Edit';
      edit.onclick = () => {
        const n = prompt('What did you do?', e.note);
        if (n && n.trim()) { e.note = n.trim(); save('entries', entries); render(); }
      };
      const del = document.createElement('button');
      del.className = 'small'; del.textContent = 'Delete';
      del.onclick = () => {
        if (confirm('Delete this entry?')) { entries = entries.filter(x => x.id !== e.id); save('entries', entries); render(); }
      };
      btns.append(edit, del);
      row.append(text, btns);
      list.append(row);
    }
  }
}
$('month').onchange = () => { $('msg').textContent = ''; renderList(); };

// ---- export ----
function buildWorkbook(ym) {
  const days = monthDays(ym);
  const rows = [['Date', 'What I did', 'Time worked', 'Decimal hours']];
  days.forEach(d => rows.push([d.date, d.notes.join(', '), hm(d.mins), d.mins / 60]));
  const first = 2, last = rows.length;
  const totalMins = days.reduce((s, d) => s + d.mins, 0);
  rows.push([]);
  rows.push(['', '', 'Total hours', null]);
  rows.push(['', '', 'Pay rate', rate]);
  rows.push(['', '', 'Total pay', null]);
  const hoursRow = last + 2, rateRow = last + 3, payRow = last + 4;
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws[`D${hoursRow}`] = { t: 'n', f: `SUM(D${first}:D${last})`, v: totalMins / 60, z: '0.00' };
  ws[`D${rateRow}`] = { t: 'n', v: rate, z: '"$"0.00' };
  ws[`D${payRow}`] = { t: 'n', f: `D${hoursRow}*D${rateRow}`, v: totalMins / 60 * rate, z: '"$"0.00' };
  for (let r = first; r <= last; r++) if (ws[`D${r}`]) ws[`D${r}`].z = '0.00';
  ws['!ref'] = `A1:D${payRow}`;
  ws['!cols'] = [{ wch: 12 }, { wch: 40 }, { wch: 14 }, { wch: 14 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, ym);
  return wb;
}

$('export').onclick = () => {
  const ym = $('month').value;
  if (!ym) { $('msg').textContent = 'Pick a month first.'; return; }
  if (!monthDays(ym).length) { $('msg').textContent = `No work logged in ${ym}.`; return; }
  XLSX.writeFile(buildWorkbook(ym), `Invoice-${ym}.xlsx`);
  $('msg').textContent = `Saved Invoice-${ym}.xlsx to Downloads.`;
};

$('rate').onchange = () => {
  const v = parseFloat($('rate').value);
  if (!(v >= 0)) { $('rate').value = rate; return; }
  rate = v;
  save('rate', rate);
  renderList();
};

// ---- init ----
$('rate').value = rate;
$('month').value = dayKey(Date.now()).slice(0, 7);
if (running === null && !pending) { /* idle */ }
render();
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
