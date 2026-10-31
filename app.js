(function () {
  const $ = (id) => document.getElementById(id);
  let tplBuf = null, tplName = 'slides', parsed = null;
  const SAMPLE = [
    '# Format: T|Type heading   ba   Q|question [tag]|A|B|C|D|answer',
    'T|Type-01: Sample heading',
    'Q|Bangladesher rajdhani kothay? [Sample 01]|Dhaka|Chattogram|Khulna|Sylhet|A',
    'Q|2 + 3 = ? [Sample 02]|4|5|6|7|B',
    'Q|Pani-r rasayonik sutro kon ti? [Sample 03]|CO2|H2O|NaCl|O2|B'
  ].join('\n');

  async function loadTemplate(buf, name) {
    try {
      const info = await MCQ.inspect(buf);
      tplBuf = buf; tplName = name.replace(/\.pptx$/i, '');
      $('qNo').value = info.questionSlide || ''; $('tNo').value = info.typeSlide || '';
      const ok = !!info.questionSlide;
      $('tplInfo').className = ok ? 'info' : 'info err';
      $('tplInfo').textContent = ok
        ? `${name}: ${info.count}ta slide. Question sample slide #${info.questionSlide}, type heading sample slide ${info.typeSlide ? '#' + info.typeSlide : '(paoa jayni)'}.`
        : `${name}: question sample slide paoa jayni. Slide e "(A) … (B) …" ar "Ans:" likha thakte hobe, nile nicher box e slide no. nijei de.`;
    } catch (e) { tplBuf = null; $('tplInfo').className = 'info err'; $('tplInfo').textContent = 'Template porte parini: ' + e.message; }
    updateGen();
  }
  function updateGen() { $('genBtn').disabled = !(tplBuf && parsed && parsed.items.length && !parsed.errors.length); }

  $('tplFile').onchange = async (e) => { const f = e.target.files[0]; if (f) { $('tplBuiltin').value = ''; await loadTemplate(await f.arrayBuffer(), f.name); } };
  $('tplBuiltin').onchange = async (e) => {
    if (!e.target.value) return;
    const url = new URL(e.target.value, location.href).href;
    const err = $('tplInfo'); err.className = 'info err';
    if (location.protocol === 'file:') { err.textContent = 'Tui file:// diye khulechis. GitHub Pages URL (https://…github.io/…) diye khul.\nChecked URL: ' + url; return; }
    try {
      const r = await fetch(url);
      if (!r.ok) { err.textContent = `Built-in template paoa jayni. HTTP ${r.status}\nChecked URL: ${url}\n` + (r.status === 404 ? 'Mane file ta repo te oi path e nai. Repo-r root e templates/demo-rtd.pptx ache kina dekh.' : ''); return; }
      await loadTemplate(await r.arrayBuffer(), url.split('/').pop());
    } catch (ex) { err.textContent = 'Network error: ' + ex.message + '\nChecked URL: ' + url; }
  };
  $('sampleBtn').onclick = () => { $('mcqText').value = SAMPLE; };
  $('checkBtn').onclick = () => {
    parsed = MCQ.parseItems($('mcqText').value);
    const q = parsed.items.filter((i) => i.kind === 'Q').length, t = parsed.items.length - q;
    const pi = $('parseInfo');
    if (parsed.errors.length) { pi.className = 'info err'; pi.textContent = parsed.errors.slice(0, 12).join('\n') + (parsed.errors.length > 12 ? `\n… aro ${parsed.errors.length - 12}ta` : ''); }
    else { pi.className = 'info'; pi.textContent = `${q}ta question, ${t}ta type heading. Total ${parsed.items.length}ta slide hobe.`; }
    const tb = document.createElement('table');
    parsed.items.forEach((it, i) => {
      const tr = tb.insertRow();
      if (it.kind === 'T') { tr.className = 'type'; const c = tr.insertCell(); c.colSpan = 3; c.textContent = it.title; }
      else { tr.insertCell().textContent = i + 1; tr.insertCell().textContent = it.q; tr.insertCell().textContent = 'Ans: ' + it.ans; }
    });
    $('preview').replaceChildren(tb);
    updateGen();
  };
  $('genBtn').onclick = async () => {
    const gi = $('genInfo'); gi.className = 'info'; gi.textContent = 'Banachhi…'; $('genBtn').disabled = true;
    try {
      const blob = await MCQ.generate(tplBuf, parsed.items, { questionSlide: +$('qNo').value || null, typeSlide: +$('tNo').value || null, type: 'blob' });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = tplName + '_MCQ.pptx'; document.body.appendChild(a); a.click(); a.remove();
      gi.textContent = 'Hoye gese! Download shuru hoyeche.';
    } catch (e) { gi.className = 'info err'; gi.textContent = 'Error: ' + e.message; }
    updateGen();
  };
})();
