/* Этап 22 (v22.0): реальные числа обоснованности демо-портфеля для презентации/сценария.
   Запуск: node scripts/_stats22.js -> scripts/data/demo_stats_v22.json */
const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');
const html = fs.readFileSync(path.join(__dirname, '..', 'output', 'versions', 'Аудит_НИОКР_модель_v22.0.html'), 'utf-8');
const vc = new VirtualConsole(); const errs = [];
vc.on('jsdomError', e => errs.push(String((e && e.message) || e)));
const dom = new JSDOM(html, { runScripts: 'dangerously', pretendToBeVisual: true, virtualConsole: vc, url: 'http://localhost/' });
const w = dom.window, api = w.niokrAPI;
const r4 = v => v == null ? null : Math.round(v * 10000) / 10000;
setTimeout(() => {
  api.demo('port');
  const projects = api.projects().map(p => {
    const J = p.J, R = p.R;
    return {
      name: p.obj.name, shifr: p.obj.shifr, IA: r4(R.IA), IC: r4(R.IC), zA: R.zA, zC: R.zC, rec: R.rec,
      n: J.n, okCount: J.okCount, verdict: J.verdict, short: J.short,
      W: r4(J.W.all.W), chi2: r4(J.W.all.chi2), sig: J.W.all.sig, mUsed: J.W.all.mUsed,
      WA: r4(J.W.A.W), WC: r4(J.W.C.W),
      ciA: J.ciA ? { lo: r4(J.ciA.lo), hi: r4(J.ciA.hi), se: r4(J.ciA.se), stable: J.ciA.stable } : null,
      ciC: J.ciC ? { lo: r4(J.ciC.lo), hi: r4(J.ciC.hi), se: r4(J.ciC.se), stable: J.ciC.stable } : null,
      jackChanges: J.jack.changes, jackN: J.jack.n, maxDIA: r4(J.jack.maxDIA), maxDIC: r4(J.jack.maxDIC),
      gates: J.gates.map(g => ({ id: g.id, ok: g.ok }))
    };
  });
  api.setModes(false, true);
  const ir = api.projects().map(p => ({ shifr: p.obj.shifr, IR: r4(p.R.IR), zR: p.R.zR }));
  api.setModes(false, false);
  const out = { version: '22.0', date: '2026-11-11', nProjects: projects.length, projects, ir, windowErrors: errs.length };
  fs.writeFileSync(path.join(__dirname, 'data', 'demo_stats_v22.json'), JSON.stringify(out, null, 1));
  console.log(JSON.stringify(out, null, 1));
  process.exit(0);
}, 400);