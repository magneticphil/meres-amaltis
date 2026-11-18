// Run with: node --test round-selection.test.cjs
// Exercise the real form handlers, queued payload and local lot attachment without network writes.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
const section = (a, b) => script.slice(script.indexOf(a), script.indexOf(b));

function setup(cycles = []) {
  const elements = new Map(), storage = new Map();
  const classes = () => ({ toggle() {}, add() {}, remove() {} });
  const element = (dataset = {}) => ({ value:'', dataset, classList:classes(),
    addEventListener(name, fn) { this[name] = fn; }, closest() { return this; } });
  const rooms = [3, 4, 5, 6, 7].map(n => element({ ps:'Salle ' + n }));
  const mothers = [element({ pn:'202601010001_A' }), element({ pn:'202601010002_A' })];
  let rounds = [];
  const get = id => {
    if (!elements.has(id)) {
      const e = element();
      Object.defineProperty(e, 'innerHTML', { set(value) {
        if (id === 'p-rounds') rounds = [...value.matchAll(/<button\b([^>]*)>/g)].map(m => {
          const b = element();
          for (const a of m[1].matchAll(/data-(pr|rn)="([^"]*)"/g)) b.dataset[a[1]] = a[2];
          return b;
        });
      } });
      elements.set(id, e);
    }
    return elements.get(id);
  };
  const ctx = vm.createContext({ console, navigator:{}, location:{ hostname:'test.invalid' }, window:{},
    document:{ getElementById:get, querySelectorAll(selector) {
      if (selector.startsWith('#p-salles')) return rooms;
      if (selector.startsWith('#p-rounds')) return rounds;
      if (selector === '[data-pn]') return mothers;
      return [];
    } }, localStorage:{ getItem:k => storage.get(k) || null, setItem:(k, v) => storage.set(k, v), removeItem:k => storage.delete(k) },
    setTimeout:() => 0, clearTimeout() {},
    ouvrir() {}, tete:() => '', pied:() => '', fermer() {}, toast() {}, scheduleFlush() {}, recalc() {},
    soumettre:async (_, fn) => fn(), nonVide(v, message) { assert.ok(String(v).trim(), message); }, pasFutur() {} });
  vm.runInContext(section('/* Adresse du script', '/* ---------- Connexion'), ctx);
  vm.runInContext(section('const qKey', 'const LIBELLE'), ctx);
  vm.runInContext(section('/* Prise de boutures', '/* Destruction */'), ctx);
  vm.runInContext("const normLot = s => String(s || '').replace(/\\s+/g, '').toUpperCase();", ctx);
  ctx.data = { moi:'Test', today:'2026-11-18', salles:rooms.map(x => x.dataset.ps), cycles,
    meres:mothers.map((m, i) => ({ numero:m.dataset.pn, statut:'active', cultivar:'Test', surnom:'M' + i, creation:'2026-01-01' })),
    lots:[], boutures:[], destructions:[] };
  vm.runInContext("S.auth = { nom:'Test' }; S.d = data; formPrise();", ctx);
  get('p-date').value = '2026-11-18'; mothers[0].value = '12'; mothers[1].value = '8';
  return { get, ctx, rounds:() => rounds,
    room(n) { get('p-salles').onclick({ target:rooms.find(x => x.dataset.ps === 'Salle ' + n) }); },
    round(n) { get('p-round').value = String(n); get('p-round').input(); },
    lot(value) { get('p-lot').value = value; get('p-lot').input(); },
    chip(predicate) { const b = rounds.find(predicate); assert.ok(b, 'Round button exists'); get('p-rounds').onclick({ target:b }); },
    async save() { await get('p-ok').onclick(); return JSON.parse(vm.runInContext('JSON.stringify(qGet())', ctx)); },
    projected() { return JSON.parse(vm.runInContext('JSON.stringify(superposer(data))', ctx)); } };
}
const cycle = (salle, n) => ({ id:'s' + salle + 'r' + n, salle:'Salle ' + salle, round:'round ' + n, debut:'2026-' + String(n).padStart(2, '0') + '-01' });

test('uncatalogued rounds persist for every destination, including S5R15', async () => {
  for (const salle of [3, 4, 5, 6, 7]) for (const round of [15, 16, 30, 101]) {
    const f = setup([cycle(salle, 12), cycle(salle, 13), cycle(salle, 14)]);
    f.room(salle); f.round(round);
    assert.equal(f.get('p-lot').value, `S${salle}R${round}`);
    const q = await f.save(), p = q[0].args[0];
    assert.equal(p.round, 'round ' + round); assert.equal(p.cycle, '');
    assert.equal(p.lignes.reduce((n, x) => n + x.nb, 0), 20);
    const d = f.projected();
    assert.equal(d.lots[0].round, p.round); assert.equal(d.lots[0].salle, p.salle);
    assert.equal(d.boutures.length, 2);
  }
});
test('older calendar rounds remain visible and save their cycle', async () => {
  const f = setup([1, 2, 3, 4, 5].map(n => cycle(5, n)));
  f.room(5); f.round(1);
  assert.ok(f.rounds().some(x => x.dataset.pr === 's5r1'));
  const p = (await f.save())[0].args[0];
  assert.equal(p.round, 'round 1'); assert.equal(p.cycle, 's5r1');
});
test('typing a lot synchronizes destination and explicit selections update the lot', async () => {
  const f = setup([cycle(5, 14), cycle(6, 4)]);
  f.lot('s5r15'); assert.equal(f.get('p-round').value, 15);
  f.chip(x => x.dataset.pr === 's5r14'); assert.equal(f.get('p-lot').value, 'S5R14');
  f.room(6); f.round(4); assert.equal(f.get('p-lot').value, 'S6R4');
  const p = (await f.save())[0].args[0]; assert.equal(p.cycle, 's6r4');
});
test('next-round shortcut and unspecified custom lots still work', async () => {
  const f = setup([cycle(5, 14)]); f.room(5); f.chip(x => x.dataset.rn === '15');
  assert.equal((await f.save())[0].args[0].round, 'round 15');
  const g = setup(); g.room(3); g.lot('LOT-ESSAI');
  assert.equal((await g.save())[0].args[0].round, '');
});
test('invalid round numbers and inconsistent structured lots cannot be queued', async () => {
  for (const n of [-1, 0, 1.5, 1e20]) {
    const f = setup(); f.room(5); f.lot('TEST'); f.round(n); f.lot('TEST');
    await assert.rejects(f.save(), /entier positif/);
  }
  const f = setup(); f.room(5); f.round(15); f.lot('S9R15');
  await assert.rejects(f.save(), /ne correspond pas/);
});
