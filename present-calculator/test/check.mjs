/**
 * 動作確認用スクリプト（Node.js）
 *
 *   node test/check.mjs <パフォーマンススケジュール.xlsx> [入場者プレゼント情報.xlsx]
 *
 * index.html 内のロジックをそのまま読み込んで、
 * ・スケジュールの読み取り（日付／スクリーン／キャパ／上映回数）
 * ・入場者プレゼント表の読み取り
 * ・満席想定での配布シミュレーション
 * が正しく動くかを確認します。
 */
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const html = fs.readFileSync(path.join(here, "..", "index.html"), "utf8");
const script = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]).join("\n");
vm.runInThisContext(script, { filename: "index.html#script" });
const P = globalThis.__PGC;

const psPath = process.argv[2];
const giftPath = process.argv[3];
if (!psPath) {
  console.error("使い方: node test/check.mjs <PS.xlsx> [プレゼント情報.xlsx]");
  process.exit(1);
}

function buf(p) {
  const b = fs.readFileSync(p);
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
}

let failed = 0;
function ok(cond, msg) {
  console.log((cond ? "  OK   " : "  NG   ") + msg);
  if (!cond) failed++;
}

console.log("== スケジュール読み込み ==");
const sc = P.parseSchedule(P.readXlsx(buf(psPath)));
console.log(`  劇場: ${sc.theater} / 期間: ${sc.start} 〜 ${sc.end} (${sc.dates.length}日)`);
console.log(`  上映回数: ${sc.perfs.length} / 総座席: ${sc.perfs.reduce((a, p) => a + p.cap, 0)} / 作品数: ${sc.titles.length}`);
console.log("  シート: " + sc.sheets.map(s => `${s.name}[${s.dates.length}日/${s.count}回]`).join(" "));
ok(sc.perfs.length > 0, "上映回を取得できた");
ok(sc.dates.length > 0, "日付を取得できた");
ok(sc.perfs.every(p => p.cap > 0), "全上映回にキャパシティがある");
ok(sc.perfs.every(p => p.screen > 0), "全上映回にスクリーン番号がある");
if (sc.warnings.length) console.log("  警告: " + sc.warnings.join(" / "));

const screens = {};
for (const p of sc.perfs) screens[p.screen] = p.cap;
console.log("  スクリーン別キャパ: " + Object.keys(screens).sort((a, b) => a - b).map(k => `SC${k}=${screens[k]}`).join(" "));

console.log("\n  作品別（上位10件）");
for (const t of sc.titles.slice(0, 10)) {
  console.log(`   ${t.label.padEnd(28)} ${String(t.shows).padStart(3)}回 ${String(t.seats).padStart(6)}席 SC${t.screens.join(",")}`);
}

console.log("\n== 先週編成の繰り返し推計 ==");
const horizon = "2026-12-31";
const projected = P.projectPerfs(sc, horizon);
ok(projected.length > sc.perfs.length, `推計込みの上映回 ${projected.length} 回`);
ok(projected.every(p => p.date <= horizon), "推計が指定期限を超えない");

if (giftPath) {
  console.log("\n== 入場者プレゼント情報の読み込み ==");
  const sheets = P.readXlsx(buf(giftPath));
  const gifts = sheets.flatMap(s => P.parseGiftSheet(s, sc.start));
  ok(gifts.length > 0, `${gifts.length}件を取得`);
  for (const g of gifts) {
    console.log(`   [${g.section || "本編"}] ${String(g.title).slice(0, 26)} | ${String(g.present).slice(0, 26)} | 数=${g.qty}${g.bundle ? `(束${g.bundle})` : ""}${g.extras.length ? ` +追加${g.extras.map(e => e.qty).join(",")}` : ""} | ${g.start || g.startText} 〜 ${g.end || g.endText}`);
  }

  console.log("\n== 満席想定シミュレーション（自動一致した作品のみ） ==");
  const opts = { base: sc.start, rate: 1, strict: false, horizon: horizon };
  for (const g of gifts) {
    if (!g.qty) continue;
    const gift = {
      title: g.title, present: g.present, qty: g.qty + g.extras.reduce((a, e) => a + e.qty, 0), per: 1,
      start: g.start, end: (g.openEnded || g.tbd) ? null : g.end,
      links: null, titleKey: P.normTitle(g.title), screens: [], include: "", exclude: "", timeFrom: "", timeTo: ""
    };
    const r = P.simulateGift(gift, projected, opts);
    console.log(`   ${String(g.title).slice(0, 24).padEnd(24)} 在庫${String(gift.qty).padStart(6)} 対象${String(r.targets).padStart(4)}回 配布${String(r.served).padStart(4)}回 残${String(r.left).padStart(6)} → ${r.lastDate || "-"} [${r.status}]`);
  }
}

console.log(failed ? `\n${failed} 件の確認に失敗しました` : "\nすべての確認に成功しました");
process.exit(failed ? 1 : 0);
