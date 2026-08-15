/**
 * 1ファイル版（standalone.html）を生成するスクリプト。
 *
 *   node build-standalone.js
 *
 * index.html / style.css / app.js を1つのHTMLにまとめる。
 * 生成物はメール添付やUSBでの配布用。中身は3ファイル版と同じ。
 */

const fs = require('fs');
const path = require('path');

const dir = __dirname;
const html = fs.readFileSync(path.join(dir, 'index.html'), 'utf8');
let css = fs.readFileSync(path.join(dir, 'style.css'), 'utf8');
const js = fs.readFileSync(path.join(dir, 'app.js'), 'utf8');

let body = html.split('<body>')[1].split('</body>')[0];
body = body.replace(/\s*<script src="app\.js"><\/script>/, '').trim();

// ダークテーマのトークンを OS設定 / 明示指定 の両方に対応させる
const darkBlock = css.match(/@media \(prefers-color-scheme: dark\) \{\n  :root \{\n([\s\S]*?)\n  \}\n\}/);
if (!darkBlock) throw new Error('ダークテーマのトークン定義が見つかりません');
const tokens = darkBlock[1];

css = css.replace(darkBlock[0], `@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
${tokens.replace(/^ {4}/gm, '    ')}
  }
}

:root[data-theme="dark"] {
${tokens.replace(/^ {4}/gm, '  ')}
}`);

const out = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>月次予算割振アプリ</title>
<style>
${css}
</style>
</head>
<body>

${body}

<script>
${js}
</script>
</body>
</html>
`;

const dest = path.join(dir, 'standalone.html');
fs.writeFileSync(dest, out);
console.log(`生成しました: ${dest} (${(out.length / 1024).toFixed(1)}KB)`);
