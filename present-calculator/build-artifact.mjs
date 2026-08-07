/**
 * index.html から、Web で配布するためのページ artifact.html を生成します。
 *
 *   node build-artifact.mjs
 *
 * ・<!DOCTYPE>/<html>/<head>/<body> を外したページ本体
 * ・index.html 自身を Base64 で同梱し、「ダウンロード」ボタンで単体HTMLとして保存できる
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const srcPath = path.join(here, "index.html");
const src = fs.readFileSync(srcPath);
const html = src.toString("utf8");
const b64 = src.toString("base64");

// <body> の中身と <style> を取り出す
const style = /<style>[\s\S]*?<\/style>/.exec(html)[0];
const body = /<body>([\s\S]*)<\/body>/.exec(html)[1];

const downloadPanel = `
<section class="card noprint" id="dlCard">
  <h2><span class="step">DL</span>ツールのダウンロード<span class="sub">保存すればオフラインで使えます</span></h2>
  <div class="body">
    <div class="row" style="gap:14px">
      <button class="primary" id="dlBtn">⬇ HTMLファイルをダウンロード（${Math.round(src.length / 1024)}KB）</button>
      <span class="hint">
        保存した <code>present-calculator.html</code> をダブルクリックすると、
        このページと同じツールがブラウザで開きます。ネット接続は不要で、読み込んだExcelは端末の外に出ません。
      </span>
    </div>
  </div>
</section>`;

const dlScript = `
<script>
(function(){
  var B64="${b64}";
  var btn=document.getElementById("dlBtn");
  if(!btn) return;
  btn.addEventListener("click",function(){
    var bin=atob(B64), buf=new Uint8Array(bin.length);
    for(var i=0;i<bin.length;i++) buf[i]=bin.charCodeAt(i);
    var url=URL.createObjectURL(new Blob([buf],{type:"text/html;charset=utf-8"}));
    var a=document.createElement("a");
    a.href=url; a.download="present-calculator.html";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function(){ URL.revokeObjectURL(url); },1000);
  });
})();
<\/script>`;

const out = `<title>入場者プレゼント 配布可能期間 計算ツール</title>
${style}
${body.replace("<main>", "<main>\n" + downloadPanel.trim())}
${dlScript}
`;

fs.writeFileSync(path.join(here, "artifact.html"), out);
console.log("artifact.html を生成しました:", Math.round(out.length / 1024) + "KB");
