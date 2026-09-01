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
    <div class="hint" id="dlMsg" style="margin-top:6px"></div>
  </div>
</section>`;

/* 保存は claude.ai の枠内なら window.claude.downloads.save() を通す。
   枠の外（保存したHTMLを直接開いた場合など）ではリンクで落とす。
   どちらかは押した瞬間に決める。読み込み時に決め打ちすると、
   まだ window.claude が入っていない一瞬に当たっただけでボタンが死ぬ。 */
const dlScript = `
<script>
(function(){
  var B64="${b64}";
  var btn=document.getElementById("dlBtn"), msg=document.getElementById("dlMsg");
  function say(t){ if(msg) msg.textContent=t||""; }
  function bytes(){
    var bin=atob(B64), buf=new Uint8Array(bin.length);
    for(var i=0;i<bin.length;i++) buf[i]=bin.charCodeAt(i);
    return buf;
  }
  function viaLink(){
    try{
      var url=URL.createObjectURL(new Blob([bytes()],{type:"text/html;charset=utf-8"}));
      var a=document.createElement("a");
      a.href=url; a.download="present-calculator.html";
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function(){ URL.revokeObjectURL(url); },1000);
      return true;
    }catch(e){ return false; }
  }
  function run(){
    var dl=(window.claude&&window.claude.downloads)||null;   /* 押した時点で見る */
    if(!dl){
      say("保存を試しています…");
      var okLink=viaLink();
      setTimeout(function(){
        say(okLink
          ? "保存されない場合は、この画面では保存できません。チャットに添付した index.html をお使いください。"
          : "この画面では保存できません。チャットに添付した index.html をお使いください。");
      },900);
      return;
    }
    if(btn) btn.disabled=true;
    say("保存の確認が出ます。「保存」を選んでください。");
    dl.save({filename:"present-calculator.html", data:bytes()}).then(function(){
      say("保存しました。ダブルクリックで開けます。");
    })["catch"](function(e){
      var c=(e&&e.code)||"unavailable";
      if(c==="extension_not_enabled"||c==="rejected_extension"){
        /* .html を保存できない環境では .txt で渡し、拡張子を直してもらう */
        return dl.save({filename:"present-calculator.html.txt", data:bytes()}).then(function(){
          say("保存しました。ファイル名の末尾「.txt」を消して present-calculator.html にすると開けます。");
        });
      }
      if(c==="declined") say("保存を取り消しました。もう一度押せばやり直せます。");
      else if(c==="rate_limited") say("少し待ってからもう一度押してください。");
      else if(c==="too_large") say("ファイルが大きすぎて保存できません。チャットの添付をお使いください。");
      else { say("保存できませんでした（"+c+"）。リンクでの保存を試します…"); viaLink(); }
    })["catch"](function(e){
      var c=(e&&e.code)||"?";
      say("保存できませんでした（"+c+"）。チャットに添付した index.html をお使いください。");
    }).then(function(){ if(btn) btn.disabled=false; });
  }
  if(btn) btn.addEventListener("click",run);
  /* スマホの「できること」シートからも呼べるようにしておく */
  window.__dlTool={ run:run, label:"ツールをダウンロード", size:${Math.round(src.length/1024)} };
})();
<\/script>`;

const out = `<title>入場者プレゼント 配布可能期間 計算ツール</title>
${style}
${body.replace("<main>", "<main>\n" + downloadPanel.trim())}
${dlScript}
`;

fs.writeFileSync(path.join(here, "artifact.html"), out);
console.log("artifact.html を生成しました:", Math.round(out.length / 1024) + "KB");
