/**
 * 現場の事故につながるバグの再現テスト。
 *   node test/bugcheck.mjs [PS.xlsx]
 * 過去に実際に起きた不具合を、修正されたままかどうか毎回確かめる。
 */
import fs from 'node:fs'; import vm from 'node:vm';
vm.runInThisContext([...fs.readFileSync(''+process.cwd()+'/index.html','utf8')
  .matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m=>m[1]).join('\n'));
const P = globalThis.__PGC;
const K = P.normTitle('T');
const mk = (date,screen,cap,min,title='T') => ({date,screen,cap,title,key:P.normTitle(title),min,projected:false});
let ng=0; const ok=(c,m)=>{ console.log((c?'  OK  ':'  NG  ')+m); if(!c) ng++; };

// A: 状態判定と掲示日付の軸が一致するか
{
  const perfs=[];
  for(let d=0;d<20;d++){ const dt='2026-08-'+String(10+d).padStart(2,'0');
    perfs.push(mk(dt,1,500,600)); perfs.push(mk(dt,9,30,900)); }
  const r=P.simulateGift({title:'T',present:'X',qty:1400,per:1,start:'2026-08-10',end:null,
    links:null,titleKey:K,screens:[],include:'',exclude:'',timeFrom:'',timeTo:''},
    perfs,{base:'2026-08-10',rate:1,strict:false,horizon:'2026-08-29'});
  ok(r.safeUntil==='2026-08-11' && r.safeDaysLeft===1, `A: 掲示日=${r.safeUntil} 残り${r.safeDaysLeft}日（lastDate=${r.lastDate}）`);
}
// C: 配布開始前の日付を掲示しない
{
  const perfs=[mk('2026-08-12',1,500,600), mk('2026-08-13',9,30,600)];
  const r=P.simulateGift({title:'T',present:'X',qty:100,per:1,start:'2026-08-12',end:null,
    links:null,titleKey:K,screens:[],include:'',exclude:'',timeFrom:'',timeTo:''},
    perfs,{base:'2026-08-10',rate:1,strict:false,horizon:'2026-08-20'});
  ok(r.safeUntil===null, `C: 開始前(8/11)を掲示しない → safeUntil=${r.safeUntil}`);
}
// D: 推計打ち切り日を確定終了日にしない
{
  const perfs=[]; for(let d=0;d<5;d++) perfs.push(mk('2026-08-'+String(10+d).padStart(2,'0'),1,100,600));
  const r=P.simulateGift({title:'T',present:'X',qty:100000,per:1,start:'2026-08-10',end:'2026-12-31',
    links:null,titleKey:K,screens:[],include:'',exclude:'',timeFrom:'',timeTo:''},
    perfs,{base:'2026-08-10',rate:1,strict:false,horizon:'2026-08-14'});
  ok(r.endedByHorizon===true, `D: 推計打切りを判別 endedByHorizon=${r.endedByHorizon} lastDate=${r.lastDate}`);
}
// B: 在庫0でも safeRate が数値
{
  const perfs=[mk('2026-08-10',1,500,600)];
  const r=P.simulateGift({title:'T',present:'X',qty:0,per:1,start:'2026-08-10',end:'2026-08-14',
    links:null,titleKey:K,screens:[],include:'',exclude:'',timeFrom:'',timeTo:''},
    perfs,{base:'2026-08-10',rate:1,strict:false,horizon:'2026-08-20'});
  ok(r.safeRate!==null, `B: 在庫0でも safeRate=${r.safeRate}（nullでない）`);
}
// H: showsLeft は起算日時点の手持ちから
{
  const perfs=[]; for(let d=0;d<10;d++) perfs.push(mk('2026-08-'+String(10+d).padStart(2,'0'),1,100,600));
  const r=P.simulateGift({title:'T',present:'X',qty:300,per:1,start:'2026-08-10',end:null,
    links:null,titleKey:K,screens:[],include:'',exclude:'',timeFrom:'',timeTo:''},
    perfs,{base:'2026-08-10',rate:1,strict:false,horizon:'2026-08-25'});
  ok(r.showsLeft===3, `H: 手持ち300個 / 1回100個 → あと約${r.showsLeft}回分（終了時残=${r.left}）`);
}
// E: 別作品の同名プレゼントが混ざらない
{
  const perfs=[mk('2026-08-10',1,300,600,'作品A'), mk('2026-08-10',2,300,700,'作品B')];
  const rs=[
    P.simulateGift({id:1,title:'作品A',present:'クリアファイル',qty:300,per:1,start:'2026-08-10',end:null,
      links:null,titleKey:P.normTitle('作品A'),screens:[],include:'',exclude:'',timeFrom:'',timeTo:''},perfs,{base:'2026-08-10',rate:1,strict:false,horizon:'2026-08-20'}),
    P.simulateGift({id:2,title:'作品B',present:'クリアファイル',qty:0,per:1,start:'2026-08-10',end:null,
      links:null,titleKey:P.normTitle('作品B'),screens:[],include:'',exclude:'',timeFrom:'',timeTo:''},perfs,{base:'2026-08-10',rate:1,strict:false,horizon:'2026-08-20'})];
  const plan=P.buildDailyPlan('2026-08-10',rs);
  const ids=plan.flatMap(p=>p.items.map(i=>i.id));
  ok(ids.includes(1)&&ids.includes(2), `E: 同名でもIDで区別 ids=[${ids}]`);
}
// N: 1回に複数配るとき、入プレがサンプリングより先に並ぶ
{
  const perfs=[mk('2026-08-10',1,100,600,'作品A')];
  const g=(id,cat,name)=>({id,cat,title:'作品A',present:name,qty:100000,per:1,start:'2026-08-10',
    end:null,links:null,titleKey:P.normTitle('作品A'),screens:[],include:'',exclude:'',timeFrom:'',timeTo:''});
  const opt={base:'2026-08-10',rate:1,strict:false,horizon:'2026-08-20'};
  /* わざとサンプリング→LV→入プレの順で渡す */
  const rs=[P.simulateGift(g(1,'sampling','チラシ'),perfs,opt),
            P.simulateGift(g(2,'event','LV特典'),perfs,opt),
            P.simulateGift(g(3,'present','クリアファイル'),perfs,opt)];
  const cats=P.buildDailyPlan('2026-08-10',rs)[0].items.map(i=>i.cat);
  ok(cats.join()==='present,sampling,event', `N: 1回の配布順 = ${cats.join(' → ')}`);
}
// O: 準備数の一覧にIDではなくプレゼント名が出る（掲示・指示表の致命的な表示崩れ）
{
  const perfs=[mk('2026-08-10',1,100,600,'作品A')];
  const r=P.simulateGift({id:77,cat:'present',title:'作品A',present:'クリアファイル',qty:1000,per:1,
    start:'2026-08-10',end:null,links:null,titleKey:P.normTitle('作品A'),
    screens:[],include:'',exclude:'',timeFrom:'',timeTo:''},perfs,
    {base:'2026-08-10',rate:1,strict:false,horizon:'2026-08-20'});
  const it=P.buildDailyPlan('2026-08-10',[r])[0].items[0];
  ok(it.name==='クリアファイル'&&it.id===77, `O: 準備数の表示名=${it.name}（id=${it.id}）`);
}
// P: 起算日より前から配っているものを「残数未確認」として立てる
{
  const perfs=[]; for(let d=0;d<10;d++) perfs.push(mk('2026-08-'+String(10+d).padStart(2,'0'),1,100,600));
  const base={title:'T',present:'X',qty:1000,per:1,end:null,links:null,titleKey:K,
    screens:[],include:'',exclude:'',timeFrom:'',timeTo:''};
  const o={base:'2026-08-10',rate:1,strict:false,horizon:'2026-08-25'};
  const before=P.simulateGift({...base,start:'2026-07-20'},perfs,o);          // 起算日より前に開始
  const counted=P.simulateGift({...base,start:'2026-07-20',counted:true},perfs,o);
  const after=P.simulateGift({...base,start:'2026-08-10'},perfs,o);           // 起算日に開始
  ok(before.needsCount===true && counted.needsCount===false && after.needsCount===false,
    `P: 残数未確認フラグ 未確認=${before.needsCount} 確認済=${counted.needsCount} 起算日開始=${after.needsCount}`);
  /* 起算日より前の上映回は数に入れない（PSにその期間のデータが無いので逆算できない） */
  ok(before.perfs.every(p=>p.date>='2026-08-10'), 'P: 起算日より前の上映回は対象に入れない');
}
// Q: 掲示に出す在庫と「約◯回分」の時点が揃っている
{
  const perfs=[]; for(let d=0;d<10;d++) perfs.push(mk('2026-08-'+String(10+d).padStart(2,'0'),1,100,600));
  const r=P.simulateGift({title:'T',present:'X',qty:1000,per:1,start:'2026-08-10',end:null,
    links:null,titleKey:K,screens:[],include:'',exclude:'',timeFrom:'',timeTo:''},
    perfs,{base:'2026-08-10',rate:1,strict:false,horizon:'2026-08-25'});
  /* left は「配り終えた後の残り」。掲示にはこれではなく gift.qty を出す。
     混ぜると「残0個（約10回分）」のような噛み合わない表示になる。 */
  ok(r.left!==r.gift.qty && r.showsLeft===Math.floor(r.gift.qty/r.needPerShow),
    `Q: left=${r.left}(終了時) / qty=${r.gift.qty}(起算日) / 約${r.showsLeft}回分は qty 基準`);
}
// R: 表をそのまま貼り付けても読み取れる
{
  const txt=fs.readFileSync(process.cwd()+'/test/fixtures/gift-paste.txt','utf8');
  const rows=P.parseGiftText(txt,'2026-08-07');
  const cat=r=>/サンプリング/.test(r.section)?'sampling':/LV|特別上映/i.test(r.section)?'event':'present';
  const n=k=>rows.filter(r=>cat(r)===k).length;
  ok(rows.length===24, `R: 24件を読み取る（実際 ${rows.length}件）`);
  ok(n('present')===13&&n('event')===9&&n('sampling')===2,
    `R: 区分の振り分け 入プレ${n('present')} / LV${n('event')} / サンプ${n('sampling')}`);
  const pau=rows.find(r=>/ダイノ・ムービー/.test(r.title));
  ok(pau&&pau.qty===2800&&pau.bundle===200&&pau.start==='2026-07-24'&&pau.openEnded,
    `R: 「2800（200）」「7月24日(金)」「無くなるまで」 → ${pau&&pau.qty}/${pau&&pau.bundle}/${pau&&pau.start}`);
  const toy=rows.find(r=>/トイ・ストーリー/.test(r.title));
  ok(toy&&toy.qty===8000&&toy.bundle===1000, `R: 桁区切り「8,000（1000）」→ ${toy&&toy.qty}`);
  /* 直前の件の備考を、次の件の作品名に取り違えないこと（現場で作品が入れ替わる事故） */
  const chii=rows.find(r=>/ちいかわ/.test(r.title));
  ok(chii&&/残数報告/.test(chii.note)&&chii.tbd, 'R: 複数行の備考は直前の件に付く／「調整中」を判別');
  ok(toy&&!/残数報告/.test(toy.title), 'R: 備考を次の件の作品名に混ぜない');
  const fam=rows.find(r=>/ファミリーシアター/.test(r.title));
  ok(fam&&/同時に配布可能。/.test(fam.note), 'R: 文で終わる備考も直前の件に付く');
  /* 【】で始まるが備考ではないもの（企画上映の作品名）を備考に吸い込まない */
  const cars=rows.find(r=>/カーズ/.test(r.title));
  ok(cars&&cars.start==='2026-09-04'&&cars.end==='2026-09-10', `R: 【企画上映】…の行を1件として読む`);
  /* 納品数が空でも、日付だけで1件として立てる */
  const doro=rows.find(r=>/ドロヘドロ/.test(r.title));
  ok(doro&&doro.qty===null&&doro.present==='入浴剤', 'R: 納品数が空でも読み取る');
  ok(rows.every(r=>r.start), 'R: 24件すべてに配布開始日が入る');
}
// S: Excelから直接コピーした形式（改行入りセルが引用符で囲まれる）も読める
{
  const q=c=>/[\n\t"]/.test(c)?'"'+c.replace(/"/g,'""')+'"':c;
  const tsv=[
    ['作品名','プレゼント内容','納品数\n（一束の数）','配布開始日','配布終了日'],
    ['オークストリートの異変\nIMAX限定','IMAX限定ビジュアルA3ポスター','1000','8/14(金)','無くなるまで'],
  ].map(r=>r.map(q).join('\t')).join('\n');
  const rows=P.parseGiftText(tsv,'2026-08-07');
  ok(rows.length===1&&rows[0].title==='オークストリートの異変 IMAX限定'&&rows[0].qty===1000,
    `S: 引用符付きTSV → ${rows.length}件 / ${rows[0]&&rows[0].title}`);
  /* 文中の引用符（手書き"災愛"）を壊さない */
  const plain='作品名\tプレゼント内容\t納品数\t配布開始日\n'
    +'『オブセッション 災愛』第２弾\t【ニッキー手書き"災愛"メッセージ入り】特製ポストカード\t1000（250）\t8/14(金)';
  const r2=P.parseGiftText(plain,'2026-08-07');
  ok(r2.length===1&&/"災愛"/.test(r2[0].present), `S: 文中の引用符を残す → ${r2[0]&&r2[0].present}`);
}
// M: 重複した上映回を二重に数えない
{
  const psPath=process.argv[2];
  if(!psPath){ console.log('  --  M: PSファイルを引数で渡すと重複除去も検査します'); }
  const b=psPath?fs.readFileSync(psPath):null;
  if(!b){ console.log(ng?`\n${ng}件 失敗`:'\nすべて期待どおり'); process.exit(ng?1:0); }
  const sheets=P.readXlsx(b.buffer.slice(b.byteOffset,b.byteOffset+b.byteLength));
  const dup=P.parseSchedule(sheets.concat(sheets));   // 同じシートを2回渡す
  const one=P.parseSchedule(sheets);
  ok(dup.perfs.length===one.perfs.length, `M: 重複入力でも上映回は ${dup.perfs.length}（単体 ${one.perfs.length}）`);
  ok(dup.warnings.some(w=>w.includes('重複')), 'M: 重複を警告に出す');
}
console.log(ng?`\n${ng}件 失敗`:'\nすべて期待どおり');
process.exit(ng?1:0);
