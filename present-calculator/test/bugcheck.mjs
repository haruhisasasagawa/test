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
