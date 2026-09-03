/**
 * 初期サンプルデータ。
 * 商品名・仕入先・単価・PI値はすべてサンプル値です。
 * 実運用時は「商品マスタ」画面またはCSV取込で自社の値に置き換えてください。
 *
 * packSize : 発注単位1つあたりの消費単位数（例: 1ケース = 500杯分 → 500）
 * price    : 発注単位あたりの仕入単価（税抜）
 * pi       : 来場者100人あたりの消費数（PI値）
 * lead     : リードタイム（発注から入荷までの日数）
 * lot      : 発注ロット（この倍数で発注する）
 */
const SEED_ITEMS = [
  // ---------- フード ----------
  { code: 'F-101', name: 'ポップコーン原料豆',        category: 'food',      vendor: 'フード商事',        unit: 'ケース', packSize: 200,  price: 12000, pi: 18,  lead: 3, lot: 1, storage: '常温' },
  { code: 'F-102', name: 'ポップコーン用オイル',      category: 'food',      vendor: 'フード商事',        unit: 'ケース', packSize: 300,  price: 8600,  pi: 18,  lead: 3, lot: 1, storage: '常温' },
  { code: 'F-103', name: 'ポップコーン用ソルト',      category: 'food',      vendor: 'フード商事',        unit: 'ケース', packSize: 400,  price: 4200,  pi: 11,  lead: 3, lot: 1, storage: '常温' },
  { code: 'F-104', name: 'キャラメルフレーバー',      category: 'food',      vendor: 'フード商事',        unit: 'ケース', packSize: 120,  price: 8400,  pi: 7,   lead: 4, lot: 1, storage: '常温' },
  { code: 'F-201', name: 'ナチョスチップス',          category: 'food',      vendor: 'フード商事',        unit: 'ケース', packSize: 100,  price: 6500,  pi: 4,   lead: 3, lot: 1, storage: '常温' },
  { code: 'F-202', name: 'チーズソース',              category: 'food',      vendor: 'フード商事',        unit: 'ケース', packSize: 80,   price: 7200,  pi: 4,   lead: 3, lot: 1, storage: '常温' },
  { code: 'F-301', name: 'ホットドッグ ソーセージ',   category: 'food',      vendor: '冷凍食品センター',  unit: 'ケース', packSize: 120,  price: 9600,  pi: 3,   lead: 4, lot: 1, storage: '冷凍' },
  { code: 'F-302', name: 'ホットドッグ バンズ',       category: 'food',      vendor: '冷凍食品センター',  unit: 'ケース', packSize: 120,  price: 4800,  pi: 3,   lead: 4, lot: 1, storage: '冷凍' },
  { code: 'F-401', name: 'フライドポテト',            category: 'food',      vendor: '冷凍食品センター',  unit: 'ケース', packSize: 90,   price: 8100,  pi: 5,   lead: 4, lot: 1, storage: '冷凍' },
  { code: 'F-402', name: 'チュリトス',                category: 'food',      vendor: '冷凍食品センター',  unit: 'ケース', packSize: 60,   price: 5400,  pi: 2,   lead: 4, lot: 1, storage: '冷凍' },

  // ---------- ドリンク ----------
  { code: 'D-101', name: 'コーラ シロップ（BIB）',    category: 'drink',     vendor: 'ドリンク物流',      unit: '箱',     packSize: 260,  price: 9800,  pi: 22,  lead: 2, lot: 1, storage: '常温' },
  { code: 'D-102', name: 'コーラ ゼロ シロップ',      category: 'drink',     vendor: 'ドリンク物流',      unit: '箱',     packSize: 260,  price: 9800,  pi: 8,   lead: 2, lot: 1, storage: '常温' },
  { code: 'D-103', name: 'ジンジャーエール シロップ', category: 'drink',     vendor: 'ドリンク物流',      unit: '箱',     packSize: 260,  price: 9500,  pi: 6,   lead: 2, lot: 1, storage: '常温' },
  { code: 'D-104', name: 'オレンジ シロップ',         category: 'drink',     vendor: 'ドリンク物流',      unit: '箱',     packSize: 240,  price: 9200,  pi: 5,   lead: 2, lot: 1, storage: '常温' },
  { code: 'D-105', name: 'ウーロン茶（リキッド）',    category: 'drink',     vendor: 'ドリンク物流',      unit: 'ケース', packSize: 220,  price: 8900,  pi: 5,   lead: 2, lot: 1, storage: '常温' },
  { code: 'D-201', name: 'アイスコーヒー リキッド',   category: 'drink',     vendor: 'ドリンク物流',      unit: 'ケース', packSize: 120,  price: 7600,  pi: 4,   lead: 2, lot: 1, storage: '冷蔵' },
  { code: 'D-202', name: '牛乳（1L）',                category: 'drink',     vendor: 'ドリンク物流',      unit: 'ケース', packSize: 12,   price: 2400,  pi: 1.5, lead: 1, lot: 1, storage: '冷蔵' },
  { code: 'D-301', name: 'ミネラルウォーター 500ml',  category: 'drink',     vendor: 'ドリンク物流',      unit: 'ケース', packSize: 24,   price: 1900,  pi: 4,   lead: 2, lot: 1, storage: '常温' },
  { code: 'D-401', name: '炭酸ガスボンベ',            category: 'drink',     vendor: 'ドリンク物流',      unit: '本',     packSize: 1,    price: 6000,  pi: 0.2, lead: 3, lot: 1, storage: '常温' },

  // ---------- 包材 ----------
  { code: 'P-101', name: 'ポップコーンカップ M',      category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 500,  price: 6500,  pi: 12,  lead: 5, lot: 1, storage: '常温' },
  { code: 'P-102', name: 'ポップコーンカップ L',      category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 400,  price: 7200,  pi: 8,   lead: 5, lot: 1, storage: '常温' },
  { code: 'P-103', name: 'ポップコーンバケツ',        category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 200,  price: 9000,  pi: 2,   lead: 7, lot: 1, storage: '常温' },
  { code: 'P-201', name: 'ドリンクカップ M',          category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 1000, price: 8500,  pi: 25,  lead: 5, lot: 1, storage: '常温' },
  { code: 'P-202', name: 'ドリンクカップ L',          category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 800,  price: 9200,  pi: 18,  lead: 5, lot: 1, storage: '常温' },
  { code: 'P-203', name: 'ドリンクリッド（蓋）',      category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 2000, price: 6800,  pi: 45,  lead: 5, lot: 1, storage: '常温' },
  { code: 'P-204', name: 'ストロー（紙・個包装）',    category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 5000, price: 5500,  pi: 42,  lead: 5, lot: 1, storage: '常温' },
  { code: 'P-205', name: 'ドリンクホルダー（2杯用）', category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 500,  price: 6000,  pi: 8,   lead: 5, lot: 1, storage: '常温' },
  { code: 'P-301', name: 'ナチョストレー',            category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 1000, price: 7000,  pi: 4,   lead: 5, lot: 1, storage: '常温' },
  { code: 'P-302', name: 'ホットドッグ用スリーブ',    category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 1000, price: 5200,  pi: 3,   lead: 5, lot: 1, storage: '常温' },
  { code: 'P-303', name: 'ポテト用カップ',            category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 1000, price: 6300,  pi: 5,   lead: 5, lot: 1, storage: '常温' },
  { code: 'P-401', name: 'ナプキン',                  category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 6000, price: 4200,  pi: 60,  lead: 5, lot: 1, storage: '常温' },
  { code: 'P-402', name: 'おしぼり',                  category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 2000, price: 4600,  pi: 10,  lead: 5, lot: 1, storage: '常温' },
  { code: 'P-403', name: '手提げ袋',                  category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 2000, price: 5800,  pi: 15,  lead: 5, lot: 1, storage: '常温' },
  { code: 'P-404', name: 'ビニール手袋',              category: 'packaging', vendor: '包材サプライ',      unit: 'ケース', packSize: 3000, price: 3900,  pi: 20,  lead: 5, lot: 1, storage: '常温' }
];

/** 曜日別の想定来場者数（日〜土）。設定画面から変更できます。 */
const SEED_ATTENDANCE = [2600, 1200, 1100, 1300, 1400, 1800, 3000];
