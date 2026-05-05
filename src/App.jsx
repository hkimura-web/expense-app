import { useState } from "react";
import * as XLSX from "xlsx";

// ─── マスターデータ ──────────────────────────────────────────────────────────

const DESTINATIONS = [
  {
    label: "三ツ沢公園陸上競技場",
    transport: [
      { name: "京浜急行", direction: "往・復", from: "上大岡", to: "横浜", amount: "456" },
      { name: "横浜市営バス", direction: "往・復", from: "横浜駅西口", to: "三ツ沢総合グランド入口", amount: "440" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
    ],
    carRoute: "自宅 ～ 三ツ沢公園陸上競技場",
  },
  {
    label: "横浜国際陸上競技場（日産スタジアム）",
    transport: [
      { name: "京浜急行", direction: "往・復", from: "上大岡", to: "横浜", amount: "456" },
      { name: "JR", direction: "往・復", from: "横浜", to: "小机", amount: "356" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
    ],
    carRoute: "自宅 ～ 横浜国際陸上競技場",
  },
  {
    label: "金沢区海の公園・八景島",
    transport: [
      { name: "シーサイドライン", direction: "往・復", from: "金沢八景", to: "海の公園柴口", amount: "540" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
    ],
    carRoute: "自宅 ～ 海の公園駐車場〈柴口〉",
  },
  {
    label: "横浜市立万騎が原中学校",
    transport: [
      { name: "京急", direction: "往・復", from: "上大岡", to: "横浜", amount: "456" },
      { name: "相鉄", direction: "往・復", from: "横浜", to: "二又川", amount: "420" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
    ],
    carRoute: "自宅 ～ 横浜市立万騎が原中学校",
  },
  {
    label: "中央大学附属横浜高等学校",
    transport: [
      { name: "横浜市営地下鉄", direction: "往・復", from: "上大岡", to: "センター北", amount: "796" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
    ],
    carRoute: "自宅 ～ 中央大学附属横浜高等学校",
  },
  {
    label: "レモンガススタジアム平塚",
    transport: [
      { name: "横浜市営地下鉄", direction: "往・復", from: "上永谷", to: "戸塚", amount: "" },
      { name: "JR", direction: "往・復", from: "戸塚", to: "平塚", amount: "" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
      { name: "", direction: "往・復", from: "", to: "", amount: "" },
    ],
    carRoute: "自宅 ～ レモンガススタジアム平塚",
  },
  { label: "その他（自由入力）", transport: null, carRoute: "" },
];

const PURPOSES = [
  "【陸上競技部・高校】横浜地区高体連陸上競技選手権大会引率",
  "【陸上競技部・高校】横浜地区高体連学校対校新人陸上競技大会引率",
  "【陸上競技部・高校】神奈川県高等学校陸上競技大会　横浜地区予選会引率",
  "【陸上競技部・高校】神奈川県高等学校新人陸上競技大会 県大会引率",
  "【陸上競技部・高校】神奈川県高等学校新人陸上競技大会横浜地区予選引率",
  "【陸上競技部・高校】神奈川県高等学校総合体育大会　関東高等学校陸上競技大会 県予選会引率",
  "【陸上競技部・中学】横浜市中学校秋季陸上競技大会引率",
  "【陸上競技部・中学】横浜市中学校総合体育大会 陸上競技の部引率",
  "【陸上競技部・中学】神奈川県中学校陸上競技選手権大会引率",
  "【陸上競技部・中学】全日本中学校通信陸上競技大会 神奈川大会 横浜地区予選会引率",
  "【陸上競技部・中学・高校】横浜市陸上競技選手権大会 兼 第５回横浜市記録会引率",
  "【陸上競技部・中学・高校】港南区陸上競技大会引率",
  "【陸上競技部・中学・高校】金沢区ロードレース大会引率",
  "【陸上競技部・高校】第1回横浜市記録会引率",
  "【陸上競技部・中学・高校】第2回横浜市記録会引率",
  "【陸上競技部・中学・高校】第3回横浜市記録会引率",
  "陸上競技部　夏合宿",
  "横浜地区高体連陸上競技専門部　顧問総会",
  "全日本中学校通信陸上競技大会 横浜地区予選打ち合わせ会",
  "その他（自由入力）",
];

const TRANSPORT_OPTIONS = [
  "京浜急行", "横浜市営バス", "横浜市営地下鉄", "JR", "相鉄", "東急",
  "シーサイドライン", "小田急", "箱根登山鉄道", "京急バス", "その他（自由入力）",
];

// 旅費規程早見表 2021年4月より
const ALLOWANCE_OPTIONS = [
  { label: "なし：0円", value: 0, group: null },
  { label: "生徒引率 校外・部活動（平日 or 土勤務日）：1,400円", value: 1400, group: "① 宿泊なし生徒引率" },
  { label: "生徒引率 校内・部活動（土休業日・日・祝）：1,400円", value: 1400, group: "① 宿泊なし生徒引率" },
  { label: "生徒引率 校外・部活動（土休業日・日・祝）：3,400円", value: 3400, group: "① 宿泊なし生徒引率" },
  { label: "宿泊引率 日当（平日・土勤務日）：1,400円", value: 1400, group: "② 宿泊あり生徒引率" },
  { label: "宿泊引率 日当（土休業日・日・祝）：3,400円", value: 3400, group: "② 宿泊あり生徒引率" },
  { label: "宿泊引率 宿泊手当（毎日）：3,400円", value: 3400, group: "② 宿泊あり生徒引率" },
  { label: "個人出張 平日 2h以上4h未満（16:10以降）：700円", value: 700, group: "③ 個人出張" },
  { label: "個人出張 平日 4時間以上（16:10以降）：1,400円", value: 1400, group: "③ 個人出張" },
  { label: "個人出張 土勤務日 2h以上4h未満（12:30以降）：700円", value: 700, group: "③ 個人出張" },
  { label: "個人出張 土勤務日 4時間以上（12:30以降）：1,400円", value: 1400, group: "③ 個人出張" },
  { label: "個人出張 土休業日・日・祝 2h以上4h未満：1,400円", value: 1400, group: "③ 個人出張" },
  { label: "個人出張 土休業日・日・祝 4時間以上：2,800円", value: 2800, group: "③ 個人出張" },
];

// ─── 初期値（木村勇人さん専用デフォルト）
const initialTrip = {
  affiliation: "六浦中学校・高等学校",
  name: "木村勇人",
  date: "", timeStart: "", timeEnd: "",
  destinationIdx: 0,
  destinationCustom: "",
  purposeIdx: 0,
  purposeCustom: "",
  transport: DESTINATIONS[0].transport.map(t => ({ ...t })),
  allowanceIdx: 3,   // 生徒引率 校外 休日 3400円
  allowanceDays: "1",
};
const initialCar = {
  dateFrom: "", dateTo: "",
  purpose: "陸上競技部　引率時　テント運搬",
  route: "自宅 ～ 三ツ沢公園陸上競技場",
  driver: "木村勇人",
  passengers: "なし",
  carType: "ホンダ ステップワゴン",
  carNumber: "横浜323 ― さ23",
};
const initialParking = {
  paymentType: "立替払い",
  employeeId: "300117",
  name: "木村勇人",
  budgetItem: "課外13",
  paymentDate: "",
  content: "陸上競技部　引率時　テント運搬のため　駐車場代",
  amount: "",
};

// ─── UI部品 ─────────────────────────────────────────────────────────────────
const BLUE = "#2563eb";
const baseInput = { width: "100%", padding: "9px 12px", border: "1.5px solid #e5e7eb", borderRadius: 8, fontSize: 14, background: "#fafafa", color: "#111", outline: "none", boxSizing: "border-box", transition: "border-color 0.15s", fontFamily: "inherit" };

function Input(props) {
  const [f, setF] = useState(false);
  return <input {...props} style={{ ...baseInput, borderColor: f ? BLUE : "#e5e7eb", background: f ? "#fff" : "#fafafa", ...props.style }} onFocus={() => setF(true)} onBlur={() => setF(false)} />;
}
function Sel({ value, onChange, children, style }) {
  return <select value={value} onChange={onChange} style={{ ...baseInput, ...style, cursor: "pointer" }}>{children}</select>;
}
function Field({ label, children, half, note }) {
  return (
    <div style={{ marginBottom: 14, width: half ? "48%" : "100%" }}>
      <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: "#6b7280", letterSpacing: "0.08em", marginBottom: 5, textTransform: "uppercase" }}>
        {label}{note && <span style={{ fontWeight: 400, marginLeft: 6, color: "#9ca3af", textTransform: "none" }}>{note}</span>}
      </label>
      {children}
    </div>
  );
}

// ─── Excel出力 ───────────────────────────────────────────────────────────────
function downloadExcel(trip, car, parking) {
  const dest = trip.destinationIdx < DESTINATIONS.length - 1 ? DESTINATIONS[trip.destinationIdx].label : trip.destinationCustom;
  const purpose = trip.purposeIdx < PURPOSES.length - 1 ? PURPOSES[trip.purposeIdx] : trip.purposeCustom;
  const allowance = ALLOWANCE_OPTIONS[trip.allowanceIdx];
  const allowanceAmt = allowance.value * (parseInt(trip.allowanceDays) || 1);
  const transportTotal = trip.transport.reduce((s, t) => s + (parseInt(t.amount) || 0), 0);
  const total = transportTotal + allowanceAmt;
  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet([
    ["出張費精算書"], [],
    ["所属", trip.affiliation], ["氏名", trip.name], ["出張日", trip.date],
    ["時間", `${trip.timeStart} ～ ${trip.timeEnd}`], ["出張先", dest], ["用件", purpose], [],
    ["交通機関", "往復", "出発", "到着", "金額(円)"],
    ...trip.transport.filter(t => t.name || t.amount).map(t => [t.name, t.direction, t.from ? `${t.from}駅` : "", t.to ? `${t.to}駅` : "", parseInt(t.amount) || 0]),
    [], ["日当区分", allowance.label, "", "", allowanceAmt], [], ["合計", "", "", "", total],
  ]);
  ws1["!cols"] = [{ wch: 22 }, { wch: 40 }, { wch: 14 }, { wch: 14 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws1, "出張費精算書");
  const ws2 = XLSX.utils.aoa_to_sheet([["自家用車使用届"], [], ["使用開始日", car.dateFrom], ["使用終了日", car.dateTo], ["使用目的", car.purpose], ["使用区間", car.route], ["運転者", car.driver], ["同乗者", car.passengers], ["使用車種", car.carType], ["車両番号", car.carNumber]]);
  ws2["!cols"] = [{ wch: 14 }, { wch: 40 }];
  XLSX.utils.book_append_sheet(wb, ws2, "自家用車使用届");
  const ws3 = XLSX.utils.aoa_to_sheet([["駐車場代精算書（現金払い）"], [], ["支払区分", parking.paymentType], ["教職員番号", parking.employeeId], ["氏名", parking.name], ["予算科目", parking.budgetItem], ["支払年月日", parking.paymentDate], ["支払内容", parking.content], ["支払金額(円)", parseInt(parking.amount) || 0]]);
  ws3["!cols"] = [{ wch: 14 }, { wch: 40 }];
  XLSX.utils.book_append_sheet(wb, ws3, "駐車場代精算書");
  XLSX.writeFile(wb, `経費精算_${trip.name}_${trip.date || new Date().toISOString().slice(0, 10)}.xlsx`);
}

// ─── 印刷 ────────────────────────────────────────────────────────────────────
function printTab(tab, trip, car, parking) {
  const dest = trip.destinationIdx < DESTINATIONS.length - 1 ? DESTINATIONS[trip.destinationIdx].label : trip.destinationCustom;
  const purpose = trip.purposeIdx < PURPOSES.length - 1 ? PURPOSES[trip.purposeIdx] : trip.purposeCustom;
  const allowance = ALLOWANCE_OPTIONS[trip.allowanceIdx];
  const allowanceAmt = allowance.value * (parseInt(trip.allowanceDays) || 1);
  const transportTotal = trip.transport.reduce((s, t) => s + (parseInt(t.amount) || 0), 0);
  const total = transportTotal + allowanceAmt;
  const h2 = s => `<h2 style="font-size:17px;font-weight:800;margin:0 0 14px;border-bottom:2px solid #1e3a8a;padding-bottom:7px;color:#1e3a8a">${s}</h2>`;
  const row = (l, v) => `<tr><td style="padding:6px 10px;font-weight:700;color:#555;white-space:nowrap;width:130px;background:#f8f9fa;border-bottom:1px solid #e5e7eb">${l}</td><td style="padding:6px 10px;border-bottom:1px solid #e5e7eb">${v || "—"}</td></tr>`;
  const tbl = rows => `<table style="width:100%;border-collapse:collapse;margin-bottom:18px;font-size:13px;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden">${rows}</table>`;
  let content = "";
  if (tab === 0) {
    const tRows = trip.transport.filter(t => t.name || t.amount).map(t =>
      `<tr><td style="padding:5px 8px;border-bottom:1px solid #f0f0f0">${t.name}</td><td style="padding:5px 8px;border-bottom:1px solid #f0f0f0">${t.direction}</td><td style="padding:5px 8px;border-bottom:1px solid #f0f0f0">${t.from}駅 ～ ${t.to}駅</td><td style="padding:5px 8px;border-bottom:1px solid #f0f0f0;text-align:right">¥${(parseInt(t.amount) || 0).toLocaleString()}</td></tr>`).join("");
    content = h2("出張費精算書") +
      tbl(row("所属", trip.affiliation) + row("氏名", trip.name) + row("出張日", trip.date) + row("時間", `${trip.timeStart} ～ ${trip.timeEnd}`) + row("出張先", dest) + row("用件", purpose)) +
      `<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:18px;border:1px solid #e5e7eb"><thead><tr style="background:#1e3a8a;color:#fff"><th style="padding:7px">交通機関</th><th style="padding:7px">往復</th><th style="padding:7px">区間</th><th style="padding:7px;text-align:right">金額</th></tr></thead><tbody>${tRows}</tbody></table>` +
      `<div style="background:#1e3a8a;color:#fff;padding:14px 18px;border-radius:8px;display:flex;justify-content:space-between;align-items:center"><span style="font-size:11px;opacity:.8">${allowance.label}</span><span>交通費¥${transportTotal.toLocaleString()} + 日当¥${allowanceAmt.toLocaleString()} <strong style="font-size:20px;margin-left:10px">= ¥${total.toLocaleString()}</strong></span></div>`;
  } else if (tab === 1) {
    content = h2("自家用車使用届") + tbl(row("使用開始日", car.dateFrom) + row("使用終了日", car.dateTo) + row("使用目的", car.purpose) + row("使用区間", car.route) + row("運転者", car.driver) + row("同乗者", car.passengers) + row("使用車種", car.carType) + row("車両番号", car.carNumber));
  } else {
    content = h2("駐車場代精算書（現金払い）") + tbl(row("支払区分", parking.paymentType) + row("教職員番号", parking.employeeId) + row("氏名", parking.name) + row("予算科目", parking.budgetItem) + row("支払年月日", parking.paymentDate) + row("支払内容", parking.content) + row("支払金額", `¥${(parseInt(parking.amount) || 0).toLocaleString()}`));
  }
  const w = window.open("", "_blank");
  w.document.write(`<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"><title>印刷</title><style>body{font-family:'Hiragino Kaku Gothic ProN',sans-serif;padding:40px;max-width:720px;margin:0 auto;color:#111}@media print{body{padding:20px}}</style></head><body><div style="font-size:11px;color:#999;margin-bottom:22px">関東学院六浦中学校・高等学校　印刷日：${new Date().toLocaleDateString("ja-JP")}</div>${content}<script>window.onload=()=>window.print()<\/script></body></html>`);
  w.document.close();
}

// ─── 出張費精算書フォーム ─────────────────────────────────────────────────────
function TripForm({ data, setData }) {
  const up = (k, v) => setData(d => ({ ...d, [k]: v }));
  const upT = (i, k, v) => setData(d => ({ ...d, transport: d.transport.map((t, idx) => idx === i ? { ...t, [k]: v } : t) }));
  const transportTotal = data.transport.reduce((s, t) => s + (parseInt(t.amount) || 0), 0);
  const allowanceInfo = ALLOWANCE_OPTIONS[data.allowanceIdx];
  const allowanceAmt = allowanceInfo.value * (parseInt(data.allowanceDays) || 1);
  const total = transportTotal + allowanceAmt;

  const handleDestChange = idx => {
    const dest = DESTINATIONS[idx];
    setData(d => ({
      ...d, destinationIdx: idx, destinationCustom: "",
      transport: dest.transport ? dest.transport.map(t => ({ ...t })) : [
        { name: "", direction: "往・復", from: "", to: "", amount: "" },
        { name: "", direction: "往・復", from: "", to: "", amount: "" },
        { name: "", direction: "往・復", from: "", to: "", amount: "" },
        { name: "", direction: "往・復", from: "", to: "", amount: "" },
      ],
    }));
  };

  const allowanceGroups = [];
  let cg = null;
  ALLOWANCE_OPTIONS.forEach((opt, idx) => {
    if (opt.group !== cg) { cg = opt.group; allowanceGroups.push({ group: opt.group, items: [] }); }
    allowanceGroups[allowanceGroups.length - 1].items.push({ ...opt, idx });
  });

  return (
    <div>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Field label="所属"><Input value={data.affiliation} onChange={e => up("affiliation", e.target.value)} /></Field>
        <Field label="氏名" half><Input value={data.name} onChange={e => up("name", e.target.value)} /></Field>
        <Field label="出張日" half><Input type="date" value={data.date} onChange={e => up("date", e.target.value)} /></Field>
        <Field label="開始時間" half><Input type="time" value={data.timeStart} onChange={e => up("timeStart", e.target.value)} /></Field>
        <Field label="終了時間" half><Input type="time" value={data.timeEnd} onChange={e => up("timeEnd", e.target.value)} /></Field>
        <Field label="出張先">
          <Sel value={data.destinationIdx} onChange={e => handleDestChange(Number(e.target.value))}>
            {DESTINATIONS.map((d, i) => <option key={i} value={i}>{d.label}</option>)}
          </Sel>
          {data.destinationIdx === DESTINATIONS.length - 1 && (
            <Input value={data.destinationCustom} onChange={e => up("destinationCustom", e.target.value)} placeholder="出張先を入力" style={{ marginTop: 6 }} />
          )}
        </Field>
        <Field label="用件">
          <Sel value={data.purposeIdx} onChange={e => setData(d => ({ ...d, purposeIdx: Number(e.target.value), purposeCustom: "" }))}>
            {PURPOSES.map((p, i) => <option key={i} value={i}>{p}</option>)}
          </Sel>
          {data.purposeIdx === PURPOSES.length - 1 && (
            <Input value={data.purposeCustom} onChange={e => up("purposeCustom", e.target.value)} placeholder="用件を入力" style={{ marginTop: 6 }} />
          )}
        </Field>
      </div>

      <div style={{ marginBottom: 6 }}>
        <label style={{ fontSize: 11, fontWeight: 700, color: "#6b7280", letterSpacing: "0.08em", textTransform: "uppercase" }}>交通費</label>
      </div>
      <div style={{ border: "1.5px solid #e5e7eb", borderRadius: 10, overflow: "hidden", marginBottom: 16 }}>
        <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 2fr 2fr 1fr", background: "#f3f4f6", padding: "8px 12px", fontSize: 11, fontWeight: 700, color: "#6b7280", letterSpacing: "0.06em" }}>
          <span>交通機関</span><span>往復</span><span>出発駅</span><span>到着駅</span><span style={{ textAlign: "right" }}>金額(円)</span>
        </div>
        {data.transport.map((t, i) => {
          const isCustom = !TRANSPORT_OPTIONS.includes(t.name) && t.name !== "";
          const selVal = isCustom ? "その他（自由入力）" : (t.name || "");
          return (
            <div key={i} style={{ borderTop: "1px solid #f0f0f0" }}>
              <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 2fr 2fr 1fr", gap: 8, padding: "8px 12px", alignItems: "center" }}>
                <Sel value={selVal} onChange={e => upT(i, "name", e.target.value === "その他（自由入力）" ? "" : e.target.value)} style={{ fontSize: 13 }}>
                  <option value="">（未選択）</option>
                  {TRANSPORT_OPTIONS.map(o => <option key={o} value={o}>{o}</option>)}
                </Sel>
                <Sel value={t.direction} onChange={e => upT(i, "direction", e.target.value)} style={{ fontSize: 12, padding: "9px 4px" }}>
                  <option>往・復</option><option>片道</option>
                </Sel>
                <Input value={t.from} onChange={e => upT(i, "from", e.target.value)} placeholder="出発駅" style={{ fontSize: 13 }} />
                <Input value={t.to} onChange={e => upT(i, "to", e.target.value)} placeholder="到着駅" style={{ fontSize: 13 }} />
                <Input type="number" value={t.amount} onChange={e => upT(i, "amount", e.target.value)} placeholder="0" style={{ fontSize: 13, textAlign: "right" }} />
              </div>
              {selVal === "その他（自由入力）" && (
                <div style={{ padding: "0 12px 8px" }}>
                  <Input value={t.name} onChange={e => upT(i, "name", e.target.value)} placeholder="交通機関名を入力" style={{ fontSize: 13 }} />
                </div>
              )}
            </div>
          );
        })}
      </div>

      <Field label="日当" note="（旅費規程早見表より選択）">
        <Sel value={data.allowanceIdx} onChange={e => up("allowanceIdx", Number(e.target.value))}>
          {allowanceGroups.map(g => g.group
            ? <optgroup key={g.group} label={`── ${g.group} ──`}>
                {g.items.map(it => <option key={it.idx} value={it.idx}>{it.label}</option>)}
              </optgroup>
            : g.items.map(it => <option key={it.idx} value={it.idx}>{it.label}</option>)
          )}
        </Sel>
      </Field>
      <Field label="日数" half>
        <Input type="number" value={data.allowanceDays} onChange={e => up("allowanceDays", e.target.value)} min="1" />
      </Field>

      <div style={{ background: "linear-gradient(135deg,#1e3a8a,#2563eb)", borderRadius: 12, padding: "18px 24px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
        <div><div style={{ color: "rgba(255,255,255,0.7)", fontSize: 11, marginBottom: 2 }}>交通費合計</div><div style={{ color: "#fff", fontSize: 16, fontWeight: 600 }}>¥{transportTotal.toLocaleString()}</div></div>
        <div style={{ color: "rgba(255,255,255,0.4)", fontSize: 20 }}>+</div>
        <div><div style={{ color: "rgba(255,255,255,0.7)", fontSize: 11, marginBottom: 2 }}>日当</div><div style={{ color: "#fff", fontSize: 16, fontWeight: 600 }}>¥{allowanceAmt.toLocaleString()}</div></div>
        <div style={{ color: "rgba(255,255,255,0.4)", fontSize: 20 }}>=</div>
        <div style={{ textAlign: "right" }}><div style={{ color: "rgba(255,255,255,0.7)", fontSize: 11, marginBottom: 2 }}>合計金額</div><div style={{ color: "#fff", fontSize: 28, fontWeight: 800 }}>¥{total.toLocaleString()}</div></div>
      </div>
    </div>
  );
}

// ─── 自家用車使用届フォーム ───────────────────────────────────────────────────
function CarForm({ data, setData }) {
  const up = (k, v) => setData(d => ({ ...d, [k]: v }));
  return (
    <div>
      <div style={{ background: "#fffbeb", border: "1.5px solid #fcd34d", borderRadius: 10, padding: "12px 16px", fontSize: 13, color: "#92400e", marginBottom: 16 }}>
        ⚠️　デフォルトは木村勇人さんのホンダ ステップワゴンです。変更が必要な場合のみ修正してください。
      </div>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Field label="使用開始日" half><Input type="date" value={data.dateFrom} onChange={e => up("dateFrom", e.target.value)} /></Field>
        <Field label="使用終了日" half><Input type="date" value={data.dateTo} onChange={e => up("dateTo", e.target.value)} /></Field>
        <Field label="使用目的"><Input value={data.purpose} onChange={e => up("purpose", e.target.value)} /></Field>
        <Field label="使用区間"><Input value={data.route} onChange={e => up("route", e.target.value)} /></Field>
        <Field label="運転者" half><Input value={data.driver} onChange={e => up("driver", e.target.value)} /></Field>
        <Field label="同乗者" half><Input value={data.passengers} onChange={e => up("passengers", e.target.value)} /></Field>
        <Field label="使用車種" half><Input value={data.carType} onChange={e => up("carType", e.target.value)} /></Field>
        <Field label="車両番号" half><Input value={data.carNumber} onChange={e => up("carNumber", e.target.value)} /></Field>
      </div>
      <div style={{ background: "#f0f9ff", border: "1.5px solid #bae6fd", borderRadius: 10, padding: "12px 16px", fontSize: 13, color: "#0369a1" }}>
        ⓘ　この届出は出張前に所属長への提出が必要です。
      </div>
    </div>
  );
}

// ─── 駐車場代精算書フォーム ───────────────────────────────────────────────────
function ParkingForm({ data, setData }) {
  const up = (k, v) => setData(d => ({ ...d, [k]: v }));
  return (
    <div>
      <div style={{ background: "#fffbeb", border: "1.5px solid #fcd34d", borderRadius: 10, padding: "12px 16px", fontSize: 13, color: "#92400e", marginBottom: 16 }}>
        ⚠️　氏名・教職員番号・支払内容はデフォルト入力済みです。<strong>支払金額と支払年月日だけ</strong>入力してください。
      </div>
      <Field label="支払区分">
        <div style={{ display: "flex", gap: 12 }}>
          {["立替払い", "仮払金"].map(opt => (
            <label key={opt} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 14 }}>
              <input type="radio" name="payType" checked={data.paymentType === opt} onChange={() => up("paymentType", opt)} style={{ accentColor: BLUE }} />
              {opt}
            </label>
          ))}
        </div>
      </Field>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Field label="教職員番号" half><Input value={data.employeeId} onChange={e => up("employeeId", e.target.value)} /></Field>
        <Field label="氏名" half><Input value={data.name} onChange={e => up("name", e.target.value)} /></Field>
        <Field label="予算科目" half><Input value={data.budgetItem} onChange={e => up("budgetItem", e.target.value)} /></Field>
        <Field label="★ 支払年月日" half>
          <Input type="date" value={data.paymentDate} onChange={e => up("paymentDate", e.target.value)}
            style={{ borderColor: !data.paymentDate ? "#f97316" : "#e5e7eb", background: !data.paymentDate ? "#fff7ed" : "#fafafa" }} />
        </Field>
        <Field label="支払内容"><Input value={data.content} onChange={e => up("content", e.target.value)} /></Field>
        <Field label="★ 支払金額(円)">
          <Input type="number" value={data.amount} onChange={e => up("amount", e.target.value)} placeholder="毎回入力"
            style={{ borderColor: !data.amount ? "#f97316" : "#e5e7eb", background: !data.amount ? "#fff7ed" : "#fafafa", fontSize: 18, fontWeight: 700 }} />
        </Field>
      </div>
      {data.amount && (
        <div style={{ background: "linear-gradient(135deg,#064e3b,#059669)", borderRadius: 12, padding: "18px 24px", marginTop: 4, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ color: "rgba(255,255,255,0.8)", fontSize: 14 }}>支払金額</span>
          <span style={{ color: "#fff", fontSize: 28, fontWeight: 800 }}>¥{parseInt(data.amount || 0).toLocaleString()}</span>
        </div>
      )}
    </div>
  );
}

// ─── メインアプリ ─────────────────────────────────────────────────────────────
const TABS = ["出張費精算書", "自家用車使用届", "駐車場代精算書"];

export default function App() {
  const [tab, setTab] = useState(0);
  const [trip, setTrip] = useState(initialTrip);
  const [car, setCar] = useState(initialCar);
  const [parking, setParking] = useState(initialParking);
  const [excelDone, setExcelDone] = useState(false);

  const handleExcel = () => { downloadExcel(trip, car, parking); setExcelDone(true); setTimeout(() => setExcelDone(false), 2000); };
  const handleReset = () => { if (tab === 0) setTrip(initialTrip); if (tab === 1) setCar(initialCar); if (tab === 2) setParking(initialParking); };

  const btn = (label, onClick, bg, color = "#fff", border = "none") => (
    <button onClick={onClick} style={{ flex: 1, padding: "13px 0", background: bg, color, border, borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", letterSpacing: "0.02em" }}
      onMouseEnter={e => e.currentTarget.style.opacity = "0.88"} onMouseLeave={e => e.currentTarget.style.opacity = "1"}>
      {label}
    </button>
  );

  return (
    <div style={{ minHeight: "100vh", background: "#f8fafc", fontFamily: "'Hiragino Kaku Gothic ProN','Noto Sans JP',sans-serif" }}>
      <div style={{ background: "#fff", borderBottom: "1px solid #e5e7eb", padding: "0 32px" }}>
        <div style={{ maxWidth: 760, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", height: 60 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 32, height: 32, background: "linear-gradient(135deg,#2563eb,#1e40af)", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center" }}>
              <span style={{ color: "#fff", fontSize: 16 }}>📋</span>
            </div>
            <span style={{ fontWeight: 800, fontSize: 15, color: "#111" }}>経費精算システム</span>
          </div>
          <span style={{ fontSize: 11, color: "#9ca3af" }}>関東学院六浦中学校・高等学校</span>
        </div>
      </div>
      <div style={{ background: "#fff", borderBottom: "1px solid #e5e7eb" }}>
        <div style={{ maxWidth: 760, margin: "0 auto", display: "flex" }}>
          {TABS.map((t, i) => (
            <button key={i} onClick={() => setTab(i)} style={{ padding: "14px 20px", border: "none", background: "none", cursor: "pointer", fontSize: 13, fontWeight: tab === i ? 700 : 500, color: tab === i ? BLUE : "#6b7280", borderBottom: tab === i ? `2.5px solid ${BLUE}` : "2.5px solid transparent", transition: "all 0.15s", fontFamily: "inherit" }}>{t}</button>
          ))}
        </div>
      </div>
      <div style={{ padding: "32px 16px 100px" }}>
        <div style={{ background: "#fff", borderRadius: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.08),0 8px 24px rgba(0,0,0,0.06)", padding: "28px 32px", maxWidth: 760, margin: "0 auto" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 24 }}>
            <div>
              <div style={{ fontSize: 20, fontWeight: 800, color: "#111" }}>{TABS[tab]}</div>
              <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 3 }}>
                {tab === 0 && "出張先・用件・交通機関・日当を選択して記入"}
                {tab === 1 && "公用出張における自家用車使用届"}
                {tab === 2 && "現金払い精算書（駐車場代）"}
              </div>
            </div>
            <button onClick={handleReset} style={{ fontSize: 12, color: "#9ca3af", background: "none", border: "1px solid #e5e7eb", borderRadius: 6, padding: "5px 12px", cursor: "pointer", fontFamily: "inherit" }}>リセット</button>
          </div>
          {tab === 0 && <TripForm data={trip} setData={setTrip} />}
          {tab === 1 && <CarForm data={car} setData={setCar} />}
          {tab === 2 && <ParkingForm data={parking} setData={setParking} />}
        </div>
      </div>
      <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: "rgba(255,255,255,0.95)", backdropFilter: "blur(8px)", borderTop: "1px solid #e5e7eb", padding: "12px 16px" }}>
        <div style={{ maxWidth: 760, margin: "0 auto", display: "flex", gap: 10 }}>
          {btn(excelDone ? "✓ ダウンロード完了" : "📊 Excelで出力（3シート）", handleExcel, excelDone ? "#059669" : "linear-gradient(135deg,#16a34a,#15803d)")}
          {btn("🖨️ 印刷 / PDF保存", () => printTab(tab, trip, car, parking), "#fff", "#1e3a8a", "2px solid #1e3a8a")}
        </div>
      </div>
    </div>
  );
}
