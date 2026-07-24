import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

/* ──────────────────────────────────────────────────────────
   設定：把這裡改成你自己的 Cloudflare Worker 網址
   部署 Worker 前先留空，app 仍可手動輸入使用
   ────────────────────────────────────────────────────────── */
const API_URL = "";   // 例："https://invoice-api.your-name.workers.dev"

const CATS = [
  { id: "reagent",  label: "試劑耗材", color: "#3F6B52" },
  { id: "antibody", label: "抗體",     color: "#6E9E7C" },
  { id: "plastic",  label: "塑膠器材", color: "#C9D64B" },
  { id: "equip",    label: "儀器設備", color: "#D9A441" },
  { id: "service",  label: "委外服務", color: "#B4553C" },
  { id: "software", label: "軟體授權", color: "#8A7BB8" },
  { id: "other",    label: "其他",     color: "#A8B3AE" },
];
const catOf = (id) => CATS.find((c) => c.id === id) || CATS[6];
const nt = (n) => "$" + Number(n || 0).toLocaleString("zh-TW");
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const today = () => new Date().toISOString().slice(0, 10);
const roc = (iso) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${Number(y) - 1911}/${m}/${d}`;
};

const KEY = "invoices-v1";

/* 影像壓縮：長邊 1500px */
async function compress(file, maxSide = 1500, quality = 0.8) {
  const dataUrl = await new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = () => rej(new Error("讀取檔案失敗"));
    r.readAsDataURL(file);
  });
  const img = await new Promise((res, rej) => {
    const i = new Image();
    i.onload = () => res(i);
    i.onerror = () => rej(new Error("影像格式無法辨識"));
    i.src = dataUrl;
  });
  const scale = Math.min(1, maxSide / Math.max(img.width, img.height));
  const cv = document.createElement("canvas");
  cv.width = Math.round(img.width * scale);
  cv.height = Math.round(img.height * scale);
  cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
  return cv.toDataURL("image/jpeg", quality);
}

const toBase64 = (file) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = () => rej(new Error("讀取檔案失敗"));
    r.readAsDataURL(file);
  });

const PROMPT = `你是台灣實驗室核銷助理。請閱讀這份發票或收據，抽取欄位後只回傳 JSON，不要有任何說明文字或 markdown 標記。多頁時以第一張發票為準。

- invoice: 發票號碼，台灣格式為兩個大寫英文字母加八位數字。找不到填空字串
- date: 發票日期，一律轉西元 YYYY-MM-DD。圖上若為民國年（115/06/10）請加 1911 換算成 2026-06-10
- amount: 總計金額，純數字。取「總計」或「應收金額」，不要取未稅金額
- item: 品名，多項用頓號連接，保留型號
- vendor: 開立發票的公司或廠商名稱
- category: 擇一填英文 id：reagent（試劑、化學品、培養基）、antibody（抗體）、plastic（離心管、吸管尖、孔盤）、equip（儀器設備）、service（委外服務、定序、合成）、software（軟體授權）、other
- confidence: 辨識可信度，0 到 1 的小數

範例：{"invoice":"ZZ19417253","date":"2026-06-10","amount":50000,"item":"抗體 Sox2、Stat3","vendor":"友和貿易","category":"antibody","confidence":0.9}

看不清或沒有的欄位填空字串或 0，不要杜撰。`;

async function readInvoice(block) {
  if (!API_URL) throw new Error("尚未設定辨識服務網址，請先部署 Worker 或改用手動輸入。");
  const resp = await fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      messages: [{ role: "user", content: [block, { type: "text", text: PROMPT }] }],
    }),
  });
  if (!resp.ok) throw new Error(`辨識服務回應 ${resp.status}`);
  const data = await resp.json();
  const text = (data.content || []).filter((c) => c.type === "text").map((c) => c.text).join("");
  const clean = text.replace(/```json|```/g, "").trim();
  const a = clean.indexOf("{"), b = clean.lastIndexOf("}");
  if (a < 0 || b < 0) throw new Error("無法解析辨識結果");
  return JSON.parse(clean.slice(a, b + 1));
}

export default function App() {
  const [rows, setRows] = useState([]);
  const [editing, setEditing] = useState(null);
  const [scanning, setScanning] = useState(false);
  const [q, setQ] = useState("");
  const [cat, setCat] = useState("all");
  const [month, setMonth] = useState("all");
  const [toast, setToast] = useState("");

  useEffect(() => {
    try {
      const r = localStorage.getItem(KEY);
      if (r) setRows(JSON.parse(r));
    } catch {}
  }, []);
  useEffect(() => {
    try { localStorage.setItem(KEY, JSON.stringify(rows)); } catch {}
  }, [rows]);

  const flash = (m) => { setToast(m); setTimeout(() => setToast(""), 2400); };

  const save = (rec) => {
    const { _scanned, _confidence, _dupe, ...clean } = rec;
    setRows((p) => p.some((r) => r.id === clean.id)
      ? p.map((r) => (r.id === clean.id ? clean : r)) : [...p, clean]);
    setEditing(null); flash("已儲存");
  };
  const remove = (id) => setRows((p) => p.filter((r) => r.id !== id));

  const months = useMemo(() =>
    [...new Set(rows.filter((r) => r.date).map((r) => r.date.slice(0, 7)))].sort().reverse(), [rows]);

  const visible = useMemo(() => {
    let v = rows;
    if (cat !== "all") v = v.filter((r) => r.category === cat);
    if (month !== "all") v = v.filter((r) => r.date?.startsWith(month));
    if (q.trim()) {
      const s = q.toLowerCase();
      v = v.filter((r) => [r.item, r.vendor, r.invoice, r.note].join(" ").toLowerCase().includes(s));
    }
    return [...v].sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  }, [rows, cat, month, q]);

  const total = visible.reduce((s, r) => s + Number(r.amount || 0), 0);

  const byCat = useMemo(() => CATS.map((c) => ({
    ...c,
    count: visible.filter((r) => r.category === c.id).length,
    value: visible.filter((r) => r.category === c.id).reduce((s, r) => s + Number(r.amount || 0), 0),
  })).filter((d) => d.count > 0), [visible]);

  const byMonth = useMemo(() => {
    const m = {};
    visible.forEach((r) => { if (r.date) m[r.date.slice(0, 7)] = (m[r.date.slice(0, 7)] || 0) + Number(r.amount || 0); });
    return Object.entries(m).sort();
  }, [visible]);

  const onScanned = ({ fields, image, dupe }) => {
    setScanning(false);
    setEditing({
      id: uid(),
      date: fields.date || today(),
      invoice: fields.invoice || "",
      amount: fields.amount || "",
      item: fields.item || "",
      vendor: fields.vendor || "",
      category: CATS.some((c) => c.id === fields.category) ? fields.category : "other",
      note: "", receipt: image,
      _scanned: true, _confidence: fields.confidence, _dupe: dupe,
    });
  };

  const exportXlsx = () => {
    if (!visible.length) return flash("沒有可匯出的資料");
    const detail = visible.map((r, i) => ({
      序號: i + 1, 日期: roc(r.date), 發票號碼: r.invoice,
      品名: r.item, 廠商: r.vendor, 分類: catOf(r.category).label,
      金額: Number(r.amount || 0), 備註: r.note,
    }));
    detail.push({}, { 品名: "合計", 金額: total });
    const ws1 = XLSX.utils.json_to_sheet(detail);
    ws1["!cols"] = [{ wch: 6 }, { wch: 12 }, { wch: 14 }, { wch: 40 },
      { wch: 18 }, { wch: 11 }, { wch: 12 }, { wch: 22 }];

    const ws2 = XLSX.utils.json_to_sheet(
      byCat.map((c) => ({ 分類: c.label, 筆數: c.count, 金額: c.value })));
    ws2["!cols"] = [{ wch: 12 }, { wch: 8 }, { wch: 13 }];

    const ws3 = XLSX.utils.json_to_sheet(
      byMonth.map(([m, v]) => ({ 月份: m, 金額: v })));
    ws3["!cols"] = [{ wch: 12 }, { wch: 13 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, "發票明細");
    XLSX.utils.book_append_sheet(wb, ws2, "分類統計");
    XLSX.utils.book_append_sheet(wb, ws3, "月份統計");
    XLSX.writeFile(wb, `發票統計_${today()}.xlsx`);
    flash("已匯出 Excel");
  };

  return (
    <div style={{ background: "#FBFAF6", minHeight: "100vh", color: "#16241F",
      fontFamily: "'Noto Sans TC','Microsoft JhengHei',system-ui,sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&family=Newsreader:opsz,wght@6..72,400;6..72,600&display=swap');
        *{box-sizing:border-box}
        body{margin:0}
        .btn{border:none;border-radius:2px;padding:9px 15px;font-size:13.5px;font-weight:500;cursor:pointer;font-family:inherit}
        .btn:hover{opacity:.85}
        .btn:focus-visible{outline:2px solid #C9D64B;outline-offset:2px}
        .inp{width:100%;padding:8px 10px;border:1px solid #DDE0D8;border-radius:2px;font-size:13.5px;font-family:inherit;background:#fff;color:#16241F}
        .inp:focus{outline:none;border-color:#3F6B52}
        .lbl{display:block;font-size:10.5px;letter-spacing:.09em;color:#7C8A85;margin-bottom:4px;font-weight:500}
        .card{background:#fff;border:1px solid #E6E8E1;border-radius:3px}
        th{font-size:10.5px;letter-spacing:.07em;color:#7C8A85;font-weight:500;text-align:left;padding:9px 10px;border-bottom:1px solid #E6E8E1;white-space:nowrap}
        td{padding:11px 10px;border-bottom:1px solid #F0F1EC;font-size:13.5px;vertical-align:top}
        tbody tr:hover{background:#FCFCF9}
        .num{font-variant-numeric:tabular-nums}
        @keyframes spin{to{transform:rotate(360deg)}}
        .spin{animation:spin 1s linear infinite}
        @media (prefers-reduced-motion:reduce){.spin{animation:none}}
        @media (max-width:760px){.hide-sm{display:none}}
      `}</style>

      <header style={{ borderBottom: "1px solid #E6E8E1", background: "#fff",
        position: "sticky", top: 0, zIndex: 20 }}>
        <div style={{ maxWidth: 1120, margin: "0 auto", padding: "14px 20px",
          display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <span style={{ width: 9, height: 26, background: "#C9D64B", transform: "skewX(-12deg)" }} />
          <h1 style={{ fontFamily: "Newsreader,serif", fontSize: 22, fontWeight: 600, margin: 0 }}>
            發票統計
          </h1>
          <div style={{ flex: 1 }} />
          <button className="btn" style={{ background: "#16241F", color: "#FBFAF6" }}
            onClick={() => setScanning(true)}>掃描發票</button>
          <button className="btn" style={{ background: "#3F6B52", color: "#fff" }}
            onClick={() => setEditing(blank())}>手動新增</button>
          <button className="btn" style={{ background: "#fff", color: "#16241F", border: "1px solid #DDE0D8" }}
            onClick={exportXlsx}>匯出 Excel</button>
        </div>
      </header>

      <main style={{ maxWidth: 1120, margin: "0 auto", padding: "22px 20px 70px" }}>
        {/* 摘要 */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))",
          gap: 12, marginBottom: 16 }}>
          {[["總金額", nt(total)], ["張數", visible.length + " 張"],
            ["平均單張", nt(visible.length ? Math.round(total / visible.length) : 0)]].map(([l, v]) => (
            <div key={l} className="card" style={{ padding: "15px 17px" }}>
              <div style={{ fontSize: 10.5, letterSpacing: ".08em", color: "#7C8A85", marginBottom: 6 }}>{l}</div>
              <div className="num" style={{ fontFamily: "Newsreader,serif", fontSize: 24, fontWeight: 600 }}>{v}</div>
            </div>
          ))}
        </div>

        {byCat.length > 0 && (
          <div className="card" style={{ padding: 17, marginBottom: 16 }}>
            <div style={{ fontSize: 10.5, letterSpacing: ".08em", color: "#7C8A85", marginBottom: 11 }}>分類占比</div>
            <div style={{ display: "flex", height: 22, borderRadius: 2, overflow: "hidden", marginBottom: 11 }}>
              {byCat.map((c) => (
                <div key={c.id} title={`${c.label} ${nt(c.value)}`}
                  style={{ width: `${(c.value / total) * 100}%`, background: c.color }} />
              ))}
            </div>
            <div style={{ display: "flex", gap: 18, flexWrap: "wrap", fontSize: 12.5 }}>
              {byCat.map((c) => (
                <span key={c.id} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span style={{ width: 9, height: 9, background: c.color, borderRadius: 1 }} />
                  {c.label} {c.count} 張 · <span className="num">{nt(c.value)}</span>
                </span>
              ))}
            </div>
          </div>
        )}

        {/* 篩選 */}
        <div style={{ display: "flex", gap: 9, alignItems: "center", marginBottom: 14, flexWrap: "wrap" }}>
          <input className="inp" style={{ width: 210 }} placeholder="搜尋品名 / 廠商 / 發票號碼"
            value={q} onChange={(e) => setQ(e.target.value)} />
          <select className="inp" style={{ width: "auto" }} value={cat} onChange={(e) => setCat(e.target.value)}>
            <option value="all">全部分類</option>
            {CATS.map((c) => <option key={c.id} value={c.id}>{c.label}</option>)}
          </select>
          <select className="inp" style={{ width: "auto" }} value={month} onChange={(e) => setMonth(e.target.value)}>
            <option value="all">全部月份</option>
            {months.map((m) => <option key={m} value={m}>{m}</option>)}
          </select>
        </div>

        <div className="card" style={{ overflowX: "auto" }}>
          {visible.length === 0 ? (
            <div style={{ padding: "60px 20px", textAlign: "center", color: "#7C8A85", fontSize: 14 }}>
              {rows.length === 0
                ? "還沒有發票。點「掃描發票」上傳圖片或 PDF，或用「手動新增」輸入。"
                : "沒有符合篩選條件的發票。"}
            </div>
          ) : (
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 700 }}>
              <thead>
                <tr>
                  <th>日期</th><th>品名</th>
                  <th className="hide-sm">廠商</th><th className="hide-sm">發票號碼</th>
                  <th style={{ textAlign: "right" }}>金額</th><th />
                </tr>
              </thead>
              <tbody>
                {visible.map((r) => (
                  <tr key={r.id}>
                    <td className="num" style={{ whiteSpace: "nowrap", color: "#7C8A85" }}>{roc(r.date) || "—"}</td>
                    <td>
                      <div style={{ display: "flex", gap: 8 }}>
                        <span style={{ width: 3, alignSelf: "stretch", background: catOf(r.category).color,
                          borderRadius: 2, flexShrink: 0 }} />
                        <div>
                          <div style={{ fontWeight: 500, marginBottom: 2 }}>{r.item || "（未填品名）"}</div>
                          <div style={{ fontSize: 12, color: "#7C8A85" }}>
                            {catOf(r.category).label}{r.note && ` · ${r.note}`}
                          </div>
                          {r.receipt && <a href={r.receipt} target="_blank" rel="noreferrer"
                            style={{ fontSize: 12, color: "#3F6B52" }}>檢視發票</a>}
                        </div>
                      </div>
                    </td>
                    <td className="hide-sm" style={{ fontSize: 12.5, color: "#7C8A85" }}>{r.vendor || "—"}</td>
                    <td className="hide-sm num" style={{ fontSize: 12.5,
                      color: r.invoice ? "#16241F" : "#C3C9BF" }}>{r.invoice || "—"}</td>
                    <td className="num" style={{ textAlign: "right", fontWeight: 500, whiteSpace: "nowrap" }}>
                      {nt(r.amount)}
                    </td>
                    <td style={{ whiteSpace: "nowrap", textAlign: "right" }}>
                      <button className="btn" style={{ background: "none", color: "#16241F", padding: "5px 8px" }}
                        onClick={() => setEditing(r)}>編輯</button>
                      <button className="btn" style={{ background: "none", color: "#B4553C", padding: "5px 6px" }}
                        onClick={() => remove(r.id)}>刪</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </main>

      {scanning && <ScanModal existing={rows} onDone={onScanned} onClose={() => setScanning(false)} />}
      {editing && <FormModal rec={editing} onSave={save} onClose={() => setEditing(null)} />}

      {toast && (
        <div style={{ position: "fixed", bottom: 26, left: "50%", transform: "translateX(-50%)",
          background: "#16241F", color: "#FBFAF6", padding: "11px 22px", borderRadius: 2,
          fontSize: 13.5, zIndex: 70 }}>{toast}</div>
      )}
    </div>
  );
}

const blank = () => ({
  id: uid(), date: today(), invoice: "", amount: "", item: "",
  vendor: "", category: "reagent", note: "", receipt: null,
});

function ScanModal({ existing, onDone, onClose }) {
  const [mode, setMode] = useState(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [preview, setPreview] = useState(null);
  const [fileName, setFileName] = useState("");
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const imgRef = useRef(null);
  const pdfRef = useRef(null);

  const stop = () => { streamRef.current?.getTracks().forEach((t) => t.stop()); streamRef.current = null; };
  useEffect(() => stop, []);

  const imgBlock = (u) => ({ type: "image",
    source: { type: "base64", media_type: "image/jpeg", data: u.split(",")[1] } });

  const process = async (block, image, label) => {
    setBusy(true); setErr(""); setPreview(image); setFileName(label || "");
    try {
      const fields = await readInvoice(block);
      const dupe = fields.invoice
        ? existing.find((r) => r.invoice && r.invoice.toUpperCase() === String(fields.invoice).toUpperCase())
        : null;
      onDone({ fields, image, dupe: dupe ? dupe.invoice : null });
    } catch (e) {
      setErr(e.message === "無法解析辨識結果"
        ? "讀不出這份檔案的內容。請確認字跡清楚，或改用手動新增。" : e.message);
      setBusy(false);
    }
  };

  const startCamera = async () => {
    setErr("");
    try {
      const s = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: "environment" }, width: { ideal: 1920 } } });
      streamRef.current = s; setMode("camera");
      setTimeout(() => { if (videoRef.current) { videoRef.current.srcObject = s; videoRef.current.play(); } }, 60);
    } catch { setErr("無法開啟相機。請確認已授權相機權限，或改用上傳檔案。"); }
  };

  const shoot = async () => {
    const v = videoRef.current; if (!v) return;
    const cv = document.createElement("canvas");
    const scale = Math.min(1, 1500 / Math.max(v.videoWidth, v.videoHeight));
    cv.width = Math.round(v.videoWidth * scale);
    cv.height = Math.round(v.videoHeight * scale);
    cv.getContext("2d").drawImage(v, 0, 0, cv.width, cv.height);
    stop(); setMode(null);
    const u = cv.toDataURL("image/jpeg", 0.8);
    await process(imgBlock(u), u);
  };

  const onImg = async (e) => {
    const f = e.target.files?.[0]; e.target.value = "";
    if (!f) return; setErr("");
    try { const u = await compress(f); await process(imgBlock(u), u); }
    catch (ex) { setErr(ex.message); setBusy(false); }
  };

  const onPdf = async (e) => {
    const f = e.target.files?.[0]; e.target.value = "";
    if (!f) return; setErr("");
    const isPdf = f.type === "application/pdf" || /\.pdf$/i.test(f.name);
    if (!isPdf) return setErr("請選擇 PDF 檔。");
    if (f.size > 4.5 * 1024 * 1024) return setErr("PDF 超過 4.5 MB，請只保留發票那幾頁。");
    try {
      const b64 = await toBase64(f);
      await process({ type: "document",
        source: { type: "base64", media_type: "application/pdf", data: b64 } }, null, f.name);
    } catch (ex) { setErr(ex.message); setBusy(false); }
  };

  return (
    <div onClick={busy ? undefined : onClose}
      style={{ position: "fixed", inset: 0, background: "rgba(22,36,31,.55)",
        display: "flex", alignItems: "center", justifyContent: "center", padding: 16, zIndex: 60 }}>
      <div onClick={(e) => e.stopPropagation()} className="card"
        style={{ width: "100%", maxWidth: 470, padding: 24 }}>
        <h2 style={{ fontFamily: "Newsreader,serif", fontSize: 19, fontWeight: 600, margin: "0 0 6px" }}>
          掃描發票
        </h2>
        <p style={{ fontSize: 12.5, color: "#7C8A85", margin: "0 0 18px", lineHeight: 1.6 }}>
          拍照、選圖片或上傳 PDF，系統會讀出發票號碼、日期、金額、品名與廠商，帶入表單後由你確認再存檔。
        </p>

        <input ref={imgRef} type="file" accept="image/*" onChange={onImg} style={{ display: "none" }} />
        <input ref={pdfRef} type="file" accept="application/pdf,.pdf" onChange={onPdf} style={{ display: "none" }} />

        {busy ? (
          <div style={{ textAlign: "center", padding: "24px 0" }}>
            {preview ? (
              <img src={preview} alt="" style={{ maxWidth: "100%", maxHeight: 180,
                objectFit: "contain", border: "1px solid #E6E8E1", marginBottom: 18 }} />
            ) : fileName ? (
              <div style={{ padding: "20px 16px", border: "1px solid #E6E8E1",
                marginBottom: 18, background: "#FCFCF9" }}>
                <div style={{ fontSize: 11, letterSpacing: ".1em", color: "#B4553C", marginBottom: 6 }}>PDF</div>
                <div style={{ fontSize: 13, wordBreak: "break-all" }}>{fileName}</div>
              </div>
            ) : null}
            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10 }}>
              <span className="spin" style={{ width: 15, height: 15, border: "2px solid #E6E8E1",
                borderTopColor: "#3F6B52", borderRadius: "50%", display: "inline-block" }} />
              <span style={{ fontSize: 13.5, color: "#7C8A85" }}>辨識中，約 5 到 10 秒</span>
            </div>
          </div>
        ) : mode === "camera" ? (
          <>
            <div style={{ position: "relative", background: "#16241F", borderRadius: 2, overflow: "hidden" }}>
              <video ref={videoRef} playsInline muted
                style={{ width: "100%", display: "block", maxHeight: "50vh", objectFit: "cover" }} />
              <div style={{ position: "absolute", inset: "9%", border: "2px solid #C9D64B",
                borderRadius: 3, pointerEvents: "none", opacity: .8 }} />
            </div>
            <div style={{ display: "flex", gap: 9, marginTop: 15, justifyContent: "flex-end" }}>
              <button className="btn" style={{ background: "none", color: "#7C8A85" }}
                onClick={() => { stop(); setMode(null); }}>返回</button>
              <button className="btn" style={{ background: "#3F6B52", color: "#fff" }} onClick={shoot}>拍攝</button>
            </div>
          </>
        ) : (
          <div style={{ display: "grid", gap: 10 }}>
            <button className="btn" onClick={startCamera}
              style={{ background: "#16241F", color: "#FBFAF6", padding: 16, fontSize: 14 }}>
              開啟相機拍攝
            </button>
            <button className="btn" onClick={() => imgRef.current?.click()}
              style={{ background: "#fff", color: "#16241F", border: "1px dashed #C3C9BF", padding: 16, fontSize: 14 }}>
              選擇圖片檔
            </button>
            <button className="btn" onClick={() => pdfRef.current?.click()}
              style={{ background: "#fff", color: "#16241F", border: "1px dashed #C3C9BF", padding: 16, fontSize: 14 }}>
              上傳 PDF
              <span style={{ display: "block", fontSize: 11.5, color: "#7C8A85", fontWeight: 400, marginTop: 3 }}>
                電子發票、報價單，上限 4.5 MB
              </span>
            </button>
          </div>
        )}

        {err && (
          <div style={{ marginTop: 15, padding: "10px 12px", background: "#FBF3F1",
            borderLeft: "3px solid #B4553C", fontSize: 12.5, color: "#B4553C", lineHeight: 1.6 }}>{err}</div>
        )}

        {!busy && mode !== "camera" && (
          <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 16 }}>
            <button className="btn" style={{ background: "none", color: "#7C8A85" }} onClick={onClose}>取消</button>
          </div>
        )}
      </div>
    </div>
  );
}

function FormModal({ rec, onSave, onClose }) {
  const [f, setF] = useState(rec);
  const [err, setErr] = useState("");
  const fileRef = useRef(null);
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));

  const pick = async (e) => {
    const file = e.target.files?.[0]; e.target.value = "";
    if (!file) return;
    try { set("receipt", await compress(file)); setErr(""); }
    catch (ex) { setErr(ex.message); }
  };

  const submit = () => {
    if (!f.item.trim()) return setErr("請填寫品名。");
    if (!Number(f.amount)) return setErr("請填寫金額。");
    onSave({ ...f, amount: Number(f.amount) });
  };

  const lowConf = f._scanned && typeof f._confidence === "number" && f._confidence < 0.75;

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(22,36,31,.42)",
      display: "flex", justifyContent: "center", padding: "5vh 16px", zIndex: 50, overflowY: "auto" }}>
      <div onClick={(e) => e.stopPropagation()} className="card"
        style={{ width: "100%", maxWidth: 520, padding: 24, height: "fit-content" }}>
        <h2 style={{ fontFamily: "Newsreader,serif", fontSize: 19, fontWeight: 600, margin: "0 0 6px" }}>
          {f._scanned ? "確認掃描結果" : rec.item ? "編輯發票" : "新增發票"}
        </h2>
        {f._scanned && (
          <p style={{ fontSize: 12.5, color: "#7C8A85", margin: "0 0 14px", lineHeight: 1.6 }}>
            以下欄位由發票影像讀出，存檔前請核對金額與發票號碼。
          </p>
        )}
        {f._dupe && (
          <div style={{ marginBottom: 14, padding: "10px 12px", background: "#FDF8EC",
            borderLeft: "3px solid #D9A441", fontSize: 12.5, color: "#8A6520", lineHeight: 1.6 }}>
            發票號碼 {f._dupe} 已經登錄過，確認後再存檔。
          </div>
        )}
        {lowConf && (
          <div style={{ marginBottom: 14, padding: "10px 12px", background: "#FBF3F1",
            borderLeft: "3px solid #B4553C", fontSize: 12.5, color: "#B4553C", lineHeight: 1.6 }}>
            影像不夠清楚，辨識可信度偏低，請逐欄核對。
          </div>
        )}

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div><label className="lbl">日期</label>
            <input className="inp" type="date" value={f.date} onChange={(e) => set("date", e.target.value)} /></div>
          <div><label className="lbl">發票號碼</label>
            <input className="inp" value={f.invoice} onChange={(e) => set("invoice", e.target.value)} /></div>
          <div style={{ gridColumn: "1/-1" }}><label className="lbl">品名</label>
            <input className="inp" value={f.item} onChange={(e) => set("item", e.target.value)} /></div>
          <div><label className="lbl">廠商</label>
            <input className="inp" value={f.vendor} onChange={(e) => set("vendor", e.target.value)} /></div>
          <div><label className="lbl">金額</label>
            <input className="inp" type="number" inputMode="numeric" value={f.amount}
              onChange={(e) => set("amount", e.target.value)} /></div>
          <div><label className="lbl">分類</label>
            <select className="inp" value={f.category} onChange={(e) => set("category", e.target.value)}>
              {CATS.map((c) => <option key={c.id} value={c.id}>{c.label}</option>)}
            </select></div>
          <div><label className="lbl">備註</label>
            <input className="inp" value={f.note} onChange={(e) => set("note", e.target.value)} /></div>
        </div>

        <div style={{ marginTop: 15 }}>
          <label className="lbl">發票影像</label>
          <input ref={fileRef} type="file" accept="image/*" onChange={pick} style={{ display: "none" }} />
          {f.receipt ? (
            <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
              <a href={f.receipt} target="_blank" rel="noreferrer">
                <img src={f.receipt} alt="發票" style={{ width: 70, height: 70,
                  objectFit: "cover", border: "1px solid #E6E8E1" }} />
              </a>
              <button className="btn" style={{ background: "none", color: "#B4553C", padding: "6px 8px" }}
                onClick={() => set("receipt", null)}>移除</button>
            </div>
          ) : (
            <button className="btn" onClick={() => fileRef.current?.click()}
              style={{ background: "#fff", color: "#16241F", border: "1px dashed #C3C9BF",
                width: "100%", padding: 15 }}>選擇影像</button>
          )}
        </div>

        {err && <div style={{ marginTop: 13, color: "#B4553C", fontSize: 13 }}>{err}</div>}

        <div style={{ display: "flex", gap: 9, marginTop: 20, justifyContent: "flex-end" }}>
          <button className="btn" style={{ background: "none", color: "#7C8A85" }} onClick={onClose}>取消</button>
          <button className="btn" style={{ background: "#3F6B52", color: "#fff" }} onClick={submit}>儲存</button>
        </div>
      </div>
    </div>
  );
}
