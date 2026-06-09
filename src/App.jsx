import React, { useState, useMemo, useEffect, useRef, useCallback } from "react";

/*
  PHẦN MỀM HỒ SƠ NGHIỆM THU — NEWTECONS / Bộ phận MEP
  • Landing page xanh + upload logo
  • App shell: header xanh + tabs + bảng hồ sơ
  • UPLOAD FORM WORD MẪU cho từng gói khác nhau:
      - App dò chữ MÀU ĐỎ và placeholder {{ten_truong}} làm trường biến đổi
      - Giữ nguyên 100% phần còn lại của form (font, bảng, logo, lề)
      - Khi xuất: thay text trong chính document.xml gốc rồi đóng gói lại .docx
  • Nếu chưa upload form: dùng mẫu mặc định (đã tạo sẵn từ form gốc).
  Thư viện: JSZip (mở/đóng gói docx).
*/

// ============================================================
//  TẢI THƯ VIỆN
// ============================================================
function useScript(src, globalName) {
  const [ready, setReady] = useState(!!window[globalName]);
  useEffect(() => {
    if (window[globalName]) { setReady(true); return; }
    const s = document.createElement("script");
    s.src = src;
    s.onload = () => setReady(true);
    document.body.appendChild(s);
  }, [src, globalName]);
  return ready;
}

// ============================================================
//  XỬ LÝ DOCX TEMPLATE
// ============================================================
const RED_RE = /^[A-Fa-f0-9]{6}$/;
function isRed(hex) {
  if (!hex || !RED_RE.test(hex)) return false;
  const R = parseInt(hex.slice(0, 2), 16);
  const G = parseInt(hex.slice(2, 4), 16);
  const B = parseInt(hex.slice(4, 6), 16);
  return R > 120 && G < 120 && B < 120;
}

// Lấy nội dung text của 1 run XML
function runText(runXml) {
  const m = runXml.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g);
  if (!m) return "";
  return m.map((t) => t.replace(/<w:t(?:\s[^>]*)?>/, "").replace(/<\/w:t>/, "")).join("");
}
function runColor(runXml) {
  const m = runXml.match(/<w:color\s+w:val="([0-9A-Fa-f]{6})"/);
  return m ? m[1] : null;
}
function decodeXml(s) {
  return s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}
function encodeXml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Lấy text của 1 ô <w:tc>
function cellText(tcXml) {
  const m = tcXml.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g);
  if (!m) return "";
  return decodeXml(m.map((t) => t.replace(/<w:t(?:\s[^>]*)?>/, "").replace(/<\/w:t>/, "")).join("")).trim();
}
// Tách các ô <w:tc>...</w:tc> ở mức cao nhất trong 1 hàng
function splitCells(trXml) {
  const cells = [];
  const re = /<w:tc>[\s\S]*?<\/w:tc>/g;
  let m;
  while ((m = re.exec(trXml))) cells.push(m[0]);
  return cells;
}
function splitRows(tblXml) {
  const rows = [];
  const re = /<w:tr\b[\s\S]*?<\/w:tr>/g;
  let m;
  while ((m = re.exec(tblXml))) rows.push({ xml: m[0], start: m.index, end: m.index + m[0].length });
  return rows;
}
const norm = (s) => s.toLowerCase().replace(/\s+/g, " ").trim();

// Tìm bảng khối lượng: <w:tbl> có hàng tiêu đề chứa "tên vật liệu" và "số lượng"
function findKlTable(xml) {
  const re = /<w:tbl>[\s\S]*?<\/w:tbl>/g;
  let m;
  while ((m = re.exec(xml))) {
    const tbl = m[0];
    const rows = splitRows(tbl);
    if (!rows.length) continue;
    const headerCells = splitCells(rows[0].xml).map((c) => norm(cellText(c)));
    const joined = headerCells.join(" | ");
    if (joined.includes("tên vật liệu") && joined.includes("số lượng")) {
      // map cột theo tiêu đề
      const colMap = {};
      headerCells.forEach((h, idx) => {
        if (h.includes("stt")) colMap.stt = idx;
        else if (h.includes("tên")) colMap.ten = idx;
        else if (h.includes("nhãn")) colMap.nhan = idx;
        else if (h.includes("số lượng")) colMap.sl = idx;
        else if (h.includes("đơn vị")) colMap.dv = idx;
        else if (h.includes("ghi chú")) colMap.ghiChu = idx;
      });
      return { tblXml: tbl, start: m.index, end: m.index + tbl.length, rows, colMap, headerCells };
    }
  }
  return null;
}

// Đọc dữ liệu các dòng (bỏ hàng tiêu đề) thành mảng kl
function readKlRows(tblInfo) {
  const { rows, colMap } = tblInfo;
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const cells = splitCells(rows[i].xml);
    const get = (k) => (colMap[k] != null && cells[colMap[k]] != null ? cellText(cells[colMap[k]]) : "");
    const stt = get("stt"), ten = get("ten"), nhan = get("nhan"), sl = get("sl"), dv = get("dv"), ghiChu = get("ghiChu");
    if (![stt, ten, nhan, sl, dv, ghiChu].some((v) => v !== "")) continue;
    const header = ten !== "" && nhan === "" && sl === "" && dv === "";
    out.push({ _id: crypto.randomUUID(), stt, ten, nhan, sl, dv, ghiChu, header });
  }
  return out;
}

function analyzeTemplate(xml) {
  const fields = [];
  const seenRedTexts = new Map();

  // ---- Bảng khối lượng (tự nhận diện) ----
  const klTable = findKlTable(xml);
  let scanXml = xml;
  let klRows = [];
  if (klTable) {
    klRows = readKlRows(klTable);
    // loại vùng bảng ra khỏi phần quét trường đỏ rời
    scanXml = xml.slice(0, klTable.start) + xml.slice(klTable.end);
  }

  // ---- placeholders {{...}} (không tính trong bảng) ----
  const phSet = new Set();
  const phRe = /\{\{\s*([#\/]?\s*[\w.\-]+)\s*\}\}/g;
  let m;
  while ((m = phRe.exec(scanXml))) {
    const raw = m[1].trim();
    if (raw.startsWith("#") || raw.startsWith("/")) continue;
    if (!phSet.has(raw)) {
      phSet.add(raw);
      fields.push({ key: "ph:" + raw, label: raw.replace(/_/g, " "), type: "placeholder", token: "{{" + raw + "}}", value: "" });
    }
  }

  // ---- run đỏ, gộp theo đoạn (ngoài bảng khối lượng) ----
  const paras = scanXml.split(/(<\/w:p>)/);
  const redChunks = [];
  const reRun = /<w:r\b[\s\S]*?<\/w:r>/g;
  let chunkText = "";
  const flush = () => { const t = chunkText.trim(); if (t) redChunks.push(t); chunkText = ""; };
  for (const para of paras) {
    const runs = para.match(reRun) || [];
    for (const run of runs) {
      const col = runColor(run);
      const txt = decodeXml(runText(run));
      if (isRed(col) && txt !== "") chunkText += txt;
      else flush();
    }
    flush();
  }
  redChunks.forEach((t) => {
    if (/^\{\{.*\}\}$/.test(t)) return;
    const key = "red:" + t;
    if (seenRedTexts.has(key)) { seenRedTexts.get(key).count++; return; }
    const f = { key, label: t.length > 40 ? t.slice(0, 40) + "…" : t, type: "red", original: t, value: t, count: 1 };
    seenRedTexts.set(key, f);
    fields.push(f);
  });

  // Bảng khối lượng: có nếu tìm thấy bảng auto HOẶC có loop placeholder {{#kl}}
  const hasKlLoop = /\{\{\s*#kl\s*\}\}/.test(xml) && /\{\{\s*\/kl\s*\}\}/.test(xml);
  const hasKlTable = !!klTable;

  return {
    fields,
    hasKlLoop,
    hasKlTable,
    klRows,
    klColMap: klTable ? klTable.colMap : null,
  };
}

// ============================================================
//  ROOT
// ============================================================
export default function App() {
  const [screen, setScreen] = useState("landing");
  const [logo, setLogo] = useState(null);
  const jszipReady = useScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js", "JSZip");
  const docxReady = useScript("https://cdnjs.cloudflare.com/ajax/libs/docx/8.5.0/docx.umd.js", "docx");
  const xlsxReady = useScript("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js", "XLSX");

  // Danh sách "gói / template"
  const [templates, setTemplates] = useState([]); // {id,name,zipB64,xml,fields,hasKlLoop}
  const [activeTpl, setActiveTpl] = useState(null);

  return (
    <>
      <style>{globalCss}</style>
      {screen === "landing" ? (
        <Landing logo={logo} setLogo={setLogo} onEnter={() => setScreen("app")} />
      ) : (
        <AppShell
          logo={logo}
          setLogo={setLogo}
          onHome={() => setScreen("landing")}
          jszipReady={jszipReady}
          docxReady={docxReady}
          xlsxReady={xlsxReady}
          templates={templates}
          setTemplates={setTemplates}
          activeTpl={activeTpl}
          setActiveTpl={setActiveTpl}
        />
      )}
    </>
  );
}

// ============================================================
//  LANDING
// ============================================================
function Landing({ logo, setLogo, onEnter }) {
  const fileRef = useRef(null);
  const onPick = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    const r = new FileReader();
    r.onload = () => setLogo(r.result);
    r.readAsDataURL(f);
  };
  return (
    <div style={L.page}>
      <span style={{ ...L.blob, width: 360, height: 360, top: 40, left: 80 }} />
      <span style={{ ...L.blob, width: 220, height: 220, top: 260, right: 220, opacity: 0.5 }} />
      <span style={{ ...L.blob, width: 300, height: 300, bottom: -40, right: 80, opacity: 0.6 }} />
      <div style={L.center}>
        <div style={L.logoWrap} onClick={() => fileRef.current?.click()} title="Bấm để tải logo">
          {logo ? <img src={logo} alt="logo" style={L.logoImg} /> : (
            <div style={L.logoPlaceholder}><div style={L.logoMark}>N</div><div style={L.logoText}>TẢI LOGO</div></div>
          )}
        </div>
        <input ref={fileRef} type="file" accept="image/*" onChange={onPick} style={{ display: "none" }} />
        <h1 style={L.h1}>Hồ sơ nghiệm thu</h1>
        <p style={L.sub}>Lập phiếu yêu cầu &amp; biên bản nghiệm thu cho mọi gói thầu — upload form mẫu riêng, app tự dò trường cần điền, xuất Word giữ nguyên định dạng.</p>
        <div style={L.btnRow}>
          <button style={L.btnWhite} onClick={onEnter}>Vào ứng dụng →</button>
          <button style={L.btnGlass} onClick={() => fileRef.current?.click()}>{logo ? "Đổi logo" : "Tải logo"}</button>
        </div>
        <div style={L.cardRow}>
          <FeatureCard icon="⤓" title="Upload form theo gói" desc="Mỗi gói một mẫu riêng" />
          <FeatureCard icon="◧" title="Tự dò trường đỏ" desc="Đỏ &amp; {{placeholder}}" />
          <FeatureCard icon="⎙" title="Xuất Word" desc="Giữ nguyên định dạng gốc" />
        </div>
      </div>
    </div>
  );
}
function FeatureCard({ icon, title, desc }) {
  return (
    <div style={L.card}>
      <div style={L.cardIcon}>{icon}</div>
      <div style={L.cardTitle}>{title}</div>
      <div style={L.cardDesc} dangerouslySetInnerHTML={{ __html: desc }} />
    </div>
  );
}

// ============================================================
//  APP SHELL
// ============================================================
function AppShell({ logo, setLogo, onHome, jszipReady, docxReady, xlsxReady, templates, setTemplates, activeTpl, setActiveTpl }) {
  const [tab, setTab] = useState("ho-so");
  const [records, setRecords] = useState({}); // { [tplId]: [hoSo,...] }
  const [expanded, setExpanded] = useState(null);
  const [exportTarget, setExportTarget] = useState(null);
  const [query, setQuery] = useState("");
  const [busy, setBusy] = useState(false);
  const logoRef = useRef(null);
  const formRef = useRef(null);

  const onPickLogo = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    const r = new FileReader();
    r.onload = () => setLogo(r.result);
    r.readAsDataURL(f);
  };

  // ---- Upload FORM mẫu (.docx) ----
  const onPickForm = async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    e.target.value = "";
    if (!window.JSZip) { alert("Đang tải thư viện, thử lại sau 1 giây."); return; }
    setBusy(true);
    try {
      const buf = await f.arrayBuffer();
      const zip = await window.JSZip.loadAsync(buf);
      const docFile = zip.file("word/document.xml");
      if (!docFile) throw new Error("File không phải Word (.docx) hợp lệ.");
      const xml = await docFile.async("string");
      const { fields, hasKlLoop, hasKlTable, klRows, klColMap } = analyzeTemplate(xml);
      const tpl = {
        id: crypto.randomUUID(),
        name: f.name.replace(/\.docx$/i, ""),
        fileName: f.name,
        zipBuf: buf,        // giữ buffer gốc để đóng gói lại
        xml,
        fields,
        hasKlLoop,
        hasKlTable,
        klRows,
        klColMap,
      };
      setTemplates((t) => [...t, tpl]);
      setActiveTpl(tpl.id);
      setRecords((r) => ({ ...r, [tpl.id]: [makeRecord(tpl)] }));
      if (fields.length === 0 && !hasKlTable && !hasKlLoop) {
        alert("Đã nạp form nhưng không tìm thấy chữ đỏ, bảng khối lượng hay {{placeholder}} nào. Bạn vẫn có thể xuất nguyên bản.");
      }
    } catch (err) {
      alert("Lỗi đọc form: " + err.message);
    } finally {
      setBusy(false);
    }
  };

  const tplList = templates;
  const active = templates.find((t) => t.id === activeTpl) || null;
  const recs = (active && records[active.id]) || [];

  const updateRec = (id, patch) =>
    setRecords((r) => ({ ...r, [active.id]: r[active.id].map((x) => (x.id === id ? { ...x, ...patch } : x)) }));
  const addRec = () => {
    const rec = makeRecord(active);
    setRecords((r) => ({ ...r, [active.id]: [...(r[active.id] || []), rec] }));
    setExpanded(rec.id);
  };
  const removeRec = (id) => {
    setRecords((r) => ({ ...r, [active.id]: r[active.id].filter((x) => x.id !== id) }));
    if (expanded === id) setExpanded(null);
  };
  const dupRec = (id) =>
    setRecords((r) => {
      const src = r[active.id].find((x) => x.id === id);
      const copy = { ...src, id: crypto.randomUUID(), values: { ...src.values }, kl: src.kl.map((k) => ({ ...k, _id: crypto.randomUUID() })) };
      return { ...r, [active.id]: [...r[active.id], copy] };
    });

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return recs;
    return recs.filter((rec) => Object.values(rec.values).some((v) => String(v).toLowerCase().includes(q)));
  }, [recs, query]);

  return (
    <div style={A.page}>
      <header style={A.header}>
        <div style={A.headLeft}>
          <button style={A.homeBtn} onClick={onHome} title="Về trang chủ">↑</button>
          <div style={A.headLogo} onClick={() => logoRef.current?.click()} title="Tải / đổi logo">
            {logo ? <img src={logo} alt="logo" style={A.headLogoImg} /> : <span style={A.headLogoMark}>N</span>}
          </div>
          <input ref={logoRef} type="file" accept="image/*" onChange={onPickLogo} style={{ display: "none" }} />
          <div>
            <div style={A.headTitle}>Bộ phận MEP</div>
            <div style={A.headSub}>Hồ sơ nghiệm thu · {active ? active.name : "chưa chọn gói"}</div>
          </div>
        </div>
        <div style={A.headActions}>
          <input ref={formRef} type="file" accept=".docx" onChange={onPickForm} style={{ display: "none" }} />
          <HeadBtn onClick={() => formRef.current?.click()} tone="amber">⤓ Upload form gói</HeadBtn>
          {active && <HeadBtn onClick={addRec} tone="green">+ Thêm hồ sơ</HeadBtn>}
          <HeadBtn onClick={() => logoRef.current?.click()}>⤒ Logo</HeadBtn>
        </div>
      </header>

      {/* Chọn gói (template) */}
      <div style={A.tplBar}>
        <span style={A.tplLabel}>GÓI:</span>
        {tplList.length === 0 && <span style={A.tplEmpty}>Chưa có form nào — bấm “Upload form gói” để thêm.</span>}
        {tplList.map((t) => (
          <button
            key={t.id}
            style={t.id === activeTpl ? A.tplChipActive : A.tplChip}
            onClick={() => { setActiveTpl(t.id); setExpanded(null); if (!records[t.id]) setRecords((r) => ({ ...r, [t.id]: [makeRecord(t)] })); }}
            title={`${t.fields.length} trường biến đổi`}
          >
            {t.name} <span style={A.tplCount}>{t.fields.length}</span>
          </button>
        ))}
      </div>

      <nav style={A.tabs}>
        <TabBtn active={tab === "ho-so"} onClick={() => setTab("ho-so")} icon="▦">Bảng hồ sơ</TabBtn>
        <TabBtn active={tab === "dashboard"} onClick={() => setTab("dashboard")} icon="▤">Dashboard</TabBtn>
      </nav>

      <main style={A.main}>
        {busy && <div style={A.busy}>Đang đọc form…</div>}

        {!active ? (
          <EmptyState onUpload={() => formRef.current?.click()} />
        ) : tab === "dashboard" ? (
          <Dashboard recs={recs} tpl={active} />
        ) : (
          <>
            <div style={A.searchWrap}>
              <span style={A.searchIcon}>⌕</span>
              <input style={A.search} placeholder="Tìm trong hồ sơ..." value={query} onChange={(e) => setQuery(e.target.value)} />
            </div>

            <div style={A.tableCard}>
              <table style={A.table}>
                <thead>
                  <tr>
                    <th style={{ ...A.th, width: 36 }}></th>
                    {active.fields.slice(0, 3).map((f) => (
                      <th key={f.key} style={A.th}>{f.label}</th>
                    ))}
                    {(active.hasKlLoop || active.hasKlTable) && <th style={{ ...A.thCenter, width: 90 }}>Dòng KL</th>}
                    <th style={{ ...A.thRight, width: 220 }}>Thao tác</th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.length === 0 && (
                    <tr><td colSpan={6} style={A.empty}>Chưa có hồ sơ. Bấm “+ Thêm hồ sơ”.</td></tr>
                  )}
                  {filtered.map((rec) => {
                    const open = expanded === rec.id;
                    return (
                      <React.Fragment key={rec.id}>
                        <tr style={open ? A.trOpen : undefined} className="hsRow">
                          <td style={A.tdToggle}>
                            <button style={A.toggleBtn} onClick={() => setExpanded(open ? null : rec.id)}>
                              <span style={{ display: "inline-block", transform: open ? "rotate(90deg)" : "none", transition: "transform .15s" }}>▶</span>
                            </button>
                          </td>
                          {active.fields.slice(0, 3).map((f) => (
                            <td key={f.key} style={A.td}>
                              <input style={A.cellInput} value={rec.values[f.key] ?? ""} onChange={(e) => updateRec(rec.id, { values: { ...rec.values, [f.key]: e.target.value } })} />
                            </td>
                          ))}
                          {(active.hasKlLoop || active.hasKlTable) && <td style={A.tdCenter}><span style={A.badge}>{rec.kl.filter((k) => !k.header).length}</span></td>}
                          <td style={A.tdRight}>
                            <button style={A.btnExport} onClick={() => setExportTarget(rec)}>Xuất Word</button>
                            <button style={A.btnIcon} title="Nhân bản" onClick={() => dupRec(rec.id)}>⧉</button>
                            <button style={A.btnIconDanger} title="Xóa" onClick={() => removeRec(rec.id)}>✕</button>
                          </td>
                        </tr>
                        {open && (
                          <tr><td colSpan={6} style={A.expandCell}>
                            <RecordEditor rec={rec} tpl={active} xlsxReady={xlsxReady} onChange={(patch) => updateRec(rec.id, patch)} />
                          </td></tr>
                        )}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </>
        )}
      </main>

      {exportTarget && active && (
        <ExportModal
          rec={exportTarget} tpl={active} docxReady={docxReady} jszipReady={jszipReady}
          onClose={() => setExportTarget(null)}
          onSave={(patch) => { updateRec(exportTarget.id, patch); setExportTarget((t) => ({ ...t, ...patch })); }}
        />
      )}
    </div>
  );
}

function blankKl(header) {
  return { _id: crypto.randomUUID(), stt: "", ten: "", nhan: "", sl: "", dv: "", ghiChu: "", header: !!header };
}
function makeRecord(tpl) {
  const values = {};
  tpl.fields.forEach((f) => { values[f.key] = f.type === "red" ? f.original : ""; });
  let kl;
  if (tpl.klRows && tpl.klRows.length) {
    kl = tpl.klRows.map((r) => ({ ...r, _id: crypto.randomUUID() }));
  } else {
    kl = [{ _id: crypto.randomUUID(), stt: "1", ten: "Vật liệu mẫu", nhan: "", sl: "100", dv: "m", ghiChu: "", header: false }];
  }
  return { id: crypto.randomUUID(), values, kl };
}

function HeadBtn({ children, onClick, tone }) {
  const map = { green: A.headBtnGreen, amber: A.headBtnAmber };
  return <button style={map[tone] || A.headBtn} onClick={onClick}>{children}</button>;
}
function TabBtn({ children, active, onClick, icon }) {
  return <button style={active ? A.tabActive : A.tab} onClick={onClick}><span style={{ marginRight: 7 }}>{icon}</span>{children}</button>;
}

function EmptyState({ onUpload }) {
  return (
    <div style={A.emptyState}>
      <div style={A.emptyIcon}>⤓</div>
      <h2 style={{ margin: "8px 0 6px", color: "#15324a" }}>Bắt đầu bằng việc upload form mẫu</h2>
      <p style={{ color: "#5d7589", maxWidth: 520, margin: "0 auto 18px", lineHeight: 1.6 }}>
        Mỗi gói thầu có một form Word riêng. Tải file .docx của gói lên — app sẽ tự nhận diện các phần
        <b style={{ color: "#c0392b" }}> chữ đỏ</b> và <b>{"{{placeholder}}"}</b> làm trường cần điền, giữ nguyên mọi định dạng khác.
      </p>
      <button style={A.btnPrimary} onClick={onUpload}>Upload form gói (.docx)</button>
      <div style={A.tipBox}>
        <b>Mẹo soạn form:</b> Tô <span style={{ color: "#c0392b" }}>đỏ</span> những chỗ thay đổi đơn lẻ (số phiếu, ngày, tên vật liệu…) — mỗi cụm đỏ thành một ô nhập.
        <br /><b>Bảng khối lượng:</b> đừng tô đỏ từng ô. Thay vào đó, ở <i>hàng vật liệu mẫu</i> đặt <code>{"{{#kl}}"}</code> ở ô đầu và <code>{"{{/kl}}"}</code> ở ô cuối,
        rồi điền các ô bằng <code>{"{{stt}}"}</code> <code>{"{{ten}}"}</code> <code>{"{{nhan}}"}</code> <code>{"{{sl}}"}</code> <code>{"{{dv}}"}</code> <code>{"{{ghiChu}}"}</code>.
        App sẽ tự nhân hàng theo số dòng bạn nhập và cho phép dán từ Excel.
      </div>
    </div>
  );
}

// ============================================================
//  EDITOR HỒ SƠ
// ============================================================
// ============================================================
const KL_COLS = ["stt", "ten", "nhan", "sl", "dv", "ghiChu"];
function RecordEditor({ rec, tpl, xlsxReady, onChange }) {
  const setVal = (key, v) => onChange({ values: { ...rec.values, [key]: v } });
  const xlsxInputRef = useRef(null);
  const [sel, setSel] = useState({ row: 0, col: 1 }); // ô đang chọn (mặc định cột Tên)

  // bảng khối lượng
  const rows = rec.kl;
  const setRows = (kl) => onChange({ kl });
  const setRow = (i, patch) => setRows(rows.map((r, idx) => (idx === i ? { ...r, ...patch } : r)));
  const addRow = () => setRows([...rows, blankKl(false)]);
  const addHeader = () => setRows([...rows, { ...blankKl(true), ten: "TIÊU ĐỀ NHÓM" }]);
  const delRow = (i) => setRows(rows.filter((_, idx) => idx !== i));
  const moveRow = (i, dir) => { const j = i + dir; if (j < 0 || j >= rows.length) return; const n = [...rows]; [n[i], n[j]] = [n[j], n[i]]; setRows(n); };

  // Dán một DẢI Ô từ Excel bắt đầu tại ô đang chọn (tràn sang phải & xuống dưới, tự thêm dòng)
  const pasteRange = (e, startRow, startCol) => {
    const text = e.clipboardData.getData("text/plain");
    if (!text) return;
    const grid = text.replace(/\r/g, "").replace(/\n$/, "").split("\n").map((l) => l.split("\t"));
    if (grid.length === 1 && grid[0].length === 1) return; // 1 ô đơn -> để input tự xử lý
    e.preventDefault();
    const next = rows.map((r) => ({ ...r }));
    grid.forEach((line, ri) => {
      const tr = startRow + ri;
      while (tr >= next.length) next.push(blankKl(false));
      line.forEach((val, ci) => {
        const tc = startCol + ci;
        if (tc >= KL_COLS.length) return;
        next[tr][KL_COLS[tc]] = String(val).trim();
      });
      // suy đoán dòng nhóm: có tên nhưng không có nhãn & số lượng
      const rr = next[tr];
      if (!rr.nhan && !rr.sl && !rr.dv && rr.ten) rr.header = true;
    });
    setRows(next);
  };

  const onCellKey = (e, ri, ci) => {
    const max = rows.length - 1;
    if (e.key === "ArrowDown") { e.preventDefault(); setSel({ row: Math.min(max, ri + 1), col: ci }); focusCell(ri + 1, ci); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setSel({ row: Math.max(0, ri - 1), col: ci }); focusCell(ri - 1, ci); }
    else if (e.key === "Enter") { e.preventDefault(); if (ri === max) addRow(); setSel({ row: ri + 1, col: ci }); setTimeout(() => focusCell(ri + 1, ci), 0); }
  };
  const focusCell = (ri, ci) => {
    const el = document.querySelector(`[data-cell="${rec.id}-${ri}-${ci}"]`);
    if (el) el.focus();
  };


  // ---- Xuất phiếu Excel mẫu để điền khối lượng ----
  const exportKlTemplate = () => {
    const XLSX = window.XLSX;
    if (!XLSX) return;
    const recName = rec.values[tpl.fields[0]?.key] || tpl.name;
    // Hàng hướng dẫn + tiêu đề + dữ liệu hiện có
    const aoa = [
      ["PHIẾU NHẬP KHỐI LƯỢNG", "", "", "", "", "", ""],
      ["Gói:", tpl.name, "", "Hồ sơ:", String(recName), "", ""],
      ["Hướng dẫn: cột Loại = 'nhom' cho dòng tiêu đề nhóm; để trống cho dòng vật liệu. Không xóa/đổi hàng tiêu đề cột bên dưới.", "", "", "", "", "", ""],
      ["STT", "Tên vật liệu", "Nhãn hiệu", "Số lượng", "Đơn vị", "Ghi chú", "Loại"],
      ...rows.map((r) => [r.stt, r.ten, r.header ? "" : r.nhan, r.header ? "" : r.sl, r.header ? "" : r.dv, r.ghiChu, r.header ? "nhom" : ""]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 6 }, { wch: 34 }, { wch: 16 }, { wch: 12 }, { wch: 10 }, { wch: 20 }, { wch: 8 }];
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: 6 } },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "KhoiLuong");
    XLSX.writeFile(wb, `KhoiLuong_${String(recName).replace(/[^\w.-]+/g, "_")}.xlsx`);
  };

  // ---- Nhập lại phiếu Excel đã điền ----
  const importKl = async (e) => {
    const f = e.target.files?.[0];
    e.target.value = "";
    if (!f) return;
    const XLSX = window.XLSX;
    if (!XLSX) { alert("Đang tải thư viện Excel, thử lại sau giây lát."); return; }
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      // tìm hàng tiêu đề cột (ô đầu = "STT")
      let hdr = aoa.findIndex((row) => String(row[0]).trim().toUpperCase() === "STT");
      if (hdr === -1) { alert("Không tìm thấy hàng tiêu đề (ô đầu là 'STT'). Hãy dùng đúng file mẫu xuất ra."); return; }
      const body = aoa.slice(hdr + 1).filter((row) => row.some((c) => String(c).trim() !== ""));
      const parsed = body.map((c) => {
        const loai = String(c[6] ?? "").trim().toLowerCase();
        const header = loai === "nhom" || loai === "nhóm" || (!String(c[2] ?? "").trim() && !String(c[3] ?? "").trim() && String(c[1] ?? "").trim() !== "" && String(c[0] ?? "").trim() === "");
        return {
          _id: crypto.randomUUID(),
          stt: String(c[0] ?? "").trim(),
          ten: String(c[1] ?? "").trim(),
          nhan: header ? "" : String(c[2] ?? "").trim(),
          sl: header ? "" : String(c[3] ?? "").trim(),
          dv: header ? "" : String(c[4] ?? "").trim(),
          ghiChu: String(c[5] ?? "").trim(),
          header,
        };
      });
      if (parsed.length === 0) { alert("File không có dòng dữ liệu nào."); return; }
      setRows(parsed);
    } catch (err) {
      alert("Lỗi đọc Excel: " + err.message);
    }
  };

  return (
    <div style={A.editor}>
      <div style={A.editorHead}><strong style={{ fontSize: 13, color: "#0e4d6b" }}>TRƯỜNG BIẾN ĐỔI ({tpl.fields.length})</strong></div>
      <div style={A.fieldGrid}>
        {tpl.fields.map((f) => (
          <Field key={f.key} label={f.label + (f.type === "placeholder" ? "  ·{{}}" : "  ·đỏ")}>
            <input style={A.input} value={rec.values[f.key] ?? ""} onChange={(e) => setVal(f.key, e.target.value)} />
          </Field>
        ))}
        {tpl.fields.length === 0 && <div style={{ color: "#7790a3", fontSize: 13 }}>Form này không có trường biến đổi — xuất nguyên bản.</div>}
      </div>

      {(tpl.hasKlLoop || tpl.hasKlTable) && (
        <>
          <div style={A.editorHead}>
            <strong style={{ fontSize: 13, color: "#0e4d6b" }}>BẢNG KHỐI LƯỢNG</strong>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button style={A.btnGhostSm} onClick={addRow}>+ Dòng</button>
              <button style={A.btnGhostSm} onClick={addHeader}>+ Tiêu đề nhóm</button>
              <button style={A.btnExcelOut} disabled={!xlsxReady} onClick={exportKlTemplate} title="Tải file Excel mẫu để điền khối lượng">⬇ Xuất Excel mẫu</button>
              <button style={A.btnExcelIn} disabled={!xlsxReady} onClick={() => xlsxInputRef.current?.click()} title="Nhập lại file Excel đã điền">⬆ Nhập Excel</button>
              <input ref={xlsxInputRef} type="file" accept=".xlsx,.xls" onChange={importKl} style={{ display: "none" }} />
            </div>
          </div>
          <div style={A.pasteHint}>
            <b>Dán dải ô:</b> bấm vào một ô rồi Ctrl+V — dữ liệu từ Excel sẽ tràn sang phải &amp; xuống dưới như trong Excel (tự thêm dòng).
            Cột: STT · Tên vật liệu · Nhãn hiệu · Số lượng · Đơn vị · Ghi chú. Dòng chỉ có tên (không nhãn/SL) sẽ thành <span style={{ color: ACCENT, fontWeight: 700 }}>tiêu đề nhóm</span>.
          </div>

          <div style={A.sheetScroll}>
            <table style={A.sheet}>
              <thead>
                <tr>
                  <th style={{ ...A.shTh, width: 56 }}>STT</th>
                  <th style={A.shTh}>Tên vật liệu</th>
                  <th style={{ ...A.shTh, width: 130 }}>Nhãn hiệu</th>
                  <th style={{ ...A.shTh, width: 90 }}>Số lượng</th>
                  <th style={{ ...A.shTh, width: 80 }}>Đơn vị</th>
                  <th style={A.shTh}>Ghi chú</th>
                  <th style={{ ...A.shTh, width: 78, borderRight: "none" }}></th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r, ri) => (
                  <tr key={r._id}>
                    {KL_COLS.map((col, ci) => {
                      const isSel = sel.row === ri && sel.col === ci;
                      const disabled = r.header && (col === "nhan" || col === "sl" || col === "dv");
                      const center = col === "stt" || col === "nhan" || col === "sl" || col === "dv";
                      return (
                        <td key={col} style={{ ...A.shTd, ...(isSel ? A.shTdSel : null) }}>
                          <input
                            data-cell={`${rec.id}-${ri}-${ci}`}
                            value={r[col]}
                            disabled={disabled}
                            onFocus={() => setSel({ row: ri, col: ci })}
                            onChange={(e) => setRow(ri, { [col]: e.target.value })}
                            onPaste={(e) => pasteRange(e, ri, ci)}
                            onKeyDown={(e) => onCellKey(e, ri, ci)}
                            style={{
                              ...A.shInput,
                              textAlign: center ? "center" : "left",
                              fontWeight: r.header && col === "ten" ? 700 : 400,
                              color: r.header ? ACCENT : "#15324a",
                              background: disabled ? "#fafbfc" : "transparent",
                            }}
                          />
                        </td>
                      );
                    })}
                    <td style={A.shActions}>
                      <button style={r.header ? A.tagBtnOn : A.tagBtn} title="Đổi dòng tiêu đề nhóm" onClick={() => setRow(ri, { header: !r.header })}>N</button>
                      <button style={A.miniBtn} onClick={() => moveRow(ri, -1)} title="Lên">↑</button>
                      <button style={A.miniBtn} onClick={() => moveRow(ri, 1)} title="Xuống">↓</button>
                      <button style={A.miniBtnDanger} onClick={() => delRow(ri)} title="Xóa dòng">✕</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ marginTop: 8 }}>
            <button style={A.btnGhostSm} onClick={addRow}>+ Thêm dòng cuối</button>
          </div>
        </>
      )}
    </div>
  );
}

function Field({ label, children }) {
  return <label style={A.field}><span style={A.fieldLabel}>{label}</span>{children}</label>;
}

// ============================================================
//  DASHBOARD
// ============================================================
function Dashboard({ recs, tpl }) {
  const total = recs.length;
  const totalKl = recs.reduce((s, r) => s + (r.kl?.filter((k) => !k.header).length || 0), 0);
  return (
    <div style={A.dashGrid}>
      <StatCard label="Số hồ sơ trong gói" value={total} unit="bộ" />
      <StatCard label="Trường biến đổi / hồ sơ" value={tpl.fields.length} unit="trường" />
      <StatCard label="Tổng dòng khối lượng" value={totalKl} unit="dòng" />
    </div>
  );
}
function StatCard({ label, value, unit }) {
  return <div style={A.statCard}><div style={A.statLabel}>{label}</div><div style={A.statValue}>{value} <span style={A.statUnit}>{unit}</span></div></div>;
}

// ============================================================
//  EXPORT MODAL
// ============================================================
function ExportModal({ rec, tpl, docxReady, jszipReady, onClose, onSave }) {
  const [form, setForm] = useState(rec);
  const [exporting, setExporting] = useState(false);
  const [err, setErr] = useState("");
  const setVal = (key, v) => setForm((f) => ({ ...f, values: { ...f.values, [key]: v } }));

  const missing = useMemo(() => tpl.fields.filter((f) => !String(form.values[f.key] ?? "").trim()).map((f) => f.label), [form, tpl]);

  const doExport = async () => {
    setErr("");
    if (!window.JSZip) { setErr("Bộ xử lý file (JSZip) chưa tải xong — kiểm tra kết nối mạng rồi thử lại."); return; }
    setExporting(true);
    try {
      onSave({ values: form.values, kl: form.kl });
      await exportFromTemplate(tpl, form);
      onClose();
    } catch (e) {
      setErr("Không xuất được: " + (e?.message || e));
      setExporting(false);
    }
  };

  return (
    <div style={A.overlay} onMouseDown={onClose}>
      <div style={A.modal} onMouseDown={(e) => e.stopPropagation()}>
        <div style={A.modalHead}>
          <h2 style={{ margin: 0, fontSize: 18, color: "#0e4d6b" }}>Nhập thông tin trước khi xuất</h2>
          <button style={A.btnIcon} onClick={onClose}>✕</button>
        </div>
        <p style={A.modalSub}>Gói: <b>{tpl.name}</b> · {tpl.fields.length} trường biến đổi (đỏ / placeholder). Phần còn lại giữ nguyên theo form.</p>
        <div style={A.modalGrid}>
          {tpl.fields.map((f) => (
            <Field key={f.key} label={f.label}>
              <input style={A.input} value={form.values[f.key] ?? ""} onChange={(e) => setVal(f.key, e.target.value)} />
            </Field>
          ))}
        </div>
        {missing.length > 0 && <div style={A.warn}>Còn trống: {missing.join(", ")} — bạn vẫn xuất được, ô trống sẽ để rỗng trong file.</div>}
        {err && <div style={A.warn}>{err}</div>}
        <div style={A.modalFoot}>
          <button style={A.btnGhost} onClick={onClose}>Hủy</button>
          <button style={!jszipReady || exporting ? A.btnDisabled : A.btnPrimary} disabled={!jszipReady || exporting} onClick={doExport}>
            {exporting ? "Đang tạo file..." : jszipReady ? "Xuất file Word (.docx)" : "Đang tải bộ xuất..."}
          </button>
        </div>
      </div>
    </div>
  );
}

// ---- Thay text trong document.xml gốc rồi đóng gói lại ----
async function exportFromTemplate(tpl, form) {
  const JSZip = window.JSZip;
  const zip = await JSZip.loadAsync(tpl.zipBuf);
  let xml = await zip.file("word/document.xml").async("string");

  // 1) Bảng khối lượng lặp {{#kl}}...{{/kl}}
  if (tpl.hasKlLoop) {
    xml = expandKlLoop(xml, form.kl);
  } else if (tpl.hasKlTable) {
    xml = rebuildKlTable(xml, form.kl);
  }

  // 2) Placeholder thường {{name}}
  tpl.fields.filter((f) => f.type === "placeholder").forEach((f) => {
    const val = encodeXml(form.values[f.key] ?? "");
    xml = xml.split(f.token).join(val);
  });

  // 3) Trường đỏ: thay text trong các run đỏ có nội dung gốc khớp.
  //    Thực hiện ở cấp run để không phá XML.
  tpl.fields.filter((f) => f.type === "red").forEach((f) => {
    const newVal = form.values[f.key] ?? "";
    if (newVal === f.original) return; // không đổi
    xml = replaceRedRunText(xml, f.original, newVal);
  });

  zip.file("word/document.xml", xml);
  const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const first = form.values[tpl.fields[0]?.key] || tpl.name;
  a.href = url;
  a.download = `${String(first).replace(/[^\w.-]+/g, "_")}.docx`;
  a.click();
  URL.revokeObjectURL(url);
}

// Thay text của 1 nhóm run đỏ liền kề (gộp lại bằng original) bằng newVal.
// Chiến lược: tìm chuỗi run đỏ liên tiếp mà nối text = original, đặt newVal vào run đầu, xóa text run sau.
function replaceRedRunText(xml, original, newVal) {
  const runRe = /<w:r\b[\s\S]*?<\/w:r>/g;
  const runs = [];
  let m;
  while ((m = runRe.exec(xml))) {
    runs.push({ start: m.index, end: m.index + m[0].length, xml: m[0] });
  }
  // tìm dãy run đỏ liên tiếp có text nối = original
  for (let i = 0; i < runs.length; i++) {
    let acc = "";
    let j = i;
    const group = [];
    while (j < runs.length) {
      const col = runColor(runs[j].xml);
      const txt = decodeXml(runText(runs[j].xml));
      if (!isRed(col) || txt === "") break;
      acc += txt;
      group.push(runs[j]);
      if (acc === original) break;
      if (!original.startsWith(acc)) { acc = ""; group.length = 0; break; }
      j++;
    }
    if (acc === original && group.length) {
      // build new XML: run đầu giữ rPr, đặt newVal; các run sau -> text rỗng
      let out = xml.slice(0, group[0].start);
      const firstRun = group[0].xml;
      const newFirst = setRunText(firstRun, newVal);
      out += newFirst;
      for (let k = 1; k < group.length; k++) {
        out += setRunText(group[k].xml, "");
      }
      out += xml.slice(group[group.length - 1].end);
      return out;
    }
  }
  return xml;
}

function setRunText(runXml, text) {
  const enc = encodeXml(text);
  // có sẵn <w:t>?
  if (/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/.test(runXml)) {
    let replaced = false;
    return runXml.replace(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g, () => {
      if (!replaced) { replaced = true; return `<w:t xml:space="preserve">${enc}</w:t>`; }
      return "";
    });
  }
  // không có -> chèn trước </w:r>
  return runXml.replace(/<\/w:r>/, `<w:t xml:space="preserve">${enc}</w:t></w:r>`);
}

// Mở rộng hàng lặp bảng khối lượng
function expandKlLoop(xml, kl) {
  const startTok = "{{#kl}}";
  const endTok = "{{/kl}}";
  const sIdx = xml.indexOf(startTok);
  const eIdx = xml.indexOf(endTok);
  if (sIdx === -1 || eIdx === -1) return xml;
  // tìm hàng <w:tr> bao quanh marker
  const trStart = xml.lastIndexOf("<w:tr", sIdx);
  const trEnd = xml.indexOf("</w:tr>", eIdx) + "</w:tr>".length;
  if (trStart === -1 || trEnd === -1) return xml;
  let template = xml.slice(trStart, trEnd).split(startTok).join("").split(endTok).join("");
  const built = kl.map((r) => {
    let row = template;
    const map = { stt: r.stt, ten: r.ten, nhan: r.header ? "" : r.nhan, sl: r.header ? "" : r.sl, dv: r.header ? "" : r.dv, ghiChu: r.ghiChu };
    Object.keys(map).forEach((k) => { row = row.split("{{" + k + "}}").join(encodeXml(map[k] ?? "")); });
    return row;
  }).join("");
  return xml.slice(0, trStart) + built + xml.slice(trEnd);
}

// Thay TEXT của 1 ô <w:tc>: đặt giá trị vào run đầu tiên (giữ rPr), xóa text các run còn lại.
// Nếu ô không có run nào -> chèn 1 run tối thiểu vào <w:p> đầu.
function setCellText(tcXml, text) {
  const enc = encodeXml(text);
  const runs = tcXml.match(/<w:r\b[\s\S]*?<\/w:r>/g);
  if (runs && runs.length) {
    let done = false;
    return tcXml.replace(/<w:r\b[\s\S]*?<\/w:r>/g, (run) => {
      // bỏ qua run không chứa <w:t> (vd: chỉ có <w:br/>)? vẫn xử lý run đầu có text được
      if (!done) {
        done = true;
        return setRunText(run, text);
      }
      // các run sau: nếu có text thì làm rỗng, giữ run không-text nguyên trạng
      if (/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/.test(run)) return setRunText(run, "");
      return run;
    });
  }
  // không có run: chèn vào <w:p> đầu tiên
  return tcXml.replace(/<\/w:p>/, `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:color w:val="C00000"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${enc}</w:t></w:r></w:p>`);
}

// Dựng lại bảng khối lượng từ các dòng người dùng, dùng hàng mẫu của form để giữ định dạng.
function rebuildKlTable(xml, kl) {
  const info = findKlTable(xml);
  if (!info) return xml;
  const { tblXml, start, end, rows, colMap } = info;
  if (rows.length < 2) return xml; // không có hàng dữ liệu mẫu

  // chọn hàng mẫu: 1 hàng dữ liệu thường + 1 hàng nhóm (nếu có)
  let dataTemplate = null, headerTemplate = null;
  for (let i = 1; i < rows.length; i++) {
    const cells = splitCells(rows[i].xml);
    const get = (k) => (colMap[k] != null && cells[colMap[k]] != null ? cellText(cells[colMap[k]]) : "");
    const isHdr = get("ten") !== "" && get("nhan") === "" && get("sl") === "";
    if (isHdr && !headerTemplate) headerTemplate = rows[i].xml;
    if (!isHdr && !dataTemplate) dataTemplate = rows[i].xml;
  }
  if (!dataTemplate) dataTemplate = rows[1].xml;
  if (!headerTemplate) headerTemplate = dataTemplate;

  const buildRow = (r) => {
    const tplRow = r.header ? headerTemplate : dataTemplate;
    const cells = splitCells(tplRow);
    const vals = {
      stt: r.stt,
      ten: r.ten,
      nhan: r.header ? "" : r.nhan,
      sl: r.header ? "" : r.sl,
      dv: r.header ? "" : r.dv,
      ghiChu: r.ghiChu,
    };
    let newRow = tplRow;
    // thay từng ô theo index cột (thay từ phải sang trái để index không lệch)
    const entries = Object.keys(vals)
      .filter((k) => colMap[k] != null)
      .map((k) => ({ k, idx: colMap[k] }))
      .sort((a, b) => b.idx - a.idx);
    for (const { k, idx } of entries) {
      const cell = cells[idx];
      if (cell == null) continue;
      const newCell = setCellText(cell, vals[k] ?? "");
      newRow = newRow.replace(cell, newCell);
    }
    return newRow;
  };

  const headerRow = rows[0].xml;
  const body = kl.map(buildRow).join("");
  const beforeRows = tblXml.slice(0, rows[0].start);
  const rebuilt = beforeRows + headerRow + body + "</w:tbl>";
  return xml.slice(0, start) + rebuilt + xml.slice(end);
}

// ============================================================
//  STYLES
// ============================================================
const BLUE = "#1390b0", BLUE_DK = "#0e6f8a", BLUE_DKR = "#0c5e75";
const INK = "#15324a", LINE = "#e6edf2", ACCENT = "#c0392b", AMBER = "#d98324";

const L = {
  page: { position: "relative", minHeight: "100vh", overflow: "hidden", background: `radial-gradient(120% 120% at 50% 0%, #1aa3c4 0%, ${BLUE} 45%, ${BLUE_DK} 100%)`, fontFamily: "'Inter', system-ui, sans-serif", display: "flex", alignItems: "center", justifyContent: "center", padding: "48px 20px" },
  blob: { position: "absolute", borderRadius: "50%", background: "rgba(255,255,255,.08)", pointerEvents: "none" },
  center: { position: "relative", zIndex: 2, textAlign: "center", color: "#fff", maxWidth: 760, width: "100%" },
  logoWrap: { width: 130, height: 130, margin: "0 auto 18px", borderRadius: 24, background: "rgba(255,255,255,.12)", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", border: "1.5px dashed rgba(255,255,255,.5)", overflow: "hidden" },
  logoImg: { width: "100%", height: "100%", objectFit: "contain", padding: 8 },
  logoPlaceholder: { textAlign: "center", color: "rgba(255,255,255,.92)" },
  logoMark: { fontSize: 46, fontWeight: 900, lineHeight: 1, fontFamily: "Georgia, serif" },
  logoText: { fontSize: 11, letterSpacing: ".18em", marginTop: 6, fontWeight: 700 },
  h1: { fontSize: 50, fontWeight: 800, margin: "6px 0 14px", letterSpacing: "-.02em" },
  sub: { fontSize: 17, lineHeight: 1.6, color: "rgba(255,255,255,.88)", maxWidth: 580, margin: "0 auto 30px" },
  btnRow: { display: "flex", gap: 14, justifyContent: "center", flexWrap: "wrap", marginBottom: 46 },
  btnWhite: { background: "#fff", color: BLUE_DK, border: "none", borderRadius: 12, padding: "15px 30px", fontSize: 16, fontWeight: 800, cursor: "pointer", fontFamily: "inherit", boxShadow: "0 8px 24px rgba(0,0,0,.18)" },
  btnGlass: { background: "rgba(255,255,255,.16)", color: "#fff", border: "1.5px solid rgba(255,255,255,.4)", borderRadius: 12, padding: "15px 26px", fontSize: 16, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" },
  cardRow: { display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16, maxWidth: 660, margin: "0 auto" },
  card: { background: "rgba(255,255,255,.12)", border: "1px solid rgba(255,255,255,.2)", borderRadius: 16, padding: "20px 18px", textAlign: "left" },
  cardIcon: { fontSize: 22, marginBottom: 10 },
  cardTitle: { fontSize: 15.5, fontWeight: 800, marginBottom: 4 },
  cardDesc: { fontSize: 13, color: "rgba(255,255,255,.82)", lineHeight: 1.45 },
};

const A = {
  page: { minHeight: "100vh", background: "#f1f5f8", fontFamily: "'Inter', system-ui, sans-serif", color: INK },
  header: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 16, flexWrap: "wrap", padding: "12px 22px", background: `linear-gradient(135deg, ${BLUE} 0%, ${BLUE_DKR} 100%)`, color: "#fff" },
  headLeft: { display: "flex", alignItems: "center", gap: 14 },
  homeBtn: { width: 34, height: 34, borderRadius: 9, border: "1px solid rgba(255,255,255,.35)", background: "rgba(255,255,255,.12)", color: "#fff", cursor: "pointer", fontSize: 16 },
  headLogo: { width: 44, height: 44, borderRadius: 10, background: "rgba(255,255,255,.15)", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", overflow: "hidden", border: "1px dashed rgba(255,255,255,.4)" },
  headLogoImg: { width: "100%", height: "100%", objectFit: "contain", padding: 4 },
  headLogoMark: { fontSize: 24, fontWeight: 900, fontFamily: "Georgia, serif" },
  headTitle: { fontSize: 18, fontWeight: 800, lineHeight: 1.1 },
  headSub: { fontSize: 12, color: "rgba(255,255,255,.85)", marginTop: 2 },
  headActions: { display: "flex", gap: 8, flexWrap: "wrap" },
  headBtn: { background: "rgba(255,255,255,.14)", color: "#fff", border: "1px solid rgba(255,255,255,.3)", borderRadius: 9, padding: "9px 15px", fontSize: 13.5, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" },
  headBtnGreen: { background: "#1f9d57", color: "#fff", border: "1px solid #1f9d57", borderRadius: 9, padding: "9px 16px", fontSize: 13.5, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" },
  headBtnAmber: { background: AMBER, color: "#fff", border: `1px solid ${AMBER}`, borderRadius: 9, padding: "9px 16px", fontSize: 13.5, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" },

  tplBar: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", padding: "10px 22px", background: "#eaf4f8", borderBottom: `1px solid ${LINE}` },
  tplLabel: { fontSize: 11, fontWeight: 800, letterSpacing: ".08em", color: "#5d7589" },
  tplEmpty: { fontSize: 13, color: "#7790a3" },
  tplChip: { background: "#fff", color: "#42627a", border: `1px solid ${LINE}`, borderRadius: 20, padding: "6px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" },
  tplChipActive: { background: BLUE, color: "#fff", border: `1px solid ${BLUE}`, borderRadius: 20, padding: "6px 14px", fontSize: 13, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" },
  tplCount: { display: "inline-block", marginLeft: 6, background: "rgba(0,0,0,.12)", borderRadius: 10, padding: "0 7px", fontSize: 11 },

  tabs: { display: "flex", gap: 4, background: "#fff", padding: "0 22px", borderBottom: `1px solid ${LINE}` },
  tab: { background: "transparent", border: "none", borderBottom: "3px solid transparent", padding: "14px 16px", fontSize: 14, fontWeight: 600, color: "#5d7589", cursor: "pointer", fontFamily: "inherit" },
  tabActive: { background: "transparent", border: "none", borderBottom: `3px solid ${BLUE}`, padding: "14px 16px", fontSize: 14, fontWeight: 800, color: BLUE_DK, cursor: "pointer", fontFamily: "inherit" },

  main: { padding: "22px", maxWidth: 1400, margin: "0 auto" },
  busy: { padding: "10px 16px", background: "#fff7e8", border: `1px solid #f0dcb4`, borderRadius: 10, color: "#8a6516", fontSize: 13, marginBottom: 14 },
  emptyState: { textAlign: "center", padding: "50px 20px", background: "#fff", border: `1px solid ${LINE}`, borderRadius: 16 },
  emptyIcon: { fontSize: 46, color: BLUE },
  tipBox: { marginTop: 26, textAlign: "left", maxWidth: 640, marginInline: "auto", background: "#f4fafc", border: `1px solid ${LINE}`, borderRadius: 12, padding: "14px 16px", fontSize: 12.5, color: "#42627a", lineHeight: 1.7 },

  searchWrap: { position: "relative", marginBottom: 16 },
  searchIcon: { position: "absolute", left: 16, top: "50%", transform: "translateY(-50%)", color: "#8aa0b2", fontSize: 17 },
  search: { width: "100%", border: `1px solid ${LINE}`, borderRadius: 12, padding: "13px 16px 13px 42px", fontSize: 14.5, fontFamily: "inherit", background: "#fff", boxSizing: "border-box" },

  tableCard: { background: "#fff", border: `1px solid ${LINE}`, borderRadius: 14, overflow: "hidden", boxShadow: "0 2px 8px rgba(20,50,74,.05)" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13.5 },
  th: { textAlign: "left", fontSize: 11, letterSpacing: ".05em", textTransform: "uppercase", color: "#7790a3", padding: "13px 14px", borderBottom: `1px solid ${LINE}`, fontWeight: 800, background: "#f8fafc" },
  thCenter: { textAlign: "center", fontSize: 11, letterSpacing: ".05em", textTransform: "uppercase", color: "#7790a3", padding: "13px 14px", borderBottom: `1px solid ${LINE}`, fontWeight: 800, background: "#f8fafc" },
  thRight: { textAlign: "right", fontSize: 11, letterSpacing: ".05em", textTransform: "uppercase", color: "#7790a3", padding: "13px 14px", borderBottom: `1px solid ${LINE}`, fontWeight: 800, background: "#f8fafc" },
  td: { padding: "8px 10px", borderBottom: `1px solid ${LINE}`, verticalAlign: "middle" },
  tdNoWrap: { padding: "10px 14px", borderBottom: `1px solid ${LINE}`, whiteSpace: "nowrap", color: "#42627a" },
  tdCenter: { padding: "10px 14px", borderBottom: `1px solid ${LINE}`, textAlign: "center" },
  tdRight: { padding: "8px 14px", borderBottom: `1px solid ${LINE}`, textAlign: "right", whiteSpace: "nowrap" },
  tdToggle: { padding: "8px 6px 8px 16px", borderBottom: `1px solid ${LINE}`, width: 36 },
  trOpen: { background: "#f4fafc" },
  empty: { padding: "40px 20px", textAlign: "center", color: "#7790a3", fontSize: 14 },
  toggleBtn: { border: "none", background: "transparent", cursor: "pointer", fontSize: 11, color: BLUE, padding: 4 },
  badge: { display: "inline-block", minWidth: 24, padding: "3px 9px", background: "#e3f1f6", borderRadius: 20, fontSize: 12, fontWeight: 800, color: BLUE_DK },
  cellInput: { width: "100%", border: "1px solid transparent", borderRadius: 7, padding: "7px 9px", fontSize: 13.5, fontFamily: "inherit", background: "transparent", color: INK },
  btnExport: { background: BLUE, color: "#fff", border: "none", borderRadius: 8, padding: "7px 13px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", marginRight: 6 },
  btnIcon: { background: "#fff", border: `1px solid ${LINE}`, borderRadius: 7, width: 30, height: 30, cursor: "pointer", color: "#5d7589", marginLeft: 4, fontSize: 13 },
  btnIconDanger: { background: "#fff", border: "1px solid #f3cdc8", borderRadius: 7, width: 30, height: 30, cursor: "pointer", color: ACCENT, marginLeft: 4, fontSize: 13 },
  expandCell: { padding: 0, background: "#f4fafc", borderBottom: `2px solid ${LINE}` },

  editor: { padding: "18px 20px 22px" },
  fieldGrid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 },
  editorHead: { display: "flex", justifyContent: "space-between", alignItems: "center", margin: "6px 0 10px" },
  pasteZone: { border: `1.5px dashed ${BLUE}`, borderRadius: 9, padding: "10px 14px", fontSize: 12, color: BLUE_DK, background: "#eef8fb", marginBottom: 12, outline: "none" },
  klTable: { width: "100%", borderCollapse: "collapse", fontSize: 13, background: "#fff" },
  klTh: { textAlign: "left", fontSize: 10.5, letterSpacing: ".03em", textTransform: "uppercase", color: "#7790a3", padding: "7px 8px", borderBottom: `1px solid ${LINE}`, fontWeight: 800 },
  klTd: { padding: "2px 4px", borderBottom: "1px solid #eef3f6" },
  klTdMini: { padding: "2px 4px", borderBottom: "1px solid #eef3f6", textAlign: "center", width: 30 },
  klTdActions: { padding: "2px 4px", borderBottom: "1px solid #eef3f6", whiteSpace: "nowrap", width: 86 },
  klRowHeader: { background: "#fbe9e7" },
  klInput: { width: "100%", border: `1px solid ${LINE}`, borderRadius: 6, padding: "6px 8px", fontSize: 13, fontFamily: "inherit", background: "#fff", boxSizing: "border-box" },

  field: { display: "flex", flexDirection: "column", gap: 5 },
  fieldLabel: { fontSize: 12, fontWeight: 700, color: "#42627a" },
  input: { border: `1px solid ${LINE}`, borderRadius: 8, padding: "10px 12px", fontSize: 14, fontFamily: "inherit", background: "#fff", boxSizing: "border-box", width: "100%" },

  btnPrimary: { background: BLUE, color: "#fff", border: "none", borderRadius: 9, padding: "12px 22px", fontSize: 14, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" },
  btnDisabled: { background: "#bcd2dd", color: "#fff", border: "none", borderRadius: 9, padding: "12px 22px", fontSize: 14, fontWeight: 800, cursor: "not-allowed", fontFamily: "inherit" },
  btnGhost: { background: "#fff", color: INK, border: `1.5px solid ${LINE}`, borderRadius: 9, padding: "11px 18px", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" },
  btnGhostSm: { background: "#fff", color: BLUE_DK, border: `1px solid ${LINE}`, borderRadius: 7, padding: "7px 12px", fontSize: 12.5, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" },
  btnExcelOut: { background: "#eaf6ee", color: "#1f7a45", border: "1px solid #bfe3cc", borderRadius: 7, padding: "7px 12px", fontSize: 12.5, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" },
  btnExcelIn: { background: "#1f9d57", color: "#fff", border: "1px solid #1f9d57", borderRadius: 7, padding: "7px 12px", fontSize: 12.5, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" },

  pasteHint: { fontSize: 12, color: "#42627a", background: "#eef8fb", border: `1px solid #cfe6ef`, borderRadius: 8, padding: "9px 13px", marginBottom: 12, lineHeight: 1.55 },
  sheetScroll: { overflowX: "auto", border: "1px solid #000", borderRadius: 2, background: "#fff" },
  sheet: { width: "100%", borderCollapse: "collapse", fontSize: 13.5, fontFamily: "'Times New Roman', Georgia, serif", tableLayout: "fixed", minWidth: 720 },
  shTh: { background: "#dfdfdf", border: "1px solid #000", padding: "6px 8px", fontSize: 13, fontWeight: 700, textAlign: "center", color: "#111", fontFamily: "'Times New Roman', Georgia, serif" },
  shTd: { border: "1px solid #000", padding: 0, position: "relative", verticalAlign: "middle" },
  shTdSel: { outline: `2px solid ${BLUE}`, outlineOffset: -2 },
  shInput: { width: "100%", border: "none", padding: "6px 8px", fontSize: 13.5, fontFamily: "'Times New Roman', Georgia, serif", background: "transparent", boxSizing: "border-box", outline: "none" },
  shActions: { border: "1px solid #000", borderLeft: "none", padding: "0 4px", whiteSpace: "nowrap", textAlign: "center", background: "#fafbfc", width: 78 },
  tagBtn: { background: "#fff", border: `1px solid ${LINE}`, borderRadius: 5, width: 22, height: 22, cursor: "pointer", color: "#9aa9b6", marginRight: 2, fontSize: 11, fontWeight: 800, padding: 0 },
  tagBtnOn: { background: ACCENT, border: `1px solid ${ACCENT}`, borderRadius: 5, width: 22, height: 22, cursor: "pointer", color: "#fff", marginRight: 2, fontSize: 11, fontWeight: 800, padding: 0 },
  miniBtn: { background: "#fff", border: `1px solid ${LINE}`, borderRadius: 5, width: 24, height: 24, cursor: "pointer", color: "#5d7589", marginLeft: 2, fontSize: 11, padding: 0 },
  miniBtnDanger: { background: "#fff", border: "1px solid #f3cdc8", borderRadius: 5, width: 24, height: 24, cursor: "pointer", color: ACCENT, marginLeft: 2, fontSize: 11, padding: 0 },

  overlay: { position: "fixed", inset: 0, background: "rgba(13,60,77,.45)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "40px 16px", zIndex: 100, overflowY: "auto" },
  modal: { background: "#fff", borderRadius: 16, width: "min(720px, 100%)", padding: "24px 26px 22px", boxShadow: "0 24px 70px rgba(13,60,77,.3)" },
  modalHead: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  modalSub: { fontSize: 13, color: "#5d7589", margin: "6px 0 18px" },
  modalGrid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 },
  modalFoot: { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 },
  warn: { marginTop: 16, padding: "10px 14px", background: "#fdf1ef", border: "1px solid #f3cdc8", borderRadius: 9, fontSize: 13, color: ACCENT, fontWeight: 700 },

  dashGrid: { display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16 },
  statCard: { background: "#fff", border: `1px solid ${LINE}`, borderRadius: 14, padding: "20px 22px", boxShadow: "0 2px 8px rgba(20,50,74,.05)" },
  statLabel: { fontSize: 12, letterSpacing: ".05em", textTransform: "uppercase", color: "#7790a3", fontWeight: 800 },
  statValue: { fontSize: 34, fontWeight: 800, color: INK, marginTop: 8 },
  statUnit: { fontSize: 15, fontWeight: 600, color: "#7790a3" },
};

const globalCss = `
  * { box-sizing: border-box; }
  body { margin: 0; }
  input:focus { outline: 2px solid ${BLUE}44; border-color: ${BLUE}; }
  button:focus-visible { outline: 2px solid ${BLUE}; outline-offset: 2px; }
  code { background: #e3eef3; padding: 1px 5px; border-radius: 4px; font-size: 11.5px; }
  tr.hsRow:hover { background: #f7fbfc; }
  tr.hsRow:hover input[style*="transparent"] { background: #fff; border-color: ${LINE}; }
  @media (max-width: 760px) {
    div[style*="grid-template-columns: 1fr 1fr"], div[style*="repeat(3, 1fr)"] { grid-template-columns: 1fr !important; }
  }
`;
