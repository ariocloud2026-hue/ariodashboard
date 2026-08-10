// ============================================================================
//  Kunlik hisobot → Telegram guruhga (rasm + matn + Excel), avtomatik.
//  Ma'lumot: Google Sheets (gviz CSV). Parslash logikasi ilovadagi bilan bir xil.
//  Ishlaydi: GitHub Actions (kunlik cron). Node 20+.
// ============================================================================

import puppeteer from "puppeteer";
import * as XLSX from "xlsx";

// ─── Sozlamalar (env orqali, aks holda standart qiymatlar) ──────────────────
const XARAJAT_ID = process.env.XARAJAT_SHEET_ID || "109Zb2H6Ba29q7vOU5cM5hmJ6N1e9HyyaPyRBJDBYlN0";
const PRIXOD_ID  = process.env.PRIXOD_SHEET_ID  || "1VUfzI8k0i9KfuZNmdEda49v8n20oWMWnme5w7QjnwC4";
const BOT_TOKEN  = process.env.8618336444:AAGDvr8lFjzARSzJnIwwTngTRbCa1Z0OeNI;
const CHAT_ID    = process.env.-1003820971069;
// Ixtiyoriy: aynan bir sanani sinash uchun (masalan "2026-08-07"). Bo'sh bo'lsa — bugun.
const FORCE_DATE = process.env.REPORT_DATE || "";

// ─── Filial varaqlari (ilovadagi SH/LB bilan bir xil) ───────────────────────
const SH = ["sotuv_0","andijon_01","fargona_02","nukus_03","qarshi_04","samarqand_05","sirdaryo_06","jizzax_07","toshkent_08","namangan_09","buxoro_010"];
const LB = {sotuv_0:"Sotuv",andijon_01:"Andijon",fargona_02:"Farg'ona",nukus_03:"Nukus",qarshi_04:"Qarshi",samarqand_05:"Samarqand",sirdaryo_06:"Sirdaryo",jizzax_07:"Jizzax",toshkent_08:"Toshkent",namangan_09:"Namangan",buxoro_010:"Buxoro"};
const FIL_ORDER = ["Sotuv","Andijon","Farg'ona","Nukus","Qarshi","Samarqand","Sirdaryo","Jizzax","Toshkent","Namangan","Buxoro"];

const MO   = ["","Yan","Fev","Mar","Apr","May","Iyn","Iyl","Avg","Sen","Okt","Noy","Dek"];
const RUMO = {"январ":1,"феврал":2,"март":3,"апрел":4,"май":5,"июн":6,"июл":7,"август":8,"сентябр":9,"октябр":10,"ноябр":11,"декабр":12};

// ─── Yordamchi funksiyalar (ilovadan aynan ko'chirilgan) ────────────────────
const pad2 = n => String(n).padStart(2,"0");

function dval(v){
  if(v instanceof Date && !isNaN(v)) return v;
  if(typeof v === "number" && v > 20000 && v < 60000){
    const d = new Date(Math.round((v - 25569) * 86400000));
    return isNaN(d) ? null : d;
  }
  if(typeof v === "string" && v.trim()){
    const s = v.trim();
    const m = s.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/);
    if(m) return new Date(+m[3], +m[2]-1, +m[1]);
    const x = new Date(s);
    if(!isNaN(x)) return x;
  }
  return null;
}

function monthInfo(moRaw, dtRaw){
  let year = null, mo = null, label = "";
  const d1 = dval(dtRaw);
  if(d1){ year = d1.getFullYear(); mo = d1.getMonth() + 1; }
  if(moRaw !== null && moRaw !== undefined && moRaw !== ""){
    if(moRaw instanceof Date && !isNaN(moRaw)){
      year = moRaw.getFullYear(); mo = moRaw.getMonth() + 1;
      label = `${MO[mo]||""} ${year}`;
    } else {
      const s = String(moRaw).trim();
      let m = s.match(/^(\d{4})-(\d{2})/);
      if(m){ year = +m[1]; mo = +m[2]; label = s; }
      else {
        m = s.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/);
        if(m){ year = +m[3]; mo = +m[2]; label = s; }
        else {
          const low = s.toLowerCase(), yr = (low.match(/\d{4}/) || [])[0];
          for(const [ru, mn] of Object.entries(RUMO)){
            if(low.includes(ru)){ mo = mn; if(yr) year = +yr; label = s; break; }
          }
        }
      }
    }
  }
  if(!label && mo) label = `${MO[mo]||""} ${year||""}`.trim();
  return { sort: (year && mo) ? `${year}-${pad2(mo)}` : "9999-99", label: label || "Noma'lum" };
}

function dayInfo(dtObj){
  if(!dtObj) return { sort:null, label:null };
  const y=dtObj.getFullYear(), m=dtObj.getMonth()+1, d=dtObj.getDate();
  return { sort:`${y}-${pad2(m)}-${pad2(d)}`, label:`${pad2(d)}.${pad2(m)}.${y}` };
}

function toNum(v){
  if(typeof v === "number") return isFinite(v) ? v : 0;
  if(v === null || v === undefined || v === "") return 0;
  let s = String(v).trim().replace(/[^\d,.\-]/g, "");
  if(!s) return 0;
  if(s.includes(",") && s.includes(".")){
    if(s.lastIndexOf(",") > s.lastIndexOf(".")) s = s.replace(/\./g,"").replace(",",".");
    else s = s.replace(/,/g,"");
  } else if(s.includes(",")){
    const after = s.split(",").pop();
    s = after.length === 3 ? s.replace(/,/g,"") : s.replace(",",".");
  }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function findHeaderRow(arr, keywords){
  const lim = Math.min(arr.length, 12);
  for(let r=0; r<lim; r++){
    const row = arr[r]; if(!row) continue;
    const joined = row.map(c=>String(c||"").toLowerCase()).join("|");
    if(keywords.some(k=>joined.includes(k))) return r;
  }
  return 0;
}

function computeUsd(sum, usd, rate, cur, fallbackRate){
  if(usd && usd > 0) return usd;
  const c = String(cur||"").toLowerCase();
  if(c.includes("$") || c.includes("usd") || c.includes("дол") || c.includes("dol")) return sum;
  const rt = (rate && rate > 0) ? rate : fallbackRate;
  return (rt && rt > 0 && sum) ? sum / rt : 0;
}

function avgRate(rows){
  const rr = rows.map(x=>x.rate).filter(x=>x>0);
  return rr.length ? rr.reduce((a,b)=>a+b,0)/rr.length : 0;
}

function parseCSVText(text){
  const out=[];
  if(!text) return out;
  let row=[], cell="", inQ=false;
  const pushRow=()=>{ row.push(cell); cell=""; if(row.some(x=>String(x).trim()!=="")) out.push(row); row=[]; };
  for(let i=0;i<text.length;i++){
    const c=text[i];
    if(inQ){
      if(c==='"'){ if(text[i+1]==='"'){ cell+='"'; i++; } else inQ=false; }
      else cell+=c;
    } else {
      if(c==='"') inQ=true;
      else if(c===",") { row.push(cell); cell=""; }
      else if(c==="\n") pushRow();
      else if(c==="\r") {}
      else cell+=c;
    }
  }
  if(cell!=="" || row.length) pushRow();
  return out;
}

// ─── Xarajat (chiqim) parslash — ilovadagi parseRows ────────────────────────
function parseRows(arr, filial, flabel){
  if(!arr || arr.length < 2) return [];
  const hr = findHeaderRow(arr, ["категор","kategor","дата","date","сумм","сум "]);
  const h  = arr[hr].map((c,i)=>({ l:String(c||"").toLowerCase().trim(), i }));
  const fi = (...kw)=>{ const f=h.find(x=>kw.some(k=>x.l.includes(k))); return f?f.i:-1; };
  const usdI = [...h].reverse().find(x=>x.l.includes("$")||x.l.includes("usd"))?.i ?? -1;
  const sumI = [...h].reverse().find(x=>(x.l.includes("сумм")||x.l.startsWith("sum")) && !x.l.includes("$") && !x.l.includes("usd"))?.i ?? -1;
  const I = { dt:fi("дата","date"), mo:fi("месяц","month"), fio:fi("фио","ф.и"), cat:fi("категор"), typ:fi("тип"), cur:fi("валют"), kurs:fi("курс","kurs","rate"), bal:fi("остаток","qoldiq","balan"), desc:fi("описан","descrip","description"), sum:sumI, usd:usdI };
  const rows = [];
  for(let i=hr+1;i<arr.length;i++){
    const r=arr[i];
    const cat = I.cat>=0 ? r[I.cat] : null;
    if(!cat || !String(cat).trim()) continue;
    const mi = monthInfo(I.mo>=0?r[I.mo]:null, I.dt>=0?r[I.dt]:null);
    const dtObj = I.dt>=0 ? dval(r[I.dt]) : null;
    const di = dayInfo(dtObj);
    rows.push({
      filial, flabel, cat:String(cat).trim(),
      month:mi.label, ms:mi.sort, day:di.sort, dayLabel:di.label,
      fio:I.fio>=0?r[I.fio]:null, typ:I.typ>=0?r[I.typ]:null, cur:I.cur>=0?r[I.cur]:null,
      rate:I.kurs>=0?toNum(r[I.kurs]):0, balance:I.bal>=0?toNum(r[I.bal]):0,
      desc:I.desc>=0?(r[I.desc]?String(r[I.desc]).trim():null):null,
      sum:toNum(I.sum>=0?r[I.sum]:0), usd:toNum(I.usd>=0?r[I.usd]:0),
    });
  }
  const fr = avgRate(rows);
  for(const x of rows){
    x.usd = computeUsd(x.sum, x.usd, x.rate, x.cur, fr);
    if(!(x.rate>0)) x.rate = (x.usd>0 && x.sum) ? (x.sum/x.usd) : fr;
  }
  return rows;
}

// ─── Prixod (kirim) parslash — ilovadagi parsePrixodRows ────────────────────
function parsePrixodRows(matrix, filialName, flabel){
  if(!matrix||matrix.length<2) return [];
  const hr=findHeaderRow(matrix, ["тип опл","категор","дата","date","клиент","менедж","сумм","остаток"]);
  const hdr=matrix[hr].map(h=>String(h||"").trim().toLowerCase());
  const fi=(...names)=>{ for(const n of names){ const idx=hdr.findIndex(h=>h.includes(n.toLowerCase())); if(idx>=0) return idx; } return -1; };
  const iDay=fi("дата","sana","date"), iMo=fi("месяц","oy","month"), iCli=fi("клиент","client","mijoz"),
        iMgr=fi("менеджер","manager","menejer"), iTyp=fi("тип опл","tip op","payment","тип"), iCat=fi("категор","kategori","cat"),
        iDesc=fi("описан","tavsif","desc"), iCur=fi("валют","valyut","curr"), iRate=fi("курс","kurs","rate"),
        iBal=fi("остаток","qoldiq","balan"), iUsd=fi("сумма $","summa $","сумма$","usd","$ "), iSum=fi("сумма","summa","sum","amount");
  const rows=[];
  for(let i=hr+1;i<matrix.length;i++){
    const row=matrix[i];
    if(!row||row.every(c=>c===null||c===undefined||c===""||(typeof c==="string"&&!c.trim()))) continue;
    const gf=idx=>idx>=0?toNum(row[idx]):0;
    const gs=idx=>idx>=0?String(row[idx]??"").trim():"";
    const usd=iUsd>=0?gf(iUsd):0, sum=gf(iSum), cur=gs(iCur);
    if(!usd&&!sum) continue;
    const dtObj=iDay>=0?dval(row[iDay]):null;
    let day="", dayLabel="";
    if(dtObj){
      const _y=dtObj.getFullYear(),_m=dtObj.getMonth()+1,_d=dtObj.getDate();
      day=`${_y}-${pad2(_m)}-${pad2(_d)}`; dayLabel=`${pad2(_d)}.${pad2(_m)}.${_y}`;
    } else {
      const dayRaw=gs(iDay); dayLabel=dayRaw;
      const dm=dayRaw.match(/(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})/);
      if(dm){ let [,a,b,c]=dm; if(c.length===2) c="20"+c; day=`${c}-${b.padStart(2,"0")}-${a.padStart(2,"0")}`; }
      else { const dm2=dayRaw.match(/(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/); if(dm2){const[,y,m,d]=dm2;day=`${y}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;} }
    }
    const mi=monthInfo(iMo>=0?row[iMo]:null, iDay>=0?row[iDay]:null);
    rows.push({ filial:filialName, flabel, day, dayLabel, month:mi.label||gs(iMo), ms:mi.sort,
      client:gs(iCli), manager:gs(iMgr), typ:gs(iTyp), cat:gs(iCat), desc:gs(iDesc), cur,
      rate:gf(iRate), balance:gf(iBal), sum, usd });
  }
  const fr=avgRate(rows);
  for(const x of rows){
    x.usd = computeUsd(x.sum, x.usd, x.rate, x.cur, fr);
    if(!(x.rate>0)) x.rate = (x.usd>0 && x.sum) ? (x.sum/x.usd) : fr;
  }
  return rows;
}

// ─── Google Sheets: bitta varaqni gviz CSV orqali olish ─────────────────────
async function fetchSheetCSV(id, sheetName){
  const url = `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
  const r = await fetch(url);
  if(!r.ok) throw new Error(`HTTP ${r.status}`);
  const t = await r.text();
  const tl = (t||"").trim().toLowerCase();
  if(!t || tl.startsWith("<!") || tl.startsWith("<html") || tl.includes("accounts.google.com") || tl.includes("signin"))
    throw new Error("NOT_PUBLIC"); // jadval ochiq emas (link bilan ko'rish yoqilmagan)
  return parseCSVText(t);
}

async function loadAll(id, parser){
  const out = [];
  await Promise.all(SH.map(async sh=>{
    try{
      const matrix = await fetchSheetCSV(id, sh);
      const rows = parser(matrix, sh, LB[sh]||sh);
      out.push(...rows);
    }catch(e){
      console.warn(`  ⚠ ${sh}: ${e.message}`);
    }
  }));
  return out;
}

// ─── Sana: Toshkent vaqti (UTC+5) ───────────────────────────────────────────
function tashkentNow(){
  const now = new Date();
  const utcMs = now.getTime() + now.getTimezoneOffset()*60000;
  return new Date(utcMs + 5*3600000); // UTC+5, DST yo'q
}
// Hisobot sanasi = KECHA (ertalab yuborilganda to'liq o'tgan kun). FORCE_DATE bo'lsa — o'sha.
function reportDate(){
  const now = tashkentNow();
  const hm = `${pad2(now.getHours())}:${pad2(now.getMinutes())}`;
  if(FORCE_DATE){
    const [y,m,d]=FORCE_DATE.split("-").map(Number);
    return { key:FORCE_DATE, label:`${pad2(d)}.${pad2(m)}.${y}`, hm };
  }
  const t = new Date(now); t.setDate(t.getDate()-1); // kecha
  return { key:`${t.getFullYear()}-${pad2(t.getMonth()+1)}-${pad2(t.getDate())}`,
           label:`${pad2(t.getDate())}.${pad2(t.getMonth()+1)}.${t.getFullYear()}`,
           hm };
}

const f2 = n => "$" + Math.round(n||0).toLocaleString("en-US");
const esc = s => String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");

// ─── Bugungi hisobni yig'ish ────────────────────────────────────────────────
function aggregate(prixod, xarajat, todayKey){
  const kRows = prixod.filter(r=>r.day===todayKey);
  const xRows = xarajat.filter(r=>r.day===todayKey);
  const kirim = kRows.reduce((s,r)=>s+(r.usd||0),0);
  const chiqim= xRows.reduce((s,r)=>s+(r.usd||0),0);
  const net = kirim - chiqim;

  const fmap={};
  const bump=(arr,key)=>arr.forEach(r=>{ const f=String(r.flabel||r.filial||"—").trim()||"—"; if(!fmap[f])fmap[f]={filial:f,kirim:0,chiqim:0}; fmap[f][key]+=r.usd||0; });
  bump(kRows,"kirim"); bump(xRows,"chiqim");
  const fils = Object.values(fmap).map(o=>({...o,net:o.kirim-o.chiqim}))
    .sort((a,b)=>{ const ia=FIL_ORDER.indexOf(a.filial), ib=FIL_ORDER.indexOf(b.filial);
      if(ia!==-1||ib!==-1) return (ia===-1?99:ia)-(ib===-1?99:ib); return b.kirim-a.kirim; });

  const catMap={}; xRows.forEach(r=>{ const c=String(r.cat||"—").trim()||"—"; catMap[c]=(catMap[c]||0)+(r.usd||0); });
  const topCats=Object.entries(catMap).map(([n,v])=>({n,v})).sort((a,b)=>b.v-a.v).slice(0,8);

  return { kRows, xRows, kirim, chiqim, net, fils, topCats };
}

// ─── Hisobot HTML (rasmga aylantiriladi) ────────────────────────────────────
function reportHTML(agg, dateLabel, genTime){
  const { kirim, chiqim, net, fils, topCats, kRows, xRows } = agg;
  const tile = (l,v,clr,bg,bd,ic)=>`
    <div style="background:${bg};border:1px solid ${bd};border-radius:14px;padding:14px 16px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
        <span style="font-size:11px;font-weight:800;color:${clr};letter-spacing:.4px">${l}</span>
        <span style="font-size:18px">${ic}</span>
      </div>
      <div style="font-size:26px;font-weight:800;color:${clr};letter-spacing:-.02em">${v}</div>
    </div>`;
  const filRows = fils.map(f=>`
    <tr style="border-bottom:1px solid #eef2f7">
      <td style="padding:8px 12px;font-weight:700;color:#1e293b">${esc(f.filial)}</td>
      <td style="padding:8px 12px;text-align:right;font-weight:700;color:#059669">${f2(f.kirim)}</td>
      <td style="padding:8px 12px;text-align:right;font-weight:700;color:#dc2626">${f2(f.chiqim)}</td>
      <td style="padding:8px 12px;text-align:right;font-weight:800;color:${f.net>=0?"#0369a1":"#b45309"}">${f2(f.net)}</td>
    </tr>`).join("");
  const catBlock = topCats.length ? `
    <div style="font-size:13px;font-weight:800;color:#0f172a;margin:0 0 8px 2px">📤 Chiqim — kategoriyalar (top ${topCats.length})</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px 18px;margin-bottom:8px">
      ${topCats.map((c,i)=>`<div style="display:flex;justify-content:space-between;gap:10px;font-size:12.5px;padding:5px 10px;background:${i%2?"#fafafa":"#f8fafc"};border-radius:6px">
        <span style="color:#334155;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(c.n)}</span>
        <span style="font-weight:700;color:#dc2626;white-space:nowrap">${f2(c.v)}</span></div>`).join("")}
    </div>` : "";
  return `<!doctype html><html><head><meta charset="utf-8">
  <style>*{box-sizing:border-box;margin:0;padding:0;font-family:system-ui,-apple-system,'Segoe UI',Roboto,Arial,sans-serif}</style></head>
  <body style="background:#eef2f7;padding:20px">
    <div id="card" style="width:680px;background:#fff;border-radius:16px;padding:24px 26px;color:#0f172a;box-shadow:0 10px 40px rgba(15,23,42,.12)">
      <div style="display:flex;align-items:flex-end;justify-content:space-between;border-bottom:3px solid #2563eb;padding-bottom:12px;margin-bottom:16px">
        <div>
          <div style="font-size:22px;font-weight:800;letter-spacing:-.02em">📊 Kunlik moliyaviy hisobot</div>
          <div style="font-size:13px;color:#64748b;margin-top:2px">Kunlik yakun · barcha filiallar</div>
        </div>
        <div style="text-align:right">
          <div style="font-size:20px;font-weight:800;color:#2563eb">${dateLabel}</div>
          <div style="font-size:11px;color:#94a3b8">${kRows.length+xRows.length} yozuv</div>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:18px">
        ${tile("KIRIM",f2(kirim),"#059669","#ecfdf5","#a7f3d0","📥")}
        ${tile("CHIQIM",f2(chiqim),"#dc2626","#fef2f2","#fecaca","📤")}
        ${tile("SOF (OSTATKA)",f2(net),net>=0?"#0369a1":"#b45309",net>=0?"#eff6ff":"#fffbeb",net>=0?"#bfdbfe":"#fde68a",net>=0?"💰":"⚠️")}
      </div>
      <div style="font-size:13px;font-weight:800;color:#0f172a;margin:0 0 8px 2px">🏢 Filiallar kesmida</div>
      <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:${topCats.length?18:6}px">
        <thead><tr style="background:#f1f5f9">
          <th style="padding:8px 12px;text-align:left;font-size:11px;color:#475569;font-weight:700;border-bottom:2px solid #e2e8f0">Filial</th>
          <th style="padding:8px 12px;text-align:right;font-size:11px;color:#475569;font-weight:700;border-bottom:2px solid #e2e8f0">Kirim</th>
          <th style="padding:8px 12px;text-align:right;font-size:11px;color:#475569;font-weight:700;border-bottom:2px solid #e2e8f0">Chiqim</th>
          <th style="padding:8px 12px;text-align:right;font-size:11px;color:#475569;font-weight:700;border-bottom:2px solid #e2e8f0">Ostatka</th>
        </tr></thead>
        <tbody>${filRows || `<tr><td colspan="4" style="padding:14px;text-align:center;color:#94a3b8">Bugun uchun yozuv yo'q</td></tr>`}</tbody>
        <tfoot><tr style="background:#f8fafc;font-weight:800;border-top:2px solid #e2e8f0">
          <td style="padding:9px 12px">JAMI</td>
          <td style="padding:9px 12px;text-align:right;color:#059669">${f2(kirim)}</td>
          <td style="padding:9px 12px;text-align:right;color:#dc2626">${f2(chiqim)}</td>
          <td style="padding:9px 12px;text-align:right;color:${net>=0?"#0369a1":"#b45309"}">${f2(net)}</td>
        </tr></tfoot>
      </table>
      ${catBlock}
      <div style="margin-top:14px;padding-top:10px;border-top:1px solid #e2e8f0;display:flex;justify-content:space-between;font-size:10.5px;color:#94a3b8">
        <span>📈 Baza — moliyaviy hisobot</span><span>Yaratildi: ${dateLabel} ${genTime}</span>
      </div>
    </div>
  </body></html>`;
}

async function renderImage(html){
  const browser = await puppeteer.launch({ headless:"new", args:["--no-sandbox","--disable-setuid-sandbox"] });
  try{
    const page = await browser.newPage();
    await page.setViewport({ width:740, height:900, deviceScaleFactor:2 });
    await page.setContent(html, { waitUntil:"networkidle0" });
    const el = await page.$("#card");
    const buf = await el.screenshot({ type:"png" });
    return buf;
  } finally { await browser.close(); }
}

// ─── Excel yasash (Umumiy + Kirim + Chiqim) ─────────────────────────────────
function buildExcel(agg, dateLabel){
  const R = Math.round;
  const wb = XLSX.utils.book_new();
  const u = [["KUNLIK HISOBOT", dateLabel],[],["", "$"],["Kirim", R(agg.kirim)],["Chiqim", R(agg.chiqim)],["Sof (ostatka)", R(agg.net)],[],
            ["Filial","Kirim $","Chiqim $","Ostatka $"]];
  agg.fils.forEach(f=>u.push([f.filial, R(f.kirim), R(f.chiqim), R(f.net)]));
  u.push(["JAMI", R(agg.kirim), R(agg.chiqim), R(agg.net)]);
  const wsU = XLSX.utils.aoa_to_sheet(u); wsU["!cols"]=[{wch:22},{wch:14},{wch:14},{wch:14}];
  XLSX.utils.book_append_sheet(wb, wsU, "Umumiy");

  const mk = (rows,name)=>{
    const a=[["Sana","Filial","Klient / FIO","Agent","Tip","Kategoriya","Izoh","Valyuta","Kurs","Summa","Summa $"]];
    rows.forEach(r=>a.push([r.dayLabel||"", r.flabel||r.filial||"", r.client||r.fio||"", r.manager||"", r.typ||"", r.cat||"", r.desc||"", r.cur||"", R(r.rate||0), r.sum||0, R(r.usd||0)]));
    const ws=XLSX.utils.aoa_to_sheet(a); ws["!cols"]=[{wch:11},{wch:12},{wch:20},{wch:16},{wch:12},{wch:20},{wch:26},{wch:8},{wch:10},{wch:13},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws, name);
  };
  mk(agg.kRows,"Kirim"); mk(agg.xRows,"Chiqim");
  return XLSX.write(wb, { type:"buffer", bookType:"xlsx" });
}

// ─── Telegram yuborish ──────────────────────────────────────────────────────
async function tg(method, form){
  const r = await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/${method}`, { method:"POST", body:form });
  const j = await r.json().catch(()=>({}));
  if(!r.ok || !j.ok) throw new Error(`${method}: ${j.description || ("HTTP "+r.status)}`);
  return j;
}

async function sendPhoto(buf, caption){
  const fd = new FormData();
  fd.append("chat_id", CHAT_ID);
  fd.append("caption", caption);
  fd.append("parse_mode", "HTML");
  fd.append("photo", new Blob([buf], { type:"image/png" }), "hisobot.png");
  return tg("sendPhoto", fd);
}
async function sendDoc(buf, filename, caption){
  const fd = new FormData();
  fd.append("chat_id", CHAT_ID);
  if(caption){ fd.append("caption", caption); fd.append("parse_mode","HTML"); }
  fd.append("document", new Blob([buf], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename);
  return tg("sendDocument", fd);
}

function captionText(agg, dateLabel){
  const arrow = agg.net>=0 ? "🟢" : "🔴";
  let s = `📊 <b>Kunlik hisobot (kecha)</b> — ${dateLabel}\n`;
  s += `\n📥 Kirim:  <b>${f2(agg.kirim)}</b>`;
  s += `\n📤 Chiqim: <b>${f2(agg.chiqim)}</b>`;
  s += `\n${arrow} Sof:    <b>${f2(agg.net)}</b>`;
  s += `\n\n🏢 <b>Filiallar</b>`;
  agg.fils.forEach(f=>{ s += `\n• ${esc(f.filial)}: 📥${f2(f.kirim)} 📤${f2(f.chiqim)} = ${f2(f.net)}`; });
  if(agg.topCats.length){
    s += `\n\n📤 <b>Chiqim kategoriyalar</b>`;
    agg.topCats.slice(0,6).forEach(c=>{ s += `\n• ${esc(c.n)}: ${f2(c.v)}`; });
  }
  return s;
}

// ─── Asosiy oqim ────────────────────────────────────────────────────────────
async function main(){
  if(!BOT_TOKEN || !CHAT_ID){
    console.error("❌ TELEGRAM_BOT_TOKEN yoki TELEGRAM_CHAT_ID yo'q (GitHub Secrets ga qo'shing).");
    process.exit(1);
  }
  const { key, label, hm } = reportDate();
  console.log(`📅 Hisobot sanasi (kecha, Toshkent): ${label} (${key})`);

  console.log("⬇️  Google Sheets o'qilmoqda…");
  const [prixod, xarajat] = await Promise.all([
    loadAll(PRIXOD_ID, parsePrixodRows),
    loadAll(XARAJAT_ID, parseRows),
  ]);
  console.log(`   Kirim: ${prixod.length} yozuv · Chiqim: ${xarajat.length} yozuv`);

  const agg = aggregate(prixod, xarajat, key);
  console.log(`   Bugun → Kirim ${f2(agg.kirim)} · Chiqim ${f2(agg.chiqim)} · Sof ${f2(agg.net)} · ${agg.fils.length} filial`);

  console.log("🖼️  Rasm chizilmoqda…");
  const png = await renderImage(reportHTML(agg, label, hm));

  console.log("📄 Excel yasalmoqda…");
  const xlsx = buildExcel(agg, label);

  console.log("📤 Telegramga yuborilmoqda…");
  await sendPhoto(png, captionText(agg, label));                       // 1) rasm + matnli caption
  await sendDoc(xlsx, `Kunlik_hisobot_${key}.xlsx`, "📎 Excel hisobot"); // 2) Excel fayl
  console.log("✅ Tayyor — guruhga yuborildi.");
}

main().catch(e=>{ console.error("❌ Xato:", e.message); process.exit(1); });
