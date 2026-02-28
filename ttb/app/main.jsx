import { useState, useRef, useCallback, useEffect, useMemo } from "react";

const ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
const DEFAULT_ROWS = 50;
const DEFAULT_COLS = 26;
const ROW_HEIGHT = 22;
const DEFAULT_COL_WIDTH = 200;
const ROW_HEADER_WIDTH = 46;

const FONTS = [
  "Calibri","Arial","Times New Roman","Courier New","Georgia","Verdana",
  "Tahoma","Trebuchet MS","Impact","Comic Sans MS","Palatino","Garamond",
  "Lucida Sans","Helvetica","Century Gothic","Franklin Gothic Medium",
  "Gill Sans","Book Antiqua","Cambria","Consolas",
];
const FONT_SIZES = [8,9,10,11,12,14,16,18,20,22,24,28,32,36,48,72];

function colLabel(i) {
  let label = "";
  while (i >= 0) { label = ALPHABET[i % 26] + label; i = Math.floor(i / 26) - 1; }
  return label;
}
function cellAddr(r, c) { return `${colLabel(c)}${r + 1}`; }
function parseAddr(addr) {
  const m = addr.toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  let c = 0;
  for (let i = 0; i < m[1].length; i++) c = c * 26 + m[1].charCodeAt(i) - 64;
  return { r: parseInt(m[2]) - 1, c: c - 1 };
}

function evalFormula(expr, lookup, depth = 0) {
  if (depth > 30) return "#CIRC!";
  expr = expr.trim();
  function getRange(a1, a2) {
    const from = parseAddr(a1), to = parseAddr(a2);
    if (!from || !to) return [];
    const vals = [];
    for (let r = Math.min(from.r, to.r); r <= Math.max(from.r, to.r); r++)
      for (let c = Math.min(from.c, to.c); c <= Math.max(from.c, to.c); c++)
        vals.push(lookup(r, c, depth + 1));
    return vals;
  }
  let processed = expr.replace(/([A-Z]+\d+):([A-Z]+\d+)/gi, (_, a, b) =>
    "__RANGE__" + JSON.stringify(getRange(a, b)) + "__END__");
  processed = processed.replace(/\b([A-Z]+\d+)\b/gi, (_, addr) => {
    const cell = parseAddr(addr); if (!cell) return addr;
    const v = lookup(cell.r, cell.c, depth + 1);
    const n = parseFloat(v); return isNaN(n) ? `"${v}"` : String(n);
  });
  function extractRangeVals(s) {
    const m = s.match(/^__RANGE__(.*?)__END__$/);
    return m ? JSON.parse(m[1]).map(v => parseFloat(v)).filter(v => !isNaN(v)) : null;
  }
  processed = processed.replace(/(SUM|AVERAGE|MAX|MIN|COUNT|IF|ROUND|ABS|SQRT|POWER|LEN|UPPER|LOWER|CONCAT|AND|OR|NOT)\s*\(/gi,
    (_, fn) => `__FN_${fn.toUpperCase()}__(`);
  try {
    const fnMap = {
      __FN_SUM__: (...a) => { let s=0; a.forEach(x=>{const rv=extractRangeVals(String(x));if(rv)s+=rv.reduce((a,b)=>a+b,0);else{const n=parseFloat(x);if(!isNaN(n))s+=n;}}); return s; },
      __FN_AVERAGE__: (...a) => { let s=0,cnt=0; a.forEach(x=>{const rv=extractRangeVals(String(x));if(rv){s+=rv.reduce((a,b)=>a+b,0);cnt+=rv.length;}else{const n=parseFloat(x);if(!isNaN(n)){s+=n;cnt++;}}}); return cnt?s/cnt:0; },
      __FN_MAX__: (...a) => { let v=[]; a.forEach(x=>{const rv=extractRangeVals(String(x));if(rv)v.push(...rv);else{const n=parseFloat(x);if(!isNaN(n))v.push(n);}}); return v.length?Math.max(...v):0; },
      __FN_MIN__: (...a) => { let v=[]; a.forEach(x=>{const rv=extractRangeVals(String(x));if(rv)v.push(...rv);else{const n=parseFloat(x);if(!isNaN(n))v.push(n);}}); return v.length?Math.min(...v):0; },
      __FN_COUNT__: (...a) => { let cnt=0; a.forEach(x=>{const rv=extractRangeVals(String(x));if(rv)cnt+=rv.length;else if(!isNaN(parseFloat(x)))cnt++;}); return cnt; },
      __FN_IF__: (c, t, f) => c ? t : (f===undefined ? 0 : f),
      __FN_ROUND__: (n, d) => Math.round(n*Math.pow(10,d||0))/Math.pow(10,d||0),
      __FN_ABS__: n => Math.abs(n), __FN_SQRT__: n => Math.sqrt(n),
      __FN_POWER__: (b, e) => Math.pow(b, e), __FN_LEN__: s => String(s).length,
      __FN_UPPER__: s => String(s).toUpperCase(), __FN_LOWER__: s => String(s).toLowerCase(),
      __FN_CONCAT__: (...a) => a.join(""), __FN_AND__: (...a) => a.every(Boolean),
      __FN_OR__: (...a) => a.some(Boolean), __FN_NOT__: v => !v,
    };
    // eslint-disable-next-line no-new-func
    const fn = new Function(...Object.keys(fnMap), `"use strict"; return (${processed});`);
    const result = fn(...Object.values(fnMap));
    if (result===undefined||result===null) return "";
    if (typeof result==="boolean") return result?"WAHR":"FALSCH";
    return result;
  } catch { return "#FEHLER!"; }
}

function createData(rows, cols) {
  return Array.from({ length: rows }, () => Array(cols).fill(""));
}

// ── .tabelle format ──────────────────────────────────────────────────────────
function serializeTabelle(data, formatting, colWidths, sheets, filename) {
  const lines = [];
  lines.push(`// ${filename}.tabelle — Format v1`);
  lines.push(`// Gespeichert: ${new Date().toLocaleString("de-DE")}`);
  lines.push(`// Zeilen hier`);
  lines.push("");

  data.forEach((row, r) => {
    const cellParts = [];
    row.forEach((cell, c) => {
      if (cell === "" && !formatting[`${r},${c}`]) return;
      const fmt = formatting[`${r},${c}`] || {};
      const styleParts = [];
      if (fmt.font || fmt.fontSize) {
        const fn = fmt.font || "Calibri";
        const fs = fmt.fontSize ? `, ${fmt.fontSize}pt` : "";
        styleParts.push(`font: {${fn}${fs}}`);
      }
      if (fmt.bold) styleParts.push("bold");
      if (fmt.italic) styleParts.push("italic");
      if (fmt.underline) styleParts.push("underline");
      if (fmt.align) styleParts.push(`align: ${fmt.align}`);
      if (fmt.color) styleParts.push(`color: ${fmt.color}`);
      if (fmt.bg) styleParts.push(`bg: ${fmt.bg}`);

      const addr = `${colLabel(c)}${r+1}`.toLowerCase();
      let s = `${addr}{{content: {${cell}}`;
      if (styleParts.length) s += `, style: [${styleParts.join(", ")}]`;
      s += `}}`;
      cellParts.push(s);
    });
    if (cellParts.length) lines.push(`   ${cellParts.join(", ")};`);
  });

  const meta = [];
  colWidths.forEach((w, i) => { if (w !== DEFAULT_COL_WIDTH) meta.push(`col_width: {${colLabel(i)}: ${w}}`); });
  sheets.forEach((s, i) => { if (i > 0 || s !== "Tabelle1") meta.push(`sheet: {${i}: ${s}}`); });

  let out = `[[\n${lines.join("\n")}\n]]`;
  if (meta.length) out += `\n\n// meta\n// ${meta.join(", ")}`;
  return out;
}

function parseTabelle(text) {
  const data = createData(DEFAULT_ROWS, DEFAULT_COLS);
  const formatting = {};
  const colWidths = Array(DEFAULT_COLS).fill(DEFAULT_COL_WIDTH);

  const bodyMatch = text.match(/\[\[([\s\S]*?)\]\]/);
  if (!bodyMatch) return { data, formatting, colWidths };
  const body = bodyMatch[1];

  const cellRe = /([a-z]+\d+)\{\{content:\s*\{([^}]*)\}(?:,\s*style:\s*\[([^\]]*)\])?\}\}/gi;
  let m;
  while ((m = cellRe.exec(body)) !== null) {
    const addr = parseAddr(m[1]); if (!addr) continue;
    const { r, c } = addr;
    while (data.length <= r) data.push(Array(DEFAULT_COLS).fill(""));
    while (data[r].length <= c) data[r].push("");
    data[r][c] = m[2].trim();
    if (m[3]) {
      const ss = m[3], fmt = {};
      const fontM = ss.match(/font:\s*\{([^,}]+)(?:,\s*(\d+)pt)?\}/);
      if (fontM) { fmt.font = fontM[1].trim(); if (fontM[2]) fmt.fontSize = parseInt(fontM[2]); }
      if (/\bbold\b/i.test(ss)) fmt.bold = true;
      if (/\bitalic\b/i.test(ss)) fmt.italic = true;
      if (/\bunderline\b/i.test(ss)) fmt.underline = true;
      const alignM = ss.match(/align:\s*(\w+)/); if (alignM) fmt.align = alignM[1];
      const colorM = ss.match(/color:\s*(#[\da-fA-F]{3,6})/); if (colorM) fmt.color = colorM[1];
      const bgM = ss.match(/bg:\s*(#[\da-fA-F]{3,6})/); if (bgM) fmt.bg = bgM[1];
      formatting[`${r},${c}`] = fmt;
    }
  }

  const cwRe = /col_width:\s*\{([A-Z]+):\s*(\d+)\}/gi;
  while ((m = cwRe.exec(text)) !== null) {
    const addr = parseAddr(m[1] + "1"); if (addr) colWidths[addr.c] = parseInt(m[2]);
  }

  return { data, formatting, colWidths };
}

export default function ExcelEditor() {
  const [rows, setRows] = useState(DEFAULT_ROWS);
  const [cols, setCols] = useState(DEFAULT_COLS);
  const [data, setData] = useState(() => createData(DEFAULT_ROWS, DEFAULT_COLS));
  const [colWidths, setColWidths] = useState(() => Array(DEFAULT_COLS).fill(DEFAULT_COL_WIDTH));
  const [selected, setSelected] = useState({ r: 0, c: 0 });
  const [selection, setSelection] = useState({ r1: 0, c1: 0, r2: 0, c2: 0 });
  const [editing, setEditing] = useState(false);
  const [editValue, setEditValue] = useState("");
  const [dragSel, setDragSel] = useState(false);
  const [dragStart, setDragStart] = useState(null);
  const [history, setHistory] = useState([]);
  const [future, setFuture] = useState([]);
  const [formatting, setFormatting] = useState({});
  const [activeSheet, setActiveSheet] = useState(0);
  const [sheets, setSheets] = useState(["Tabelle1"]);
  const [findOpen, setFindOpen] = useState(false);
  const [findVal, setFindVal] = useState("");
  const [replaceVal, setReplaceVal] = useState("");
  const [contextMenu, setContextMenu] = useState(null);
  const [filename, setFilename] = useState("Mappe1");
  const [selFont, setSelFont] = useState("Calibri");
  const [selSize, setSelSize] = useState(11);

  const inputRef = useRef(null);
  const containerRef = useRef(null);
  const resizingRef = useRef(null);

  useEffect(() => { if (editing && inputRef.current) inputRef.current.focus(); }, [editing]);

  const fmtKey = (r, c) => `${r},${c}`;
  const getFmt = (r, c) => formatting[fmtKey(r, c)] || {};

  function evalCellDeep(r, c, d, depth) {
    const raw = d[r]?.[c] ?? "";
    if (typeof raw === "string" && raw.startsWith("="))
      return evalFormula(raw.slice(1), (lr, lc, dp2) => evalCellDeep(lr, lc, d, dp2), depth);
    return raw;
  }
  const evalCell = useCallback((r, c, d = data) => {
    const raw = d[r]?.[c] ?? "";
    if (typeof raw === "string" && raw.startsWith("="))
      return evalFormula(raw.slice(1), (lr, lc, depth) => evalCellDeep(lr, lc, d, depth), 0);
    return raw;
  }, [data]);

  const displayVal = (r, c) => {
    const v = evalCell(r, c);
    if (v===null||v===undefined||v==="") return "";
    if (typeof v==="number"&&!Number.isInteger(v)) return parseFloat(v.toFixed(10)).toString();
    return String(v);
  };

  const pushHistory = (d, fmt) => { setHistory(h=>[...h.slice(-100),{data:d.map(r=>[...r]),fmt:{...fmt}}]); setFuture([]); };
  const undo = () => { if(!history.length)return; const prev=history[history.length-1]; setFuture(f=>[{data:data.map(r=>[...r]),fmt:{...formatting}},...f]); setHistory(h=>h.slice(0,-1)); setData(prev.data); setFormatting(prev.fmt); };
  const redo = () => { if(!future.length)return; setHistory(h=>[...h,{data:data.map(r=>[...r]),fmt:{...formatting}}]); const next=future[0]; setFuture(f=>f.slice(1)); setData(next.data); setFormatting(next.fmt); };

  const commitEdit = (val) => {
    const v = val!==undefined?val:editValue;
    pushHistory(data,formatting);
    setData(d=>{const nd=d.map(r=>[...r]);nd[selected.r][selected.c]=v;return nd;});
    setEditing(false);
  };
  const startEdit = (r, c, initial) => {
    setSelected({r,c}); setEditing(true);
    setEditValue(initial!==undefined?initial:data[r][c]);
    const fmt=getFmt(r,c);
    setSelFont(fmt.font||"Calibri"); setSelSize(fmt.fontSize||11);
  };
  const clearRange = () => {
    pushHistory(data,formatting);
    const {r1,r2,c1,c2}=selection;
    setData(d=>{const nd=d.map(r=>[...r]);for(let r=Math.min(r1,r2);r<=Math.max(r1,r2);r++)for(let c=Math.min(c1,c2);c<=Math.max(c1,c2);c++)nd[r][c]="";return nd;});
  };
  const applyFmt = (patch) => {
    const {r1,r2,c1,c2}=selection;
    setFormatting(f=>{const nf={...f};for(let r=Math.min(r1,r2);r<=Math.max(r1,r2);r++)for(let c=Math.min(c1,c2);c<=Math.max(c1,c2);c++)nf[fmtKey(r,c)]={...(nf[fmtKey(r,c)]||{}),...patch};return nf;});
  };
  const isInSel = (r, c) => { const {r1,r2,c1,c2}=selection; return r>=Math.min(r1,r2)&&r<=Math.max(r1,r2)&&c>=Math.min(c1,c2)&&c<=Math.max(c1,c2); };
  const navigate = (dr, dc, extend) => {
    const nr=Math.max(0,Math.min(rows-1,selected.r+dr)),nc=Math.max(0,Math.min(cols-1,selected.c+dc));
    if(extend){setSelection(s=>({...s,r2:nr,c2:nc}));}else{setSelected({r:nr,c:nc});setSelection({r1:nr,c1:nc,r2:nr,c2:nc});}
  };

  const handleKeyDown = (e) => {
    if(findOpen) return;
    if(editing){
      if(e.key==="Enter"){e.preventDefault();commitEdit();navigate(1,0);}
      else if(e.key==="Escape"){setEditing(false);}
      else if(e.key==="Tab"){e.preventDefault();commitEdit();navigate(0,e.shiftKey?-1:1);}
      return;
    }
    const sh=e.shiftKey;
    if(e.key==="ArrowUp"){e.preventDefault();navigate(-1,0,sh);}
    else if(e.key==="ArrowDown"){e.preventDefault();navigate(1,0,sh);}
    else if(e.key==="ArrowLeft"){e.preventDefault();navigate(0,-1,sh);}
    else if(e.key==="ArrowRight"){e.preventDefault();navigate(0,1,sh);}
    else if(e.key==="Tab"){e.preventDefault();navigate(0,sh?-1:1);}
    else if(e.key==="Enter"){e.preventDefault();startEdit(selected.r,selected.c);}
    else if(e.key==="F2"){startEdit(selected.r,selected.c);}
    else if((e.key==="Delete"||e.key==="Backspace")&&!editing){clearRange();}
    else if(e.key==="z"&&(e.ctrlKey||e.metaKey)){e.preventDefault();undo();}
    else if(e.key==="y"&&(e.ctrlKey||e.metaKey)){e.preventDefault();redo();}
    else if(e.key==="b"&&(e.ctrlKey||e.metaKey)){e.preventDefault();applyFmt({bold:!getFmt(selected.r,selected.c).bold});}
    else if(e.key==="i"&&(e.ctrlKey||e.metaKey)){e.preventDefault();applyFmt({italic:!getFmt(selected.r,selected.c).italic});}
    else if(e.key==="u"&&(e.ctrlKey||e.metaKey)){e.preventDefault();applyFmt({underline:!getFmt(selected.r,selected.c).underline});}
    else if(e.key==="f"&&(e.ctrlKey||e.metaKey)){e.preventDefault();setFindOpen(true);}
    else if(e.key==="c"&&(e.ctrlKey||e.metaKey)){copySelection();}
    else if(e.key.length===1&&!e.ctrlKey&&!e.metaKey){startEdit(selected.r,selected.c,e.key);}
  };

  const copySelection = () => {
    const {r1,r2,c1,c2}=selection;
    const lines=[];
    for(let r=Math.min(r1,r2);r<=Math.max(r1,r2);r++){const row=[];for(let c=Math.min(c1,c2);c<=Math.max(c1,c2);c++)row.push(displayVal(r,c));lines.push(row.join("\t"));}
    navigator.clipboard?.writeText(lines.join("\n"));
  };
  const handlePaste = (e) => {
    const text=e.clipboardData?.getData("text");if(!text)return;e.preventDefault();
    const lines=text.split("\n").map(l=>l.split("\t"));
    pushHistory(data,formatting);
    setData(d=>{const nd=d.map(r=>[...r]);lines.forEach((line,dr)=>line.forEach((val,dc)=>{const r=selected.r+dr,c=selected.c+dc;if(r<rows&&c<cols)nd[r][c]=val;}));return nd;});
  };

  const insertRow = () => { pushHistory(data,formatting);setData(d=>{const nd=[...d];nd.splice(selected.r,0,Array(cols).fill(""));return nd;});setRows(n=>n+1); };
  const deleteRowFn = () => { if(rows<=1)return;pushHistory(data,formatting);setData(d=>d.filter((_,i)=>i!==selected.r));setRows(n=>n-1); };
  const insertCol = () => { pushHistory(data,formatting);setData(d=>d.map(r=>{const nr=[...r];nr.splice(selected.c,0,"");return nr;}));setColWidths(w=>{const nw=[...w];nw.splice(selected.c,0,DEFAULT_COL_WIDTH);return nw;});setCols(n=>n+1); };
  const deleteColFn = () => { if(cols<=1)return;pushHistory(data,formatting);setData(d=>d.map(r=>r.filter((_,i)=>i!==selected.c)));setColWidths(w=>w.filter((_,i)=>i!==selected.c));setCols(n=>n-1); };
  const addRow = () => { setData(d=>[...d,Array(cols).fill("")]);setRows(n=>n+1); };

  const handleColResizeStart = (e, c) => {
    e.preventDefault();
    resizingRef.current={c,startX:e.clientX,startW:colWidths[c]};
    const onMove=(ev)=>{const{c:rc,startX,startW}=resizingRef.current;setColWidths(w=>{const nw=[...w];nw[rc]=Math.max(30,startW+ev.clientX-startX);return nw;});};
    const onUp=()=>{document.removeEventListener("mousemove",onMove);document.removeEventListener("mouseup",onUp);};
    document.addEventListener("mousemove",onMove);document.addEventListener("mouseup",onUp);
  };

  const saveTabelle = () => {
    const text=serializeTabelle(data,formatting,colWidths,sheets,filename);
    const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([text],{type:"text/plain"}));a.download=`${filename}.tabelle`;a.click();
  };
  const loadTabelle = (e) => {
    const file=e.target.files[0];if(!file)return;
    const name=file.name.replace(/\.tabelle$/i,"");
    const reader=new FileReader();
    reader.onload=(ev)=>{
      const{data:nd,formatting:nfmt,colWidths:nw}=parseTabelle(ev.target.result);
      pushHistory(data,formatting);setFilename(name);setData(nd);setFormatting(nfmt);setColWidths(nw);
      setRows(nd.length);setCols(nd[0]?.length||DEFAULT_COLS);
    };
    reader.readAsText(file);e.target.value="";
  };
  const exportCSV = () => {
    const csv=data.map((row,r)=>row.map((_,c)=>`"${String(displayVal(r,c)).replace(/"/g,'""')}"`).join(",")).join("\n");
    const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv"}));a.download=`${filename}.csv`;a.click();
  };
  const importCSV = (e) => {
    const file=e.target.files[0];if(!file)return;
    const reader=new FileReader();
    reader.onload=(ev)=>{
      const lines=ev.target.result.split("\n").filter(Boolean);
      const parsed=lines.map(l=>l.split(",").map(c=>c.replace(/^"|"$/g,"").replace(/""/g,'"')));
      const nc=Math.max(...parsed.map(r=>r.length)),nr=parsed.length;
      pushHistory(data,formatting);setRows(nr);setCols(nc);setColWidths(Array(nc).fill(DEFAULT_COL_WIDTH));
      setData(parsed.map(r=>{while(r.length<nc)r.push("");return r;}));
    };
    reader.readAsText(file);e.target.value="";
  };
  const findNext = () => {
    if(!findVal)return;
    for(let r=0;r<rows;r++)for(let c=0;c<cols;c++)
      if((r>selected.r||(r===selected.r&&c>selected.c))&&String(data[r][c]).toLowerCase().includes(findVal.toLowerCase()))
        {setSelected({r,c});setSelection({r1:r,c1:c,r2:r,c2:c});return;}
  };
  const replaceAll = () => { pushHistory(data,formatting);setData(d=>d.map(row=>row.map(cell=>typeof cell==="string"?cell.replaceAll(findVal,replaceVal):cell))); };

  useEffect(()=>{
    const fmt=getFmt(selected.r,selected.c);
    setSelFont(fmt.font||"Calibri");setSelSize(fmt.fontSize||11);
  },[selected.r,selected.c,formatting]);

  const sumBar = useMemo(()=>{
    const{r1,r2,c1,c2}=selection;const vals=[];
    for(let r=Math.min(r1,r2);r<=Math.max(r1,r2);r++)for(let c=Math.min(c1,c2);c<=Math.max(c1,c2);c++){const v=parseFloat(displayVal(r,c));if(!isNaN(v))vals.push(v);}
    if(!vals.length)return "";const sum=vals.reduce((a,b)=>a+b,0);
    return `Anzahl: ${vals.length}  Ø ${(sum/vals.length).toFixed(2)}  Σ ${sum.toFixed(2)}`;
  },[selection,data]);

  const TB = ({onClick,active,title,children,style,disabled})=>(
    <button onClick={onClick} title={title} disabled={disabled} style={{
      background:active?"#cce0ff":"#f8f9fa",border:"1px solid #d0d3d8",borderRadius:"3px",
      padding:"2px 7px",fontSize:"12px",cursor:disabled?"not-allowed":"pointer",
      fontFamily:"'Inter',sans-serif",color:active?"#0056d6":"#2c2c2c",
      fontWeight:active?"700":"400",minWidth:"24px",height:"24px",
      display:"flex",alignItems:"center",justifyContent:"center",opacity:disabled?0.4:1,...style
    }}>{children}</button>
  );

  const formulaDisplay = editing ? editValue : (data[selected.r]?.[selected.c] ?? "");

  return (
    <div
      style={{fontFamily:"'Inter',sans-serif",background:"#fff",minHeight:"100vh",display:"flex",flexDirection:"column",fontSize:"13px",color:"#1f1f1f"}}
      onKeyDown={handleKeyDown} onPaste={handlePaste} tabIndex={0} ref={containerRef}
      onMouseUp={()=>setDragSel(false)}
      onContextMenu={e=>{e.preventDefault();setContextMenu({x:e.clientX,y:e.clientY});}}
      onClick={()=>setContextMenu(null)}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; font-family: 'Inter', sans-serif; } :focus { outline: none; }
        .xl-cell { border-right:1px solid #d0d3d8; border-bottom:1px solid #d0d3d8; padding:0 6px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; height:${ROW_HEIGHT}px; line-height:${ROW_HEIGHT}px; cursor:default; position:relative; font-family:'Inter',sans-serif; }
        .xl-cell:hover { background:#e8f0fe !important; }
        .xl-sel { background:#cce0ff !important; }
        .xl-active { background:#fff !important; outline:2px solid #217346 !important; outline-offset:-2px; z-index:2; }
        .xl-hdr { background:#f0f2f5; color:#5f6368; font-size:11px; text-align:center; border-right:1px solid #d0d3d8; border-bottom:2px solid #bbb; user-select:none; height:22px; line-height:22px; font-weight:600; position:relative; font-family:'Inter',sans-serif; }
        .xl-rhdr { background:#f0f2f5; color:#5f6368; font-size:11px; text-align:right; padding-right:6px; border-right:2px solid #bbb; border-bottom:1px solid #d0d3d8; user-select:none; width:${ROW_HEADER_WIDTH}px; min-width:${ROW_HEADER_WIDTH}px; height:${ROW_HEIGHT}px; line-height:${ROW_HEIGHT}px; font-family:'Inter',sans-serif; }
        .col-rh { position:absolute; right:0; top:0; width:5px; height:100%; cursor:col-resize; z-index:10; }
        .col-rh:hover { background:#217346; }
        .xl-cell input { width:100%; height:100%; border:none; background:transparent; font-size:inherit; padding:0 6px; outline:none; color:inherit; font-family:inherit; }
        .rtab { padding:4px 12px; font-size:12px; cursor:pointer; border:none; background:transparent; color:#444; font-family:'Inter',sans-serif; }
        .rtab.active { border-bottom:2px solid #217346; color:#217346; font-weight:600; }
        ::-webkit-scrollbar{width:10px;height:10px;} ::-webkit-scrollbar-track{background:#f0f0f0;} ::-webkit-scrollbar-thumb{background:#bbb;border-radius:5px;}
        .ctx { position:fixed; background:#fff; border:1px solid #ccc; box-shadow:2px 4px 12px rgba(0,0,0,.15); border-radius:4px; z-index:9999; min-width:170px; padding:4px 0; font-family:'Inter',sans-serif; }
        .cxi { padding:6px 16px; font-size:12px; cursor:pointer; font-family:'Inter',sans-serif; } .cxi:hover { background:#e8f0fe; }
        .cxs { height:1px; background:#eee; margin:3px 0; }
        select.rsel { height:24px; font-size:12px; border:1px solid #d0d3d8; border-radius:3px; background:#f8f9fa; color:#2c2c2c; padding:0 2px; cursor:pointer; font-family:'Inter',sans-serif; }
      `}</style>

      {/* Title Bar */}
      <div style={{background:"#217346",color:"#fff",padding:"4px 16px",fontSize:"12px",display:"flex",alignItems:"center",gap:"8px"}}>
        <span style={{fontSize:"15px"}}>📊</span>
        <input value={filename} onChange={e=>setFilename(e.target.value)}
          style={{background:"transparent",border:"none",color:"#fff",fontSize:"13px",fontWeight:600,outline:"none",width:180}} />
        <span style={{opacity:.55,fontSize:"11px",marginLeft:-4}}>.tabelle</span>
      </div>

      {/* Ribbon */}
      <div style={{background:"#f3f3f3",borderBottom:"1px solid #d0d3d8"}}>
        <div style={{display:"flex"}}>
          {["Start","Einfügen","Ansicht"].map((t,i)=>(
            <button key={t} className={`rtab${i===0?" active":""}`}>{t}</button>
          ))}
        </div>
        <div style={{display:"flex",gap:"3px",padding:"4px 8px",flexWrap:"wrap",alignItems:"center",borderTop:"1px solid #e0e0e0"}}>
          <TB onClick={undo} title="Rückgängig (Ctrl+Z)" disabled={!history.length}>↩</TB>
          <TB onClick={redo} title="Wiederholen (Ctrl+Y)" disabled={!future.length}>↪</TB>
          <div style={{width:1,height:20,background:"#d0d3d8",margin:"0 2px"}} />

          {/* Font family selector */}
          <select className="rsel" value={selFont} style={{width:150,fontFamily:selFont}}
            onChange={e=>{setSelFont(e.target.value);applyFmt({font:e.target.value});}} title="Schriftart">
            {FONTS.map(f=><option key={f} value={f} style={{fontFamily:f}}>{f}</option>)}
          </select>

          {/* Font size selector */}
          <select className="rsel" value={selSize} style={{width:54}}
            onChange={e=>{const s=parseInt(e.target.value);setSelSize(s);applyFmt({fontSize:s});}} title="Schriftgröße (pt)">
            {FONT_SIZES.map(s=><option key={s} value={s}>{s}</option>)}
          </select>

          <div style={{width:1,height:20,background:"#d0d3d8",margin:"0 2px"}} />
          <TB onClick={()=>applyFmt({bold:!getFmt(selected.r,selected.c).bold})} active={getFmt(selected.r,selected.c).bold} title="Fett (Ctrl+B)" style={{fontWeight:"bold",fontFamily:"serif",fontSize:13}}>B</TB>
          <TB onClick={()=>applyFmt({italic:!getFmt(selected.r,selected.c).italic})} active={getFmt(selected.r,selected.c).italic} title="Kursiv (Ctrl+I)" style={{fontStyle:"italic"}}>K</TB>
          <TB onClick={()=>applyFmt({underline:!getFmt(selected.r,selected.c).underline})} active={getFmt(selected.r,selected.c).underline} title="Unterstrichen (Ctrl+U)" style={{textDecoration:"underline"}}>U</TB>
          <div style={{width:1,height:20,background:"#d0d3d8",margin:"0 2px"}} />
          <TB onClick={()=>applyFmt({align:"left"})} active={getFmt(selected.r,selected.c).align==="left"} title="Links">≡L</TB>
          <TB onClick={()=>applyFmt({align:"center"})} active={getFmt(selected.r,selected.c).align==="center"} title="Mitte">≡M</TB>
          <TB onClick={()=>applyFmt({align:"right"})} active={getFmt(selected.r,selected.c).align==="right"} title="Rechts">≡R</TB>
          <div style={{width:1,height:20,background:"#d0d3d8",margin:"0 2px"}} />
          <label title="Hintergrundfarbe" style={{display:"flex",alignItems:"center",gap:2,fontSize:11,cursor:"pointer",background:"#f8f9fa",border:"1px solid #d0d3d8",borderRadius:3,height:24,padding:"0 5px"}}>
            🎨<input type="color" style={{width:18,height:18,border:"none",cursor:"pointer",padding:0}} onChange={e=>applyFmt({bg:e.target.value})} />
          </label>
          <label title="Textfarbe" style={{display:"flex",alignItems:"center",gap:2,fontSize:11,cursor:"pointer",background:"#f8f9fa",border:"1px solid #d0d3d8",borderRadius:3,height:24,padding:"0 5px"}}>
            A<input type="color" defaultValue="#000000" style={{width:18,height:18,border:"none",cursor:"pointer",padding:0}} onChange={e=>applyFmt({color:e.target.value})} />
          </label>
          <div style={{width:1,height:20,background:"#d0d3d8",margin:"0 2px"}} />
          <TB onClick={insertRow} title="Zeile einfügen">+Zeile</TB>
          <TB onClick={deleteRowFn} title="Zeile löschen">−Zeile</TB>
          <TB onClick={insertCol} title="Spalte einfügen">+Spalte</TB>
          <TB onClick={deleteColFn} title="Spalte löschen">−Spalte</TB>
          <div style={{width:1,height:20,background:"#d0d3d8",margin:"0 2px"}} />
          {/* .tabelle buttons */}
          <TB onClick={saveTabelle} title="Als .tabelle speichern" style={{background:"#e6f4ec",color:"#1a6336",borderColor:"#a8d5b8",fontWeight:700,gap:4}}>💾 .tabelle</TB>
          <label style={{display:"flex",alignItems:"center"}}>
            <TB title=".tabelle öffnen" style={{background:"#e6f4ec",color:"#1a6336",borderColor:"#a8d5b8",fontWeight:700,gap:4,cursor:"pointer"}}>
              📂 .tabelle
              <input type="file" accept=".tabelle" style={{display:"none"}} onChange={loadTabelle} />
            </TB>
          </label>
          <div style={{width:1,height:20,background:"#d0d3d8",margin:"0 2px"}} />
          <TB onClick={exportCSV} title="CSV exportieren">↓ CSV</TB>
          <label style={{display:"flex",alignItems:"center"}}>
            <TB title="CSV importieren" style={{cursor:"pointer"}}>↑ CSV<input type="file" accept=".csv" style={{display:"none"}} onChange={importCSV} /></TB>
          </label>
          <TB onClick={()=>setFindOpen(o=>!o)} active={findOpen} title="Suchen & Ersetzen (Ctrl+F)">🔍</TB>
        </div>
      </div>

      {/* Find & Replace */}
      {findOpen&&(
        <div style={{background:"#fffbea",borderBottom:"1px solid #e8d060",padding:"5px 12px",display:"flex",gap:8,alignItems:"center",fontSize:12}}>
          <strong>Suchen:</strong>
          <input value={findVal} onChange={e=>setFindVal(e.target.value)} onKeyDown={e=>e.key==="Enter"&&findNext()} style={{border:"1px solid #ccc",borderRadius:3,padding:"2px 6px",width:130}} />
          <strong>Ersetzen:</strong>
          <input value={replaceVal} onChange={e=>setReplaceVal(e.target.value)} style={{border:"1px solid #ccc",borderRadius:3,padding:"2px 6px",width:130}} />
          <button onClick={findNext} style={{background:"#217346",color:"#fff",border:"none",borderRadius:3,padding:"3px 10px",cursor:"pointer",fontSize:12}}>Weiter</button>
          <button onClick={replaceAll} style={{background:"#e08010",color:"#fff",border:"none",borderRadius:3,padding:"3px 10px",cursor:"pointer",fontSize:12}}>Alle ersetzen</button>
          <button onClick={()=>setFindOpen(false)} style={{background:"#e0e0e0",border:"none",borderRadius:3,padding:"3px 10px",cursor:"pointer",fontSize:12}}>✕</button>
        </div>
      )}

      {/* Formula Bar */}
      <div style={{background:"#fff",borderBottom:"1px solid #d0d3d8",display:"flex",alignItems:"center"}}>
        <div style={{background:"#f0f2f5",borderRight:"1px solid #d0d3d8",padding:"2px 8px",minWidth:68,textAlign:"center",fontSize:12,fontWeight:700,color:"#217346",height:27,lineHeight:"23px",fontFamily:"monospace"}}>
          {cellAddr(selected.r,selected.c)}
        </div>
        <div style={{width:28,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,color:"#aaa",borderRight:"1px solid #d0d3d8",height:27}}>ƒx</div>
        <input
          value={formulaDisplay}
          onChange={e=>{setEditValue(e.target.value);setEditing(true);}}
          onBlur={()=>{if(editing)commitEdit(editValue);}}
          onKeyDown={e=>{if(e.key==="Enter"){e.preventDefault();commitEdit(editValue);}if(e.key==="Escape"){setEditing(false);}}}
          style={{flex:1,border:"none",outline:"none",padding:"0 10px",fontSize:13,fontFamily:"'Inter',sans-serif",height:27}}
          placeholder="Wert oder Formel (z.B. =SUM(A1:B5))…"
        />
      </div>

      {/* Spreadsheet */}
      <div style={{flex:1,overflow:"auto"}}>
        <table style={{borderCollapse:"collapse",tableLayout:"fixed"}}>
          <colgroup>
            <col style={{width:ROW_HEADER_WIDTH}} />
            {Array.from({length:cols},(_,c)=><col key={c} style={{width:colWidths[c]}} />)}
          </colgroup>
          <thead>
            <tr>
              <th className="xl-hdr" style={{position:"sticky",top:0,left:0,zIndex:4,borderRight:"2px solid #bbb",background:"#e8eaf0"}} />
              {Array.from({length:cols},(_,c)=>(
                <th key={c} className="xl-hdr" style={{
                  position:"sticky",top:0,zIndex:3,width:colWidths[c],
                  background:isInSel(0,c)?"#c8dff0":selected.c===c?"#d8eddd":"#f0f2f5",
                  color:selected.c===c?"#217346":undefined,
                  fontWeight:selected.c===c?700:600
                }}>
                  {colLabel(c)}
                  <div className="col-rh" onMouseDown={e=>handleColResizeStart(e,c)} />
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({length:rows},(_,r)=>(
              <tr key={r}>
                <td className="xl-rhdr" style={{
                  position:"sticky",left:0,zIndex:1,
                  background:isInSel(r,0)?"#c8dff0":selected.r===r?"#d8eddd":"#f0f2f5",
                  color:selected.r===r?"#217346":undefined,
                  fontWeight:selected.r===r?700:400
                }}>{r+1}</td>
                {Array.from({length:cols},(_,c)=>{
                  const isActive=selected.r===r&&selected.c===c;
                  const inSel=isInSel(r,c);
                  const isEdit=editing&&isActive;
                  const fmt=getFmt(r,c);
                  const dv=displayVal(r,c);
                  const isNum=!isNaN(parseFloat(dv))&&dv!=="";
                  const cf=fmt.font||"Calibri";
                  const cs=fmt.fontSize||13;
                  return (
                    <td key={c}
                      className={`xl-cell${isActive?" xl-active":inSel?" xl-sel":""}`}
                      style={{
                        width:colWidths[c],fontFamily:`"${cf}",Calibri,Arial,sans-serif`,fontSize:cs,
                        fontWeight:fmt.bold?"bold":"normal",fontStyle:fmt.italic?"italic":"normal",
                        textDecoration:fmt.underline?"underline":"none",
                        textAlign:fmt.align||(isNum?"right":"left"),
                        color:fmt.color||"#1f1f1f",
                        background:isActive?"#fff":inSel?"#cce0ff":(fmt.bg||"#fff"),
                      }}
                      onMouseDown={()=>{if(editing)commitEdit();setSelected({r,c});setSelection({r1:r,c1:c,r2:r,c2:c});setDragSel(true);setDragStart({r,c});}}
                      onMouseEnter={()=>{if(dragSel&&dragStart)setSelection({r1:dragStart.r,c1:dragStart.c,r2:r,c2:c});}}
                      onDoubleClick={()=>startEdit(r,c)}
                    >
                      {isEdit?(
                        <input ref={inputRef} value={editValue}
                          onChange={e=>setEditValue(e.target.value)}
                          onBlur={()=>commitEdit()}
                          onKeyDown={e=>{
                            if(e.key==="Enter"){e.stopPropagation();commitEdit();navigate(1,0);}
                            if(e.key==="Escape"){e.stopPropagation();setEditing(false);}
                            if(e.key==="Tab"){e.preventDefault();e.stopPropagation();commitEdit();navigate(0,e.shiftKey?-1:1);}
                          }}
                          style={{fontFamily:`"${cf}",Calibri,Arial,sans-serif`,fontSize:cs,fontWeight:fmt.bold?"bold":"normal",fontStyle:fmt.italic?"italic":"normal",textAlign:fmt.align||(isNum?"right":"left")}}
                        />
                      ):dv}
                    </td>
                  );
                })}
              </tr>
            ))}
            <tr>
              <td colSpan={cols+1} onClick={addRow} style={{textAlign:"center",color:"#bbb",fontSize:11,padding:"4px",cursor:"pointer",borderTop:"1px solid #eee"}}>+ Zeile</td>
            </tr>
          </tbody>
        </table>
      </div>

      {/* Sheet Tabs + Status */}
      <div style={{background:"#f0f2f5",borderTop:"1px solid #d0d3d8",display:"flex",justifyContent:"space-between",alignItems:"stretch"}}>
        <div style={{display:"flex",alignItems:"stretch"}}>
          {sheets.map((s,i)=>(
            <div key={i} onClick={()=>setActiveSheet(i)} style={{
              padding:"4px 18px",fontSize:12,cursor:"pointer",borderRight:"1px solid #d0d3d8",
              background:i===activeSheet?"#fff":"transparent",
              color:i===activeSheet?"#217346":"#555",
              fontWeight:i===activeSheet?600:400,
              borderTop:i===activeSheet?"2px solid #217346":"2px solid transparent",
              display:"flex",alignItems:"center"
            }}>{s}</div>
          ))}
          <button onClick={()=>{const n=`Tabelle${sheets.length+1}`;setSheets(s=>[...s,n]);setActiveSheet(sheets.length);}} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,padding:"0 12px",color:"#888"}}>+</button>
        </div>
        <div style={{padding:"4px 16px",fontSize:11,color:"#666",fontFamily:"monospace",display:"flex",alignItems:"center"}}>{sumBar}</div>
      </div>

      {/* Context Menu */}
      {contextMenu&&(
        <div className="ctx" style={{left:contextMenu.x,top:contextMenu.y}}>
          <div className="cxi" onClick={()=>{copySelection();setContextMenu(null);}}>📋 Kopieren</div>
          <div className="cxs"/>
          <div className="cxi" onClick={()=>{insertRow();setContextMenu(null);}}>↕ Zeile einfügen</div>
          <div className="cxi" onClick={()=>{deleteRowFn();setContextMenu(null);}}>✕ Zeile löschen</div>
          <div className="cxi" onClick={()=>{insertCol();setContextMenu(null);}}>↔ Spalte einfügen</div>
          <div className="cxi" onClick={()=>{deleteColFn();setContextMenu(null);}}>✕ Spalte löschen</div>
          <div className="cxs"/>
          <div className="cxi" onClick={()=>{clearRange();setContextMenu(null);}}>🗑 Inhalt löschen</div>
          <div className="cxs"/>
          <div className="cxi" onClick={()=>{saveTabelle();setContextMenu(null);}}>💾 Als .tabelle speichern</div>
        </div>
      )}
    </div>
  );
}
