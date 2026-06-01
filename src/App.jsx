import { useState, useCallback, useMemo, useEffect } from "react";
import { Store, Users, Calendar, Sparkles, AlertTriangle, CheckCircle2, ChevronRight, ChevronLeft, Search, Building2, MapPin, Clock, UserCheck, Settings, FileSpreadsheet, FileDown, Hash, Info } from "lucide-react";
import * as XLSX from 'xlsx';

// ═══════════════════════════════════════════════════════════════
// APOLLO SHIFT LIBRARY — 所有班次代碼與時段定義
// ═══════════════════════════════════════════════════════════════
const APOLLO_SHIFTS = {
  // 全日班 (正職) — 早班類
  "A013": { desc:"營0900-1800", start:"09:00", end:"18:00", category:"早", type:"FT" },
  "B47":  { desc:"營0930-1830", start:"09:30", end:"18:30", category:"早", type:"FT" },
  "B54":  { desc:"營1000-1900", start:"10:00", end:"19:00", category:"早", type:"FT" },
  "B32":  { desc:"營1030-1930", start:"10:30", end:"19:30", category:"早", type:"FT" },
  "B49":  { desc:"營1100-2000", start:"11:00", end:"20:00", category:"早", type:"FT" },
  "B44":  { desc:"營1130-2030", start:"11:30", end:"20:30", category:"早", type:"FT" },
  // 全日班 (正職) — 中班類
  "B33":  { desc:"營1200-2100", start:"12:00", end:"21:00", category:"中", type:"FT" },
  "B30":  { desc:"營1230-2130", start:"12:30", end:"21:30", category:"中", type:"FT" },
  // 全日班 (正職) — 晚班類
  "B34":  { desc:"營1300-2200", start:"13:00", end:"22:00", category:"晚", type:"FT" },
  "A006": { desc:"營1330-2230", start:"13:30", end:"22:30", category:"晚", type:"FT" },
  "B46":  { desc:"營1330-2230", start:"13:30", end:"22:30", category:"晚", type:"FT" },
  "A020": { desc:"營1400-2300", start:"14:00", end:"23:00", category:"晚", type:"FT" },
  "B58":  { desc:"營1430-2330", start:"14:30", end:"23:30", category:"晚", type:"FT" },
  "A010": { desc:"營1500-2400", start:"15:00", end:"24:00", category:"晚", type:"FT" },
  "B38":  { desc:"營1530-0030", start:"15:30", end:"00:30", category:"晚", type:"FT" },
  // 兼職
  "B40":  { desc:"兼職早-0.5", start:"", end:"", category:"PT早", type:"PT" },
  "B15":  { desc:"兼職中-1",   start:"", end:"", category:"PT中", type:"PT" },
  "B012": { desc:"兼職晚-0.5", start:"", end:"", category:"PT晚", type:"PT" },
  "B23":  { desc:"兼職晚-1",   start:"", end:"", category:"PT晚", type:"PT" },
  "B52":  { desc:"兼職-0",     start:"", end:"", category:"PT彈性", type:"PT" },
  "B018": { desc:"正常班-輪班", start:"", end:"", category:"輪班", type:"FT" },
};

// 內部班別 → Apollo 代碼的優先順序對應
// 系統會依此順序在「該門市可用班別」中找第一個可用的代碼
const APOLLO_PRIORITY = {
  "早": ["B32","B33","B49","B54","B47","B44","A013","B30","B34"],
  "中": ["B30","B33","B34","A020","A006","B32"],
  "晚": ["B34","A006","A020","B46","B58","A010","B38","B33","B30"],
  "A":  ["B40","B52","B15","B012"],
  "B":  ["B15","B52","B40","B012"],
  "C":  ["B15","B23","B012","B52"],
  "D":  ["B23","B012","B52","B15"],
};

// ═══════════════════════════════════════════════════════════════
// STORE DATA — 55 stores with Apollo unitCode & usableShifts
// ═══════════════════════════════════════════════════════════════
const STORES = [
  // 南西區
  { code:"AS57", unitCode:"D021007", name:"南西AS店", mgr:"含含", area:"中山區", district:"南西區", model:6, ft:6, pt:0, floor:"1F", type:"街邊店", usableShifts:["B33","B34"] },
  { code:"AS24", unitCode:"D021006", name:"南西LADY店", mgr:"含含", area:"中山區", district:"南西區", model:6, ft:6, pt:0, floor:"1F", type:"街邊店", usableShifts:["B33","B34"] },
  { code:"AS40", unitCode:"D021002", name:"中山一店", mgr:"含含", area:"中山區", district:"南西區", model:7, ft:6, pt:0, floor:"2F", type:"街邊店", usableShifts:["B33","B34"] },
  { code:"AS62", unitCode:"", name:"南西三店", mgr:"含含", area:"中山區", district:"南西區", model:6, ft:6, pt:0, floor:"2F", type:"街邊店", usableShifts:[] },
  { code:"ASIP", unitCode:"D021004", name:"南西IP店", mgr:"含含", area:"中山區", district:"南西區", model:5, ft:4, pt:1, floor:"1F", type:"街邊店", usableShifts:["B012","B23","B33","B34","B52"] },
  { code:"AS88", unitCode:"D021001", name:"士林一店", mgr:"律廷", area:"士林區", district:"南西區", model:6, ft:5, pt:0, floor:"2F", type:"街邊店", usableShifts:["A010"] },
  { code:"AS07", unitCode:"D021021", name:"南西大碼店", mgr:"RIISA", area:"中山區", district:"南西區", model:5, ft:3, pt:1, floor:"1F", type:"街邊店", usableShifts:["A006","B012","B15","B30","B32","B33","B34","B52"] },
  // 敦南區
  { code:"AS92", unitCode:"D021016", name:"信義ATT店", mgr:"LULU", area:"信義區", district:"敦南區", model:10, ft:7, pt:2, floor:"1F", type:"百貨店", usableShifts:["A006","A020","B012","B23","B54"] },
  { code:"AS37", unitCode:"D021012", name:"南港CITYLINK店", mgr:"LULU", area:"南港區", district:"敦南區", model:3, ft:3, pt:0, floor:"1F", type:"百貨店", usableShifts:["A020","B30","B32","B34"] },
  { code:"AS10", unitCode:"D021009", name:"大安店", mgr:"LULU", area:"大安區", district:"敦南區", model:5, ft:4, pt:1, floor:"1F", type:"街邊店", usableShifts:["A006","A020","B33","B34"] },
  { code:"AS04", unitCode:"D021014", name:"敦南AS店", mgr:"樺樺", area:"大安區", district:"敦南區", model:5, ft:5, pt:0, floor:"2F", type:"街邊店", usableShifts:["A006"] },
  { code:"AS05", unitCode:"D021013", name:"敦南LADY店", mgr:"樺樺", area:"大安區", district:"敦南區", model:5, ft:4, pt:0, floor:"2F", type:"街邊店", usableShifts:["A006","A020","B33"] },
  { code:"AS77", unitCode:"D021010", name:"忠孝CHIC.A店", mgr:"樺樺", area:"大安區", district:"敦南區", model:3, ft:2, pt:1, floor:"1F", type:"街邊店", usableShifts:["A006","B012","B15"] },
  { code:"AS49", unitCode:"D021023", name:"敦南大碼店", mgr:"RIISA", area:"大安區", district:"敦南區", model:4, ft:2, pt:2, floor:"1F", type:"街邊店", usableShifts:["A006","B012","B15","B34"] },
  // 西門區
  { code:"AS61", unitCode:"D021019", name:"西門AS店", mgr:"LULU", area:"萬華區", district:"西門區", model:6, ft:5, pt:1, floor:"2F", type:"街邊店", usableShifts:["A020","B012","B23","B33","B34","B52"] },
  { code:"AS70", unitCode:"D021018", name:"西門三店", mgr:"LULU", area:"萬華區", district:"西門區", model:6, ft:5, pt:1, floor:"2F", type:"街邊店", usableShifts:["A020","B012","B23","B33","B52"] },
  { code:"AS12", unitCode:"D021025", name:"西門大碼店", mgr:"RIISA", area:"萬華區", district:"西門區", model:3, ft:3, pt:1, floor:"1F", type:"街邊店", usableShifts:["A006","A020","B34"] },
  // 新北區
  { code:"AS71", unitCode:"D021017", name:"中和環球店", mgr:"律廷", area:"中和區", district:"新北區", model:3, ft:3, pt:1, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"AT02", unitCode:"D021028", name:"板橋遠百LADY店", mgr:"律廷", area:"板橋區", district:"新北區", model:3, ft:3, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"AS73", unitCode:"D021020", name:"新莊宏匯店", mgr:"LULU", area:"新莊區", district:"新北區", model:7, ft:5, pt:0, floor:"1F", type:"百貨店", usableShifts:["A006","A020","B30","B32","B34"] },
  { code:"AS68", unitCode:"D021024", name:"樹林大碼店", mgr:"RIISA", area:"樹林區", district:"新北區", model:3, ft:3, pt:1, floor:"1F", type:"街邊店", usableShifts:["B012","B15","B32","B34"] },
  // 桃園區
  { code:"AS86", unitCode:"D021022", name:"統領大碼店", mgr:"RIISA", area:"桃園", district:"桃園區", model:3, ft:4, pt:0, floor:"1F", type:"百貨店", usableShifts:["B30","B32"] },
  { code:"AT04", unitCode:"D021026", name:"大江大碼店", mgr:"RIISA", area:"桃園", district:"桃園區", model:3, ft:3, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"AS89", unitCode:"D022002", name:"桃園大江店", mgr:"Cindy", area:"桃園", district:"桃園區", model:7, ft:7, pt:0, floor:"1F", type:"百貨店", usableShifts:["B30","B32","B33","B34"] },
  { code:"AS63", unitCode:"D022006", name:"華泰OUTLET店", mgr:"Cindy", area:"桃園", district:"桃園區", model:6, ft:5, pt:1, floor:"1F", type:"百貨店", usableShifts:["B30","B32","B34"] },
  { code:"AS95", unitCode:"D022004", name:"桃園台茂店", mgr:"Cindy", area:"桃園", district:"桃園區", model:5, ft:5, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B33","B34"] },
  { code:"AS72", unitCode:"D022005", name:"桃園統領店", mgr:"Cindy", area:"桃園", district:"桃園區", model:5, ft:4, pt:1, floor:"1F", type:"百貨店", usableShifts:["B012","B23","B30","B32"] },
  { code:"AS96", unitCode:"D022001", name:"桃園台茂AS店", mgr:"Cindy", area:"桃園", district:"桃園區", model:3, ft:3, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B33","B34"] },
  { code:"AS35", unitCode:"D022008", name:"新竹巨城LADY店", mgr:"Cindy", area:"新竹", district:"桃園區", model:6, ft:5, pt:1, floor:"1F", type:"百貨店", usableShifts:["B012","B23","B30","B32"] },
  // 台中區
  { code:"AS15", unitCode:"D022011", name:"逢甲店", mgr:"小LU", area:"逢甲區", district:"台中區", model:5, ft:5, pt:0, floor:"1F", type:"街邊店", usableShifts:["A006","A010","B32","B34","B38"] },
  { code:"AS44", unitCode:"D022010", name:"文心大碼店", mgr:"小LU", area:"南屯區", district:"台中區", model:3, ft:4, pt:0, floor:"1F", type:"街邊店", usableShifts:["B30","B32","B34"] },
  { code:"AS29", unitCode:"D022016", name:"一中大碼店", mgr:"小LU", area:"一中區", district:"台中區", model:5, ft:4, pt:0, floor:"2F", type:"街邊店", usableShifts:["A006","A010","A020","B30","B34","B44"] },
  { code:"AS17", unitCode:"D022014", name:"一中店", mgr:"77", area:"一中區", district:"台中區", model:6, ft:5, pt:1, floor:"2F", type:"街邊店", usableShifts:["A006","A020","B012","B15","B34"] },
  { code:"AS25", unitCode:"D022015", name:"一中LADY店", mgr:"77", area:"一中區", district:"台中區", model:3, ft:4, pt:0, floor:"1F", type:"街邊店", usableShifts:["A006","A020","B32"] },
  { code:"AS51", unitCode:"D022018", name:"麗寶OUTLET店", mgr:"77", area:"后里區", district:"台中區", model:4, ft:4, pt:0, floor:"1F", type:"百貨店", usableShifts:["B30","B32","B34"] },
  { code:"AS74", unitCode:"", name:"勤美誠品店", mgr:"77", area:"西區", district:"台中區", model:5, ft:4, pt:0, floor:"1F", type:"百貨店", usableShifts:[] },
  { code:"AS66", unitCode:"D022017", name:"台中遠百店", mgr:"77", area:"西屯區", district:"台中區", model:5, ft:4, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"AS52", unitCode:"D022022", name:"中港新光店", mgr:"泡泡", area:"西屯區", district:"台中區", model:5, ft:5, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"AS69", unitCode:"D022009", name:"文心LADY店", mgr:"泡泡", area:"南屯區", district:"台中區", model:4, ft:4, pt:0, floor:"1F", type:"街邊店", usableShifts:["B30","B32","B34"] },
  { code:"ASTH", unitCode:"D022023", name:"台中漢神店", mgr:"泡泡", area:"西屯區", district:"台中區", model:6, ft:5, pt:1, floor:"1F", type:"百貨店", usableShifts:["A006","B32","B33","B34"] },
  // 嘉義區
  { code:"AS84", unitCode:"D023009", name:"嘉義秀泰店", mgr:"泡泡", area:"嘉義", district:"嘉義區", model:5, ft:4, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"AS93", unitCode:"D023008", name:"嘉義文化一店", mgr:"泡泡", area:"嘉義", district:"嘉義區", model:6, ft:4, pt:1, floor:"2F", type:"街邊店", usableShifts:["A020","B012","B15","B32","B34","B52"] },
  { code:"AS97", unitCode:"D023015", name:"嘉義大碼店", mgr:"泡泡", area:"嘉義", district:"嘉義區", model:3, ft:3, pt:0, floor:"1F", type:"街邊店", usableShifts:["B32","B34"] },
  // 台南區
  { code:"AT03", unitCode:"D023016", name:"新台南FOCUS店", mgr:"姚", area:"台南", district:"台南區", model:8, ft:8, pt:0, floor:"1F", type:"百貨店", usableShifts:["B30","B32","B34"] },
  { code:"AT05", unitCode:"D023017", name:"南紡店", mgr:"姚", area:"台南", district:"台南區", model:5, ft:4, pt:0, floor:"1F", type:"百貨店", usableShifts:["B30","B32","B34"] },
  { code:"AS38", unitCode:"D023004", name:"南紡LADY店", mgr:"姚", area:"台南", district:"台南區", model:3, ft:3, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"ATMI", unitCode:"D023018", name:"台南三井OUTLET店", mgr:"姚", area:"台南", district:"台南區", model:5, ft:5, pt:0, floor:"1F", type:"百貨店", usableShifts:["B30","B32","B34"] },
  // 夢時代區
  { code:"AS45", unitCode:"D023010", name:"高雄夢時代店", mgr:"姚", area:"高雄", district:"夢時代區", model:10, ft:9, pt:1, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"AS64", unitCode:"D023012", name:"夢時代大碼店", mgr:"姚", area:"高雄", district:"夢時代區", model:5, ft:5, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"AS82", unitCode:"D023011", name:"高雄夢時代CHIC.A店", mgr:"姚", area:"高雄", district:"夢時代區", model:4, ft:4, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  // 北高雄區
  { code:"AS46", unitCode:"D023006", name:"岡山LADY店", mgr:"姚", area:"高雄", district:"北高雄區", model:5, ft:5, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"AS59", unitCode:"D023007", name:"左營新光店", mgr:"姚", area:"高雄", district:"北高雄區", model:5, ft:5, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"AS83", unitCode:"D023001", name:"三多遠百店", mgr:"姚", area:"高雄", district:"北高雄區", model:4, ft:4, pt:0, floor:"1F", type:"百貨店", usableShifts:["B30","B32","B34"] },
  { code:"AS81", unitCode:"D023014", name:"高雄漢神LADY店", mgr:"姚", area:"高雄", district:"北高雄區", model:5, ft:4, pt:1, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
  { code:"AS79", unitCode:"D023013", name:"高雄漢神CHIC.A店", mgr:"姚", area:"高雄", district:"北高雄區", model:3, ft:3, pt:0, floor:"1F", type:"百貨店", usableShifts:["B32","B34"] },
];
const DISTRICTS = ["南西區","敦南區","西門區","新北區","桃園區","台中區","台南區","嘉義區","夢時代區","北高雄區"];

// ═══════════════════════════════════════════════════════════════
// HELPER: get Apollo code for store + internal shift
// ═══════════════════════════════════════════════════════════════
function getApolloCode(internalShift, store) {
  if (!store?.usableShifts?.length) return "";
  if (!internalShift || ["休","例"].includes(internalShift)) return "";
  const priority = APOLLO_PRIORITY[internalShift] || [];
  for (const code of priority) {
    if (store.usableShifts.includes(code)) return code;
  }
  return store.usableShifts[0] || "";
}

// ═══════════════════════════════════════════════════════════════
// STAFFING MODEL (依組別 + 店型)
// ═══════════════════════════════════════════════════════════════
const getModel = (sz, type) => {
  const base = {3:{wd:2,we:3},4:{wd:2,we:3},5:{wd:3,we:4},6:{wd:4,we:5},7:{wd:4,we:6},8:{wd:5,we:6},9:{wd:5,we:7},10:{wd:6,we:8}};
  const b = base[sz] || base[3];
  const e = type==="百貨店" ? "11:00-20:00" : "12:00-21:00";
  const m = {
    3:[{name:"早",time:e,ppl:1},{name:"晚",time:"13:00-22:00",ppl:1}],
    4:[{name:"早",time:e,ppl:1},{name:"晚",time:"13:00-22:00",ppl:1}],
    5:[{name:"早",time:e,ppl:1},{name:"晚",time:"13:00-22:00",ppl:2}],
    6:[{name:"早",time:e,ppl:1},{name:"晚",time:"13:00-22:00",ppl:3}],
    7:[{name:"早",time:e,ppl:1},{name:"晚",time:"13:00-22:00",ppl:3}],
    8:[{name:"早",time:e,ppl:2},{name:"晚",time:"13:00-22:00",ppl:3}],
    9:[{name:"早",time:e,ppl:1},{name:"晚",time:"13:00-22:00",ppl:4}],
    10:[{name:"早",time:e,ppl:2},{name:"晚",time:"13:00-22:00",ppl:4}],
  };
  const mw = {
    3:[{name:"早",time:e,ppl:1},{name:"晚",time:"13:00-22:00",ppl:2}],
    4:[{name:"早",time:e,ppl:1},{name:"晚",time:"13:00-22:00",ppl:2}],
    5:[{name:"早",time:e,ppl:1},{name:"晚",time:"13:00-22:00",ppl:3}],
    6:[{name:"早",time:e,ppl:1},{name:"中",time:"12:30-21:30",ppl:1},{name:"晚",time:"13:00-22:00",ppl:3}],
    7:[{name:"早",time:e,ppl:1},{name:"中",time:"12:30-21:30",ppl:1},{name:"晚",time:"13:00-22:00",ppl:4}],
    8:[{name:"早",time:e,ppl:2},{name:"中",time:"12:30-21:30",ppl:1},{name:"晚",time:"13:00-22:00",ppl:3}],
    9:[{name:"早",time:e,ppl:1},{name:"中",time:"12:00-21:00",ppl:1},{name:"晚",time:"13:00-22:00",ppl:5}],
    10:[{name:"早",time:e,ppl:2},{name:"中",time:"12:00-21:00",ppl:1},{name:"晚",time:"14:30-23:30",ppl:5}],
  };
  return { weekday:b.wd, weekend:b.we, shifts:{weekday:m[sz]||m[3], weekend:mw[sz]||mw[3]} };
};

const PT_SLOTS = [
  { key:"A", label:"A時段", time:"10:30-18:30", totalHours:8, paidHours:7 },
  { key:"B", label:"B時段", time:"12:00-18:00", totalHours:6, paidHours:6 },
  { key:"C", label:"C時段", time:"12:00-21:00", totalHours:9, paidHours:8 },
  { key:"D", label:"D時段", time:"13:00-22:00", totalHours:9, paidHours:8 },
];
const DEFAULT_FT_SALARY = 40000;
const DEFAULT_PT_HOURLY = 220;

const HOLIDAYS_2026 = {
  "2026-01-01":"元旦","2026-02-15":"小年夜","2026-02-16":"除夕","2026-02-17":"春節初一",
  "2026-02-18":"春節初二","2026-02-19":"春節初三","2026-02-28":"228紀念日",
  "2026-04-04":"兒童節","2026-04-05":"清明節","2026-05-01":"勞動節",
  "2026-06-19":"端午節","2026-09-25":"中秋節","2026-09-28":"教師節",
  "2026-10-10":"國慶日","2026-10-25":"光復節","2026-12-25":"行憲紀念日"
};
const BIWEEK_ANCHOR = new Date(2025, 11, 28);
const getBiweek = (s) => { const d = new Date(s); const diff = Math.floor((d-BIWEEK_ANCHOR)/86400000); return diff<0?0:Math.floor(diff/14)+1; };
const getBiweekRange = (bw) => { const s = new Date(BIWEEK_ANCHOR); s.setDate(s.getDate()+(bw-1)*14); const e = new Date(s); e.setDate(e.getDate()+13); return {start:s, end:e}; };

const LEAVE_TYPES = [
  {key:"特",label:"特休",color:"#6366f1"},{key:"婚",label:"婚假",color:"#ec4899"},{key:"喪",label:"喪假",color:"#6b7280"},
  {key:"事",label:"事假",color:"#f59e0b"},{key:"病",label:"病假",color:"#10b981"},{key:"產",label:"產假",color:"#8b5cf6"},
  {key:"檢",label:"產檢",color:"#a78bfa"},{key:"公",label:"公假",color:"#3b82f6"},
];
const SHIFT_META = {
  "早":{bg:"#dbeafe",fg:"#1e40af"},"中":{bg:"#fef3c7",fg:"#92400e"},"晚":{bg:"#ede9fe",fg:"#5b21b6"},
  "A":{bg:"#e0f2fe",fg:"#0c4a6e"},"B":{bg:"#fce7f3",fg:"#9d174d"},"C":{bg:"#fed7aa",fg:"#9a3412"},"D":{bg:"#d1fae5",fg:"#065f46"},
  "休":{bg:"#dcfce7",fg:"#166534"},"例":{bg:"#bbf7d0",fg:"#14532d"},
};
const DAYS_TW = ["日","一","二","三","四","五","六"];
const daysInMonth = (y,m) => new Date(y,m+1,0).getDate();
const dateKey = (y,m,d) => `${y}-${String(m+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`;

// 預設 Apollo 匯出欄位 (符合分頁2格式)
const APOLLO_EXPORT_HEADERS = ["員工編號","員工姓名","單位代碼","單位名稱","日期","星期","班次代碼","班次名稱","上班時間","下班時間","狀態代碼","備註"];
const TEMPLATE_FILES = ["班表匯入範本阿波蘿對應值_V2.xlsx", "班表匯入範本05月份門市營業部.xlsx"];

export default function SchedulingApp() {
  const [step, setStep] = useState(0);
  const [storeType, setStoreType] = useState(null);
  const [selectedStore, setSelectedStore] = useState(null);
  const [year, setYear] = useState(2026);
  const [month, setMonth] = useState(4);
  const [staffList, setStaffList] = useState([]);
  const [newName, setNewName] = useState("");
  const [newEmpCode, setNewEmpCode] = useState("");
  const [newType, setNewType] = useState("ft");
  const [newPtSlots, setNewPtSlots] = useState([]);
  const [sharedPool, setSharedPool] = useState([]);
  const [newPoolName, setNewPoolName] = useState("");
  const [newPoolDistrict, setNewPoolDistrict] = useState("");
  const [newPoolSlots, setNewPoolSlots] = useState([]);
  const [designatedOff, setDesignatedOff] = useState({});
  const [schedule, setSchedule] = useState({});
  const [leaveData, setLeaveData] = useState({});
  const [searchTerm, setSearchTerm] = useState("");
  const [areaFilter, setAreaFilter] = useState("全部");
  const [autoGenerated, setAutoGenerated] = useState(false);
  const [ftSalary, setFtSalary] = useState(DEFAULT_FT_SALARY);
  const [ptHourly, setPtHourly] = useState(DEFAULT_PT_HOURLY);
  const [showExportSettings, setShowExportSettings] = useState(false);
  const [exportHeaders, setExportHeaders] = useState(APOLLO_EXPORT_HEADERS);
  const [templateInfo, setTemplateInfo] = useState(null);
  const [statusCodes, setStatusCodes] = useState({"特":"ANN","婚":"MAR","喪":"BER","事":"PER","病":"SIK","產":"MAT","檢":"PRE","公":"OFF"});

// ─── 載入 Apollo 範本 (部署版：從 public 資料夾抓 template.xlsx) ───
  useEffect(() => {
    async function load() {
      try {
        const res = await fetch('/template.xlsx');
        if (!res.ok) return; // 沒有範本檔也沒關係，會用內建預設欄位
        const data = await res.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array' });
        const targetSheet = wb.SheetNames[1] || wb.SheetNames[0]; // 分頁2優先
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[targetSheet], { header: 1, defval: "" });
        if (rows.length > 0) {
          let hi = 0;
          for (let i = 0; i < Math.min(rows.length, 5); i++) {
            const f = (rows[i] || []).filter(c => c !== "" && c != null).length;
            if (f >= 2) { hi = i; break; }
          }
          const headers = (rows[hi] || []).map(h => (h ?? "").toString().trim()).filter(h => h);
          if (headers.length > 0) {
            setExportHeaders(headers);
            setTemplateInfo({ filename: 'template.xlsx', sheet: targetSheet, headers, sampleRow: rows[hi + 1] || null });
          }
        }
      } catch (err) { console.log("Template load skipped — 使用內建預設欄位"); }
    }
    load();
  }, []);
  const store = selectedStore ? STORES.find(s => s.code === selectedStore) : null;
  const model = store ? getModel(store.model, store.type) : null;

  const monthDates = useMemo(() => {
    const arr = []; const dim = daysInMonth(year, month);
    for (let d=1; d<=dim; d++) {
      const dk = dateKey(year, month, d); const dow = new Date(year, month, d).getDay();
      arr.push({ day:d, dateKey:dk, dow, biweek:getBiweek(dk), isWeekend:dow===0||dow===6, isHoliday:!!HOLIDAYS_2026[dk], holidayName:HOLIDAYS_2026[dk]||null });
    }
    return arr;
  }, [year, month]);

  const biweeksInMonth = useMemo(() => Array.from(new Set(monthDates.map(d => d.biweek))).sort((a,b) => a-b), [monthDates]);
  const areas = useMemo(() => ["全部", ...Array.from(new Set(STORES.filter(x => !storeType || x.type === storeType).map(x => x.area))).sort()], [storeType]);
  const filteredStores = useMemo(() => STORES.filter(s => {
    if (storeType && s.type !== storeType) return false;
    if (areaFilter !== "全部" && s.area !== areaFilter) return false;
    if (searchTerm && !s.name.includes(searchTerm) && !s.code.includes(searchTerm) && !(s.unitCode||"").includes(searchTerm) && !s.mgr.includes(searchTerm)) return false;
    return true;
  }), [storeType, areaFilter, searchTerm]);

  const addStaff = useCallback(() => {
    if (!newName.trim()) { alert("請輸入姓名"); return; }
    setStaffList(p => [...p, {
      id:`s_${Date.now()}`, empCode:newEmpCode.trim(), name:newName.trim(), type:newType,
      ptSlots: newType==="pt"?(newPtSlots.length>0?[...newPtSlots]:["D"]):null,
    }]);
    setNewName(""); setNewEmpCode(""); setNewPtSlots([]);
  }, [newName, newEmpCode, newType, newPtSlots]);

  const updateStaff = useCallback((id, field, value) => setStaffList(p => p.map(s => s.id===id?{...s,[field]:value}:s)), []);
  const removeStaff = useCallback(id => {
    setStaffList(p => p.filter(s => s.id !== id));
    const clean = (st, set) => { const n = {}; Object.keys(st).forEach(k => { if (!k.startsWith(id+"_")) n[k] = st[k]; }); set(n); };
    clean(schedule, setSchedule); clean(leaveData, setLeaveData); clean(designatedOff, setDesignatedOff);
  }, [schedule, leaveData, designatedOff]);

  const addPool = useCallback(() => {
    if (!newPoolName.trim() || !newPoolDistrict) return;
    setSharedPool(p => [...p, {id:`p_${Date.now()}`, name:newPoolName.trim(), district:newPoolDistrict, slots:newPoolSlots.length>0?[...newPoolSlots]:["D"]}]);
    setNewPoolName(""); setNewPoolSlots([]);
  }, [newPoolName, newPoolDistrict, newPoolSlots]);
  const removePool = useCallback(id => setSharedPool(p => p.filter(x => x.id !== id)), []);
  const countDes = useCallback(sid => monthDates.filter(d => designatedOff[`${sid}_${d.dateKey}`]).length, [designatedOff, monthDates]);
  const toggleDes = useCallback((sid, dk) => setDesignatedOff(p => { const k = `${sid}_${dk}`; const n = {...p}; if (n[k]) delete n[k]; else n[k] = true; return n; }), []);

  // ─── 自動排班 ───
  const generateSchedule = useCallback(() => {
    if (!model || staffList.length === 0) return;
    const newSch = {};
    const ft = staffList.filter(s => s.type === "ft");
    const pt = staffList.filter(s => s.type === "pt");
    const rest = {};
    ft.forEach(p => { rest[p.id] = new Set(); monthDates.forEach(d => { if (designatedOff[`${p.id}_${d.dateKey}`]) rest[p.id].add(d.dateKey); }); });

    biweeksInMonth.forEach(bw => {
      const bd = monthDates.filter(d => d.biweek === bw);
      const target = bd.length >= 14 ? 4 : Math.max(1, Math.round(4*bd.length/14));
      ft.forEach(p => {
        let need = target - bd.filter(d => rest[p.id].has(d.dateKey)).length;
        if (need <= 0) return;
        const cands = bd.filter(d => !rest[p.id].has(d.dateKey)).map(d => {
          let sc = 0;
          if (d.dow === 1 || d.dow === 2) sc += 3;
          if (d.dow === 3) sc += 2; if (d.dow === 4) sc += 1;
          sc -= ft.filter(o => rest[o.id].has(d.dateKey)).length * 2;
          return {d, sc};
        }).sort((a,b) => b.sc - a.sc);
        for (let i=0; i<cands.length && need>0; i++) { rest[p.id].add(cands[i].d.dateKey); need--; }
      });
    });

    ft.forEach(p => { let s=0; for (let i=0; i<monthDates.length; i++) { const d=monthDates[i]; if (rest[p.id].has(d.dateKey)) s=0; else { s++; if (s>=7) { rest[p.id].add(d.dateKey); s=0; } } } });

    ft.forEach(p => {
      biweeksInMonth.forEach(bw => {
        const br = monthDates.filter(d => d.biweek === bw && rest[p.id].has(d.dateKey)).sort((a,b) => a.day - b.day);
        br.forEach((d, i) => { newSch[`${p.id}_${d.dateKey}`] = i < 2 ? "例" : "休"; });
      });
    });

    monthDates.forEach(d => {
      const shifts = d.isWeekend ? model.shifts.weekend : model.shifts.weekday;
      const work = ft.filter(p => !rest[p.id].has(d.dateKey)).sort((a,b) => a.id.localeCompare(b.id));
      let idx = 0;
      shifts.forEach(sh => { for (let k=0; k<sh.ppl && idx<work.length; k++) { newSch[`${work[idx].id}_${d.dateKey}`] = sh.name; idx++; } });
      while (idx < work.length) { const fb = shifts.find(s => s.name==="晚") || shifts[0]; newSch[`${work[idx].id}_${d.dateKey}`] = fb.name; idx++; }
    });

    pt.forEach(p => {
      const slot = (p.ptSlots && p.ptSlots[0]) || "D";
      monthDates.forEach(d => { newSch[`${p.id}_${d.dateKey}`] = (d.isWeekend || d.isHoliday) ? slot : "休"; });
    });

    setSchedule(newSch); setAutoGenerated(true);
  }, [model, staffList, monthDates, designatedOff, biweeksInMonth]);

  const violations = useMemo(() => {
    const v = [];
    if (!model || staffList.length === 0) return v;
    staffList.filter(s => s.type === "ft").forEach(s => {
      biweeksInMonth.forEach(bw => {
        const bd = monthDates.filter(d => d.biweek === bw); if (bd.length < 7) return;
        let r=0, li=0, x=0;
        bd.forEach(d => { const val = schedule[`${s.id}_${d.dateKey}`]; const lv = leaveData[`${s.id}_${d.dateKey}`]; if (val==="例") {r++; li++;} else if (val==="休") {r++; x++;} else if (lv) r++; });
        const need = bd.length >= 14 ? 4 : Math.max(2, Math.round(4*bd.length/14));
        if (r < need) v.push({level:"high", msg:`${s.name} 雙周${bw}僅${r}天休息（需${need}天）`});
        else if (bd.length >= 14 && (li<2 || x<2)) v.push({level:"med", msg:`${s.name} 雙周${bw}例${li}/休${x}（需2+2）`});
      });
      let st=0, ss=null;
      for (let i=0; i<monthDates.length; i++) {
        const d = monthDates[i]; const val = schedule[`${s.id}_${d.dateKey}`]; const lv = leaveData[`${s.id}_${d.dateKey}`];
        if (val==="休"||val==="例"||lv) {st=0; ss=null;}
        else if (val) { st++; if (ss===null) ss=d.day; if (st>6) { v.push({level:"high", msg:`${s.name} 連續工作${st}天(${ss}~${d.day}日)`}); st=0; ss=null; } }
      }
    });
    monthDates.forEach(d => {
      const need = d.isWeekend ? model.weekend : model.weekday; let w=0;
      staffList.forEach(s => { const val = schedule[`${s.id}_${d.dateKey}`]; const lv = leaveData[`${s.id}_${d.dateKey}`]; if (!lv && val && !["休","例"].includes(val)) w++; });
      if (w>0 && w<need) v.push({level:"high", msg:`${month+1}/${d.day}(${DAYS_TW[d.dow]}) 在班${w}人 < 需${need}人`});
    });
    return v;
  }, [schedule, leaveData, staffList, monthDates, biweeksInMonth, model, month]);

  const costBreakdown = useMemo(() => {
    const rows = staffList.map(s => {
      if (s.type === "ft") return {id:s.id, name:s.name, empCode:s.empCode, type:"ft", cost:ftSalary, detail:`月底薪 NT$${ftSalary.toLocaleString()}`};
      const sc = {A:0,B:0,C:0,D:0}; let h = 0;
      monthDates.forEach(d => { const v = schedule[`${s.id}_${d.dateKey}`]; const slot = PT_SLOTS.find(p => p.key === v); if (slot) { sc[v]++; h += slot.paidHours; } });
      const cost = h * ptHourly;
      const parts = PT_SLOTS.filter(p => sc[p.key]>0).map(p => `${p.key}×${sc[p.key]}(${p.paidHours}hr)`);
      const detail = parts.length ? `${parts.join(" + ")} = ${h}hr × NT$${ptHourly} = NT$${cost.toLocaleString()}` : "尚未排班";
      return {id:s.id, name:s.name, empCode:s.empCode, type:"pt", cost, detail, totalHours:h};
    });
    const ftCost = rows.filter(r => r.type==="ft").reduce((a,r) => a+r.cost, 0);
    const ptCost = rows.filter(r => r.type==="pt").reduce((a,r) => a+r.cost, 0);
    const ptH = rows.filter(r => r.type==="pt").reduce((a,r) => a+(r.totalHours||0), 0);
    return {rows, ftCost, ptCost, totalCost:ftCost+ptCost, ftCount:rows.filter(r => r.type==="ft").length, ptCount:rows.filter(r => r.type==="pt").length, ptTotalHours:ptH};
  }, [staffList, schedule, monthDates, ftSalary, ptHourly]);

  // ═══════════════════════════════════════════════════════════════
  // EXPORT — 使用 Apollo 單位代碼與班別代碼
  // ═══════════════════════════════════════════════════════════════
  const downloadSchedule = useCallback(() => {
    if (!store || staffList.length === 0) { alert("請先設定門市與員工"); return; }
    if (!store.unitCode) {
      if (!confirm(`⚠️ ${store.name} 尚未在 Apollo 系統建檔，匯出檔的單位代碼欄位將為空白。\n\n建議先到 Apollo 系統建檔後再匯出。\n\n仍要繼續匯出嗎？`)) return;
    }
    if (store.usableShifts.length === 0 && staffList.some(s => s.type === "ft")) {
      if (!confirm(`⚠️ ${store.name} 沒有設定可使用班別，匯出檔的班次代碼欄位將為空白。\n\n仍要繼續嗎？`)) return;
    }

    const rows = [];
    staffList.forEach((s, idx) => {
      const empCode = s.empCode || `${store.code}-${String(idx+1).padStart(3,"0")}`;
      monthDates.forEach(d => {
        const key = `${s.id}_${d.dateKey}`;
        const shiftVal = schedule[key]; const leaveVal = leaveData[key];

        let apolloCode = "", shiftName = "", startT = "", endT = "", statusCode = "";
        if (leaveVal) {
          statusCode = statusCodes[leaveVal] || "";
        } else if (shiftVal && !["休","例"].includes(shiftVal)) {
          apolloCode = getApolloCode(shiftVal, store);
          const info = APOLLO_SHIFTS[apolloCode];
          if (info) { shiftName = info.desc; startT = info.start; endT = info.end; }
        }

        const base = {
          "員工編號": empCode,
          "員工姓名": s.name,
          "單位代碼": store.unitCode || "",
          "單位名稱": store.name,
          "日期": d.dateKey,
          "星期": DAYS_TW[d.dow],
          "班次代碼": apolloCode,
          "班次名稱": shiftName,
          "上班時間": startT,
          "下班時間": endT,
          "狀態代碼": statusCode,
          "備註": d.isHoliday ? d.holidayName : "",
        };

        // 對映到實際範本 headers (含模糊比對)
        const final = {};
        exportHeaders.forEach(h => {
          if (base[h] !== undefined) { final[h] = base[h]; return; }
          const mk = Object.keys(base).find(bk =>
            h.includes(bk) || bk.includes(h) ||
            (h.includes("編號") && bk.includes("編號")) ||
            (h.includes("姓名") && bk.includes("姓名")) ||
            (h.includes("單位") && bk.includes("單位")) ||
            (h.includes("班次") && bk.includes("班次")) ||
            (h.includes("狀態") && bk.includes("狀態")) ||
            (h.includes("日期") && bk.includes("日期")) ||
            (h.includes("上班") && bk.includes("上班")) ||
            (h.includes("下班") && bk.includes("下班")) ||
            (h.includes("倉庫") && bk.includes("單位"))
          );
          final[h] = mk ? base[mk] : "";
        });
        rows.push(final);
      });
    });

    const ws = XLSX.utils.json_to_sheet(rows, {header:exportHeaders});
    ws['!cols'] = exportHeaders.map(h => ({wch:Math.max((h||"").length*2+4, 12)}));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "阿波蘿匯入");
    XLSX.writeFile(wb, `Apollo匯入_${store.name}_${year}${String(month+1).padStart(2,"0")}.xlsx`);
  }, [store, staffList, schedule, leaveData, monthDates, year, month, exportHeaders, statusCodes]);

  // ─── Styles ───
  const css = {
    app: {fontFamily:"'Noto Sans TC', 'Microsoft JhengHei', sans-serif", minHeight:"100vh", background:"linear-gradient(135deg,#0f172a 0%,#1e1b4b 50%,#0f172a 100%)", color:"#e2e8f0"},
    header: {background:"linear-gradient(90deg,#7c3aed 0%,#2563eb 100%)", padding:"16px 24px", display:"flex", alignItems:"center", justifyContent:"space-between", boxShadow:"0 4px 20px rgba(124,58,237,0.3)", position:"sticky", top:0, zIndex:20, flexWrap:"wrap", gap:"10px"},
    logo: {fontSize:"20px", fontWeight:800, display:"flex", alignItems:"center", gap:"8px"},
    subtitle: {fontSize:"12px", color:"rgba(255,255,255,0.75)", marginTop:"2px"},
    nav: {display:"flex", gap:"6px", alignItems:"center", flexWrap:"wrap"},
    navBtn: (a) => ({padding:"8px 14px", borderRadius:"8px", border:"none", cursor:"pointer", fontSize:"13px", fontWeight:600, background:a?"rgba(255,255,255,0.25)":"rgba(255,255,255,0.08)", color:"#fff", display:"inline-flex", alignItems:"center", gap:"4px"}),
    container: {maxWidth:"1400px", margin:"0 auto", padding:"20px"},
    card: {background:"rgba(30,41,59,0.85)", borderRadius:"14px", border:"1px solid rgba(148,163,184,0.15)", padding:"20px", marginBottom:"16px", boxShadow:"0 8px 32px rgba(0,0,0,0.2)"},
    grid3: {display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))", gap:"12px"},
    typeCard: (s) => ({padding:"28px 24px", borderRadius:"16px", border:s?"2px solid #7c3aed":"1px solid rgba(148,163,184,0.2)", background:s?"linear-gradient(135deg,rgba(124,58,237,0.2),rgba(37,99,235,0.2))":"rgba(15,23,42,0.5)", cursor:"pointer", textAlign:"center"}),
    storeCard: (s, ok) => ({padding:"14px", borderRadius:"10px", border:s?"2px solid #7c3aed":ok?"1px solid rgba(148,163,184,0.15)":"1px solid rgba(239,68,68,0.3)", background:s?"rgba(124,58,237,0.15)":ok?"rgba(15,23,42,0.5)":"rgba(239,68,68,0.05)", cursor:"pointer"}),
    badge: (bg) => ({display:"inline-block", padding:"2px 8px", borderRadius:"4px", fontSize:"11px", fontWeight:600, background:bg, color:"#fff", marginRight:"4px"}),
    input: {padding:"9px 12px", borderRadius:"8px", border:"1px solid rgba(148,163,184,0.3)", background:"rgba(15,23,42,0.8)", color:"#e2e8f0", fontSize:"14px", outline:"none", width:"100%"},
    select: {padding:"9px 12px", borderRadius:"8px", border:"1px solid rgba(148,163,184,0.3)", background:"rgba(15,23,42,0.8)", color:"#e2e8f0", fontSize:"13px", outline:"none", cursor:"pointer"},
    btn: (bg, dis) => ({padding:"9px 18px", borderRadius:"8px", border:"none", background:dis?"rgba(100,116,139,0.4)":(bg||"#7c3aed"), color:"#fff", cursor:dis?"not-allowed":"pointer", fontSize:"13px", fontWeight:600, display:"inline-flex", alignItems:"center", gap:"6px", opacity:dis?0.5:1}),
    chk: (a, c) => ({padding:"6px 12px", borderRadius:"8px", border:`1.5px solid ${a?c:"rgba(148,163,184,0.3)"}`, background:a?c+"22":"transparent", color:a?c:"#94a3b8", cursor:"pointer", fontSize:"12px", fontWeight:600}),
    stepDot: (a, d) => ({width:"28px", height:"28px", borderRadius:"50%", display:"flex", alignItems:"center", justifyContent:"center", fontSize:"12px", fontWeight:700, background:d?"#10b981":a?"#7c3aed":"rgba(148,163,184,0.2)", color:"#fff"}),
  };

  const steps = [
    {n:0, label:"選店型", icon:Building2}, {n:1, label:"選門市", icon:Store}, {n:2, label:"人員設定", icon:Users},
    {n:3, label:"兼職/共用池", icon:UserCheck}, {n:4, label:"指定休假", icon:Calendar}, {n:5, label:"排班產生", icon:Sparkles},
  ];
  const StepInd = () => (
    <div style={{display:"flex", alignItems:"center", justifyContent:"center", gap:"6px", padding:"12px 0", flexWrap:"wrap"}}>
      {steps.map((s, i) => (
        <div key={s.n} style={{display:"flex", alignItems:"center", gap:"6px"}}>
          <div style={css.stepDot(step===s.n, step>s.n)}>{step>s.n ? <CheckCircle2 size={14}/> : <s.icon size={14}/>}</div>
          <span style={{fontSize:"12px", color:step===s.n?"#c4b5fd":step>s.n?"#10b981":"#64748b", fontWeight:step===s.n?700:500}}>{s.label}</span>
          {i<steps.length-1 && <ChevronRight size={14} style={{color:"#475569"}}/>}
        </div>
      ))}
    </div>
  );

  // ─── STEP 0 ───
  if (step === 0) {
    const noApolloCount = STORES.filter(s => !s.unitCode).length;
    return (
      <div style={css.app}>
        <div style={css.header}>
          <div><div style={css.logo}><Sparkles size={22}/> 門市智慧排班系統 v15</div><div style={css.subtitle}>勞基法§36 合規 ｜ Apollo 打卡系統整合版</div></div>
        </div>
        <StepInd/>
        <div style={{maxWidth:"1400px", margin:"0 auto", padding:"0 20px"}}>
          <div style={{background:"rgba(16,185,129,0.1)", border:"1px solid rgba(16,185,129,0.3)", padding:"10px 14px", borderRadius:"10px", marginBottom:"12px"}}>
            <div style={{fontSize:"12px", color:"#34d399", display:"flex", alignItems:"center", gap:"8px", flexWrap:"wrap"}}>
              <CheckCircle2 size={14}/> <strong>Apollo 系統已整合</strong>
              <span>共 {STORES.length} 間門市</span>
              <span>{STORES.length - noApolloCount} 間已建檔（含 Apollo 單位代碼）</span>
              {noApolloCount > 0 && <span style={{color:"#fbbf24"}}>{noApolloCount} 間待建檔</span>}
              <span>共 {Object.keys(APOLLO_SHIFTS).length} 個班次代碼可選</span>
              {templateInfo && <span style={{color:"#a5b4fc"}}>匯出範本：{templateInfo.headers.length} 欄</span>}
            </div>
          </div>
        </div>
        <div style={css.container}>
          <div style={css.card}>
            <h2 style={{margin:"0 0 6px", color:"#f1f5f9"}}>第一步：選擇店型</h2>
            <p style={{margin:"0 0 20px", color:"#94a3b8", fontSize:"13px"}}>百貨店 11:00-22:00；街邊店 12:00-22:00。班別代碼會依各門市可用清單自動選擇。</p>
            <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:"16px"}}>
              {[
                {key:"百貨店", icon:Building2, count:STORES.filter(s => s.type==="百貨店").length, color:"#7c3aed"},
                {key:"街邊店", icon:Store, count:STORES.filter(s => s.type==="街邊店").length, color:"#2563eb"},
              ].map(t => (
                <div key={t.key} style={css.typeCard(storeType===t.key)} onClick={() => setStoreType(t.key)}>
                  <t.icon size={48} style={{color:t.color, marginBottom:"12px"}}/>
                  <div style={{fontSize:"22px", fontWeight:800, color:"#f1f5f9", marginBottom:"4px"}}>{t.key}</div>
                  <div style={{fontSize:"11px", color:"#64748b"}}>共 {t.count} 間門市</div>
                </div>
              ))}
            </div>
          </div>
          {storeType && (
            <div style={{textAlign:"center", padding:"8px"}}>
              <button style={css.btn("#7c3aed")} onClick={() => setStep(1)}>已選 {storeType} — 下一步 <ChevronRight size={16}/></button>
            </div>
          )}
        </div>
      </div>
    );
  }

  // ─── STEP 1 ───
  if (step === 1) {
    return (
      <div style={css.app}>
        <div style={css.header}>
          <div><div style={css.logo}><Store size={22}/> 選擇門市 — {storeType}</div><div style={css.subtitle}>共 {filteredStores.length} 間可選 ｜ 紅色框 = 尚未建檔 Apollo</div></div>
          <div style={css.nav}><button style={css.navBtn(false)} onClick={() => setStep(0)}><ChevronLeft size={14}/> 店型</button></div>
        </div>
        <StepInd/>
        <div style={css.container}>
          <div style={css.card}>
            <div style={{display:"flex", gap:"10px", marginBottom:"16px", flexWrap:"wrap", alignItems:"center"}}>
              <div style={{flex:"1", minWidth:"240px", position:"relative"}}>
                <Search size={16} style={{position:"absolute", left:"10px", top:"50%", transform:"translateY(-50%)", color:"#64748b"}}/>
                <input style={{...css.input, paddingLeft:"34px"}} placeholder="搜尋店名/單位代碼(D02xxxx)/區主管" value={searchTerm} onChange={e => setSearchTerm(e.target.value)}/>
              </div>
              <select style={css.select} value={areaFilter} onChange={e => setAreaFilter(e.target.value)}>
                {areas.map(a => <option key={a} value={a}>{a}</option>)}
              </select>
            </div>
            <div style={css.grid3}>
              {filteredStores.map(s => {
                const hasApollo = !!s.unitCode;
                return (
                  <div key={s.code} style={css.storeCard(selectedStore === s.code, hasApollo)} onClick={() => setSelectedStore(s.code)}>
                    <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"6px"}}>
                      <span style={{fontWeight:700, fontSize:"15px"}}>{s.name}</span>
                      {hasApollo ? <span style={{fontSize:"11px", color:"#34d399", fontFamily:"monospace", padding:"1px 6px", background:"rgba(16,185,129,0.15)", borderRadius:"4px"}}>{s.unitCode}</span>
                                 : <span style={{fontSize:"10px", color:"#fca5a5"}}>未建檔</span>}
                    </div>
                    <div style={{display:"flex", gap:"4px", flexWrap:"wrap", marginBottom:"6px"}}>
                      <span style={css.badge("#7c3aed")}>{s.model}人組</span>
                      <span style={css.badge("#0891b2")}>{s.area}</span>
                      <span style={css.badge("#059669")}>{s.mgr}</span>
                      {s.pt>0 && <span style={css.badge("#f59e0b")}>兼{s.pt}</span>}
                    </div>
                    {s.usableShifts.length > 0 && (
                      <div style={{fontSize:"10px", color:"#a5b4fc", marginBottom:"4px"}}>
                        可用班別：<span style={{fontFamily:"monospace", color:"#cbd5e1"}}>{s.usableShifts.join(" / ")}</span>
                      </div>
                    )}
                    <div style={{fontSize:"11px", color:"#94a3b8"}}>正職 {s.ft} + 兼職 {s.pt} ｜ {s.floor} ｜ {s.district}</div>
                  </div>
                );
              })}
            </div>
          </div>
          {selectedStore && (
            <div style={{textAlign:"center", padding:"8px"}}>
              <button style={css.btn("#7c3aed")} onClick={() => setStep(2)}>選定 {store.name} — 進入人員設定 <ChevronRight size={16}/></button>
            </div>
          )}
        </div>
      </div>
    );
  }

  // ─── STEP 2 ───
  if (step === 2) {
    return (
      <div style={css.app}>
        <div style={css.header}>
          <div><div style={css.logo}><Users size={22}/> 人員配置 — {store?.name}</div><div style={css.subtitle}>單位代碼: {store?.unitCode || "未建檔"} ｜ 可用班別: {store?.usableShifts.join("/") || "無"}</div></div>
          <div style={css.nav}>
            <button style={css.navBtn(false)} onClick={() => setStep(1)}><ChevronLeft size={14}/> 門市</button>
            <button style={css.navBtn(false)} onClick={() => setStep(3)} disabled={staffList.length===0}>兼職設定 <ChevronRight size={14}/></button>
          </div>
        </div>
        <StepInd/>
        <div style={css.container}>
          <div style={css.card}>
            <h3 style={{margin:"0 0 14px", color:"#c4b5fd", display:"flex", alignItems:"center", gap:"6px"}}><Settings size={16}/> 排班模組與 Apollo 對應</h3>
            <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:"12px", marginBottom:"12px"}}>
              <div style={{padding:"14px", borderRadius:"10px", background:"rgba(99,102,241,0.1)", border:"1px solid rgba(99,102,241,0.25)"}}>
                <div style={{fontSize:"11px", color:"#94a3b8"}}>平日需求 {model?.weekday} 人</div>
                <div style={{fontSize:"11px", color:"#64748b", marginTop:"6px"}}>{model?.shifts.weekday.map(s => `${s.name}${s.ppl}人 (${s.time})`).join(" ｜ ")}</div>
              </div>
              <div style={{padding:"14px", borderRadius:"10px", background:"rgba(236,72,153,0.1)", border:"1px solid rgba(236,72,153,0.25)"}}>
                <div style={{fontSize:"11px", color:"#94a3b8"}}>假日需求 {model?.weekend} 人</div>
                <div style={{fontSize:"11px", color:"#64748b", marginTop:"6px"}}>{model?.shifts.weekend.map(s => `${s.name}${s.ppl}人 (${s.time})`).join(" ｜ ")}</div>
              </div>
            </div>
            {store?.usableShifts.length > 0 && (
              <div style={{padding:"10px 14px", borderRadius:"8px", background:"rgba(16,185,129,0.08)", border:"1px solid rgba(16,185,129,0.25)"}}>
                <div style={{fontSize:"12px", color:"#34d399", marginBottom:"4px", fontWeight:600}}>📋 本店 Apollo 可用班別對應</div>
                <div style={{display:"flex", flexWrap:"wrap", gap:"6px", marginTop:"6px"}}>
                  {store.usableShifts.map(code => {
                    const info = APOLLO_SHIFTS[code];
                    if (!info) return null;
                    return (
                      <div key={code} style={{padding:"4px 10px", borderRadius:"6px", background:"rgba(15,23,42,0.6)", border:"1px solid rgba(148,163,184,0.2)", fontSize:"11px"}}>
                        <span style={{fontFamily:"monospace", color:"#fbbf24", fontWeight:700}}>{code}</span>
                        <span style={{color:"#94a3b8", marginLeft:"6px"}}>{info.desc}</span>
                        {info.start && <span style={{color:"#34d399", marginLeft:"4px", fontSize:"10px"}}>({info.start}-{info.end})</span>}
                      </div>
                    );
                  })}
                </div>
                <div style={{fontSize:"10px", color:"#94a3b8", marginTop:"8px"}}>系統會在內部排班別（早/中/晚）時，自動從上述清單挑選最匹配的代碼填入匯出檔。</div>
              </div>
            )}
            {(!store?.usableShifts.length) && (
              <div style={{padding:"10px 14px", borderRadius:"8px", background:"rgba(239,68,68,0.1)", border:"1px solid rgba(239,68,68,0.3)", fontSize:"12px", color:"#fca5a5"}}>
                ⚠️ 本店尚未在 Apollo 系統中設定可使用班別，匯出檔的班次代碼欄位將為空白。請聯絡資訊部建檔。
              </div>
            )}
          </div>

          <div style={css.card}>
            <h3 style={{margin:"0 0 12px", color:"#c4b5fd", display:"flex", alignItems:"center", gap:"6px"}}>
              <Users size={16}/> 新增人員
              <span style={{fontSize:"11px", fontWeight:400, color:"#94a3b8", marginLeft:"6px"}}>工號用於 Apollo 系統識別</span>
            </h3>
            <div style={{display:"flex", gap:"8px", alignItems:"center", flexWrap:"wrap"}}>
              <div style={{position:"relative"}}>
                <Hash size={14} style={{position:"absolute", left:"10px", top:"50%", transform:"translateY(-50%)", color:"#fbbf24"}}/>
                <input style={{...css.input, width:"160px", paddingLeft:"30px"}} placeholder="工號" value={newEmpCode} onChange={e => setNewEmpCode(e.target.value)} onKeyDown={e => e.key==="Enter" && addStaff()}/>
              </div>
              <input style={{...css.input, width:"160px"}} placeholder="姓名" value={newName} onChange={e => setNewName(e.target.value)} onKeyDown={e => e.key==="Enter" && addStaff()}/>
              <select style={css.select} value={newType} onChange={e => { setNewType(e.target.value); setNewPtSlots([]); }}>
                <option value="ft">正職</option><option value="pt">兼職</option>
              </select>
              {newType === "pt" && (
                <div style={{display:"flex", gap:"4px", flexWrap:"wrap", alignItems:"center"}}>
                  <span style={{fontSize:"12px", color:"#94a3b8"}}>時段:</span>
                  {PT_SLOTS.map(s => (
                    <button key={s.key} type="button" style={css.chk(newPtSlots.includes(s.key), "#f59e0b")} onClick={() => setNewPtSlots(p => p.includes(s.key) ? p.filter(x => x !== s.key) : [...p, s.key])}>{s.label}</button>
                  ))}
                </div>
              )}
              <button style={css.btn("#7c3aed")} onClick={addStaff}>＋ 新增</button>
            </div>
            {staffList.length > 0 && (
              <div style={{marginTop:"18px"}}>
                <div style={{fontSize:"12px", color:"#94a3b8", marginBottom:"8px"}}>
                  {staffList.filter(s => s.type==="ft").length} 位正職 ／ {staffList.filter(s => s.type==="pt").length} 位兼職
                </div>
                <div style={{display:"flex", flexDirection:"column", gap:"6px"}}>
                  {staffList.map(s => (
                    <div key={s.id} style={{display:"flex", alignItems:"center", gap:"10px", padding:"8px 12px", borderRadius:"8px", background:s.type==="ft"?"rgba(99,102,241,0.12)":"rgba(245,158,11,0.12)", border:`1px solid ${s.type==="ft"?"rgba(99,102,241,0.3)":"rgba(245,158,11,0.3)"}`}}>
                      <input value={s.empCode || ""} placeholder="工號未填" onChange={e => updateStaff(s.id, "empCode", e.target.value)} style={{...css.input, padding:"4px 8px", fontSize:"12px", width:"110px", background:"rgba(15,23,42,0.7)"}}/>
                      <input value={s.name} onChange={e => updateStaff(s.id, "name", e.target.value)} style={{...css.input, padding:"4px 8px", fontSize:"13px", fontWeight:600, width:"120px", background:"rgba(15,23,42,0.7)"}}/>
                      <span style={css.badge(s.type==="ft"?"#6366f1":"#f59e0b")}>{s.type==="ft"?"正職":"兼職"}</span>
                      {s.type==="pt" && s.ptSlots && <span style={{fontSize:"11px", color:"#fbbf24"}}>{s.ptSlots.join(",")}</span>}
                      <button style={{background:"none", border:"none", color:"#ef4444", cursor:"pointer", fontSize:"18px", marginLeft:"auto"}} onClick={() => removeStaff(s.id)}>×</button>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
          {staffList.length > 0 && (
            <div style={{textAlign:"center", padding:"8px"}}>
              <button style={css.btn("#7c3aed")} onClick={() => setStep(3)}>已設定 {staffList.length} 位 — 下一步 <ChevronRight size={16}/></button>
            </div>
          )}
        </div>
      </div>
    );
  }

  // ─── STEP 3 ───
  if (step === 3) {
    return (
      <div style={css.app}>
        <div style={css.header}>
          <div><div style={css.logo}><UserCheck size={22}/> 兼職時段與共用池</div></div>
          <div style={css.nav}>
            <button style={css.navBtn(false)} onClick={() => setStep(2)}><ChevronLeft size={14}/> 人員</button>
            <button style={css.navBtn(false)} onClick={() => setStep(4)}>指定休假 <ChevronRight size={14}/></button>
          </div>
        </div>
        <StepInd/>
        <div style={css.container}>
          <div style={css.card}>
            <h3 style={{margin:"0 0 12px", color:"#c4b5fd", display:"flex", alignItems:"center", gap:"6px"}}><Clock size={16}/> 四種兼職時段</h3>
            <div style={{display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))", gap:"10px"}}>
              {PT_SLOTS.map(s => (
                <div key={s.key} style={{padding:"12px", borderRadius:"10px", background:"rgba(245,158,11,0.08)", border:"1px solid rgba(245,158,11,0.2)"}}>
                  <div style={{fontWeight:700, color:"#fbbf24", fontSize:"15px"}}>{s.label} — {s.time}</div>
                  <div style={{fontSize:"11px", color:"#94a3b8"}}>支薪 {s.paidHours}h</div>
                </div>
              ))}
            </div>
          </div>
          <div style={css.card}>
            <h3 style={{margin:"0 0 12px", color:"#c4b5fd", display:"flex", alignItems:"center", gap:"6px"}}><MapPin size={16}/> 共用人力池（{store?.district}）</h3>
            <div style={{display:"flex", gap:"8px", flexWrap:"wrap", alignItems:"center"}}>
              <input style={{...css.input, width:"160px"}} placeholder="兼職姓名" value={newPoolName} onChange={e => setNewPoolName(e.target.value)}/>
              <select style={css.select} value={newPoolDistrict} onChange={e => setNewPoolDistrict(e.target.value)}>
                <option value="">-- 商圈 --</option>
                {DISTRICTS.map(d => <option key={d} value={d}>{d}</option>)}
              </select>
              <div style={{display:"flex", gap:"4px", flexWrap:"wrap"}}>
                {PT_SLOTS.map(s => (
                  <button key={s.key} type="button" style={css.chk(newPoolSlots.includes(s.key), "#f59e0b")} onClick={() => setNewPoolSlots(p => p.includes(s.key) ? p.filter(x => x !== s.key) : [...p, s.key])}>{s.label}</button>
                ))}
              </div>
              <button style={css.btn("#f59e0b")} onClick={addPool}>＋ 加入</button>
            </div>
            {sharedPool.length > 0 && (
              <div style={{marginTop:"16px", display:"flex", flexWrap:"wrap", gap:"8px"}}>
                {sharedPool.map(p => {
                  const sd = p.district === store?.district;
                  return (
                    <div key={p.id} style={{display:"flex", alignItems:"center", gap:"6px", padding:"8px 12px", borderRadius:"8px", background:sd?"rgba(16,185,129,0.15)":"rgba(100,116,139,0.15)", border:`1px solid ${sd?"rgba(16,185,129,0.4)":"rgba(100,116,139,0.3)"}`}}>
                      <span style={{fontWeight:600, fontSize:"13px"}}>{p.name}</span>
                      <span style={css.badge(sd?"#10b981":"#64748b")}>{p.district}</span>
                      {sd && <span style={{fontSize:"10px", color:"#10b981"}}>✓ 可支援</span>}
                      <button style={{background:"none", border:"none", color:"#ef4444", cursor:"pointer"}} onClick={() => removePool(p.id)}>×</button>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
          <div style={{textAlign:"center", padding:"8px"}}>
            <button style={css.btn("#7c3aed")} onClick={() => setStep(4)}>下一步 <ChevronRight size={16}/></button>
          </div>
        </div>
      </div>
    );
  }

  // ─── STEP 4 ───
  if (step === 4) {
    const ft = staffList.filter(s => s.type === "ft");
    return (
      <div style={css.app}>
        <div style={css.header}>
          <div><div style={css.logo}><Calendar size={22}/> 正職指定休假（最多 3 日）</div></div>
          <div style={css.nav}>
            <button style={css.navBtn(false)} onClick={() => setStep(3)}><ChevronLeft size={14}/> 共用池</button>
            <select style={{...css.select, minWidth:"80px"}} value={year} onChange={e => setYear(+e.target.value)}>{[2025,2026,2027].map(y => <option key={y} value={y}>{y}</option>)}</select>
            <select style={{...css.select, minWidth:"80px"}} value={month} onChange={e => setMonth(+e.target.value)}>{["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"].map((m,i) => <option key={i} value={i}>{m}</option>)}</select>
            <button style={css.navBtn(false)} onClick={() => setStep(5)}>產生班表 <ChevronRight size={14}/></button>
          </div>
        </div>
        <StepInd/>
        <div style={css.container}>
          <div style={css.card}>
            <h3 style={{margin:"0 0 14px", color:"#c4b5fd"}}>{year} 年 {month+1} 月</h3>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%", borderCollapse:"separate", borderSpacing:0, fontSize:"12px"}}>
                <thead>
                  <tr>
                    <th style={{padding:"8px", background:"rgba(124,58,237,0.2)", color:"#c4b5fd", textAlign:"left", position:"sticky", left:0, zIndex:2, minWidth:"80px"}}>人員</th>
                    <th style={{padding:"8px", background:"rgba(124,58,237,0.2)", color:"#c4b5fd"}}>已選</th>
                    {monthDates.map(d => (
                      <th key={d.day} style={{padding:"4px 2px", background:d.isHoliday?"rgba(239,68,68,0.2)":d.isWeekend?"rgba(236,72,153,0.15)":"rgba(124,58,237,0.2)", color:"#c4b5fd", minWidth:"32px", fontSize:"10px"}}>
                        <div>{d.day}</div><div style={{fontSize:"9px", color:d.isWeekend?"#f472b6":"#94a3b8"}}>{DAYS_TW[d.dow]}</div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {ft.map(s => {
                    const cnt = countDes(s.id);
                    return (
                      <tr key={s.id}>
                        <td style={{padding:"6px 8px", position:"sticky", left:0, background:"rgba(30,41,59,0.95)", zIndex:1, fontWeight:600, borderBottom:"1px solid rgba(148,163,184,0.1)"}}>{s.name}</td>
                        <td style={{padding:"6px 8px", textAlign:"center", borderBottom:"1px solid rgba(148,163,184,0.1)", color:cnt>3?"#ef4444":cnt===3?"#10b981":"#fbbf24", fontWeight:700}}>{cnt}/3</td>
                        {monthDates.map(d => {
                          const des = designatedOff[`${s.id}_${d.dateKey}`];
                          return (
                            <td key={d.day} style={{padding:"2px", textAlign:"center", borderBottom:"1px solid rgba(148,163,184,0.1)"}}>
                              <button style={{width:"28px", height:"28px", borderRadius:"4px", border:"1px solid transparent", background:des?"#f59e0b":d.isWeekend?"rgba(236,72,153,0.1)":"rgba(15,23,42,0.3)", color:des?"#fff":"#475569", cursor:"pointer", fontSize:"11px", fontWeight:600}} onClick={() => toggleDes(s.id, d.dateKey)}>{des?"休":""}</button>
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
          <div style={{textAlign:"center", padding:"8px"}}>
            <button style={css.btn("#7c3aed")} onClick={() => setStep(5)}>產生排班表 <Sparkles size={16}/></button>
          </div>
        </div>
      </div>
    );
  }

  // ─── STEP 5 ───
  const monthNames = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"];
  const highV = violations.filter(v => v.level === "high");
  const medV = violations.filter(v => v.level === "med");
  const cellStyle = (val, lv, isDes) => {
    let bg = "rgba(15,23,42,0.3)", fg = "#475569", label = "";
    if (lv) { const lt = LEAVE_TYPES.find(l => l.key === lv); bg = lt ? lt.color+"33" : "#333"; fg = lt ? lt.color : "#fff"; label = lv; }
    else if (val && SHIFT_META[val]) { bg = SHIFT_META[val].bg; fg = SHIFT_META[val].fg; label = val; }
    else if (val) { label = val; }
    return { bg, fg, label, border: isDes ? "2px solid #f59e0b" : "1px solid transparent" };
  };

  return (
    <div style={css.app}>
      <div style={css.header}>
        <div>
          <div style={css.logo}><Calendar size={22}/> {store?.name} — {year}年{monthNames[month]}班表</div>
          <div style={css.subtitle}>單位代碼 {store?.unitCode || "—"} ｜ 可用班別 {store?.usableShifts.join("/") || "—"}</div>
        </div>
        <div style={css.nav}>
          <button style={css.navBtn(false)} onClick={() => setStep(4)}><ChevronLeft size={14}/> 指定休</button>
          <select style={{...css.select, minWidth:"80px"}} value={month} onChange={e => setMonth(+e.target.value)}>
            {monthNames.map((m,i) => <option key={i} value={i}>{m}</option>)}
          </select>
          <button style={css.btn("#059669")} onClick={generateSchedule}><Sparkles size={14}/> {autoGenerated?"重新":""}自動排班</button>
          <button style={css.btn("#2563eb")} onClick={downloadSchedule}><FileDown size={14}/> 下載 Apollo 匯入檔</button>
          <button style={css.navBtn(false)} onClick={() => setShowExportSettings(v => !v)}><Settings size={14}/> 匯出設定</button>
        </div>
      </div>
      <StepInd/>

      <div style={css.container}>
        {showExportSettings && (
          <div style={{...css.card, border:"1px solid rgba(37,99,235,0.4)", background:"rgba(37,99,235,0.06)"}}>
            <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"14px"}}>
              <h3 style={{margin:0, color:"#93c5fd", fontSize:"15px", display:"flex", alignItems:"center", gap:"6px"}}><FileSpreadsheet size={16}/> Apollo 匯出設定</h3>
              <button style={{...css.btn("rgba(100,116,139,0.4)"), padding:"4px 10px"}} onClick={() => setShowExportSettings(false)}>收合</button>
            </div>
            <div style={{padding:"10px 12px", borderRadius:"8px", background:"rgba(15,23,42,0.5)", marginBottom:"10px", fontSize:"12px"}}>
              <div style={{fontWeight:700, color:"#c4b5fd", marginBottom:"6px"}}>本店 Apollo 對應</div>
              <div style={{color:"#94a3b8"}}>單位代碼：<strong style={{color:"#34d399", fontFamily:"monospace"}}>{store?.unitCode || "未建檔"}</strong></div>
              <div style={{color:"#94a3b8"}}>可用班別：<span style={{fontFamily:"monospace", color:"#fbbf24"}}>{store?.usableShifts.join(" / ") || "無"}</span></div>
              <div style={{color:"#94a3b8", marginTop:"6px", fontSize:"11px"}}>
                內部「早班」→ Apollo <strong style={{color:"#34d399"}}>{getApolloCode("早", store) || "—"}</strong>
                {model?.shifts.weekend.some(s => s.name==="中") && <> ／「中班」→ Apollo <strong style={{color:"#34d399"}}>{getApolloCode("中", store) || "—"}</strong></>}
                {" "}／「晚班」→ Apollo <strong style={{color:"#34d399"}}>{getApolloCode("晚", store) || "—"}</strong>
              </div>
            </div>
            <div style={{padding:"10px 12px", borderRadius:"8px", background:"rgba(15,23,42,0.5)", marginBottom:"10px"}}>
              <div style={{fontSize:"12px", fontWeight:700, color:"#c4b5fd", marginBottom:"6px"}}>📋 匯出欄位（{exportHeaders.length} 欄）</div>
              <div style={{fontSize:"11px", color:"#cbd5e1", fontFamily:"monospace", padding:"6px 10px", background:"rgba(0,0,0,0.3)", borderRadius:"4px"}}>
                {exportHeaders.join(" ｜ ")}
              </div>
              {templateInfo && <div style={{fontSize:"10px", color:"#34d399", marginTop:"4px"}}>✅ 已從範本 {templateInfo.filename} 讀取</div>}
            </div>
            <div style={{padding:"10px 12px", borderRadius:"8px", background:"rgba(15,23,42,0.5)", marginBottom:"14px"}}>
              <div style={{fontSize:"12px", fontWeight:700, color:"#c4b5fd", marginBottom:"6px"}}>🏖️ 狀態代碼（可修改）</div>
              <div style={{display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))", gap:"6px"}}>
                {Object.entries(statusCodes).map(([k,v]) => {
                  const lt = LEAVE_TYPES.find(l => l.key === k);
                  return (
                    <div key={k} style={{display:"flex", alignItems:"center", gap:"4px"}}>
                      <span style={{...css.badge(lt?.color||"#6b7280"), minWidth:"40px", textAlign:"center"}}>{lt?.label||k}</span>
                      <input value={v} onChange={e => setStatusCodes(p => ({...p, [k]:e.target.value}))} style={{...css.input, padding:"2px 6px", fontSize:"11px", width:"60px"}}/>
                    </div>
                  );
                })}
              </div>
            </div>
            <button style={css.btn("#2563eb")} onClick={downloadSchedule}><FileDown size={14}/> 下載 Apollo 匯入檔</button>
          </div>
        )}

        <div style={{...css.card, padding:"12px 16px"}}>
          <div style={{display:"flex", gap:"10px", flexWrap:"wrap", alignItems:"center"}}>
            <span style={{fontSize:"12px", color:"#94a3b8", fontWeight:600}}>圖例：</span>
            {Object.entries(SHIFT_META).map(([k,v]) => (
              <span key={k} style={{padding:"3px 10px", borderRadius:"6px", background:v.bg, color:v.fg, fontSize:"11px", fontWeight:600}}>{k}</span>
            ))}
            <span style={{marginLeft:"auto", fontSize:"11px", color:"#94a3b8"}}>滑鼠停留班別格 = 顯示 Apollo 代碼</span>
          </div>
        </div>

        {violations.length > 0 && (
          <div style={{...css.card, background:highV.length>0?"rgba(239,68,68,0.1)":"rgba(245,158,11,0.1)", border:`1px solid ${highV.length>0?"rgba(239,68,68,0.3)":"rgba(245,158,11,0.3)"}`, padding:"14px 18px"}}>
            <div style={{display:"flex", alignItems:"center", gap:"8px", marginBottom:"8px"}}>
              <AlertTriangle size={18} style={{color:highV.length>0?"#fca5a5":"#fbbf24"}}/>
              <span style={{fontWeight:700, fontSize:"14px", color:highV.length>0?"#fca5a5":"#fbbf24"}}>合規檢查：{highV.length} 嚴重 / {medV.length} 警告</span>
            </div>
            <div style={{maxHeight:"80px", overflow:"auto"}}>
              {violations.slice(0,8).map((v,i) => <div key={i} style={{fontSize:"12px", color:v.level==="high"?"#fca5a5":"#fbbf24", padding:"2px 0"}}>{v.level==="high"?"🔴":"🟡"} {v.msg}</div>)}
            </div>
          </div>
        )}

        <div style={{...css.card, padding:0, overflow:"auto", maxHeight:"calc(100vh - 320px)"}}>
          <table style={{width:"100%", borderCollapse:"separate", borderSpacing:0, fontSize:"12px"}}>
            <thead>
              <tr>
                <th style={{padding:"8px", background:"rgba(124,58,237,0.3)", color:"#c4b5fd", fontWeight:700, position:"sticky", top:0, left:0, zIndex:3, borderBottom:"2px solid #7c3aed", textAlign:"left", minWidth:"100px"}}>人員</th>
                {monthDates.map(d => (
                  <th key={d.day} style={{padding:"6px 4px", background:d.isHoliday?"rgba(239,68,68,0.25)":d.isWeekend?"rgba(236,72,153,0.2)":"rgba(124,58,237,0.2)", color:"#c4b5fd", position:"sticky", top:0, zIndex:2, borderBottom:"2px solid #7c3aed", textAlign:"center", minWidth:"42px"}}>
                    <div style={{fontSize:"13px"}}>{d.day}</div>
                    <div style={{fontSize:"10px", color:d.isWeekend?"#f472b6":"#94a3b8"}}>{DAYS_TW[d.dow]}</div>
                    <div style={{fontSize:"8px", color:"#64748b"}}>雙{d.biweek}</div>
                  </th>
                ))}
                <th style={{padding:"8px", background:"rgba(124,58,237,0.3)", color:"#c4b5fd", position:"sticky", top:0, zIndex:2, minWidth:"70px", borderBottom:"2px solid #7c3aed"}}>統計</th>
              </tr>
            </thead>
            <tbody>
              {staffList.map(s => {
                let w=0, r=0, lv=0;
                monthDates.forEach(d => {
                  const val = schedule[`${s.id}_${d.dateKey}`]; const lvv = leaveData[`${s.id}_${d.dateKey}`];
                  if (lvv) lv++; else if (val==="休"||val==="例") r++; else if (val) w++;
                });
                return (
                  <tr key={s.id}>
                    <td style={{padding:"6px 8px", textAlign:"left", borderBottom:"1px solid rgba(148,163,184,0.08)", position:"sticky", left:0, background:"rgba(30,41,59,0.95)", zIndex:1, whiteSpace:"nowrap"}}>
                      <div style={{display:"flex", alignItems:"center", gap:"6px"}}>
                        <span style={{width:"6px", height:"6px", borderRadius:"50%", background:s.type==="ft"?"#818cf8":"#f59e0b"}}/>
                        <span style={{fontWeight:600}}>{s.name}</span>
                        {s.empCode && <span style={{fontSize:"9px", color:"#fbbf24", padding:"1px 4px", background:"rgba(245,158,11,0.15)", borderRadius:"3px"}}>{s.empCode}</span>}
                      </div>
                    </td>
                    {monthDates.map(d => {
                      const key = `${s.id}_${d.dateKey}`; const val = schedule[key]; const lvv = leaveData[key]; const des = designatedOff[key];
                      const cs = cellStyle(val, lvv, des);
                      const apolloCode = (val && !lvv && !["休","例"].includes(val)) ? getApolloCode(val, store) : "";
                      return (
                        <td key={d.day} style={{padding:"2px", textAlign:"center", borderBottom:"1px solid rgba(148,163,184,0.08)"}}>
                          <button style={{width:"100%", height:"30px", borderRadius:"5px", background:cs.bg, color:cs.fg, fontSize:"11px", fontWeight:700, cursor:"pointer", border:cs.border}}
                            title={apolloCode ? `${cs.label} → Apollo: ${apolloCode}` : cs.label}
                            onClick={() => {
                              const cycle = s.type==="pt" ? ["A","B","C","D","休","例",""] : ["早","中","晚","休","例",""];
                              const idx = cycle.indexOf(val || ""); const next = cycle[(idx+1) % cycle.length] || undefined;
                              setSchedule(p => ({...p, [key]: next}));
                              if (lvv) setLeaveData(p => { const n = {...p}; delete n[key]; return n; });
                            }}
                            onContextMenu={e => {
                              e.preventDefault();
                              const choice = prompt(`${s.name} ${month+1}/${d.day}：\n${LEAVE_TYPES.map((l,i) => `${i+1}. ${l.label}`).join("\n")}\n0. 清除\n\n編號:`);
                              if (choice === "0") setLeaveData(p => { const n = {...p}; delete n[key]; return n; });
                              else {
                                const ci = parseInt(choice) - 1;
                                if (ci >= 0 && ci < LEAVE_TYPES.length) {
                                  setLeaveData(p => ({...p, [key]: LEAVE_TYPES[ci].key}));
                                  setSchedule(p => { const n = {...p}; delete n[key]; return n; });
                                }
                              }
                            }}>{cs.label || "·"}</button>
                        </td>
                      );
                    })}
                    <td style={{padding:"4px 8px", textAlign:"center", borderBottom:"1px solid rgba(148,163,184,0.08)", fontSize:"11px"}}>
                      <div style={{display:"flex", flexDirection:"column", gap:"1px"}}>
                        <span style={{color:"#818cf8"}}>班{w}</span>
                        <span style={{color:"#34d399"}}>休{r}</span>
                        {lv>0 && <span style={{color:"#f59e0b"}}>假{lv}</span>}
                      </div>
                    </td>
                  </tr>
                );
              })}
              <tr>
                <td style={{padding:"8px", textAlign:"left", position:"sticky", left:0, background:"rgba(30,41,59,0.98)", zIndex:1, fontWeight:700, color:"#c4b5fd", borderTop:"2px solid #7c3aed"}}>📊 在班</td>
                {monthDates.map(d => {
                  let cnt = 0;
                  staffList.forEach(s => {
                    const val = schedule[`${s.id}_${d.dateKey}`]; const lvv = leaveData[`${s.id}_${d.dateKey}`];
                    if (!lvv && val && !["休","例"].includes(val)) cnt++;
                  });
                  const need = d.isWeekend ? model?.weekend : model?.weekday;
                  const ok = cnt >= (need || 0);
                  return (
                    <td key={d.day} style={{padding:"4px 2px", textAlign:"center", borderTop:"2px solid #7c3aed", background:cnt>0&&!ok?"rgba(239,68,68,0.08)":cnt>0?"rgba(16,185,129,0.05)":"transparent"}}>
                      <div style={{fontWeight:800, fontSize:"14px", color:cnt===0?"#64748b":ok?"#34d399":"#ef4444"}}>{cnt}</div>
                      <div style={{fontSize:"9px", color:"#64748b"}}>需{need}</div>
                    </td>
                  );
                })}
                <td style={{borderTop:"2px solid #7c3aed"}}></td>
              </tr>
            </tbody>
          </table>
        </div>

        <div style={css.card}>
          <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"14px", flexWrap:"wrap", gap:"8px"}}>
            <h3 style={{margin:0, color:"#fbbf24", fontSize:"15px"}}>💰 本月人事成本</h3>
            <div style={{display:"flex", gap:"10px", alignItems:"center", fontSize:"12px"}}>
              <label style={{color:"#94a3b8"}}>正職月薪 NT$
                <input type="number" value={ftSalary} onChange={e => setFtSalary(Math.max(0, +e.target.value || 0))} style={{...css.input, width:"90px", padding:"4px 8px", display:"inline-block", marginLeft:"4px"}}/>
              </label>
              <label style={{color:"#94a3b8"}}>兼職時薪 NT$
                <input type="number" value={ptHourly} onChange={e => setPtHourly(Math.max(0, +e.target.value || 0))} style={{...css.input, width:"70px", padding:"4px 8px", display:"inline-block", marginLeft:"4px"}}/>
              </label>
            </div>
          </div>
          <div style={{display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))", gap:"12px", marginBottom:"16px"}}>
            <div style={{padding:"16px", borderRadius:"12px", background:"rgba(99,102,241,0.15)", border:"1px solid rgba(99,102,241,0.3)"}}>
              <div style={{fontSize:"11px", color:"#94a3b8"}}>正職 ({costBreakdown.ftCount} 人)</div>
              <div style={{fontSize:"22px", fontWeight:800, color:"#818cf8"}}>NT$ {costBreakdown.ftCost.toLocaleString()}</div>
            </div>
            <div style={{padding:"16px", borderRadius:"12px", background:"rgba(245,158,11,0.15)", border:"1px solid rgba(245,158,11,0.3)"}}>
              <div style={{fontSize:"11px", color:"#94a3b8"}}>兼職 ({costBreakdown.ptCount} 人)</div>
              <div style={{fontSize:"22px", fontWeight:800, color:"#fbbf24"}}>NT$ {costBreakdown.ptCost.toLocaleString()}</div>
              <div style={{fontSize:"10px", color:"#64748b"}}>{costBreakdown.ptTotalHours}hr × NT${ptHourly}</div>
            </div>
            <div style={{padding:"16px", borderRadius:"12px", background:"rgba(16,185,129,0.2)", border:"1px solid rgba(16,185,129,0.4)"}}>
              <div style={{fontSize:"11px", color:"#94a3b8"}}>總計</div>
              <div style={{fontSize:"26px", fontWeight:900, color:"#34d399"}}>NT$ {costBreakdown.totalCost.toLocaleString()}</div>
            </div>
          </div>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%", fontSize:"12px", borderCollapse:"collapse"}}>
              <thead><tr style={{background:"rgba(124,58,237,0.15)"}}>
                <th style={{padding:"8px 10px", textAlign:"left", color:"#c4b5fd"}}>工號</th>
                <th style={{padding:"8px 10px", textAlign:"left", color:"#c4b5fd"}}>姓名</th>
                <th style={{padding:"8px 10px", textAlign:"center", color:"#c4b5fd"}}>類型</th>
                <th style={{padding:"8px 10px", textAlign:"left", color:"#c4b5fd"}}>計算明細</th>
                <th style={{padding:"8px 10px", textAlign:"right", color:"#c4b5fd"}}>金額</th>
              </tr></thead>
              <tbody>
                {costBreakdown.rows.map(r => (
                  <tr key={r.id} style={{borderBottom:"1px solid rgba(148,163,184,0.1)"}}>
                    <td style={{padding:"8px 10px", fontSize:"11px", color:"#fbbf24", fontFamily:"monospace"}}>{r.empCode || "—"}</td>
                    <td style={{padding:"8px 10px", fontWeight:600}}>{r.name}</td>
                    <td style={{padding:"8px 10px", textAlign:"center"}}>
                      <span style={{padding:"2px 8px", borderRadius:"4px", fontSize:"10px", background:r.type==="ft"?"rgba(99,102,241,0.25)":"rgba(245,158,11,0.25)", color:r.type==="ft"?"#a5b4fc":"#fbbf24"}}>{r.type==="ft"?"正職":"兼職"}</span>
                    </td>
                    <td style={{padding:"8px 10px", color:"#94a3b8", fontSize:"11px", fontFamily:"monospace"}}>{r.detail}</td>
                    <td style={{padding:"8px 10px", textAlign:"right", fontWeight:700, color:r.type==="ft"?"#818cf8":"#fbbf24"}}>{r.cost.toLocaleString()}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}
