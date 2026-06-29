import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

// ─── CONSTANTS ───────────────────────────────────────────────────────────────

const PUDU_ITEMS = [
  { id: "CC1", name: "Pudu CC1", category: "Robot" },
  { id: "CC1-PRO", name: "Pudu CC1 Pro", category: "Robot" },
  { id: "BELLABOT", name: "BellaBot", category: "Robot" },
  { id: "BELLABOT-PRO", name: "BellaBot Pro", category: "Robot" },
  { id: "KETTYBOT", name: "KettyBot", category: "Robot" },
  { id: "KETTYBOT-PRO", name: "KettyBot Pro", category: "Robot" },
  { id: "HOLABOT", name: "HolaBot", category: "Robot" },
  { id: "HOLABOT-PRO", name: "HolaBot Pro", category: "Robot" },
  { id: "SWIFTBOT", name: "SwiftBot", category: "Robot" },
  { id: "SWIFTBOT-PRO", name: "SwiftBot Pro", category: "Robot" },
  { id: "MATRADEE", name: "Matradee", category: "Robot" },
  { id: "MATRADEE-L", name: "Matradee L", category: "Robot" },
  { id: "MATRADEE-X", name: "Matradee X", category: "Robot" },
  { id: "PADBOT-T1", name: "PadBot T1", category: "Robot" },
  { id: "PADBOT-U1", name: "PadBot U1", category: "Robot" },
  { id: "PADBOT-P3", name: "PadBot P3", category: "Robot" },
  { id: "AIMBOT", name: "AimBot", category: "Robot" },
  { id: "PUDUCTOR2", name: "PuDuctor 2", category: "Robot" },
  { id: "RIDGEBACK", name: "RidgeBack", category: "Robot" },
  { id: "RIDGEBACK-S", name: "RidgeBack S", category: "Robot" },
  { id: "ACC-TRAY-S", name: "Tray Set (Small)", category: "Accessory" },
  { id: "ACC-TRAY-L", name: "Tray Set (Large)", category: "Accessory" },
  { id: "ACC-HANDLE", name: "Handle Bar", category: "Accessory" },
  { id: "ACC-BUMPER", name: "Bumper Guard", category: "Accessory" },
  { id: "ACC-CHARGER-CC1", name: "Charger CC1", category: "Accessory" },
  { id: "ACC-CHARGER-BELLA", name: "Charger BellaBot", category: "Accessory" },
  { id: "ACC-CHARGER-KETTY", name: "Charger KettyBot", category: "Accessory" },
  { id: "ACC-CHARGER-HOLA", name: "Charger HolaBot", category: "Accessory" },
  { id: "ACC-CHARGER-SWIFT", name: "Charger SwiftBot", category: "Accessory" },
  { id: "ACC-CHARGER-MAT", name: "Charger Matradee", category: "Accessory" },
  { id: "ACC-DISINFECT", name: "Disinfection Module", category: "Accessory" },
  { id: "ACC-AD-SCREEN", name: "Advertisement Screen", category: "Accessory" },
  { id: "ACC-THERMAL", name: "Thermal Camera Module", category: "Accessory" },
  { id: "ACC-SPEAKER", name: "External Speaker", category: "Accessory" },
  { id: "ACC-TABLET-10", name: "Tablet 10in", category: "Accessory" },
  { id: "ACC-TABLET-13", name: "Tablet 13in", category: "Accessory" },
  { id: "ACC-WIFI-EXT", name: "WiFi Extender", category: "Accessory" },
  { id: "ACC-LIDAR", name: "LiDAR Sensor", category: "Accessory" },
  { id: "ACC-CAMERA-360", name: "360 Camera Module", category: "Accessory" },
  { id: "ACC-DOCK", name: "Charging Dock", category: "Accessory" },
  { id: "ACC-UV-LAMP", name: "UV Lamp Module", category: "Accessory" },
  { id: "ACC-CASE-S", name: "Carry Case (Small)", category: "Accessory" },
  { id: "ACC-CASE-L", name: "Carry Case (Large)", category: "Accessory" },
  { id: "ACC-MOUNT-WALL", name: "Wall Mount Bracket", category: "Accessory" },
  { id: "ACC-CABLE-SET", name: "Cable Set", category: "Accessory" },
  { id: "ACC-REMOTE", name: "Remote Controller", category: "Accessory" },
  { id: "ACC-BATTERY-EXT", name: "Extended Battery Pack", category: "Accessory" },
  { id: "ACC-FILTER", name: "Air Filter Module", category: "Accessory" },
  { id: "ACC-LABEL-SET", name: "Label and Signage Set", category: "Accessory" },
  { id: "ACC-TROLLEY", name: "Transport Trolley", category: "Accessory" },
  // Additional models from inventory tracking report
  { id: "T300-BLACK", name: "T300 (Black)", category: "Robot" },
  { id: "T300-LIFTING", name: "T300 - Lifting", category: "Robot" },
  { id: "T300-TOWING", name: "T300 - Towing", category: "Robot" },
  { id: "T300-CONVEYOR", name: "T300 - Conveyor", category: "Robot" },
  { id: "T300", name: "T300 Standard", category: "Robot" },
  { id: "T600-UNDERRIDE", name: "T600 Underride", category: "Robot" },
  { id: "T600-WTIDL2", name: "T600 WTIDL2", category: "Robot" },
  { id: "T600-STANDARD", name: "T600 Standard", category: "Robot" },
  { id: "T150", name: "T150", category: "Robot" },
  { id: "MT1", name: "MT1", category: "Robot" },
  { id: "MT1-MAX", name: "MT1 MAX", category: "Robot" },
  { id: "MT1-VAC", name: "MT1 VAC", category: "Robot" },
  { id: "MT1-VAC-CHARGING", name: "MT1 VAC Charging Station", category: "Accessory" },
  { id: "PUDUBOT2", name: "Pudubot 2", category: "Robot" },
  { id: "FLASHBOT", name: "Flashbot", category: "Robot" },
  { id: "WATER-TANK", name: "Mobile Water Station", category: "Accessory" },
];

// SAP Stock Categories — mirrors your SAP setup
const SAP_STOCK_CATEGORIES = [
  { id: "STOCK", label: "Stock", desc: "Physical warehouse stock — unrestricted", color: "#16a34a", bg: "#dcfce7" },
  { id: "GIT", label: "GIT", desc: "Goods in Transit — ordered, not yet received", color: "#d97706", bg: "#fef3c7" },
  { id: "DEMO", label: "Demo", desc: "Demo units — allocated for sales/marketing use", color: "#7c3aed", bg: "#ede9fe" },
];

const MOVEMENT_TYPES = {
  GR: [
    { value: "GR_SUPPLIER", label: "Goods Receipt – Supplier Delivery", color: "#16a34a", sapCat: "STOCK" },
    { value: "GR_GIT_RECEIVED", label: "GIT Received – Now in Warehouse", color: "#15803d", sapCat: "STOCK" },
    { value: "GR_DEMO_RETURN_SALES", label: "Return – Demo by Sales Personnel", color: "#166534", sapCat: "DEMO" },
    { value: "GR_DEMO_RETURN_ENG", label: "Return – POC by Engineer", color: "#166534", sapCat: "DEMO" },
    { value: "GR_LOAN_RETURN_CUST", label: "Return – Loan from Customer", color: "#14532d", sapCat: "DEMO" },
    { value: "GR_EXHIBITION_RETURN", label: "Return – Exhibition / Marketing", color: "#4d7c0f", sapCat: "DEMO" },
  ],
  GI: [
    { value: "GI_CUSTOMER_DELIVERY", label: "Goods Issue – Customer Delivery (Sale)", color: "#dc2626", sapCat: "STOCK" },
    { value: "GI_LOAN_CUSTOMER", label: "Loan Out – Customer Trial", color: "#b91c1c", sapCat: "DEMO" },
    { value: "GI_DEMO_SALES", label: "Demo Out – Sales Personnel", color: "#991b1b", sapCat: "DEMO" },
    { value: "GI_POC_ENGINEER", label: "POC Out – Engineer", color: "#7f1d1d", sapCat: "DEMO" },
    { value: "GI_EXHIBITION", label: "Exhibition Out – Marketing", color: "#c2410c", sapCat: "DEMO" },
    { value: "GI_TRANSFER_GIT", label: "Transfer to GIT – Pending Shipment", color: "#b45309", sapCat: "GIT" },
  ],
};

const ALL_MOVEMENT_TYPES = [...MOVEMENT_TYPES.GR, ...MOVEMENT_TYPES.GI];
const REGIONS = ["Kuala Lumpur", "Selangor", "Penang", "Johor Bahru", "Sabah", "Sarawak", "Singapore", "Thailand", "Indonesia", "Vietnam", "Other"];
const ROLES = ["Sales Personnel", "Engineer", "Marketing", "Logistics Admin", "Management"];
const ENTITIES = ["Entity A", "Entity B"];


// ─── STORAGE KEYS ─────────────────────────────────────────────────────────────
const SK = { MOV: "pudu_v2_movements", PERS: "pudu_v2_personnel", SAP: "pudu_v2_sap_snapshots", GIT: "pudu_v2_git", FLEET: "pudu_v2_fleet", LOCK: "pudu_v2_locked_months", SETUP: "pudu_v2_setup", USERS: "pudu_v2_users", PENDING: "pudu_v2_pending", ITEMS: "pudu_v2_item_master", CUSTOMERS: "pudu_v2_customer_master" };

// Default setup — all configurable by admin
const DEFAULT_SETUP = {
  appName: "Robotics Fleet & Demo Tracker",
  companyName: "Your Company Name",
  entities: ["Entity A", "Entity B"],
  regions: ["Head Office", "Branch Office", "Regional Office", "Warehouse", "Other"],
  grTypes: [
    { value: "GR_SUPPLIER",           label: "Goods Receipt – Supplier Delivery",    sapCat: "STOCK", active: true },
    { value: "GR_GIT_RECEIVED",       label: "GIT Received – Now in Warehouse",       sapCat: "STOCK", active: true },
    { value: "GR_DEMO_RETURN_SALES",  label: "Return – Demo by Sales Personnel",      sapCat: "DEMO",  active: true },
    { value: "GR_DEMO_RETURN_ENG",    label: "Return – POC by Engineer",              sapCat: "DEMO",  active: true },
    { value: "GR_LOAN_RETURN_CUST",   label: "Return – Loan from Customer",           sapCat: "DEMO",  active: true },
    { value: "GR_EXHIBITION_RETURN",  label: "Return – Exhibition / Marketing",       sapCat: "DEMO",  active: true },
  ],
  giTypes: [
    { value: "GI_CUSTOMER_DELIVERY",  label: "Goods Issue – Customer Delivery (Sale)", sapCat: "STOCK", active: true },
    { value: "GI_LOAN_CUSTOMER",      label: "Loan Out – Customer Trial",              sapCat: "DEMO",  active: true },
    { value: "GI_DEMO_SALES",         label: "Demo Out – Sales Personnel",             sapCat: "DEMO",  active: true },
    { value: "GI_POC_ENGINEER",       label: "POC Out – Engineer",                     sapCat: "DEMO",  active: true },
    { value: "GI_EXHIBITION",         label: "Exhibition Out – Marketing",             sapCat: "DEMO",  active: true },
    { value: "GI_TRANSFER_GIT",       label: "Transfer to GIT – Pending Shipment",    sapCat: "GIT",   active: true },
  ],
};

// ─── ITEM MASTER (SAP Item Code + Description) ─────────────────────────────
const DEFAULT_ITEMS = [
  { id:"ITM001", sapCode:"CC1",          name:"Pudu CC1",                    category:"Robot",      active:true },
  { id:"ITM002", sapCode:"CC1-PRO",      name:"Pudu CC1 Pro",                category:"Robot",      active:true },
  { id:"ITM003", sapCode:"BELLABOT",     name:"BellaBot",                    category:"Robot",      active:true },
  { id:"ITM004", sapCode:"BELLABOT-PRO", name:"BellaBot Pro",                category:"Robot",      active:true },
  { id:"ITM005", sapCode:"T300",         name:"PUDU T300",                   category:"Robot",      active:true },
  { id:"ITM006", sapCode:"T300-BLACK",   name:"PUDU T300 (BLACK) WTIDL1",    category:"Robot",      active:true },
  { id:"ITM007", sapCode:"T300-LIFTING", name:"PUDU T300 - Lifting",         category:"Robot",      active:true },
  { id:"ITM008", sapCode:"T300-TOWING",  name:"PUDU T300 - Towing",          category:"Robot",      active:true },
  { id:"ITM009", sapCode:"T300-CONVEYOR",name:"PUDU T300 - Conveyor",        category:"Robot",      active:true },
  { id:"ITM010", sapCode:"T600",         name:"PUDU T600",                   category:"Robot",      active:true },
  { id:"ITM011", sapCode:"T600-UNDER",   name:"PUDU T600 UNDERRIDE WTIDU2",  category:"Robot",      active:true },
  { id:"ITM012", sapCode:"T600-STD",     name:"PUDU T600 Standard",          category:"Robot",      active:true },
  { id:"ITM013", sapCode:"T600-WTIDL2",  name:"PUDU T600 WTIDL2",            category:"Robot",      active:true },
  { id:"ITM014", sapCode:"MT1",          name:"PUDU MT1",                    category:"Robot",      active:true },
  { id:"ITM015", sapCode:"MT1-MAX",      name:"PUDU MT1 MAX",                category:"Robot",      active:true },
  { id:"ITM016", sapCode:"MT1-VAC",      name:"PUDU MT1 VAC",                category:"Robot",      active:true },
  { id:"ITM017", sapCode:"PUDUBOT2",     name:"Pudubot 2",                   category:"Robot",      active:true },
  { id:"ITM018", sapCode:"FLASHBOT",     name:"Flashbot",                    category:"Robot",      active:true },
  { id:"ITM019", sapCode:"T150",         name:"PUDU T150",                   category:"Robot",      active:true },
  { id:"ITM020", sapCode:"HOLABOT",      name:"HolaBot",                     category:"Robot",      active:true },
  { id:"ITM021", sapCode:"PUDUCTOR2",    name:"PuDuctor 2",                  category:"Robot",      active:true },
  { id:"ITM022", sapCode:"MT1-VAC-CHG",  name:"MT1 VAC Charging Station",    category:"Accessory",  active:true },
  { id:"ITM023", sapCode:"CHG-STD",      name:"Standard Charging Station",   category:"Accessory",  active:true },
  { id:"ITM024", sapCode:"WATER-TANK",   name:"Mobile Water Station",        category:"Accessory",  active:true },
  { id:"ITM025", sapCode:"WATER-STN",    name:"Mobile Water Tank",           category:"Accessory",  active:true },
  { id:"ITM026", sapCode:"DOCK-STN",     name:"CC1 Docking Station",         category:"Accessory",  active:true },
  { id:"ITM027", sapCode:"WST-CC1PRO",   name:"Working Station for CC1 Pro", category:"Accessory",  active:true },
  { id:"ITM028", sapCode:"BATT-CHG",     name:"Battery Charger",             category:"Accessory",  active:true },
  { id:"ITM029", sapCode:"TROLLEY",      name:"Transport Trolley",           category:"Accessory",  active:true },
];

// ─── CUSTOMER MASTER ─────────────────────────────────────────────────────────
const DEFAULT_CUSTOMERS = [
  { id:"CUS001", name:"Agilent Technologies",      country:"Malaysia", active:true },
  { id:"CUS002", name:"PT Flex Batam",             country:"Indonesia", active:true },
  { id:"CUS003", name:"SFP Automation",            country:"Malaysia", active:true },
  { id:"CUS004", name:"IOI Properties",            country:"Malaysia", active:true },
  { id:"CUS005", name:"Sunway Medical Centre",     country:"Malaysia", active:true },
  { id:"CUS006", name:"SHRC",                      country:"Malaysia", active:true },
  { id:"CUS007", name:"Automex",                   country:"Malaysia", active:true },
];

// ─── ROLE PERMISSIONS ─────────────────────────────────────────────────────────
// Role permission templates
const ROLE_PERMS = {
  "Super Admin": { "Stock Position":"edit","Fleet Calendar":"edit","Dashboard":"edit","Log Movement":"edit","Movement Log":"edit","GIT Register":"edit","Pending Approvals":"edit","Admin":"edit","Setup":"edit","Import SAP":"edit","Backup & Restore":"edit","Export Excel":"edit","User Guide":"view","Version History":"edit" },
  "Logistics Admin": { "Stock Position":"view","Fleet Calendar":"view","Dashboard":"view","Log Movement":"edit","Movement Log":"edit","GIT Register":"edit","Pending Approvals":"edit","Admin":"view","Setup":"none","Import SAP":"edit","Backup & Restore":"none","Export Excel":"edit","User Guide":"view","Version History":"none" },
};

const DEFAULT_USERS = [
  { id: 1, name: "Super Admin", pin: "9999", role: "Super Admin", entity: "", region: "", active: true,
    permissions: ROLE_PERMS["Super Admin"] },
  { id: 2, name: "Logistics Admin", pin: "1234", role: "Logistics Admin", entity: "", region: "", active: true,
    permissions: ROLE_PERMS["Logistics Admin"] },
];

const TAB_LIST = ["Stock Position","Fleet Calendar","Dashboard","Log Movement","Movement Log","GIT Register","Pending Approvals","Admin","Setup","Import SAP","Backup & Restore","Export Excel","User Guide","Version History"];

// ─── SAMPLE DATA ─────────────────────────────────────────────────────────────
const SAMPLE_PERSONNEL = [
];

const SAMPLE_SAP = [
];

const SAMPLE_GIT = [
];

const SAMPLE_MOVEMENTS = [
];

// ─── SAMPLE FLEET UNITS ──────────────────────────────────────────────────────
const SAMPLE_FLEET = [
];


// ─── STORAGE (localStorage — works on Vercel, no Claude dependency) ──────────
async function load(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (raw) return JSON.parse(raw);
    localStorage.setItem(key, JSON.stringify(fallback));
    return fallback;
  } catch { return fallback; }
}
async function save(key, data) {
  try { localStorage.setItem(key, JSON.stringify(data)); } catch (e) { console.error(e); }
}

// ─── UTILITIES ────────────────────────────────────────────────────────────────
// confirm helper — works in sandbox and browser
const appConfirm = (msg) => { try { return window.confirm(msg); } catch(e) { return true; } };
const today = () => new Date().toISOString().split("T")[0];
const fmt = (d) => d ? new Date(d).toLocaleDateString("en-MY", { day: "2-digit", month: "short", year: "numeric" }) : "—";
const isOverdue = (m) => m.status !== "Returned" && m.expectedReturn && new Date(m.expectedReturn) < new Date();
const getDir = (t) => t?.startsWith("GR") ? "GR" : "GI";
const getMovLabel = (t) => ALL_MOVEMENT_TYPES.find(m => m.value === t)?.label || t;
const getItemName = (id) => PUDU_ITEMS.find(i => i.id === id)?.name || id;
const getItemCat = (id) => PUDU_ITEMS.find(i => i.id === id)?.category || "";
const getSapCatInfo = (id) => SAP_STOCK_CATEGORIES.find(c => c.id === id) || {};

// ─── PUBLIC FORM (for Sales, Marketing, AE — no login needed) ────────────────
function PublicForm({ movements, setMovements, personnel, setup, pendingItems, setPendingItems, itemMaster, customers }) {
  const blankF = { type:"", sapCat:"", itemId:"", itemDesc:"", qty:1, serialNo:"",
    custodianName:"", entity:"", location:"", region:"", remark:"",
    date:today(), expectedReturn:"", customerName:"", eventName:"", notes:"" };
  const [form, setForm] = useState(blankF);
  const [submitted, setSubmitted] = useState(false);
  const [error, setError] = useState("");
  const [dir, setDir] = useState("");

  const cfg = setup || DEFAULT_SETUP;
  const giTypes  = getActiveGI(cfg);
  const grTypes  = getActiveGR(cfg);
  const entities = getEntities(cfg);
  const regions  = getRegions(cfg);
  const activeTypes = dir === "GR" ? grTypes : dir === "GI" ? giTypes : [];

  function setF(k,v){ setForm(f=>({...f,[k]:v})); }
  function onTypeChange(val){
    const all=[...grTypes,...giTypes];
    const def=all.find(t=>t.value===val);
    setForm(f=>({...f,type:val,sapCat:def?.sapCat||""}));
  }

  const isGI = dir==="GI";
  const needsReturn   = isGI; // Always show for GI
  const needsCustomer = ["GI_CUSTOMER_DELIVERY","GI_LOAN_CUSTOMER"].includes(form.type);
  const needsEvent    = form.type==="GI_EXHIBITION";

  async function handleSubmit(){
    const resolvedName = form.custodianName==="__other__" ? (form.custodianNameOther||"").trim() : form.custodianName;
    if(!form.type||!form.itemId||!resolvedName||!form.date){
      setError("Please fill in: Movement Type, Item, Your Name, and Date."); return;
    }
    if(!form.serialNo){ setError("Serial Number is required — Fleet Calendar cannot track the unit without it."); return; }
    if(!form.entity){ setError("Please select an Entity (company / division)."); return; }
    if(isGI && !form.expectedReturn){ setError("Expected Return Date is required for all outgoing movements."); return; }
    setError("");
    const newPending = {
      ...form, pendingId:"PF_"+Date.now(),
      custodian:form.custodianName==="__other__" ? (form.custodianNameOther||form.custodianName) : form.custodianName,
      qty:Number(form.qty)||1,
      isReversal:false, reversalOf:null,
      approvalStatus:"pending",
      submittedAt:new Date().toISOString(),
      source:"public_form",
    };
    const updated=[newPending,...(pendingItems||[])];
    setPendingItems(updated);
    await save(SK.PENDING,updated);
    setSubmitted(true);
  }

  const REMARKS=["KL Demo","PEN Demo","KL Stock","PEN Stock","KL DEMO","PEN DEMO","POC","Exhibition","Loan","In Transit","Collected Back","Reserved for Next Demo"];

  const field=(label,children,required=false)=>(
    <div style={{marginBottom:14}}>
      <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>
        {label}{required&&<span style={{color:"#dc2626"}}> *</span>}
      </label>
      {children}
    </div>
  );
  const inp=(label,key,type="text",ph="",required=false)=>field(label,
    <input type={type} value={form[key]} onChange={e=>setF(key,e.target.value)} placeholder={ph}
      style={{width:"100%",border:"1px solid #d1d5db",borderRadius:6,padding:"9px 10px",fontSize:14}} />, required);

  if(submitted) return(
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:16,padding:36,maxWidth:480,width:"100%",textAlign:"center"}}>
        <div style={{fontSize:52,marginBottom:12}}>✅</div>
        <div style={{fontSize:22,fontWeight:700,color:"#0D1B3E",marginBottom:8}}>Submitted!</div>
        <div style={{fontSize:14,color:"#6b7280",marginBottom:20}}>Pending approval by Logistics Admin. It will appear in the system once approved.</div>
        <div style={{background:"#f9fafb",borderRadius:8,padding:14,marginBottom:20,textAlign:"left",fontSize:13,lineHeight:1.8}}>
          <div><strong>Direction:</strong> {dir==="GI"?"📤 Going OUT":"📥 Coming IN"}</div>
          <div><strong>Item:</strong> {getItemName(form.itemId)}</div>
          {form.itemDesc&&<div><strong>Description:</strong> {form.itemDesc}</div>}
          <div><strong>Serial No:</strong> {form.serialNo}</div>
          <div><strong>Entity:</strong> {form.entity}</div>
          <div><strong>Date:</strong> {fmt(form.date)}</div>
          {form.expectedReturn&&<div><strong>Expected Return:</strong> {fmt(form.expectedReturn)}</div>}
          {form.remark&&<div><strong>Remark:</strong> {form.remark}</div>}
          {form.location&&<div><strong>Location:</strong> {form.location}</div>}
        </div>
        <button onClick={()=>{setForm(blankF);setDir("");setSubmitted(false);}}
          style={{background:"#0D1B3E",color:"#fff",border:"none",borderRadius:8,padding:"12px 32px",fontWeight:700,fontSize:15,cursor:"pointer"}}>
          Submit Another
        </button>
      </div>
    </div>
  );

  return(
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)",padding:"20px 16px",fontFamily:"'Inter','Segoe UI',sans-serif"}}>
      <div style={{maxWidth:560,margin:"0 auto"}}>
        {/* Header */}
        <div style={{textAlign:"center",marginBottom:20}}>
          <div style={{fontSize:36,marginBottom:6}}>🤖</div>
          <div style={{color:"#fff",fontSize:20,fontWeight:700}}>{cfg.appName||"Fleet Tracker"}</div>
          <div style={{color:"#C8EEF5",fontSize:13,marginTop:4}}>Movement Submission Form</div>
          <div style={{color:"#9ca3af",fontSize:11,marginTop:2}}>Fill in all fields — Logistics Admin will approve before it goes into the system</div>
        </div>

        <div style={{background:"#fff",borderRadius:14,padding:24}}>
          {error&&<div style={{background:"#fee2e2",border:"1px solid #fca5a5",borderRadius:8,padding:"10px 14px",color:"#991b1b",fontSize:13,marginBottom:16,fontWeight:500}}>{error}</div>}

          {/* STEP 1: Direction */}
          <div style={{marginBottom:18}}>
            <div style={{fontSize:11,fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:1,marginBottom:8}}>Step 1 — Movement Direction</div>
            <div style={{display:"flex",gap:10}}>
              {[["GR","📥 Unit Coming IN","Return / Receipt","#16a34a","#dcfce7"],
                ["GI","📤 Unit Going OUT","Demo / Loan / POC / Exhibition","#dc2626","#fee2e2"]].map(([d,title,sub,col,bg])=>(
                <button key={d} onClick={()=>{setDir(d);setForm(f=>({...f,type:"",sapCat:""}));}}
                  style={{flex:1,padding:"14px 8px",borderRadius:10,border:`2px solid ${dir===d?col:"#e5e7eb"}`,
                    background:dir===d?bg:"#f9fafb",cursor:"pointer",textAlign:"center"}}>
                  <div style={{fontWeight:700,fontSize:14,color:dir===d?col:"#374151"}}>{title}</div>
                  <div style={{fontSize:11,color:dir===d?col:"#9ca3af",marginTop:3}}>{sub}</div>
                </button>
              ))}
            </div>
          </div>

          {dir&&(<>
            {/* STEP 2: Movement Details */}
            <div style={{borderTop:"1px solid #f3f4f6",paddingTop:16,marginBottom:4}}>
              <div style={{fontSize:11,fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:1,marginBottom:14}}>Step 2 — Movement Details</div>
            </div>

            {field("Movement Type",
              <select value={form.type} onChange={e=>onTypeChange(e.target.value)}
                style={{width:"100%",border:"1px solid #d1d5db",borderRadius:6,padding:"9px 10px",fontSize:14}}>
                <option value="">— Select type —</option>
                {activeTypes.map(t=><option key={t.value} value={t.value}>{t.label}</option>)}
              </select>, true)}

            {field("Item / SAP Code + Description",
              <select value={form.itemId} onChange={e=>{
                  const sel=(itemMaster||DEFAULT_ITEMS).find(i=>i.sapCode===e.target.value);
                  setForm(f=>({...f, itemId:e.target.value, itemDesc:sel?sel.name:f.itemDesc}));
                }}
                style={{width:"100%",border:"1px solid #d1d5db",borderRadius:6,padding:"9px 10px",fontSize:14}}>
                <option value="">— Select item —</option>
                {["Robot","Accessory","Spare Part","Other"].map(cat=>{
                  const catItems=(itemMaster||DEFAULT_ITEMS).filter(i=>i.category===cat&&i.active!==false);
                  return catItems.length>0?(<optgroup key={cat} label={cat}>
                    {catItems.map(i=><option key={i.id} value={i.sapCode}>[{i.sapCode}] {i.name}</option>)}
                  </optgroup>):null;
                })}
              </select>, true)}

            {field("Item Description",
              <input value={form.itemDesc||""} onChange={e=>setF("itemDesc",e.target.value)}
                placeholder="Auto-filled — or type manually"
                style={{width:"100%",border:"1px solid #d1d5db",borderRadius:6,padding:"9px 10px",fontSize:14}} />)}

            {inp("Serial Number","serialNo","text","e.g. 8260B5318060021",true)}

            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              {inp("Quantity","qty","number","1")}
              {inp("Date","date","date","",true)}
            </div>

            {needsReturn&&(
              <div style={{background:"#fef3c7",border:"1px solid #fbbf24",borderRadius:8,padding:"10px 14px",marginBottom:14}}>
                {inp("Expected Return Date ⚠️ Required for all OUT movements","expectedReturn","date","",true)}
              </div>
            )}

            {/* STEP 3: Who & Where */}
            <div style={{borderTop:"1px solid #f3f4f6",paddingTop:16,marginBottom:4}}>
              <div style={{fontSize:11,fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:1,marginBottom:14}}>Step 3 — Who & Where</div>
            </div>

            {field("Your Name",
              <select value={form.custodianName} onChange={e=>{
                  const p = (personnel||[]).find(x=>x.name===e.target.value);
                  setForm(f=>({...f,
                    custodianName:e.target.value,
                    entity: p?.entity||f.entity,
                    region: p?.region||f.region,
                  }));
                }}
                style={{width:"100%",border:"1px solid #d1d5db",borderRadius:6,padding:"9px 10px",fontSize:14}}>
                <option value="">— Select your name —</option>
                {(personnel||[]).filter(p=>p.active!==false).map(p=>(
                  <option key={p.id} value={p.name}>{p.name} ({p.role||""})</option>
                ))}
                <option value="__other__">Not listed — type below</option>
              </select>, true)}
            {form.custodianName==="__other__"&&inp("Your Full Name (specify)","custodianNameOther","text","Type your full name here",true)}

            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              {field("Entity",
                <select value={form.entity} onChange={e=>setF("entity",e.target.value)}
                  style={{width:"100%",border:"1px solid #d1d5db",borderRadius:6,padding:"9px 10px",fontSize:14}}>
                  <option value="">— Select —</option>
                  {entities.map(e=><option key={e} value={e}>{e}</option>)}
                </select>, true)}
              {field("Region",
                <select value={form.region} onChange={e=>setF("region",e.target.value)}
                  style={{width:"100%",border:"1px solid #d1d5db",borderRadius:6,padding:"9px 10px",fontSize:14}}>
                  <option value="">— Select —</option>
                  {regions.map(r=><option key={r} value={r}>{r}</option>)}
                </select>)}
            </div>

            {inp("Current Location / Site","location","text","e.g. KL Office, TMS Office, PT Flex Batam")}

            {field("Remark / Status",
              <select value={form.remark} onChange={e=>setF("remark",e.target.value)}
                style={{width:"100%",border:"1px solid #d1d5db",borderRadius:6,padding:"9px 10px",fontSize:14}}>
                <option value="">— Select remark —</option>
                {REMARKS.map(r=><option key={r} value={r}>{r}</option>)}
              </select>)}

            {needsCustomer&&field("Customer Name",
              <select value={form.customerName} onChange={e=>setF("customerName",e.target.value)}
                style={{width:"100%",border:"1px solid #d1d5db",borderRadius:6,padding:"9px 10px",fontSize:14}}>
                <option value="">— Select customer —</option>
                {(customers||DEFAULT_CUSTOMERS).filter(c=>c.active!==false).map(c=>(
                  <option key={c.id} value={c.name}>{c.name}{c.country?" ("+c.country+")":""}</option>
                ))}
                <option value="__other__">Other (not in list)</option>
              </select>)}
            {needsCustomer&&form.customerName==="__other__"&&inp("Customer Name (specify)","customerNameOther","text","Type customer name here")}
            {needsEvent&&inp("Event / Exhibition Name","eventName","text","e.g. Automex 2026")}

            {field("Notes",
              <textarea value={form.notes} onChange={e=>setF("notes",e.target.value)} rows={2}
                placeholder="Any additional remarks..."
                style={{width:"100%",border:"1px solid #d1d5db",borderRadius:6,padding:"9px 10px",fontSize:14,resize:"vertical"}} />)}

            <button onClick={handleSubmit}
              style={{width:"100%",background:"linear-gradient(135deg,#0D1B3E,#4AACCC)",color:"#fff",border:"none",
                borderRadius:10,padding:"15px",fontWeight:700,fontSize:16,cursor:"pointer",marginTop:4}}>
              Submit for Approval →
            </button>
          </>)}
        </div>
        <div style={{textAlign:"center",marginTop:14,color:"#C8EEF5",fontSize:11}}>
          Contact your Logistics Admin for support
        </div>
      </div>
    </div>
  );
}
function getMovLabel2(typeValue, setup) {
  const all = [...getActiveGR(setup), ...getActiveGI(setup)];
  return all.find(t => t.value === typeValue)?.label || typeValue;
}


// ─── PENDING APPROVALS ────────────────────────────────────────────────────────
function PendingApprovals({ pendingItems, setPendingItems, movements, setMovements }) {
  const [filter, setFilter] = useState("pending");
  const [rejectReason, setRejectReason] = useState("");
  const [rejectingId, setRejectingId] = useState(null);

  // Reload from storage on mount to pick up submissions from other devices
  useEffect(() => {
    load(SK.PENDING, []).then(fresh => {
      if (fresh && fresh.length !== pendingItems.length) {
        setPendingItems(fresh);
      }
    });
  }, []);

  async function refreshPending() {
    const fresh = await load(SK.PENDING, []);
    setPendingItems(fresh);
  }

  async function approve(item) {
    const approved = {
      ...item,
      id: Date.now(),
      status: "Active",
      approvedAt: new Date().toISOString(),
      approvalStatus: "approved",
      pendingId: undefined,
    };
    // Remove pendingId from the movement record
    delete approved.pendingId;
    delete approved.approvalStatus;
    const updatedMov = [approved, ...movements];
    setMovements(updatedMov);
    await save(SK.MOV, updatedMov);
    const updatedPending = pendingItems.map(p =>
      p.pendingId === item.pendingId
        ? { ...p, approvalStatus: "approved", approvedAt: new Date().toISOString() }
        : p
    );
    setPendingItems(updatedPending);
    await save(SK.PENDING, updatedPending);
    setRejectingId(null);
  }

  async function reject(item, reason) {
    const updatedPending = pendingItems.map(p =>
      p.pendingId === item.pendingId
        ? { ...p, approvalStatus: "rejected", rejectedAt: new Date().toISOString(), rejectReason: reason || "Rejected" }
        : p
    );
    setPendingItems(updatedPending);
    await save(SK.PENDING, updatedPending);
    setRejectingId(null);
    setRejectReason("");
  }

  const filtered = pendingItems.filter(p =>
    filter === "all" ? true : (p.approvalStatus || "pending") === filter
  );

  const pendingCount = pendingItems.filter(p => (p.approvalStatus || "pending") === "pending").length;

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      {/* Header */}
      <div style={{ background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)", borderRadius:10, padding:"16px 20px", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <div>
          <div style={{ color:"#fff", fontWeight:700, fontSize:16 }}>Pending Approvals</div>
          <div style={{ color:"#C8EEF5", fontSize:12, marginTop:2 }}>Review and approve movements submitted by field staff via the public form</div>
        </div>
        <div style={{ display:"flex", gap:10, alignItems:"center" }}>
          {pendingCount > 0 && (
            <div style={{ background:"#dc2626", color:"#fff", borderRadius:20, padding:"6px 16px", fontWeight:700, fontSize:14 }}>
              {pendingCount} pending
            </div>
          )}
          <button onClick={refreshPending}
            style={{ background:"rgba(255,255,255,0.15)", color:"#C8EEF5", border:"none", borderRadius:6, padding:"6px 14px", fontWeight:600, fontSize:12, cursor:"pointer" }}>
            🔄 Refresh
          </button>
        </div>
      </div>

      {/* Info box */}
      <div style={{ background:"#f0f9ff", border:"1px solid #bae6fd", borderRadius:8, padding:"12px 16px", fontSize:13, color:"#0369a1" }}>
        <strong>How this works:</strong> When Sales, Marketing or AE staff submit the Public Movement Form, entries appear here first. 
        Review each entry, then click <strong>✓ Approve</strong> to post it into the system, or <strong>✗ Reject</strong> to discard it with a reason. 
        Only approved entries appear in the Movement Log and Fleet Calendar.
      </div>

      {/* Filter tabs */}
      <div style={{ display:"flex", gap:8 }}>
        {[["pending","⏳ Pending",pendingCount],["approved","✓ Approved",null],["rejected","✗ Rejected",null],["all","All",null]].map(([val,label,count]) => (
          <button key={val} onClick={() => setFilter(val)}
            style={{ padding:"7px 16px", borderRadius:6, border:"1px solid", cursor:"pointer", fontWeight:600, fontSize:13,
              borderColor: filter===val ? "#0D1B3E" : "#d1d5db",
              background: filter===val ? "#0D1B3E" : "#fff",
              color: filter===val ? "#fff" : "#374151" }}>
            {label}{count > 0 ? ` (${count})` : ""}
          </button>
        ))}
      </div>

      {/* List */}
      {filtered.length === 0 ? (
        <div style={{ padding:40, textAlign:"center", color:"#9ca3af", background:"#fff", borderRadius:10, border:"1px solid #e5e7eb" }}>
          <div style={{ fontSize:28, marginBottom:8 }}>
            {filter === "pending" ? "✅" : "📋"}
          </div>
          <div style={{ fontWeight:600, fontSize:14, color:"#374151", marginBottom:4 }}>
            {filter === "pending" ? "No pending submissions" : `No ${filter} entries`}
          </div>
          <div style={{ fontSize:13 }}>
            {filter === "pending" ? "All form submissions have been reviewed." : "Nothing to show here."}
          </div>
        </div>
      ) : (
        <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
          {filtered.map(item => {
            const status = item.approvalStatus || "pending";
            const statusColors = { pending:"#d97706", approved:"#16a34a", rejected:"#dc2626" };
            const statusBgs   = { pending:"#fef3c7", approved:"#dcfce7", rejected:"#fee2e2" };
            const isPending = status === "pending";
            return (
              <div key={item.pendingId} style={{ background:"#fff", border:`1px solid ${statusColors[status]}40`, borderRadius:10, overflow:"hidden" }}>
                {/* Status bar */}
                <div style={{ background:statusBgs[status], padding:"8px 16px", display:"flex", alignItems:"center", gap:10, borderBottom:`1px solid ${statusColors[status]}30` }}>
                  <Badge label={status.toUpperCase()} bg={statusColors[status]} />
                  <span style={{ fontSize:12, color:"#6b7280" }}>Submitted {fmt(item.submittedAt?.split("T")[0] || item.date)} by <strong>{item.custodianName}</strong></span>
                  {item.approvedAt && <span style={{ fontSize:11, color:"#16a34a", marginLeft:"auto" }}>Approved {fmt(item.approvedAt.split("T")[0])}</span>}
                  {item.rejectedAt && <span style={{ fontSize:11, color:"#dc2626", marginLeft:"auto" }}>Rejected: {item.rejectReason}</span>}
                </div>

                {/* Content */}
                <div style={{ padding:"14px 16px" }}>
                  <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))", gap:10, marginBottom:12 }}>
                    {[
                      ["Direction", getDir(item.type) === "GR" ? "📥 GR (IN)" : "📤 GI (OUT)"],
                      ["Movement Type", getMovLabel(item.type)],
                      ["Item / Model", getItemName(item.itemId)],
                      ["Serial No.", item.serialNo || "—"],
                      ["Qty", item.qty],
                      ["Entity", item.entity || "—"],
                      ["Region", item.region || "—"],
                      ["Location", item.location || "—"],
                      ["Date", fmt(item.date)],
                      item.expectedReturn ? ["Expected Return", fmt(item.expectedReturn)] : null,
                      item.customerName ? ["Customer", item.customerName] : null,
                      item.eventName ? ["Event", item.eventName] : null,
                    ].filter(Boolean).map(([label, val]) => (
                      <div key={label} style={{ background:"#f9fafb", borderRadius:6, padding:"8px 10px" }}>
                        <div style={{ fontSize:10, color:"#9ca3af", fontWeight:600, textTransform:"uppercase", letterSpacing:0.5, marginBottom:2 }}>{label}</div>
                        <div style={{ fontSize:13, color:"#374151", fontWeight:500 }}>{val}</div>
                      </div>
                    ))}
                  </div>
                  {item.notes && (
                    <div style={{ background:"#f9fafb", borderRadius:6, padding:"8px 10px", fontSize:12, color:"#6b7280", marginBottom:12 }}>
                      <strong>Notes:</strong> {item.notes}
                    </div>
                  )}

                  {/* Action buttons for pending items */}
                  {isPending && (
                    <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                      <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
                        <button onClick={() => approve(item)}
                          style={{ background:"#16a34a", color:"#fff", border:"none", borderRadius:8, padding:"11px 28px", fontWeight:700, fontSize:14, cursor:"pointer" }}>
                          ✓ Approve & Post to System
                        </button>
                        <button onClick={() => setRejectingId(rejectingId === item.pendingId ? null : item.pendingId)}
                          style={{ background:"#fee2e2", color:"#dc2626", border:"1px solid #fca5a5", borderRadius:8, padding:"11px 24px", fontWeight:700, fontSize:14, cursor:"pointer" }}>
                          ✗ Reject
                        </button>
                      </div>
                      {rejectingId === item.pendingId && (
                        <div style={{ display:"flex", gap:8, alignItems:"center", background:"#fff7ed", border:"1px solid #fed7aa", borderRadius:8, padding:"10px 12px" }}>
                          <input value={rejectReason} onChange={e => setRejectReason(e.target.value)}
                            placeholder="Reason for rejection (optional)..."
                            style={{ flex:1, border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13 }} />
                          <button onClick={() => reject(item, rejectReason || "Rejected by Logistics Admin")}
                            style={{ background:"#dc2626", color:"#fff", border:"none", borderRadius:6, padding:"8px 16px", fontWeight:700, fontSize:13, cursor:"pointer" }}>
                            Confirm Reject
                          </button>
                          <button onClick={() => { setRejectingId(null); setRejectReason(""); }}
                            style={{ background:"#f3f4f6", color:"#374151", border:"none", borderRadius:6, padding:"8px 12px", fontWeight:600, fontSize:13, cursor:"pointer" }}>
                            Cancel
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}


// ─── SETUP HELPERS ───────────────────────────────────────────────────────────
const getActiveGR = (setup) => (setup?.grTypes || DEFAULT_SETUP.grTypes).filter(t => t.active !== false);
const getActiveGI = (setup) => (setup?.giTypes || DEFAULT_SETUP.giTypes).filter(t => t.active !== false);
const getActiveTypes = (setup) => [...getActiveGR(setup), ...getActiveGI(setup)];
const getEntities = (setup) => setup?.entities || DEFAULT_SETUP.entities;
const getRegions  = (setup) => setup?.regions  || DEFAULT_SETUP.regions;

// ─── MINI COMPONENTS ─────────────────────────────────────────────────────────
function Badge({ label, bg, color }) {
  return <span style={{ display: "inline-block", padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 600, color: color || "#fff", background: bg || "#0D1B3E", whiteSpace: "nowrap" }}>{label}</span>;
}
function SectionHead({ title }) {
  return <div style={{ fontSize: 12, fontWeight: 700, color: "#4AACCC", letterSpacing: 1, textTransform: "uppercase", marginBottom: 10, borderBottom: "2px solid #C8EEF5", paddingBottom: 5 }}>{title}</div>;
}
function StatCard({ label, value, sub, accent, onClick }) {
  return (
    <div onClick={onClick} style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: "14px 18px", borderLeft: `4px solid ${accent || "#4AACCC"}`, cursor: onClick ? "pointer" : "default" }}>
      <div style={{ fontSize: 24, fontWeight: 700, color: "#0D1B3E" }}>{value}</div>
      <div style={{ fontSize: 12, fontWeight: 600, color: "#374151", marginTop: 2 }}>{label}</div>
      {sub && <div style={{ fontSize: 11, color: "#6b7280", marginTop: 3 }}>{sub}</div>}
    </div>
  );
}

// ─── STOCK POSITION (NEW MAIN TAB) ───────────────────────────────────────────
function StockPosition({ movements, sapSnapshots, gitItems }) {
  const [selMonth, setSelMonth] = useState(new Date().toISOString().slice(0, 7));
  const [selItem, setSelItem] = useState("ALL");

  // Get latest SAP snapshot for selected month
  const sapMonth = sapSnapshots.filter(s => s.month === selMonth);
  // All items that appear in SAP or movements
  const allItemIds = [...new Set([...sapMonth.map(s => s.itemId), ...movements.map(m => m.itemId)])];
  const filtered = selItem === "ALL" ? allItemIds : [selItem];

  // Demo units currently in field from movement tracker
  const demoInField = {};
  movements.forEach(m => {
    if (m.sapCat === "DEMO" && getDir(m.type) === "GI" && m.status !== "Returned") {
      demoInField[m.itemId] = (demoInField[m.itemId] || 0) + m.qty;
    }
  });

  // GIT items not yet received
  const gitPending = {};
  gitItems.filter(g => g.status !== "Received").forEach(g => {
    gitPending[g.itemId] = (gitPending[g.itemId] || 0) + g.qty;
  });

  const rows = filtered.map(itemId => {
    const snap = sapMonth.find(s => s.itemId === itemId) || {};
    const sapStock = snap.sapStock || 0;
    const sapGit = snap.sapGit || 0;
    const sapDemo = snap.sapDemo || 0;
    const trackerDemo = demoInField[itemId] || 0;
    const gitTracker = gitPending[itemId] || 0;
    const effectiveAvail = sapStock - trackerDemo;
    return { itemId, sapStock, sapGit, sapDemo, trackerDemo, gitTracker, effectiveAvail, notes: snap.notes || "" };
  }).filter(r => r.sapStock > 0 || r.sapGit > 0 || r.sapDemo > 0 || r.trackerDemo > 0 || r.gitTracker > 0);

  const totalStock = rows.reduce((s, r) => s + r.sapStock, 0);
  const totalGit = rows.reduce((s, r) => s + r.sapGit, 0);
  const totalDemo = rows.reduce((s, r) => s + r.sapDemo, 0);
  const totalTrackerDemo = rows.reduce((s, r) => s + r.trackerDemo, 0);
  const totalEffective = rows.reduce((s, r) => s + r.effectiveAvail, 0);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
        <SectionHead title="SAP Stock Categories — What Each Means" />
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12 }}>
          {SAP_STOCK_CATEGORIES.map(c => (
            <div key={c.id} style={{ background: c.bg, border: `1px solid ${c.color}30`, borderRadius: 8, padding: 12 }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: c.color }}>{c.label}</div>
              <div style={{ fontSize: 12, color: "#374151", marginTop: 4 }}>{c.desc}</div>
            </div>
          ))}
        </div>
        <div style={{ marginTop: 12, padding: "10px 14px", background: "#f0f9ff", borderRadius: 8, fontSize: 12, color: "#0369a1" }}>
          <strong>Effective Available = SAP Stock − Demo Units In Field (from this tracker)</strong><br />
          This tells you how many units you can actually sell or deploy right now.
        </div>
      </div>

      {/* Filters */}
      <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Month (SAP Snapshot)</label>
          <input type="month" value={selMonth} onChange={e => setSelMonth(e.target.value)}
            style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Item</label>
          <select value={selItem} onChange={e => setSelItem(e.target.value)}
            style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }}>
            <option value="ALL">All Items</option>
            {PUDU_ITEMS.map(i => <option key={i.id} value={i.id}>{i.name}</option>)}
          </select>
        </div>
      </div>

      {/* KPI summary */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(140px,1fr))", gap: 12 }}>
        <StatCard label="SAP Stock" value={totalStock} sub="Warehouse (unrestricted)" accent="#16a34a" />
        <StatCard label="SAP GIT" value={totalGit} sub="In transit (SAP)" accent="#d97706" />
        <StatCard label="SAP Demo" value={totalDemo} sub="Demo pool (SAP)" accent="#7c3aed" />
        <StatCard label="Demo In Field" value={totalTrackerDemo} sub="From this tracker" accent="#4AACCC" />
        <StatCard label="Effective Available" value={totalEffective} sub="Stock minus in-field" accent={totalEffective < 0 ? "#dc2626" : "#0D1B3E"} />
      </div>

      {/* Main table */}
      <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead style={{ background: "#0D1B3E" }}>
            <tr>
              {["Item / Model", "SAP STOCK", "SAP GIT", "SAP DEMO", "Tracker: Demo In Field", "Effective Available", "GIT Pending (Tracker)", "Notes"].map((h, i) => (
                <th key={h} style={{ padding: "10px 10px", textAlign: i > 0 ? "center" : "left", fontWeight: 600, color: "#fff", fontSize: 12, whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 && (
              <tr><td colSpan={8} style={{ padding: 24, textAlign: "center", color: "#9ca3af" }}>
                No SAP snapshot data for {selMonth}. Go to Admin → SAP Snapshots to add data.
              </td></tr>
            )}
            {rows.map(r => (
              <tr key={r.itemId} style={{ borderTop: "1px solid #f3f4f6", background: r.effectiveAvail < 0 ? "#fff1f2" : "transparent" }}>
                <td style={{ padding: "9px 10px", fontWeight: 600, color: "#0D1B3E" }}>{getItemName(r.itemId)}</td>
                <td style={{ padding: "9px 10px", textAlign: "center" }}><Badge label={r.sapStock} bg="#16a34a" /></td>
                <td style={{ padding: "9px 10px", textAlign: "center" }}><Badge label={r.sapGit} bg={r.sapGit > 0 ? "#d97706" : "#9ca3af"} /></td>
                <td style={{ padding: "9px 10px", textAlign: "center" }}><Badge label={r.sapDemo} bg={r.sapDemo > 0 ? "#7c3aed" : "#9ca3af"} /></td>
                <td style={{ padding: "9px 10px", textAlign: "center" }}><Badge label={r.trackerDemo} bg={r.trackerDemo > 0 ? "#4AACCC" : "#9ca3af"} /></td>
                <td style={{ padding: "9px 10px", textAlign: "center" }}>
                  <Badge label={r.effectiveAvail} bg={r.effectiveAvail < 0 ? "#dc2626" : r.effectiveAvail === 0 ? "#9ca3af" : "#0D1B3E"} />
                </td>
                <td style={{ padding: "9px 10px", textAlign: "center" }}>{r.gitTracker > 0 ? <Badge label={r.gitTracker} bg="#d97706" /> : "—"}</td>
                <td style={{ padding: "9px 10px", fontSize: 12, color: "#6b7280" }}>{r.notes}</td>
              </tr>
            ))}
          </tbody>
          {rows.length > 0 && (
            <tfoot style={{ background: "#f9fafb" }}>
              <tr style={{ borderTop: "2px solid #e5e7eb" }}>
                <td style={{ padding: "9px 10px", fontWeight: 700 }}>TOTAL</td>
                <td style={{ padding: "9px 10px", textAlign: "center", fontWeight: 700, color: "#16a34a" }}>{totalStock}</td>
                <td style={{ padding: "9px 10px", textAlign: "center", fontWeight: 700, color: "#d97706" }}>{totalGit}</td>
                <td style={{ padding: "9px 10px", textAlign: "center", fontWeight: 700, color: "#7c3aed" }}>{totalDemo}</td>
                <td style={{ padding: "9px 10px", textAlign: "center", fontWeight: 700, color: "#4AACCC" }}>{totalTrackerDemo}</td>
                <td style={{ padding: "9px 10px", textAlign: "center", fontWeight: 700, color: totalEffective < 0 ? "#dc2626" : "#0D1B3E" }}>{totalEffective}</td>
                <td colSpan={2} />
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  );
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({ movements, personnel }) {
  const now = new Date();
  const thisMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  const monthMov = movements.filter(m => m.date?.startsWith(thisMonth));
  const grCount = monthMov.filter(m => getDir(m.type) === "GR").length;
  const giCount = monthMov.filter(m => getDir(m.type) === "GI").length;
  const inField = movements.filter(m => getDir(m.type) === "GI" && m.status !== "Returned");
  const overdue = movements.filter(m => isOverdue(m));

  const byCustodian = {}, byRole = {}, byModel = {}, byRegion = {}, bySapCat = { STOCK: 0, GIT: 0, DEMO: 0 };
  inField.forEach(m => {
    byCustodian[m.custodian] = (byCustodian[m.custodian] || 0) + m.qty;
    byModel[m.itemId] = (byModel[m.itemId] || 0) + m.qty;
    byRegion[m.region] = (byRegion[m.region] || 0) + m.qty;
    if (m.sapCat) bySapCat[m.sapCat] = (bySapCat[m.sapCat] || 0) + m.qty;
  });
  movements.forEach(m => { const p = personnel.find(x => x.name === m.custodian); byRole[p?.role || "Unknown"] = (byRole[p?.role || "Unknown"] || 0) + 1; });

  const topModels = Object.entries(byModel).sort((a, b) => b[1] - a[1]).slice(0, 8);
  const barMax = Math.max(...topModels.map(([, v]) => v), 1);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {movements.length === 0 && (
        <div style={{ background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)", borderRadius:12, padding:"24px 28px", color:"#fff", textAlign:"center" }}>
          <div style={{ fontSize:32, marginBottom:8 }}>🤖</div>
          <div style={{ fontSize:18, fontWeight:700, marginBottom:6 }}>Welcome to Robotics Fleet & Demo Tracker</div>
          <div style={{ fontSize:13, color:"#C8EEF5", marginBottom:16 }}>No data yet. Start by logging your first movement.</div>
          <div style={{ display:"flex", gap:12, justifyContent:"center", flexWrap:"wrap" }}>
            {[
              ["1. Add your team","Go to Admin → Personnel → add each staff member"],
              ["2. Log a movement","Go to Log Movement → select GI (OUT) → fill in the form"],
              ["3. See it here","Come back to Dashboard — all charts update instantly"],
            ].map(([step, desc]) => (
              <div key={step} style={{ background:"rgba(255,255,255,0.1)", borderRadius:8, padding:"12px 16px", maxWidth:180, textAlign:"left" }}>
                <div style={{ fontWeight:700, fontSize:12, color:"#4AACCC", marginBottom:4 }}>{step}</div>
                <div style={{ fontSize:11, color:"#C8EEF5" }}>{desc}</div>
              </div>
            ))}
          </div>
        </div>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(150px,1fr))", gap: 12 }}>
        <StatCard label="GR This Month" value={grCount} sub="Receipts" accent="#16a34a" />
        <StatCard label="GI This Month" value={giCount} sub="Issues" accent="#dc2626" />
        <StatCard label="Units In Field" value={inField.reduce((s, m) => s + m.qty, 0)} sub="Deployed / On loan" accent="#4AACCC" />
        <StatCard label="Overdue Returns" value={overdue.length} sub="Past expected date" accent="#E07820" />
        <StatCard label="Total Movements" value={movements.length} sub="All time" accent="#0D1B3E" />
      </div>

      {/* SAP Category breakdown of in-field units */}
      <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
        <SectionHead title="In-Field Units by SAP Category" />
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          {SAP_STOCK_CATEGORIES.map(c => (
            <div key={c.id} style={{ background: c.bg, border: `1px solid ${c.color}40`, borderRadius: 8, padding: "12px 20px", textAlign: "center", minWidth: 100 }}>
              <div style={{ fontSize: 24, fontWeight: 700, color: c.color }}>{bySapCat[c.id] || 0}</div>
              <div style={{ fontSize: 12, color: "#374151", fontWeight: 600 }}>{c.label}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
          <SectionHead title="Units In Field by Custodian" />
          {Object.keys(byCustodian).length === 0 ? <p style={{ color: "#9ca3af", fontSize: 13 }}>None currently in field.</p> :
            Object.entries(byCustodian).sort((a, b) => b[1] - a[1]).map(([name, qty]) => (
              <div key={name} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: "1px solid #f3f4f6" }}>
                <span style={{ fontSize: 13 }}>{name}</span>
                <Badge label={qty + " unit" + (qty > 1 ? "s" : "")} bg="#0D1B3E" />
              </div>
            ))}
        </div>
        <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
          <SectionHead title="Units In Field by Region" />
          {Object.keys(byRegion).length === 0 ? <p style={{ color: "#9ca3af", fontSize: 13 }}>None deployed.</p> :
            Object.entries(byRegion).sort((a, b) => b[1] - a[1]).map(([r, qty]) => (
              <div key={r} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: "1px solid #f3f4f6" }}>
                <span style={{ fontSize: 13 }}>{r}</span>
                <Badge label={qty} bg="#4AACCC" />
              </div>
            ))}
        </div>
      </div>

      <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
        <SectionHead title="Top Models In Field" />
        {topModels.length === 0 ? <p style={{ color: "#9ca3af", fontSize: 13 }}>No units in field.</p> :
          topModels.map(([id, qty]) => (
            <div key={id} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
              <div style={{ width: 140, fontSize: 12, color: "#374151", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{getItemName(id)}</div>
              <div style={{ flex: 1, background: "#f3f4f6", borderRadius: 4, height: 14 }}>
                <div style={{ width: `${(qty / barMax) * 100}%`, background: "linear-gradient(90deg,#4AACCC,#0D1B3E)", height: "100%", borderRadius: 4 }} />
              </div>
              <div style={{ width: 20, fontSize: 12, fontWeight: 700, color: "#0D1B3E" }}>{qty}</div>
            </div>
          ))}
      </div>

      {overdue.length > 0 && (
        <div style={{ background: "#fff7ed", border: "1px solid #fed7aa", borderRadius: 10, padding: 16 }}>
          <SectionHead title="Overdue Returns" />
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ color: "#6b7280" }}>
              {["Item", "SAP Cat", "Custodian", "Region", "Expected Return", "Days Overdue"].map(h => (
                <th key={h} style={{ padding: "4px 8px", fontWeight: 600, textAlign: "left" }}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {overdue.map(m => {
                const days = Math.floor((new Date() - new Date(m.expectedReturn)) / 86400000);
                const cat = getSapCatInfo(m.sapCat);
                return (
                  <tr key={m.id} style={{ borderTop: "1px solid #fed7aa" }}>
                    <td style={{ padding: "6px 8px", fontWeight: 600 }}>{getItemName(m.itemId)}</td>
                    <td style={{ padding: "6px 8px" }}><Badge label={m.sapCat || "—"} bg={cat.color || "#9ca3af"} /></td>
                    <td style={{ padding: "6px 8px" }}>{m.custodian}</td>
                    <td style={{ padding: "6px 8px" }}>{m.region}</td>
                    <td style={{ padding: "6px 8px", color: "#dc2626" }}>{fmt(m.expectedReturn)}</td>
                    <td style={{ padding: "6px 8px", fontWeight: 700, color: "#E07820" }}>{days}d</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* ── GR / GI ACTIVITY ANALYSIS ── */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        {/* GR Summary */}
        <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
          <SectionHead title={"GR Summary — " + thisMonth + " (" + monthMov.filter(m=>getDir(m.type)==="GR").reduce((s,m)=>s+m.qty,0) + " units)"} />
          {(() => {
            const grByType = {};
            monthMov.filter(m=>getDir(m.type)==="GR").forEach(m=>{ grByType[m.type]=(grByType[m.type]||0)+m.qty; });
            return Object.keys(grByType).length === 0
              ? <p style={{ color:"#9ca3af",fontSize:13 }}>No GR movements this month.</p>
              : Object.entries(grByType).sort((a,b)=>b[1]-a[1]).map(([t,qty])=>(
                <div key={t} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:"1px solid #f3f4f6" }}>
                  <span style={{ fontSize:12,color:"#374151",flex:1,paddingRight:8 }}>{MOVEMENT_TYPES.GR.find(x=>x.value===t)?.label||t}</span>
                  <Badge label={qty+" unit"+(qty>1?"s":"")} bg="#16a34a" />
                </div>
              ));
          })()}
        </div>
        {/* GI Summary */}
        <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
          <SectionHead title={"GI Summary — " + thisMonth + " (" + monthMov.filter(m=>getDir(m.type)==="GI").reduce((s,m)=>s+m.qty,0) + " units)"} />
          {(() => {
            const giByType = {};
            monthMov.filter(m=>getDir(m.type)==="GI").forEach(m=>{ giByType[m.type]=(giByType[m.type]||0)+m.qty; });
            return Object.keys(giByType).length === 0
              ? <p style={{ color:"#9ca3af",fontSize:13 }}>No GI movements this month.</p>
              : Object.entries(giByType).sort((a,b)=>b[1]-a[1]).map(([t,qty])=>(
                <div key={t} style={{ display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:"1px solid #f3f4f6" }}>
                  <span style={{ fontSize:12,color:"#374151",flex:1,paddingRight:8 }}>{MOVEMENT_TYPES.GI.find(x=>x.value===t)?.label||t}</span>
                  <Badge label={qty+" unit"+(qty>1?"s":"")} bg="#dc2626" />
                </div>
              ));
          })()}
        </div>
      </div>


    </div>
  );
}

// ─── LOG MOVEMENT ─────────────────────────────────────────────────────────────
function LogMovement({ movements, setMovements, personnel, lockedMonths, fleetUnits, setFleetUnits, itemMaster, customers }) {
  const blank = { dir: "", type: "", sapCat: "", itemId: "", qty: 1, serialNo: "", custodian: "", entity: "", location: "", region: "", date: today(), expectedReturn: "", customerName: "", eventName: "", notes: "", status: "Active" };
  const [form, setForm] = useState(blank);
  const [saved, setSaved] = useState(false);

  const dir = form.dir || getDir(form.type);
  const needsReturn = ["GI_LOAN_CUSTOMER","GI_DEMO_SALES","GI_POC_ENGINEER","GI_EXHIBITION"].includes(form.type);
  const needsCustomer = ["GI_CUSTOMER_DELIVERY","GI_LOAN_CUSTOMER"].includes(form.type);
  const needsEvent = form.type === "GI_EXHIBITION";
  const movTypeDef = ALL_MOVEMENT_TYPES.find(t => t.value === form.type);

  function set(k, v) { setForm(f => ({ ...f, [k]: v })); }
  function onTypeChange(val) {
    const def = ALL_MOVEMENT_TYPES.find(t => t.value === val);
    setForm(f => ({ ...f, type: val, sapCat: def?.sapCat || "" }));
  }

  async function handleSubmit() {
    if (!form.type || !form.itemId || !form.custodian || !form.date) { alert("Fill in: Movement Type, Item, Logged By, and Date."); return; }
    const movMonth = form.date.slice(0, 7);
    if (lockedMonths && lockedMonths.includes(movMonth)) {
      alert("Month " + movMonth + " is locked. No new entries can be posted to a locked month. Please contact Finance Admin to unlock if needed.");
      return;
    }
    const newMov = { ...form, id: Date.now(), qty: Number(form.qty) || 1, createdAt: new Date().toISOString() };
    const updated = [newMov, ...movements];
    setMovements(updated);
    await save(SK.MOV, updated);

    // Auto-register unit in Fleet Calendar if serial number provided
    if (form.serialNo && form.serialNo.trim() && fleetUnits !== undefined) {
      const existing = fleetUnits.find(u => u.serialNo === form.serialNo.trim());
      if (!existing) {
        // New unit — register it automatically
        const newUnit = {
          id: Date.now() + 1,
          entity: form.entity || "",
          model: getItemName(form.itemId),
          itemId: form.itemId,
          itemDesc: form.itemDesc || "",
          variant: "",
          serialNo: form.serialNo.trim(),
          location: form.location || "",
          remark: form.remark || "",
        };
        const updatedFleet = [...fleetUnits, newUnit];
        setFleetUnits(updatedFleet);
        await save(SK.FLEET, updatedFleet);
      } else {
        // Update existing unit location and remark from latest movement
        const updatedFleet = fleetUnits.map(u =>
          u.serialNo === form.serialNo.trim()
            ? { ...u, location: form.location || u.location, remark: form.remark || u.remark, entity: form.entity || u.entity, itemDesc: form.itemDesc || u.itemDesc }
            : u
        );
        setFleetUnits(updatedFleet);
        await save(SK.FLEET, updatedFleet);
      }
    }

    setForm(blank);
    setSaved(true);
    setTimeout(() => setSaved(false), 3000);
  }

  const inp = (label, key, type = "text", extra = {}) => (
    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
      <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>{label}</label>
      <input type={type} value={form[key]} onChange={e => set(key, e.target.value)}
        style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 10px", fontSize: 13 }} {...extra} />
    </div>
  );

  return (
    <div style={{ maxWidth: 680, display: "flex", flexDirection: "column", gap: 18 }}>
      {saved && <div style={{ background: "#dcfce7", border: "1px solid #86efac", borderRadius: 8, padding: "10px 16px", color: "#166534", fontWeight: 600, fontSize: 13 }}>Movement logged successfully.</div>}

      <div>
        <SectionHead title="Direction" />
        <div style={{ display: "flex", gap: 10 }}>
          {["GR", "GI"].map(d => (
            <button key={d} onClick={() => setForm(f => ({ ...f, dir: d, type: "", sapCat: "", itemId: "" }))}
              style={{ padding: "10px 28px", borderRadius: 8, fontWeight: 700, fontSize: 14, cursor: "pointer", border: "2px solid", borderColor: dir === d ? (d === "GR" ? "#16a34a" : "#dc2626") : "#d1d5db", background: dir === d ? (d === "GR" ? "#dcfce7" : "#fee2e2") : "#fff", color: dir === d ? (d === "GR" ? "#166534" : "#991b1b") : "#6b7280" }}>
              {d === "GR" ? "📥 Goods Receipt (IN)" : "📤 Goods Issue (OUT)"}
            </button>
          ))}
        </div>
      </div>

      <div>
        <SectionHead title="Movement Details" />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div style={{ gridColumn: "1/-1", display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Movement Type *</label>
            <select value={form.type} onChange={e => onTypeChange(e.target.value)}
              style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 10px", fontSize: 13 }}>
              <option value="">— Select —</option>
              {(dir === "GR" ? MOVEMENT_TYPES.GR : dir === "GI" ? MOVEMENT_TYPES.GI : ALL_MOVEMENT_TYPES).map(t => (
                <option key={t.value} value={t.value}>{t.label}</option>
              ))}
            </select>
          </div>
          {form.sapCat && (
            <div style={{ gridColumn: "1/-1", padding: "8px 12px", borderRadius: 6, background: getSapCatInfo(form.sapCat).bg || "#f9fafb", border: `1px solid ${getSapCatInfo(form.sapCat).color || "#e5e7eb"}40`, fontSize: 12 }}>
              <strong>SAP Category: {form.sapCat}</strong> — {getSapCatInfo(form.sapCat).desc}
            </div>
          )}
          <div style={{ gridColumn: "1/-1", display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Item / SAP Code + Description *</label>
            <select value={form.itemId} onChange={e => {
                const selected = (itemMaster||[]).find(i => i.sapCode === e.target.value) || (itemMaster||[]).find(i => i.id === e.target.value);
                set("itemId", e.target.value);
                if (selected) {
                  setForm(f => ({ ...f, itemId: e.target.value, itemDesc: selected.name }));
                }
              }}
              style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 10px", fontSize: 13 }}>
              <option value="">— Select Item —</option>
              {["Robot","Accessory","Spare Part","Other"].map(cat => {
                const catItems = (itemMaster||[]).filter(i => i.category===cat && i.active!==false);
                return catItems.length > 0 ? (
                  <optgroup key={cat} label={cat}>
                    {catItems.map(i => <option key={i.id} value={i.sapCode}>[{i.sapCode}] {i.name}</option>)}
                  </optgroup>
                ) : null;
              })}
            </select>
          </div>
          <div style={{ gridColumn: "1/-1", display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Item Description <span style={{ color:"#9ca3af", fontSize:10 }}>(auto-fills from item master, editable)</span></label>
            <input value={form.itemDesc||""} onChange={e => set("itemDesc", e.target.value)}
              placeholder="Auto-filled from item master — or type manually"
              style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 10px", fontSize: 13 }} />
          </div>
          {inp("Serial No. / Asset Tag *", "serialNo", "text", { placeholder: "e.g. 8260B5318060021" })}
          {inp("Quantity", "qty", "number", { min: 1 })}
          {inp("Date *", "date", "date")}
          {needsReturn && inp("Expected Return Date", "expectedReturn", "date")}
        </div>
      </div>

      <div>
        <SectionHead title="Custodian & Location" />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Logged By / Custodian *</label>
            <select value={form.custodian} onChange={e => {
                const p = personnel.find(x=>x.name===e.target.value);
                setForm(f=>({...f, custodian:e.target.value, entity:p?.entity||f.entity, region:p?.region||f.region}));
              }}
              style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 10px", fontSize: 13 }}>
              <option value="">— Select Person —</option>
              {personnel.map(p => <option key={p.id} value={p.name}>{p.name} ({p.role} — {p.entity||""})</option>)}
            </select>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Region</label>
            <select value={form.region} onChange={e => set("region", e.target.value)}
              style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 10px", fontSize: 13 }}>
              <option value="">— Select Region —</option>
              {REGIONS.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Entity</label>
            <select value={form.entity} onChange={e => set("entity", e.target.value)}
              style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 10px", fontSize: 13 }}>
              <option value="">— Select Entity —</option>
              {ENTITIES.map(e => <option key={e} value={e}>{e}</option>)}
            </select>
          </div>
          {inp("Current Location / Site *", "location", "text", { placeholder: "e.g. KL Office, UDA Warehouse, TMS Office, PT Flex Batam" })}
          <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
            <label style={{ fontSize:12, fontWeight:600, color:"#374151" }}>Remark / Status <span style={{ color:"#9ca3af", fontSize:10 }}>(shown in Fleet Calendar)</span></label>
            <select value={form.remark} onChange={e => set("remark", e.target.value)}
              style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"8px 10px", fontSize:13 }}>
              <option value="">— Select Remark —</option>
              <option value="KL Demo">KL Demo</option>
              <option value="PEN Demo">PEN Demo</option>
              <option value="KL Stock">KL Stock</option>
              <option value="PEN Stock">PEN Stock</option>
              <option value="KL DEMO">KL DEMO</option>
              <option value="PEN DEMO">PEN DEMO</option>
              <option value="POC">POC</option>
              <option value="Exhibition">Exhibition</option>
              <option value="Loan">Loan</option>
              <option value="In Transit">In Transit</option>
              <option value="Collected Back">Collected Back</option>
              <option value="Reserved for Next Demo">Reserved for Next Demo</option>
            </select>
          </div>
          {needsCustomer && (
            <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
              <label style={{ fontSize:12, fontWeight:600, color:"#374151" }}>Customer Name</label>
              <select value={form.customerName} onChange={e => set("customerName", e.target.value)}
                style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"8px 10px", fontSize:13 }}>
                <option value="">— Select Customer —</option>
                {(customers||[]).filter(c => c.active!==false).map(c => (
                  <option key={c.id} value={c.name}>{c.name}{c.country ? " ("+c.country+")" : ""}</option>
                ))}
                <option value="__other__">Other (type below)</option>
              </select>
              {form.customerName === "__other__" && (
                <input value={form.customerNameOther||""} onChange={e => set("customerNameOther", e.target.value)}
                  placeholder="Type customer name..."
                  style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"8px 10px", fontSize:13, marginTop:4 }} />
              )}
            </div>
          )}
          {needsEvent && inp("Event / Exhibition Name", "eventName", "text", { placeholder: "e.g. MIHAS 2026" })}
          <div style={{ gridColumn: "1/-1", display: "flex", flexDirection: "column", gap: 4 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Notes</label>
            <textarea value={form.notes} onChange={e => set("notes", e.target.value)} rows={2}
              placeholder="Brief purpose..." style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 10px", fontSize: 13, resize: "vertical" }} />
          </div>
        </div>
      </div>
      <button onClick={handleSubmit} style={{ background: "linear-gradient(135deg,#0D1B3E,#4AACCC)", color: "#fff", border: "none", borderRadius: 8, padding: "12px 32px", fontWeight: 700, fontSize: 15, cursor: "pointer", alignSelf: "flex-start" }}>
        Log Movement
      </button>
    </div>
  );
}

// ─── MOVEMENT LOG ─────────────────────────────────────────────────────────────
function MovementLog({ movements, setMovements, lockedMonths }) {
  const [fDir, setFDir] = useState("ALL"), [fType, setFType] = useState("ALL"), [fSapCat, setFSapCat] = useState("ALL"), [fMonth, setFMonth] = useState(""), [search, setSearch] = useState("");

  // Check if a movement's month is locked
  const isMonthLocked = (date) => {
    if (!date) return false;
    const m = date.slice(0, 7);
    return lockedMonths.includes(m);
  };

  const filtered = useMemo(() => movements.filter(m => {
    if (fDir !== "ALL" && getDir(m.type) !== fDir) return false;
    if (fType !== "ALL" && m.type !== fType) return false;
    if (fSapCat !== "ALL" && m.sapCat !== fSapCat) return false;
    if (fMonth && !m.date?.startsWith(fMonth)) return false;
    if (search && !getItemName(m.itemId).toLowerCase().includes(search.toLowerCase()) && !m.custodian?.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  }), [movements, fDir, fType, fSapCat, fMonth, search]);

  async function markReturned(id) {
    const m = movements.find(x => x.id === id);
    if (m && isMonthLocked(m.date)) { alert("This record is in a locked month (" + m.date.slice(0,7) + "). It cannot be amended. Post a reversal instead."); return; }
    const updated = movements.map(m => m.id === id ? { ...m, status: "Returned", actualReturn: today() } : m);
    setMovements(updated); await save(SK.MOV, updated);
  }

  async function reverseMovement(m) {
    if (isMonthLocked(m.date)) {
      // Post a reversal entry (opposite direction, same data, tagged as reversal)
      const reversalType = getDir(m.type) === "GI"
        ? (MOVEMENT_TYPES.GR.find(t => t.sapCat === m.sapCat) || MOVEMENT_TYPES.GR[0]).value
        : (MOVEMENT_TYPES.GI.find(t => t.sapCat === m.sapCat) || MOVEMENT_TYPES.GI[0]).value;
      const reversal = {
        ...m,
        id: Date.now(),
        type: reversalType,
        date: today(),
        status: "Active",
        isReversal: true,
        reversalOf: m.id,
        reversalDate: today(),
        notes: "[REVERSAL of " + fmt(m.date) + " entry] " + (m.notes || ""),
        createdAt: new Date().toISOString(),
      };
      const updated = [...movements, reversal];
      setMovements(updated); await save(SK.MOV, updated);
      alert("Reversal posted for " + getItemName(m.itemId) + ". A new offsetting entry has been created with today's date. Both records are preserved for audit trail.");
    } else {
      // Unlocked month — still post reversal (no delete/amend policy)
      const reversal = {
        ...m,
        id: Date.now(),
        type: getDir(m.type) === "GI"
          ? (MOVEMENT_TYPES.GR.find(t => t.sapCat === m.sapCat) || MOVEMENT_TYPES.GR[0]).value
          : (MOVEMENT_TYPES.GI.find(t => t.sapCat === m.sapCat) || MOVEMENT_TYPES.GI[0]).value,
        date: today(),
        status: "Active",
        isReversal: true,
        reversalOf: m.id,
        reversalDate: today(),
        notes: "[REVERSAL] " + (m.notes || ""),
        createdAt: new Date().toISOString(),
      };
      const updated = [...movements, reversal];
      setMovements(updated); await save(SK.MOV, updated);
      alert("Reversal posted. Original record is preserved for audit trail.");
    }
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <select value={fDir} onChange={e => { setFDir(e.target.value); setFType("ALL"); }} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }}>
          <option value="ALL">All Directions</option><option value="GR">GR (In)</option><option value="GI">GI (Out)</option>
        </select>
        <select value={fType} onChange={e => setFType(e.target.value)} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }}>
          <option value="ALL">All Types</option>
          {(fDir === "GR" ? MOVEMENT_TYPES.GR : fDir === "GI" ? MOVEMENT_TYPES.GI : ALL_MOVEMENT_TYPES).map(t => <option key={t.value} value={t.value}>{t.label}</option>)}
        </select>
        <select value={fSapCat} onChange={e => setFSapCat(e.target.value)} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }}>
          <option value="ALL">All SAP Categories</option>
          {SAP_STOCK_CATEGORIES.map(c => <option key={c.id} value={c.id}>{c.label} — {c.desc}</option>)}
        </select>
        <input type="month" value={fMonth} onChange={e => setFMonth(e.target.value)} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
        <input placeholder="Search..." value={search} onChange={e => setSearch(e.target.value)} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13, minWidth: 180 }} />
        <span style={{ fontSize: 12, color: "#6b7280", alignSelf: "center" }}>{filtered.length} records</span>
      </div>
      <div style={{ overflowX: "auto", background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead style={{ background: "#f9fafb" }}>
            <tr>{["Dir", "SAP Cat", "Movement Type", "Item", "S/N", "Qty", "Custodian", "Region", "Location", "Remark", "Date", "Exp.Return", "Status", "Actions"].map(h => (
              <th key={h} style={{ padding: "9px 8px", textAlign: "left", fontWeight: 600, color: "#374151", borderBottom: "1px solid #e5e7eb", whiteSpace: "nowrap" }}>{h}</th>
            ))}</tr>
          </thead>
          <tbody>
            {filtered.length === 0 && <tr><td colSpan={12} style={{ padding: 24, textAlign: "center", color: "#9ca3af" }}>No records.</td></tr>}
            {filtered.map(m => {
              const over = isOverdue(m);
              const cat = getSapCatInfo(m.sapCat);
              const locked = isMonthLocked(m.date);
              const isRev = m.isReversal;
              return (
                <tr key={m.id} style={{ borderTop: "1px solid #f3f4f6", background: isRev ? "#f0fdf4" : over ? "#fff7ed" : locked ? "#fefce8" : "transparent" }}>
                  <td style={{ padding: "7px 8px" }}>
                    <div style={{ display:"flex", flexDirection:"column", gap:2 }}>
                      <Badge label={getDir(m.type)} bg={getDir(m.type) === "GR" ? "#16a34a" : "#dc2626"} />
                      {isRev && <Badge label="REV" bg="#7c3aed" />}
                      {locked && <Badge label="🔒" bg="#d97706" />}
                    </div>
                  </td>
                  <td style={{ padding: "7px 8px" }}><Badge label={m.sapCat || "—"} bg={cat.color || "#9ca3af"} /></td>
                  <td style={{ padding: "7px 8px", fontSize: 11, color: "#374151", maxWidth: 150 }}>{getMovLabel(m.type)}</td>
                  <td style={{ padding: "7px 8px", fontWeight: 600, color: "#0D1B3E", whiteSpace: "nowrap" }}>{getItemName(m.itemId)}</td>
                  <td style={{ padding: "7px 8px", color: "#6b7280", fontSize: 11 }}>{m.serialNo || "—"}</td>
                  <td style={{ padding: "7px 8px", textAlign: "center", fontWeight: 700 }}>{m.qty}</td>
                  <td style={{ padding: "7px 8px", whiteSpace: "nowrap" }}>{m.custodian}</td>
                  <td style={{ padding: "7px 8px", whiteSpace: "nowrap" }}>{m.region}</td>
                  <td style={{ padding: "7px 8px", fontSize:11, color:"#6b7280", maxWidth:100, overflow:"hidden", textOverflow:"ellipsis" }}>{m.location || "—"}</td>
                  <td style={{ padding: "7px 8px" }}>{m.remark ? <Badge label={m.remark} bg="#4AACCC" /> : "—"}</td>
                  <td style={{ padding: "7px 8px", whiteSpace: "nowrap" }}>{fmt(m.date)}</td>
                  <td style={{ padding: "7px 8px", whiteSpace: "nowrap", color: over ? "#dc2626" : "#374151" }}>{over ? "⚠ " : ""}{fmt(m.expectedReturn)}</td>
                  <td style={{ padding: "7px 8px" }}><Badge label={m.status === "Returned" ? "Returned" : over ? "Overdue" : "Active"} bg={m.status === "Returned" ? "#16a34a" : over ? "#E07820" : "#4AACCC"} /></td>
                  <td style={{ padding: "7px 8px", whiteSpace: "nowrap" }}>
                    {!isRev && m.status !== "Returned" && getDir(m.type) === "GI" && !locked && (
                      <button onClick={() => markReturned(m.id)} style={{ background: "#dcfce7", border: "1px solid #86efac", color: "#166534", borderRadius: 4, padding: "3px 7px", fontSize: 11, cursor: "pointer", marginRight: 4 }}>Return</button>
                    )}
                    {!isRev && (
                      <button onClick={() => reverseMovement(m)} style={{ background: "#ede9fe", border: "1px solid #c4b5fd", color: "#7c3aed", borderRadius: 4, padding: "3px 7px", fontSize: 11, cursor: "pointer" }}>
                        {locked ? "🔒 Reverse" : "↩ Reverse"}
                      </button>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
function GITRegister({ gitItems, setGitItems }) {
  const blank = { poNo: "", itemId: "", qty: 1, supplier: "", shipDate: "", etaDate: "", vesselRef: "", status: "On Order", notes: "" };
  const [form, setForm] = useState(blank);
  const [adding, setAdding] = useState(false);

  function setF(k, v) { setForm(f => ({ ...f, [k]: v })); }
  async function addGit() {
    if (!form.poNo || !form.itemId) { alert("PO Number and Item are required."); return; }
    const updated = [{ ...form, id: Date.now(), qty: Number(form.qty) || 1 }, ...gitItems];
    setGitItems(updated); await save(SK.GIT, updated); setForm(blank); setAdding(false);
  }
  async function updateStatus(id, status) {
    const updated = gitItems.map(g => g.id === id ? { ...g, status, ...(status === "Received" ? { receivedDate: today() } : {}) } : g);
    setGitItems(updated); await save(SK.GIT, updated);
  }
  async function del(id) {
    if (!appConfirm("Delete GIT record?")) return;
    const updated = gitItems.filter(g => g.id !== id); setGitItems(updated); await save(SK.GIT, updated);
  }

  const statusColors = { "On Order": "#7c3aed", "Shipped": "#d97706", "Customs": "#dc2626", "Received": "#16a34a" };
  const pending = gitItems.filter(g => g.status !== "Received");
  const received = gitItems.filter(g => g.status === "Received");

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 10, padding: 14, fontSize: 13 }}>
        <strong style={{ color: "#92400e" }}>What is GIT (Goods in Transit)?</strong>
        <p style={{ margin: "6px 0 0", color: "#78350f" }}>
          Units you have ordered from Pudu HQ but not yet received physically in your warehouse. These should appear in SAP as GIT stock. Track them here separately — they must NOT be mixed with warehouse stock or demo fleet until the physical GR is done in SAP.
        </p>
      </div>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ display: "flex", gap: 12 }}>
          <StatCard label="In Transit / On Order" value={pending.reduce((s, g) => s + g.qty, 0)} sub={pending.length + " PO lines"} accent="#d97706" />
          <StatCard label="Received (this cycle)" value={received.reduce((s, g) => s + g.qty, 0)} sub={received.length + " lines completed"} accent="#16a34a" />
        </div>
        <button onClick={() => setAdding(a => !a)} style={{ background: "#0D1B3E", color: "#fff", border: "none", borderRadius: 8, padding: "10px 20px", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
          {adding ? "Cancel" : "+ Add GIT Entry"}
        </button>
      </div>

      {adding && (
        <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
          <SectionHead title="New GIT Entry" />
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
            {[["PO Number *", "poNo"], ["Supplier", "supplier"], ["Vessel / Ref No.", "vesselRef"]].map(([l, k]) => (
              <div key={k} style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>{l}</label>
                <input value={form[k]} onChange={e => setF(k, e.target.value)} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
              </div>
            ))}
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Item *</label>
              <select value={form.itemId} onChange={e => setF("itemId", e.target.value)} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }}>
                <option value="">— Select —</option>
                <optgroup label="Robots">{PUDU_ITEMS.filter(i => i.category === "Robot").map(i => <option key={i.id} value={i.id}>{i.name}</option>)}</optgroup>
                <optgroup label="Accessories">{PUDU_ITEMS.filter(i => i.category === "Accessory").map(i => <option key={i.id} value={i.id}>{i.name}</option>)}</optgroup>
              </select>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Qty</label>
              <input type="number" min={1} value={form.qty} onChange={e => setF("qty", e.target.value)} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Status</label>
              <select value={form.status} onChange={e => setF("status", e.target.value)} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }}>
                {["On Order","Shipped","Customs","Received"].map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
            {[["Ship Date", "shipDate"], ["ETA Date", "etaDate"]].map(([l, k]) => (
              <div key={k} style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>{l}</label>
                <input type="date" value={form[k]} onChange={e => setF(k, e.target.value)} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
              </div>
            ))}
            <div style={{ gridColumn: "1/-1", display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Notes</label>
              <input value={form.notes} onChange={e => setF("notes", e.target.value)} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
            </div>
          </div>
          <button onClick={addGit} style={{ marginTop: 12, background: "#16a34a", color: "#fff", border: "none", borderRadius: 6, padding: "9px 20px", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>Save GIT Entry</button>
        </div>
      )}

      <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead style={{ background: "#fef3c7" }}>
            <tr>{["PO No.", "Item", "Qty", "Supplier", "Vessel Ref", "Ship Date", "ETA", "Status", "Notes", "Actions"].map(h => (
              <th key={h} style={{ padding: "9px 10px", textAlign: "left", fontWeight: 600, color: "#92400e", borderBottom: "1px solid #fde68a", whiteSpace: "nowrap" }}>{h}</th>
            ))}</tr>
          </thead>
          <tbody>
            {gitItems.length === 0 && <tr><td colSpan={10} style={{ padding: 20, textAlign: "center", color: "#9ca3af" }}>No GIT entries.</td></tr>}
            {gitItems.map(g => (
              <tr key={g.id} style={{ borderTop: "1px solid #f3f4f6", background: g.status === "Received" ? "#f0fdf4" : "transparent" }}>
                <td style={{ padding: "8px 10px", fontWeight: 600 }}>{g.poNo}</td>
                <td style={{ padding: "8px 10px", color: "#0D1B3E", fontWeight: 600 }}>{getItemName(g.itemId)}</td>
                <td style={{ padding: "8px 10px", textAlign: "center", fontWeight: 700 }}>{g.qty}</td>
                <td style={{ padding: "8px 10px", color: "#6b7280" }}>{g.supplier}</td>
                <td style={{ padding: "8px 10px", color: "#6b7280", fontSize: 12 }}>{g.vesselRef || "—"}</td>
                <td style={{ padding: "8px 10px" }}>{fmt(g.shipDate)}</td>
                <td style={{ padding: "8px 10px" }}>{fmt(g.etaDate)}</td>
                <td style={{ padding: "8px 10px" }}><Badge label={g.status} bg={statusColors[g.status] || "#9ca3af"} /></td>
                <td style={{ padding: "8px 10px", fontSize: 12, color: "#6b7280" }}>{g.notes}</td>
                <td style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>
                  {g.status !== "Received" && (
                    <select value={g.status} onChange={e => updateStatus(g.id, e.target.value)} style={{ fontSize: 11, border: "1px solid #d1d5db", borderRadius: 4, padding: "3px 6px", marginRight: 4 }}>
                      {["On Order","Shipped","Customs","Received"].map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                  )}
                  <button onClick={() => del(g.id)} style={{ background: "#fee2e2", border: "1px solid #fca5a5", color: "#991b1b", borderRadius: 4, padding: "3px 7px", fontSize: 11, cursor: "pointer" }}>Del</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── ADMIN ─────────────────────────────────────────────────────────────────────
function Admin({ movements, setMovements, personnel, setPersonnel, sapSnapshots, setSapSnapshots,
    lockedMonths, setLockedMonths, setup, setSetup, currentUser,
    itemMaster, setItemMaster, customers, setCustomers,
    fleetUnits, setFleetUnits, versions, setVersions, gitItems, setGitItems }) {

  const [tab, setTab] = useState("staff");
  const s = setup || DEFAULT_SETUP;
  const isSuperAdmin = currentUser && currentUser.role === "Super Admin";

  // Staff state
  const blankP = { name:"", role:"", entity:"", region:"", email:"", phone:"", active:true };
  const [newP, setNewP] = useState(blankP);
  const [editId, setEditId] = useState(null);
  const [editData, setEditData] = useState({});
  const [staffSearch, setStaffSearch] = useState("");

  // Lists state
  const SEP = "\n";
  const [listDrafts, setListDrafts] = useState({
    entities: (s.entities || DEFAULT_SETUP.entities || []).join("\n"),
    regions:  (s.regions  || DEFAULT_SETUP.regions  || []).join("\n"),
    roles:    (s.roles    || DEFAULT_SETUP.roles    || []).join("\n"),
    remarks:  (s.remarks  || DEFAULT_SETUP.remarks  || []).join("\n"),
  });
  const [savedListKey, setSavedListKey] = useState("");

  // SAP state
  const [snapForm, setSnapForm] = useState({ month: new Date().toISOString().slice(0,7), itemId:"", sapStock:0, sapGit:0, sapDemo:0 });

  // Backup state
  const [backupDone, setBackupDone] = useState(false);
  const [restoreStatus, setRestoreStatus] = useState("idle");
  const [restorePreview, setRestorePreview] = useState(null);
  const [restoreError, setRestoreError] = useState("");

  const roles    = s.roles    || DEFAULT_SETUP.roles    || [];
  const entities = s.entities || DEFAULT_SETUP.entities || [];
  const regions  = s.regions  || DEFAULT_SETUP.regions  || [];
  const remarks  = s.remarks  || DEFAULT_SETUP.remarks  || [];

  const filteredStaff = personnel.filter(p =>
    !staffSearch ||
    p.name.toLowerCase().includes(staffSearch.toLowerCase()) ||
    (p.role||"").toLowerCase().includes(staffSearch.toLowerCase()) ||
    (p.entity||"").toLowerCase().includes(staffSearch.toLowerCase())
  );

  const SEL = { width:"100%", border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13 };

  // Staff actions
  async function addStaff() {
    if (!newP.name.trim()) { alert("Name is required."); return; }
    const updated = [...personnel, { ...newP, id: Date.now() }];
    setPersonnel(updated); await save(SK.PERS, updated); setNewP(blankP);
  }
  async function saveEdit(id) {
    const updated = personnel.map(p => p.id === id ? { ...p, ...editData } : p);
    setPersonnel(updated); await save(SK.PERS, updated); setEditId(null);
  }
  async function toggleActive(id) {
    const updated = personnel.map(p => p.id === id ? { ...p, active: !p.active } : p);
    setPersonnel(updated); await save(SK.PERS, updated);
  }
  async function removePerson(id) {
    if (!appConfirm("Remove this staff member?")) return;
    const updated = personnel.filter(p => p.id !== id);
    setPersonnel(updated); await save(SK.PERS, updated);
  }

  // Lists actions
  function saveList(key) {
    const values = listDrafts[key].split("\n").map(x=>x.trim()).filter(Boolean);
    const updated = { ...s, [key]: values };
    setSetup(updated); save(SK.SETUP, updated);
    setSavedListKey(key); setTimeout(()=>setSavedListKey(""), 2000);
  }

  // SAP actions
  async function addSnapshot() {
    if (!snapForm.itemId || !snapForm.month) { alert("Month and Item required."); return; }
    const exists = sapSnapshots.find(x => x.month===snapForm.month && x.itemId===snapForm.itemId);
    let updated;
    if (exists) updated = sapSnapshots.map(x => x.month===snapForm.month && x.itemId===snapForm.itemId ? {...x,...snapForm,sapStock:Number(snapForm.sapStock),sapGit:Number(snapForm.sapGit),sapDemo:Number(snapForm.sapDemo)} : x);
    else updated = [...sapSnapshots, {...snapForm, id:Date.now(), sapStock:Number(snapForm.sapStock), sapGit:Number(snapForm.sapGit), sapDemo:Number(snapForm.sapDemo)}];
    setSapSnapshots(updated); await save(SK.SAP, updated);
  }
  async function delSnapshot(id) {
    const updated = sapSnapshots.filter(x => x.id !== id);
    setSapSnapshots(updated); await save(SK.SAP, updated);
  }

  // Lock actions
  async function toggleLock(month) {
    const updated = (lockedMonths||[]).includes(month)
      ? (lockedMonths||[]).filter(m => m !== month)
      : [...(lockedMonths||[]), month];
    setLockedMonths(updated); await save(SK.LOCK, updated);
  }

  // Backup actions
  function doBackup() {
    const payload = {
      exportedAt: new Date().toISOString(),
      appVersion: "FleetTrack Pro",
      exportedBy: currentUser ? currentUser.name : "Admin",
      data: { movements:movements||[], personnel:personnel||[], sapSnapshots:sapSnapshots||[],
              gitItems:gitItems||[], fleetUnits:fleetUnits||[], versions:versions||[], lockedMonths:lockedMonths||[] }
    };
    const blob = new Blob([JSON.stringify(payload,null,2)], {type:"application/json"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href=url; a.download="FleetTrack_Backup_"+today()+".json"; a.click();
    URL.revokeObjectURL(url);
    setBackupDone(true); setTimeout(()=>setBackupDone(false), 3000);
  }
  function handleRestoreFile(e) {
    if (!isSuperAdmin) { alert("Only Super Admin can restore."); return; }
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try { const d = JSON.parse(ev.target.result); setRestorePreview(d); setRestoreStatus("preview"); }
      catch(err) { setRestoreError("Invalid file: "+err.message); setRestoreStatus("error"); }
    };
    reader.readAsText(file);
  }
  async function confirmRestore() {
    if (!isSuperAdmin || !restorePreview) return;
    const d = restorePreview.data || restorePreview;
    if (d.movements)    { setMovements(d.movements);        await save(SK.MOV,   d.movements); }
    if (d.personnel)    { setPersonnel(d.personnel);        await save(SK.PERS,  d.personnel); }
    if (d.sapSnapshots) { setSapSnapshots(d.sapSnapshots);  await save(SK.SAP,   d.sapSnapshots); }
    if (d.gitItems)     { setGitItems(d.gitItems);          await save(SK.GIT,   d.gitItems); }
    if (d.fleetUnits)   { setFleetUnits(d.fleetUnits);      await save(SK.FLEET, d.fleetUnits); }
    if (d.versions)     { setVersions(d.versions);          await save(VERSION_KEY, d.versions); }
    if (d.lockedMonths) { setLockedMonths(d.lockedMonths);  await save(SK.LOCK,  d.lockedMonths); }
    setRestoreStatus("done"); setRestorePreview(null);
  }

  const tabs = [
    ["staff","👥 Staff Master"],["items","📦 Item Master"],["customers","🏢 Customer Master"],
    ["lists","📋 Dropdown Lists"],["sap","📊 SAP Snapshots"],["lock","🔒 Month Lock"],["backup","💾 Backup & Restore"]
  ];

  const months = Array.from({length:12},(_,i)=>{ const d=new Date(); d.setMonth(d.getMonth()-i); return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0"); });

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div style={{ background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)", borderRadius:10, padding:"14px 20px" }}>
        <div style={{ color:"#fff", fontWeight:700, fontSize:16 }}>Admin Panel</div>
        <div style={{ color:"#C8EEF5", fontSize:12, marginTop:2 }}>Master data, settings and backup</div>
      </div>
      <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
        {tabs.map(([val,lbl])=>(
          <button key={val} onClick={()=>setTab(val)} style={{ padding:"8px 14px", borderRadius:6, border:"1px solid", cursor:"pointer", fontWeight:600, fontSize:12, borderColor:tab===val?"#0D1B3E":"#e5e7eb", background:tab===val?"#0D1B3E":"#fff", color:tab===val?"#fff":"#374151" }}>
            {lbl}
          </button>
        ))}
      </div>

      {/* ── STAFF MASTER ── */}
      {tab==="staff" && (
        <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
          <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:16 }}>
            <SectionHead title="Add Staff Member" />
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))", gap:10, marginBottom:12 }}>
              <div>
                <label style={{ fontSize:11, fontWeight:600, color:"#374151", display:"block", marginBottom:3 }}>Full Name *</label>
                <input value={newP.name} onChange={e=>setNewP(p=>({...p,name:e.target.value}))} placeholder="e.g. Ahmad Faisal" style={SEL} />
              </div>
              <div>
                <label style={{ fontSize:11, fontWeight:600, color:"#374151", display:"block", marginBottom:3 }}>Role</label>
                <select value={newP.role||""} onChange={e=>setNewP(p=>({...p,role:e.target.value}))} style={SEL}>
                  <option value="">— Role —</option>
                  {roles.map(r=><option key={r} value={r}>{r}</option>)}
                </select>
              </div>
              <div>
                <label style={{ fontSize:11, fontWeight:600, color:"#374151", display:"block", marginBottom:3 }}>Entity</label>
                <select value={newP.entity||""} onChange={e=>setNewP(p=>({...p,entity:e.target.value}))} style={SEL}>
                  <option value="">— Entity —</option>
                  {entities.map(e=><option key={e} value={e}>{e}</option>)}
                </select>
              </div>
              <div>
                <label style={{ fontSize:11, fontWeight:600, color:"#374151", display:"block", marginBottom:3 }}>Region</label>
                <select value={newP.region||""} onChange={e=>setNewP(p=>({...p,region:e.target.value}))} style={SEL}>
                  <option value="">— Region —</option>
                  {regions.map(r=><option key={r} value={r}>{r}</option>)}
                </select>
              </div>
              <div>
                <label style={{ fontSize:11, fontWeight:600, color:"#374151", display:"block", marginBottom:3 }}>Email</label>
                <input type="email" value={newP.email||""} onChange={e=>setNewP(p=>({...p,email:e.target.value}))} placeholder="name@company.com" style={SEL} />
              </div>
              <div>
                <label style={{ fontSize:11, fontWeight:600, color:"#374151", display:"block", marginBottom:3 }}>Phone</label>
                <input value={newP.phone||""} onChange={e=>setNewP(p=>({...p,phone:e.target.value}))} placeholder="+60 12-345 6789" style={SEL} />
              </div>
            </div>
            <button onClick={addStaff} style={{ background:"#0D1B3E", color:"#fff", border:"none", borderRadius:8, padding:"10px 24px", fontWeight:700, fontSize:13, cursor:"pointer" }}>
              + Add Staff
            </button>
          </div>
          <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, overflow:"hidden" }}>
            <div style={{ padding:"10px 14px", borderBottom:"1px solid #f3f4f6", display:"flex", gap:10, alignItems:"center" }}>
              <input value={staffSearch} onChange={e=>setStaffSearch(e.target.value)} placeholder="Search name, role, entity..." style={{ flex:1, border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13 }} />
              <span style={{ fontSize:12, color:"#9ca3af", whiteSpace:"nowrap" }}>{filteredStaff.length} / {personnel.length}</span>
            </div>
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13, minWidth:700 }}>
                <thead style={{ background:"#0D1B3E" }}>
                  <tr>{["Name","Role","Entity","Region","Email","Phone","Status",""].map(h=><th key={h} style={{ padding:"9px 12px", textAlign:"left", color:"#C8EEF5", fontWeight:600, fontSize:11, whiteSpace:"nowrap" }}>{h}</th>)}</tr>
                </thead>
                <tbody>
                  {filteredStaff.length===0 && <tr><td colSpan={8} style={{ padding:28, textAlign:"center", color:"#9ca3af" }}>No staff yet. Add using the form above.</td></tr>}
                  {filteredStaff.map(p=>(
                    <tr key={p.id} style={{ borderTop:"1px solid #f3f4f6", background:p.active===false?"#fafafa":"#fff" }}>
                      {editId===p.id ? (
                        <>
                          <td style={{ padding:"6px 8px" }}><input value={editData.name!==undefined?editData.name:p.name} onChange={e=>setEditData(d=>({...d,name:e.target.value}))} style={{ width:"100%", border:"1px solid #d1d5db", borderRadius:4, padding:"5px 8px", fontSize:12 }} /></td>
                          <td style={{ padding:"6px 8px" }}><select value={editData.role!==undefined?editData.role:p.role||""} onChange={e=>setEditData(d=>({...d,role:e.target.value}))} style={SEL}><option value="">— Role —</option>{roles.map(r=><option key={r} value={r}>{r}</option>)}</select></td>
                          <td style={{ padding:"6px 8px" }}><select value={editData.entity!==undefined?editData.entity:p.entity||""} onChange={e=>setEditData(d=>({...d,entity:e.target.value}))} style={SEL}><option value="">— Entity —</option>{entities.map(e=><option key={e} value={e}>{e}</option>)}</select></td>
                          <td style={{ padding:"6px 8px" }}><select value={editData.region!==undefined?editData.region:p.region||""} onChange={e=>setEditData(d=>({...d,region:e.target.value}))} style={SEL}><option value="">— Region —</option>{regions.map(r=><option key={r} value={r}>{r}</option>)}</select></td>
                          <td style={{ padding:"6px 8px" }}><input type="email" value={editData.email!==undefined?editData.email:p.email||""} onChange={e=>setEditData(d=>({...d,email:e.target.value}))} style={{ width:"100%", border:"1px solid #d1d5db", borderRadius:4, padding:"5px 8px", fontSize:12 }} /></td>
                          <td style={{ padding:"6px 8px" }}><input value={editData.phone!==undefined?editData.phone:p.phone||""} onChange={e=>setEditData(d=>({...d,phone:e.target.value}))} style={{ width:"100%", border:"1px solid #d1d5db", borderRadius:4, padding:"5px 8px", fontSize:12 }} /></td>
                          <td style={{ padding:"6px 8px" }}><Badge label={p.active===false?"Inactive":"Active"} bg={p.active===false?"#9ca3af":"#16a34a"} /></td>
                          <td style={{ padding:"6px 8px" }}>
                            <div style={{ display:"flex", gap:4 }}>
                              <button onClick={()=>saveEdit(p.id)} style={{ background:"#16a34a", color:"#fff", border:"none", borderRadius:4, padding:"4px 10px", fontSize:12, cursor:"pointer", fontWeight:600 }}>Save</button>
                              <button onClick={()=>setEditId(null)} style={{ background:"#f3f4f6", color:"#374151", border:"none", borderRadius:4, padding:"4px 8px", fontSize:12, cursor:"pointer" }}>Cancel</button>
                            </div>
                          </td>
                        </>
                      ) : (
                        <>
                          <td style={{ padding:"8px 12px", fontWeight:600, color:p.active===false?"#9ca3af":"#0D1B3E" }}>{p.name}</td>
                          <td style={{ padding:"8px 12px", color:"#6b7280" }}>{p.role||"—"}</td>
                          <td style={{ padding:"8px 12px", color:"#6b7280" }}>{p.entity||"—"}</td>
                          <td style={{ padding:"8px 12px", color:"#6b7280" }}>{p.region||"—"}</td>
                          <td style={{ padding:"8px 12px", color:"#6b7280", fontSize:11 }}>{p.email||"—"}</td>
                          <td style={{ padding:"8px 12px", color:"#6b7280", fontSize:11 }}>{p.phone||"—"}</td>
                          <td style={{ padding:"8px 12px" }}><Badge label={p.active===false?"Inactive":"Active"} bg={p.active===false?"#9ca3af":"#16a34a"} /></td>
                          <td style={{ padding:"8px 12px" }}>
                            <div style={{ display:"flex", gap:4 }}>
                              <button onClick={()=>{ setEditId(p.id); setEditData({...p}); }} style={{ background:"#f0f9ff", color:"#0369a1", border:"1px solid #bae6fd", borderRadius:4, padding:"3px 8px", fontSize:11, cursor:"pointer" }}>Edit</button>
                              <button onClick={()=>toggleActive(p.id)} style={{ background:p.active===false?"#dcfce7":"#fff7ed", color:p.active===false?"#16a34a":"#d97706", border:"none", borderRadius:4, padding:"3px 8px", fontSize:11, cursor:"pointer" }}>{p.active===false?"Activate":"Deactivate"}</button>
                              <button onClick={()=>removePerson(p.id)} style={{ background:"#fee2e2", color:"#dc2626", border:"none", borderRadius:4, padding:"3px 8px", fontSize:11, cursor:"pointer" }}>Remove</button>
                            </div>
                          </td>
                        </>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ── ITEM MASTER ── */}
      {tab==="items" && <ItemMasterSetup itemMaster={itemMaster} setItemMaster={setItemMaster} />}

      {/* ── CUSTOMER MASTER ── */}
      {tab==="customers" && <CustomerMasterSetup customers={customers} setCustomers={setCustomers} />}

      {/* ── DROPDOWN LISTS ── */}
      {tab==="lists" && (
        <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
          <div style={{ background:"#f0f9ff", border:"1px solid #bae6fd", borderRadius:8, padding:"10px 14px", fontSize:13, color:"#0369a1" }}>
            These lists power all dropdowns across the app. One item per line.
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))", gap:14 }}>
            {[["entities","Entities / Divisions"],["regions","Regions / Offices"],["roles","Staff Roles"],["remarks","Remarks / Status Labels"]].map(([key,label])=>(
              <div key={key} style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:14 }}>
                <div style={{ fontWeight:700, fontSize:13, color:"#0D1B3E", marginBottom:6 }}>{label}</div>
                {savedListKey===key && <div style={{ background:"#dcfce7", color:"#166534", borderRadius:4, padding:"2px 8px", fontSize:11, marginBottom:6, display:"inline-block" }}>Saved</div>}
                <textarea value={listDrafts[key]} onChange={e=>setListDrafts(d=>({...d,[key]:e.target.value}))} rows={6}
                  style={{ width:"100%", border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13, resize:"vertical", fontFamily:"inherit" }} />
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginTop:8 }}>
                  <span style={{ fontSize:11, color:"#9ca3af" }}>{listDrafts[key].split("\n").filter(x=>x.trim()).length} items</span>
                  <button onClick={()=>saveList(key)} style={{ background:"#0D1B3E", color:"#fff", border:"none", borderRadius:6, padding:"6px 14px", fontWeight:700, fontSize:12, cursor:"pointer" }}>Save</button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ── SAP SNAPSHOTS ── */}
      {tab==="sap" && (
        <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
          <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:16 }}>
            <SectionHead title="Add / Update SAP Snapshot" />
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))", gap:10, marginBottom:12 }}>
              <div>
                <label style={{ fontSize:11, fontWeight:600, color:"#374151", display:"block", marginBottom:3 }}>Month</label>
                <input type="month" value={snapForm.month} onChange={e=>setSnapForm(f=>({...f,month:e.target.value}))} style={SEL} />
              </div>
              <div>
                <label style={{ fontSize:11, fontWeight:600, color:"#374151", display:"block", marginBottom:3 }}>Item</label>
                <select value={snapForm.itemId} onChange={e=>setSnapForm(f=>({...f,itemId:e.target.value}))} style={SEL}>
                  <option value="">— Select —</option>
                  {(itemMaster||DEFAULT_ITEMS).filter(i=>i.active!==false).map(i=><option key={i.id} value={i.id}>[{i.sapCode}] {i.name}</option>)}
                </select>
              </div>
              {[["SAP Stock","sapStock"],["SAP GIT","sapGit"],["SAP Demo","sapDemo"]].map(([lbl,key])=>(
                <div key={key}>
                  <label style={{ fontSize:11, fontWeight:600, color:"#374151", display:"block", marginBottom:3 }}>{lbl}</label>
                  <input type="number" min="0" value={snapForm[key]} onChange={e=>setSnapForm(f=>({...f,[key]:e.target.value}))} style={SEL} />
                </div>
              ))}
            </div>
            <button onClick={addSnapshot} style={{ background:"#0D1B3E", color:"#fff", border:"none", borderRadius:8, padding:"9px 22px", fontWeight:700, fontSize:13, cursor:"pointer" }}>Save Snapshot</button>
          </div>
          <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, overflow:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
              <thead style={{ background:"#0D1B3E" }}><tr>{["Month","Item","SAP Stock","SAP GIT","SAP Demo",""].map(h=><th key={h} style={{ padding:"9px 12px", textAlign:"left", color:"#C8EEF5", fontWeight:600, fontSize:11 }}>{h}</th>)}</tr></thead>
              <tbody>
                {sapSnapshots.length===0 && <tr><td colSpan={6} style={{ padding:24, textAlign:"center", color:"#9ca3af" }}>No snapshots. Import via Import SAP or add manually above.</td></tr>}
                {sapSnapshots.map(snap=>(
                  <tr key={snap.id} style={{ borderTop:"1px solid #f3f4f6" }}>
                    <td style={{ padding:"8px 12px", fontWeight:600 }}>{snap.month}</td>
                    <td style={{ padding:"8px 12px" }}>{getItemName(snap.itemId)}</td>
                    <td style={{ padding:"8px 12px" }}>{snap.sapStock||0}</td>
                    <td style={{ padding:"8px 12px" }}>{snap.sapGit||0}</td>
                    <td style={{ padding:"8px 12px" }}>{snap.sapDemo||0}</td>
                    <td style={{ padding:"8px 12px" }}><button onClick={()=>delSnapshot(snap.id)} style={{ background:"#fee2e2", color:"#dc2626", border:"none", borderRadius:4, padding:"3px 8px", fontSize:11, cursor:"pointer" }}>Delete</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* ── MONTH LOCK ── */}
      {tab==="lock" && (
        <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:16 }}>
          <SectionHead title="Month Lock" />
          <div style={{ fontSize:13, color:"#6b7280", marginBottom:14 }}>Locked months prevent new movements. SAP Import is always exempt.</div>
          <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
            {months.map(m=>{
              const locked=(lockedMonths||[]).includes(m);
              return (
                <button key={m} onClick={()=>toggleLock(m)} style={{ padding:"10px 18px", borderRadius:8, border:"2px solid "+(locked?"#dc2626":"#e5e7eb"), background:locked?"#fee2e2":"#fff", color:locked?"#dc2626":"#374151", fontWeight:600, fontSize:13, cursor:"pointer" }}>
                  {m}{locked?" 🔒":""}
                </button>
              );
            })}
          </div>
        </div>
      )}

      {/* ── BACKUP & RESTORE ── */}
      {tab==="backup" && (
        <div style={{ display:"flex", flexDirection:"column", gap:20, maxWidth:680 }}>
          <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:20 }}>
            <SectionHead title="Download Backup" />
            <div style={{ fontSize:13, color:"#6b7280", marginBottom:14 }}>Exports all data as a JSON file. Available to all admins.</div>
            <button onClick={doBackup} style={{ background:"#0D1B3E", color:"#fff", border:"none", borderRadius:8, padding:"12px 28px", fontWeight:700, fontSize:14, cursor:"pointer" }}>
              Download Backup
            </button>
            {backupDone && <div style={{ marginTop:10, color:"#16a34a", fontWeight:600, fontSize:13 }}>Backup downloaded successfully.</div>}
          </div>
          <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:20 }}>
            <SectionHead title="Restore from Backup" />
            {!isSuperAdmin ? (
              <div style={{ background:"#fef3c7", border:"1px solid #fbbf24", borderRadius:8, padding:"12px 16px", fontSize:13, color:"#92400e", fontWeight:500 }}>
                Restore is restricted to Super Admin only. Logistics Admin can download backups but cannot restore.
              </div>
            ) : (
              <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
                <div style={{ background:"#fee2e2", border:"1px solid #fca5a5", borderRadius:8, padding:"10px 14px", fontSize:13, color:"#991b1b" }}>
                  Warning: Restoring will overwrite ALL current data. Download a backup first.
                </div>
                <div>
                  <label style={{ fontSize:12, fontWeight:600, color:"#374151", display:"block", marginBottom:6 }}>Select Backup File (.json)</label>
                  <input type="file" accept=".json" onChange={handleRestoreFile} style={{ fontSize:13 }} />
                </div>
                {restoreStatus==="error" && <div style={{ background:"#fee2e2", color:"#991b1b", borderRadius:8, padding:10, fontSize:13 }}>{restoreError}</div>}
                {restoreStatus==="done"  && <div style={{ background:"#dcfce7", color:"#166534", borderRadius:8, padding:10, fontSize:13, fontWeight:600 }}>Data restored successfully.</div>}
                {restoreStatus==="preview" && restorePreview && (
                  <div style={{ background:"#f0f9ff", border:"1px solid #bae6fd", borderRadius:10, padding:16 }}>
                    <div style={{ fontWeight:700, color:"#0369a1", marginBottom:10 }}>Backup Preview</div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:14 }}>
                      {[
                        ["Exported",restorePreview.exportedAt ? new Date(restorePreview.exportedAt).toLocaleString() : "Unknown"],
                        ["Exported By",restorePreview.exportedBy||"Unknown"],
                        ["Movements",(restorePreview.data ? restorePreview.data.movements||[] : []).length+" records"],
                        ["Staff",(restorePreview.data ? restorePreview.data.personnel||[] : []).length+" records"],
                        ["Fleet Units",(restorePreview.data ? restorePreview.data.fleetUnits||[] : []).length+" units"],
                        ["GIT Items",(restorePreview.data ? restorePreview.data.gitItems||[] : []).length+" records"],
                      ].map(([lbl,val])=>(
                        <div key={lbl} style={{ background:"#fff", borderRadius:6, padding:"8px 12px", border:"1px solid #e0f2fe" }}>
                          <div style={{ fontSize:10, color:"#0369a1", fontWeight:600, marginBottom:2 }}>{lbl}</div>
                          <div style={{ fontSize:13, fontWeight:600, color:"#0D1B3E" }}>{val}</div>
                        </div>
                      ))}
                    </div>
                    <div style={{ display:"flex", gap:10 }}>
                      <button onClick={confirmRestore} style={{ background:"#dc2626", color:"#fff", border:"none", borderRadius:8, padding:"10px 24px", fontWeight:700, fontSize:13, cursor:"pointer" }}>
                        Confirm Restore (Overwrites All Data)
                      </button>
                      <button onClick={()=>{ setRestoreStatus("idle"); setRestorePreview(null); }} style={{ background:"#f3f4f6", color:"#374151", border:"none", borderRadius:8, padding:"10px 20px", fontWeight:600, fontSize:13, cursor:"pointer" }}>
                        Cancel
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

function SapImport({ sapSnapshots, setSapSnapshots, gitItems, setGitItems }) {
  const [status, setStatus] = useState("idle"); // idle | parsing | preview | done | error
  const [errorMsg, setErrorMsg] = useState("");
  const [preview, setPreview] = useState({ snapshots: [], git: [] });
  const [fileName, setFileName] = useState("");

  // Valid pudu model IDs for validation
  const validIds = new Set(PUDU_ITEMS.map(i => i.id));

  function reset() { setStatus("idle"); setPreview({ snapshots: [], git: [] }); setFileName(""); setErrorMsg(""); }

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setStatus("parsing");
    setErrorMsg("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      // ── Sheet 1: SAP_Stock_Snapshot ──────────────────────────────
      const snapshots = [];
      const snSheet = wb.Sheets["SAP_Stock_Snapshot"];
      if (!snSheet) throw new Error("Sheet 'SAP_Stock_Snapshot' not found. Check sheet name.");
      const snRows = XLSX.utils.sheet_to_json(snSheet, { header: 1, defval: "" });
      // Find header row (row with "month" or "Month")
      let snHdrIdx = snRows.findIndex(r => r.some(c => String(c).toLowerCase().includes("month") && r.some(x => String(x).toLowerCase().includes("pudu"))));
      if (snHdrIdx === -1) snHdrIdx = 2; // fallback to row index 2 (row 3)
      const snHdr = snRows[snHdrIdx].map(h => String(h).toLowerCase().split("\n").join(" ").trim());

      const snCol = (name) => snHdr.findIndex(h => h.includes(name));
      const iMonth = snCol("month"), iPudu = snCol("pudu model"), iStock = snCol("sap stock"), iGit = snCol("sap git"), iDemo = snCol("sap demo"), iNotes = snCol("remark");
      if (iMonth === -1 || iPudu === -1 || iStock === -1) throw new Error("Required columns not found in SAP_Stock_Snapshot. Ensure Month, Pudu Model Code, SAP Stock columns exist.");
      for (let r = snHdrIdx + 1; r < snRows.length; r++) {
        const row = snRows[r];
        const month = String(row[iMonth] || "").trim();
        const puduCode = String(row[iPudu] || "").trim().toUpperCase();
        const stock = parseInt(row[iStock]) || 0;
        const git = iGit >= 0 ? (parseInt(row[iGit]) || 0) : 0;
        const demo = iDemo >= 0 ? (parseInt(row[iDemo]) || 0) : 0;
        const notes = iNotes >= 0 ? String(row[iNotes] || "").trim() : "";
        if (!month || !puduCode || month.length !== 7 || month[4] !== '-') continue;
        const warning = !validIds.has(puduCode) ? "⚠ Model code not matched — will still import" : "";
        snapshots.push({ month, itemId: puduCode, sapStock: stock, sapGit: git, sapDemo: demo, notes, warning });
      }

      // ── Sheet 2: SAP_GIT_Register ────────────────────────────────
      const gitRows = [];
      const gitSheet = wb.Sheets["SAP_GIT_Register"];
      if (gitSheet) {
        const gitData = XLSX.utils.sheet_to_json(gitSheet, { header: 1, defval: "" });
        let gitHdrIdx = gitData.findIndex(r => r.some(c => String(c).toLowerCase().includes("po")));
        if (gitHdrIdx === -1) gitHdrIdx = 2;
        const gitHdr = gitData[gitHdrIdx].map(h => String(h).toLowerCase().split("\n").join(" ").trim());

        const gc = (name) => gitHdr.findIndex(h => h.includes(name));
        const iPo = gc("po"), iPuduG = gc("pudu model"), iQtyO = gc("qty ordered"), iQtyR = gc("qty received"), iSupp = gc("supplier"), iVessel = gc("vessel"), iEta = gc("expected"), iStat = gc("status"), iDesc = gc("material desc");
        for (let r = gitHdrIdx + 1; r < gitData.length; r++) {
          const row = gitData[r];
          const poNo = String(row[iPo] || "").trim();
          const puduCode = String(row[iPuduG >= 0 ? iPuduG : 0] || "").trim().toUpperCase();
          const qtyO = parseInt(row[iQtyO >= 0 ? iQtyO : 0]) || 0;
          const qtyR = iQtyR >= 0 ? (parseInt(row[iQtyR]) || 0) : 0;
          const supplier = iSupp >= 0 ? String(row[iSupp] || "").trim() : "";
          const vessel = iVessel >= 0 ? String(row[iVessel] || "").trim() : "";
          const eta = iEta >= 0 ? String(row[iEta] || "").trim() : "";
          const gitStatus = iStat >= 0 ? String(row[iStat] || "On Order").trim() : "On Order";
          const desc = iDesc >= 0 ? String(row[iDesc] || "").trim() : "";
          if (!poNo || !puduCode || qtyO === 0) continue;
          if (gitStatus === "Received" && qtyO === qtyR) continue; // skip fully received
          const warning = !validIds.has(puduCode) ? "⚠ Model code not matched" : "";
          gitRows.push({ poNo, itemId: puduCode, qty: qtyO - qtyR, supplier, vesselRef: vessel, etaDate: eta, status: gitStatus, notes: desc, warning, shipDate: "" });
        }
      }

      if (snapshots.length === 0 && gitRows.length === 0) throw new Error("No valid data rows found. Check that data starts from row 4 and Month is in YYYY-MM format.");
      setPreview({ snapshots, git: gitRows });
      setStatus("preview");
    } catch (err) {
      setErrorMsg(err.message || "Unknown error reading file.");
      setStatus("error");
    }
    e.target.value = "";
  }

  async function confirmImport() {
    // Merge snapshots: update existing month+item or add new
    let newSnaps = [...sapSnapshots];
    preview.snapshots.forEach(s => {
      const idx = newSnaps.findIndex(x => x.month === s.month && x.itemId === s.itemId);
      const entry = { id: idx >= 0 ? newSnaps[idx].id : Date.now() + Math.random(), month: s.month, itemId: s.itemId, sapStock: s.sapStock, sapGit: s.sapGit, sapDemo: s.sapDemo, notes: s.notes };
      if (idx >= 0) newSnaps[idx] = entry; else newSnaps = [entry, ...newSnaps];
    });
    setSapSnapshots(newSnaps);
    await save(SK.SAP, newSnaps);

    // Merge GIT: match by PO number, update or add
    let newGit = [...gitItems];
    preview.git.forEach(g => {
      const idx = newGit.findIndex(x => x.poNo === g.poNo && x.itemId === g.itemId);
      const entry = { id: idx >= 0 ? newGit[idx].id : Date.now() + Math.random(), ...g };
      if (idx >= 0) newGit[idx] = entry; else newGit = [entry, ...newGit];
    });
    setGitItems(newGit);
    await save(SK.GIT, newGit);
    setStatus("done");
  }

  const warnColor = "#92400e", warnBg = "#fef3c7";

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16, maxWidth: 900 }}>
      {/* Header */}
      <div style={{ background: "linear-gradient(135deg,#0D1B3E,#1a3a6b)", borderRadius: 10, padding: "16px 20px", color: "#fff" }}>
        <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 4 }}>Import from SAP</div>
        <div style={{ fontSize: 12, color: "#C8EEF5" }}>Upload your SAP inventory Excel file — Stock Snapshot and GIT Register update automatically. Demo movements are NOT affected.</div>
      </div>

      {/* What this imports */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
        {[
          ["📦 Stock Snapshot", "Updates Stock Position tab with SAP Stock, GIT, Demo qty per model per month", "#16a34a", "#dcfce7"],
          ["🚢 GIT Register", "Adds or updates open PO lines from SAP. Fully received POs are skipped.", "#d97706", "#fef3c7"],
          ["✋ Demo Unchanged", "Demo unit movements (loan, POC, exhibition) are NOT touched — those stay manual only.", "#4AACCC", "#f0f9ff"],
        ].map(([title, desc, col, bg]) => (
          <div key={title} style={{ background: bg, border: `1px solid ${col}40`, borderRadius: 8, padding: 12 }}>
            <div style={{ fontWeight: 700, fontSize: 13, color: col, marginBottom: 4 }}>{title}</div>
            <div style={{ fontSize: 12, color: "#374151" }}>{desc}</div>
          </div>
        ))}
      </div>

      {/* Step 1 — Download template */}
      <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
        <SectionHead title="Step 1 — Download the SAP Import Template" />
        <p style={{ fontSize: 13, color: "#374151", margin: "0 0 12px" }}>
          Download the template, fill in your SAP data (or replace the dummy rows with a paste from your SAP report), then upload below.
        </p>
        <div style={{ background: "#f9fafb", border: "1px dashed #d1d5db", borderRadius: 8, padding: 14, fontSize: 12, color: "#6b7280" }}>
          <strong style={{ color: "#0D1B3E" }}>Template contains 2 sheets:</strong>
          <ul style={{ margin: "6px 0 0", paddingLeft: 18 }}>
            <li><strong>SAP_Stock_Snapshot</strong> — Month-end closing stock per model (Stock / GIT / Demo qty). Maps to MB52 or your inventory audit report.</li>
            <li><strong>SAP_GIT_Register</strong> — Open purchase orders in transit. Maps to ME2M or ME2N open PO list.</li>
            <li><strong>INSTRUCTIONS</strong> — Column mapping guide and Pudu model code reference.</li>
          </ul>
          <div style={{ marginTop: 10, padding: "8px 12px", background: warnBg, borderRadius: 6, color: warnColor }}>
            <strong>Important:</strong> Do NOT rename the sheet tabs or column headers in Row 3. The app reads them exactly as named. Only replace the data rows.
          </div>
        </div>
      </div>

      {/* Step 2 — Upload */}
      <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
        <SectionHead title="Step 2 — Upload Your Filled SAP File" />
        {status === "idle" || status === "error" ? (
          <div>
            <label style={{
              display: "inline-flex", alignItems: "center", gap: 8,
              background: "#0D1B3E", color: "#fff", borderRadius: 8,
              padding: "11px 24px", fontWeight: 600, fontSize: 14, cursor: "pointer"
            }}>
              📂 Choose SAP Excel File (.xlsx)
              <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
            {status === "error" && (
              <div style={{ marginTop: 12, background: "#fee2e2", border: "1px solid #fca5a5", borderRadius: 8, padding: "10px 14px", color: "#991b1b", fontSize: 13 }}>
                <strong>Error reading file:</strong> {errorMsg}
              </div>
            )}
          </div>
        ) : status === "parsing" ? (
          <div style={{ color: "#6b7280", fontSize: 13 }}>Reading file: {fileName}...</div>
        ) : status === "done" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ background: "#dcfce7", border: "1px solid #86efac", borderRadius: 8, padding: "12px 16px", color: "#166534", fontWeight: 600, fontSize: 14 }}>
              Import complete — {preview.snapshots.length} stock snapshot rows and {preview.git.length} GIT rows imported successfully.
            </div>
            <button onClick={reset} style={{ alignSelf: "flex-start", background: "#f9fafb", border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 16px", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
              Import Another File
            </button>
          </div>
        ) : null}
      </div>

      {/* Preview */}
      {status === "preview" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <div style={{ background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 10, padding: 14, display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
            <div>
              <div style={{ fontWeight: 700, fontSize: 14, color: "#0D1B3E" }}>Preview — {fileName}</div>
              <div style={{ fontSize: 12, color: "#6b7280", marginTop: 2 }}>
                {preview.snapshots.length} stock rows · {preview.git.length} GIT rows detected. Review below then confirm.
              </div>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={reset} style={{ background: "#fff", border: "1px solid #d1d5db", borderRadius: 6, padding: "9px 16px", fontWeight: 600, fontSize: 13, cursor: "pointer", color: "#374151" }}>
                Cancel
              </button>
              <button onClick={confirmImport} style={{ background: "linear-gradient(135deg,#166534,#16a34a)", color: "#fff", border: "none", borderRadius: 6, padding: "9px 20px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                Confirm Import
              </button>
            </div>
          </div>

          {/* Stock Snapshot Preview */}
          {preview.snapshots.length > 0 && (
            <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 14 }}>
              <SectionHead title={"Stock Snapshot — " + preview.snapshots.length + " rows"} />
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ background: "#f9fafb" }}>
                    <tr>{["Month","Pudu Model Code","Item Name","SAP Stock","SAP GIT","SAP Demo","Notes","Status"].map(h=>(
                      <th key={h} style={{ padding:"7px 10px",textAlign:"left",fontWeight:600,color:"#374151",borderBottom:"1px solid #e5e7eb",whiteSpace:"nowrap" }}>{h}</th>
                    ))}</tr>
                  </thead>
                  <tbody>
                    {preview.snapshots.map((s, i) => (
                      <tr key={i} style={{ borderTop:"1px solid #f3f4f6", background: s.warning ? warnBg : "transparent" }}>
                        <td style={{ padding:"6px 10px",fontWeight:600 }}>{s.month}</td>
                        <td style={{ padding:"6px 10px",fontFamily:"monospace",fontSize:11 }}>{s.itemId}</td>
                        <td style={{ padding:"6px 10px" }}>{getItemName(s.itemId)}</td>
                        <td style={{ padding:"6px 10px",textAlign:"center",color:"#16a34a",fontWeight:700 }}>{s.sapStock}</td>
                        <td style={{ padding:"6px 10px",textAlign:"center",color:"#d97706",fontWeight:700 }}>{s.sapGit}</td>
                        <td style={{ padding:"6px 10px",textAlign:"center",color:"#7c3aed",fontWeight:700 }}>{s.sapDemo}</td>
                        <td style={{ padding:"6px 10px",color:"#6b7280",fontSize:11 }}>{s.notes}</td>
                        <td style={{ padding:"6px 10px" }}>
                          {s.warning
                            ? <Badge label="No match" bg="#d97706" />
                            : <Badge label="Ready" bg="#16a34a" />}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* GIT Preview */}
          {preview.git.length > 0 && (
            <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 14 }}>
              <SectionHead title={"GIT Register — " + preview.git.length + " rows"} />
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead style={{ background: "#fef3c7" }}>
                    <tr>{["PO No.","Item","Qty","Supplier","Vessel Ref","ETA","Status",""].map(h=>(
                      <th key={h} style={{ padding:"7px 10px",textAlign:"left",fontWeight:600,color:"#92400e",borderBottom:"1px solid #fde68a",whiteSpace:"nowrap" }}>{h}</th>
                    ))}</tr>
                  </thead>
                  <tbody>
                    {preview.git.map((g, i) => {
                      const sc = { "On Order":"#7c3aed","Shipped":"#d97706","Customs":"#dc2626","Received":"#16a34a" };
                      return (
                        <tr key={i} style={{ borderTop:"1px solid #f3f4f6", background: g.warning ? warnBg : "transparent" }}>
                          <td style={{ padding:"6px 10px",fontWeight:600 }}>{g.poNo}</td>
                          <td style={{ padding:"6px 10px" }}>{getItemName(g.itemId)}</td>
                          <td style={{ padding:"6px 10px",textAlign:"center",fontWeight:700 }}>{g.qty}</td>
                          <td style={{ padding:"6px 10px",color:"#6b7280" }}>{g.supplier}</td>
                          <td style={{ padding:"6px 10px",fontSize:11,color:"#6b7280" }}>{g.vesselRef||"—"}</td>
                          <td style={{ padding:"6px 10px" }}>{g.etaDate||"—"}</td>
                          <td style={{ padding:"6px 10px" }}><Badge label={g.status} bg={sc[g.status]||"#9ca3af"} /></td>
                          <td style={{ padding:"6px 10px" }}>{g.warning ? <Badge label="No match" bg="#d97706" /> : <Badge label="Ready" bg="#16a34a" />}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}


// EXPORT EXCEL
function ExportExcel({ movements, personnel, sapSnapshots, gitItems }) {
  const [filters, setFilters] = useState({ dateFrom:"",dateTo:"",direction:"ALL",movType:"ALL",region:"ALL",custodian:"ALL",status:"ALL",itemCategory:"ALL",sapCat:"ALL" });
  const [exported, setExported] = useState(false);
  function setF(k,v){setFilters(f=>({...f,[k]:v}));}

  function applyFilters(data){
    return data.filter(m=>{
      if(filters.dateFrom && m.date<filters.dateFrom) return false;
      if(filters.dateTo && m.date>filters.dateTo) return false;
      if(filters.direction!=="ALL" && getDir(m.type)!==filters.direction) return false;
      if(filters.movType!=="ALL" && m.type!==filters.movType) return false;
      if(filters.region!=="ALL" && m.region!==filters.region) return false;
      if(filters.custodian!=="ALL" && m.custodian!==filters.custodian) return false;
      if(filters.status!=="ALL" && (m.status||"Active")!==filters.status) return false;
      if(filters.itemCategory!=="ALL" && getItemCat(m.itemId)!==filters.itemCategory) return false;
      if(filters.sapCat!=="ALL" && m.sapCat!==filters.sapCat) return false;
      return true;
    });
  }

  function styleWs(ws, hdrColor){
    const range=XLSX.utils.decode_range(ws["!ref"]||"A1");
    const cols=[];
    for(let C=range.s.c;C<=range.e.c;C++){
      const hCell=ws[XLSX.utils.encode_cell({r:0,c:C})];
      if(hCell) hCell.s={font:{bold:true,color:{rgb:"FFFFFF"}},fill:{fgColor:{rgb:hdrColor}},alignment:{horizontal:"center",wrapText:true}};
      let max=10;
      for(let R=0;R<=range.e.r;R++){const cell=ws[XLSX.utils.encode_cell({r:R,c:C})];if(cell?.v)max=Math.min(40,Math.max(max,String(cell.v).length+2));}
      cols.push({wch:max});
    }
    ws["!cols"]=cols;
    return ws;
  }

  function doExport(){
    const filtered=applyFilters(movements);
    if(filtered.length===0){alert("No records match filters.");return;}
    const wb=XLSX.utils.book_new();

    const detailRows=filtered.map(m=>({
      "Date":m.date,"Direction":getDir(m.type),"SAP Category":m.sapCat||"",
      "Movement Type":getMovLabel(m.type),"Item":getItemName(m.itemId),
      "Item Category":getItemCat(m.itemId),"Item Code":m.itemId,
      "Serial No":m.serialNo||"","Qty":m.qty,"Custodian":m.custodian||"",
      "Region":m.region||"","Customer/Event":m.customerName||m.eventName||"",
      "Expected Return":m.expectedReturn||"","Actual Return":m.actualReturn||"",
      "Remark":m.remark||"","Location":m.location||"","Item Description":m.itemDesc||"","Status":isOverdue(m)?"Overdue":(m.status||"Active"),"Notes":m.notes||"",
    }));
    XLSX.utils.book_append_sheet(wb,styleWs(XLSX.utils.json_to_sheet(detailRows),"0D1B3E"),"Movement Detail");

    const byCat={};
    filtered.forEach(m=>{const k=(m.sapCat||"Unknown")+"_"+getDir(m.type);if(!byCat[k])byCat[k]={"SAP Category":m.sapCat,"Direction":getDir(m.type),"Transactions":0,"Total Qty":0};byCat[k]["Transactions"]++;byCat[k]["Total Qty"]+=m.qty;});
    XLSX.utils.book_append_sheet(wb,styleWs(XLSX.utils.json_to_sheet(Object.values(byCat)),"7c3aed"),"By SAP Category");

    const byType={};
    filtered.forEach(m=>{if(!byType[m.type])byType[m.type]={"Direction":getDir(m.type),"SAP Category":m.sapCat||"","Movement Type":getMovLabel(m.type),"Transactions":0,"Total Qty":0};byType[m.type]["Transactions"]++;byType[m.type]["Total Qty"]+=m.qty;});
    XLSX.utils.book_append_sheet(wb,styleWs(XLSX.utils.json_to_sheet(Object.values(byType)),"166534"),"By Movement Type");

    const byModel={};
    filtered.forEach(m=>{if(!byModel[m.itemId])byModel[m.itemId]={"Item":getItemName(m.itemId),"Category":getItemCat(m.itemId),"SAP Category":m.sapCat||"","GR Qty":0,"GI Qty":0};if(getDir(m.type)==="GR")byModel[m.itemId]["GR Qty"]+=m.qty;else byModel[m.itemId]["GI Qty"]+=m.qty;});
    const modelRows=Object.values(byModel).map(r=>({...r,"Net Movement":r["GR Qty"]-r["GI Qty"]}));
    XLSX.utils.book_append_sheet(wb,styleWs(XLSX.utils.json_to_sheet(modelRows),"1a3a6b"),"By Model");

    const byPerson={};
    filtered.forEach(m=>{const p=personnel.find(x=>x.name===m.custodian);if(!byPerson[m.custodian])byPerson[m.custodian]={"Name":m.custodian,"Role":p?.role||"","Region":p?.region||"","Total":0,"GR":0,"GI":0,"In Field":0,"Overdue":0};byPerson[m.custodian]["Total"]++;if(getDir(m.type)==="GR")byPerson[m.custodian]["GR"]++;else{byPerson[m.custodian]["GI"]++;if(m.status!=="Returned")byPerson[m.custodian]["In Field"]++;if(isOverdue(m))byPerson[m.custodian]["Overdue"]++;}});
    XLSX.utils.book_append_sheet(wb,styleWs(XLSX.utils.json_to_sheet(Object.values(byPerson)),"E07820"),"Personnel Activity");

    const overdueRows=movements.filter(m=>isOverdue(m)).map(m=>({
      "Item":getItemName(m.itemId),"SAP Category":m.sapCat||"","Serial No":m.serialNo||"",
      "Custodian":m.custodian,"Region":m.region,"Date Out":m.date,
      "Expected Return":m.expectedReturn,"Days Overdue":Math.floor((new Date()-new Date(m.expectedReturn))/86400000),
      "Customer/Event":m.customerName||m.eventName||"","Notes":m.notes||"",
    }));
    if(overdueRows.length>0) XLSX.utils.book_append_sheet(wb,styleWs(XLSX.utils.json_to_sheet(overdueRows),"991b1b"),"Overdue Returns");

    const snapMonth=filters.dateFrom?filters.dateFrom.slice(0,7):new Date().toISOString().slice(0,7);
    const snapRows=sapSnapshots.filter(s=>s.month===snapMonth).map(s=>{
      const inField=movements.filter(m=>m.itemId===s.itemId&&m.sapCat==="DEMO"&&getDir(m.type)==="GI"&&m.status!=="Returned").reduce((sum,m)=>sum+m.qty,0);
      return{"Month":s.month,"Item":getItemName(s.itemId),"SAP Stock":s.sapStock,"SAP GIT":s.sapGit,"SAP Demo":s.sapDemo,"Tracker Demo In Field":inField,"Effective Available":s.sapStock-inField,"Notes":s.notes};
    });
    if(snapRows.length>0) XLSX.utils.book_append_sheet(wb,styleWs(XLSX.utils.json_to_sheet(snapRows),"16a34a"),"Stock Position");

    const gitRows=gitItems.map(g=>({
      "PO No":g.poNo,"Item":getItemName(g.itemId),"Qty":g.qty,"Supplier":g.supplier,
      "Vessel Ref":g.vesselRef,"Ship Date":g.shipDate,"ETA":g.etaDate,"Status":g.status,"Notes":g.notes,
    }));
    if(gitRows.length>0) XLSX.utils.book_append_sheet(wb,styleWs(XLSX.utils.json_to_sheet(gitRows),"d97706"),"GIT Register");

    XLSX.utils.book_append_sheet(wb,styleWs(XLSX.utils.json_to_sheet([
      {"Filter":"Date From","Value":filters.dateFrom||"All"},
      {"Filter":"Date To","Value":filters.dateTo||"All"},
      {"Filter":"Direction","Value":filters.direction},
      {"Filter":"SAP Category","Value":filters.sapCat},
      {"Filter":"Movement Type","Value":filters.movType==="ALL"?"All":getMovLabel(filters.movType)},
      {"Filter":"Region","Value":filters.region},
      {"Filter":"Custodian","Value":filters.custodian},
      {"Filter":"Records Exported","Value":filtered.length},
      {"Filter":"Exported On","Value":new Date().toLocaleDateString("en-MY")},
    ]),"4AACCC"),"Export Criteria");

    XLSX.writeFile(wb,"PuduStock_"+new Date().toISOString().slice(0,10)+".xlsx");
    setExported(true);setTimeout(()=>setExported(false),4000);
  }

  const previewCount=applyFilters(movements).length;
  const uniqRegions=[...new Set(movements.map(m=>m.region).filter(Boolean))].sort();
  const uniqCust=[...new Set(movements.map(m=>m.custodian).filter(Boolean))].sort();
  const sel=(label,key,opts)=>(
    <div style={{ display:"flex",flexDirection:"column",gap:4 }}>
      <label style={{ fontSize:12,fontWeight:600,color:"#374151" }}>{label}</label>
      <select value={filters[key]} onChange={e=>setF(key,e.target.value)} style={{ border:"1px solid #d1d5db",borderRadius:6,padding:"7px 10px",fontSize:13 }}>
        {opts.map(([v,l])=><option key={v} value={v}>{l}</option>)}
      </select>
    </div>
  );

  return(
    <div style={{ maxWidth:760,display:"flex",flexDirection:"column",gap:18 }}>
      {exported && <div style={{ background:"#dcfce7",border:"1px solid #86efac",borderRadius:8,padding:"10px 16px",color:"#166534",fontWeight:600,fontSize:13 }}>Excel downloaded successfully.</div>}
      <div style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:18 }}>
        <SectionHead title="Filter Criteria — leave blank to export all" />
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12 }}>
          {[["dateFrom","Date From","date"],["dateTo","Date To","date"]].map(([k,l,t])=>(
            <div key={k} style={{ display:"flex",flexDirection:"column",gap:4 }}>
              <label style={{ fontSize:12,fontWeight:600,color:"#374151" }}>{l}</label>
              <input type={t} value={filters[k]} onChange={e=>setF(k,e.target.value)} style={{ border:"1px solid #d1d5db",borderRadius:6,padding:"7px 10px",fontSize:13 }} />
            </div>
          ))}
          {sel("SAP Category","sapCat",[["ALL","All Categories"],...SAP_STOCK_CATEGORIES.map(c=>[c.id,c.label+" - "+c.desc])])}
          {sel("Direction","direction",[["ALL","All"],["GR","GR - In"],["GI","GI - Out"]])}
          {sel("Movement Type","movType",[["ALL","All Types"],...ALL_MOVEMENT_TYPES.map(t=>[t.value,t.label])])}
          {sel("Region","region",[["ALL","All Regions"],...uniqRegions.map(r=>[r,r])])}
          {sel("Custodian","custodian",[["ALL","All Personnel"],...uniqCust.map(c=>[c,c])])}
          {sel("Status","status",[["ALL","All"],["Active","Active"],["Returned","Returned"]])}
          {sel("Item Category","itemCategory",[["ALL","All"],["Robot","Robot"],["Accessory","Accessory"]])}
        </div>
      </div>
      <div style={{ background:"#f0f9ff",border:"1px solid #bae6fd",borderRadius:10,padding:14,display:"flex",alignItems:"center",justifyContent:"space-between" }}>
        <div>
          <div style={{ fontWeight:700,fontSize:15,color:"#0D1B3E" }}>{previewCount} records matched</div>
          <div style={{ fontSize:12,color:"#6b7280",marginTop:2 }}>Up to 9 sheets: Detail, SAP Category, By Type, By Model, Personnel, Overdue, Stock Position, GIT Register, Criteria</div>
        </div>
        <button onClick={doExport} style={{ background:previewCount===0?"#e5e7eb":"linear-gradient(135deg,#166534,#16a34a)",color:previewCount===0?"#9ca3af":"#fff",border:"none",borderRadius:8,padding:"12px 28px",fontWeight:700,fontSize:14,cursor:previewCount===0?"not-allowed":"pointer" }}>
          Download Excel
        </button>
      </div>
    </div>
  );
}

// USER GUIDE
function UserGuide(){
  const howtoSteps=[
    {who:"Admin - One-Time Setup",color:"#0D1B3E",bg:"#e8edf7",steps:["Go to Admin tab, then Personnel sub-tab","Add each team member: name, role (Sales/Engineer/Marketing/Logistics Admin), and region","Delete sample names (Ahmad Farid etc.) once your real team is added","After each SAP month-end close, go to Admin then SAP Snapshots and enter the closing stock per item: SAP Stock, GIT qty, and Demo qty - this feeds the Stock Position tab"]},
    {who:"Sales Personnel - Issue Unit OUT",color:"#991b1b",bg:"#fff1f2",steps:["Click Log Movement tab","Click the Goods Issue (OUT) button","Select Demo Out - Sales Personnel as the movement type","The SAP Category auto-fills as DEMO","Select the Robot model from the Item dropdown","Fill in Serial No (e.g. PD-BB-2026-001), Qty as 1, Date, and Expected Return date","Select your name from Custodian dropdown and your Region","Add brief Notes such as Demo at Pavilion KL - F&B client","Click Log Movement - done","When you return the unit: Log Movement - Goods Receipt IN - Return Demo by Sales Personnel using the same serial number"]},
    {who:"Engineer - POC Unit OUT and Return",color:"#166534",bg:"#f0fdf4",steps:["Click Log Movement tab","Click Goods Issue (OUT)","Select POC Out - Engineer","SAP Category auto-fills as DEMO","Fill in Customer Name (the company you are visiting for POC)","Fill Serial No, Date, Expected Return","Click Log Movement","To return: Log Movement - Goods Receipt IN - Return POC by Engineer - same serial number"]},
    {who:"Marketing - Exhibition OUT and Return",color:"#7c3aed",bg:"#f5f3ff",steps:["Click Log Movement tab","Click Goods Issue (OUT)","Select Exhibition Out - Marketing","Fill in Event Name (e.g. Trade Show 2026)","Fill Serial No, Qty (can be more than 1 for exhibition), Date, Expected Return","To return after event: Log Movement - Goods Receipt IN - Return Exhibition Marketing"]},
    {who:"Logistics Admin - GIT Tracking",color:"#d97706",bg:"#fffbeb",steps:["Go to GIT Register tab - this tracks units ordered but not yet received","When PO is confirmed: click Add GIT Entry, fill PO number, item, qty, supplier, ETA date, set Status to On Order","When Pudu ships: update Status to Shipped, add vessel reference number","When goods arrive at port or customs: update to Customs","When goods are physically received in warehouse and GR is posted in SAP: update to Received","GIT units must NEVER appear in warehouse stock until SAP GR is done"]},
    {who:"Logistics Admin - SAP Import",color:"#0369a1",bg:"#f0f9ff",steps:["Go to Import SAP tab","Download the SAP Import Template (Excel file)","Fill in SAP_Stock_Snapshot sheet with month-end closing stock per model","Fill in SAP_GIT_Register sheet with open POs from SAP ME2M or ME2N","Do NOT rename sheet tabs or column headers in Row 3","Save the file and click Choose SAP Excel File to upload","Review the preview table - green Ready means matched, orange means model code needs checking","Click Confirm Import - Stock Position and GIT Register update automatically"]},
    {who:"Logistics Admin - Month End Process",color:"#5b21b6",bg:"#f5f3ff",steps:["After SAP month-end close, go to Import SAP tab and upload your SAP report","OR go to Admin then SAP Snapshots to manually key in closing stock per model","Go to Stock Position tab - it now shows SAP vs Tracker side by side with Effective Available stock","Go to Admin then Monthly Summary - review all GR and GI movements for the month","Use this summary to verify against SAP postings","Go to Export Excel - set date range, download the report","Use By Movement Type sheet for SAP reconciliation entries","Check Overdue Returns in Dashboard - follow up with custodians"]},
  ];

  return(
    <div style={{ maxWidth:820,display:"flex",flexDirection:"column",gap:16 }}>
      <div style={{ background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)",borderRadius:12,padding:"18px 22px",color:"#fff" }}>
        <div style={{ fontSize:17,fontWeight:700,marginBottom:3 }}>How to Use This App</div>
        <div style={{ fontSize:13,color:"#C8EEF5" }}>Step-by-step guide for all roles: Sales, Engineer, Marketing, Logistics Admin</div>
      </div>
      {howtoSteps.map((s,i)=>(
        <div key={i} style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,overflow:"hidden" }}>
          <div style={{ background:s.bg,borderBottom:"3px solid "+s.color,padding:"10px 18px" }}>
            <span style={{ background:s.color,color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700 }}>{s.who}</span>
          </div>
          <div style={{ padding:"12px 18px" }}>
            {s.steps.map((step,j)=>(
              <div key={j} style={{ padding:"5px 0 5px 14px",borderLeft:"2px solid #e5e7eb",marginBottom:4,fontSize:13,color:"#374151" }}>
                {j+1}. {step}
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}


// ─── FLEET CALENDAR ───────────────────────────────────────────────────────────
function FleetCalendar({ movements, fleetUnits, setFleetUnits }) {
  const now = new Date();
  const [selMonth, setSelMonth] = useState(now.getFullYear() + "-" + String(now.getMonth()+1).padStart(2,"0"));
  const [selEntity, setSelEntity] = useState("ALL");
  const [editUnit, setEditUnit] = useState(null);
  const [showAddUnit, setShowAddUnit] = useState(false);
  const [newUnit, setNewUnit] = useState({ entity:ENTITIES[0], model:"", variant:"", serialNo:"", location:"", remark:"" });

  // Days in selected month
  const [yr, mo] = selMonth.split("-").map(Number);
  const daysInMonth = new Date(yr, mo, 0).getDate();
  const days = Array.from({ length: daysInMonth }, (_, i) => i + 1);

  // Filter units by entity
  const filteredUnits = selEntity === "ALL" ? fleetUnits : fleetUnits.filter(u => u.entity === selEntity);

  // Group by entity for display
  const grouped = {};
  filteredUnits.forEach(u => {
    if (!grouped[u.entity]) grouped[u.entity] = [];
    grouped[u.entity].push(u);
  });

  // Build status map from movements: unitId+day -> { status, label }
  // status: "demo" = red, "reserved" = orange, "available" = green
  function getDayStatus(unit, day) {
    const dateStr = `${yr}-${String(mo).padStart(2,"0")}-${String(day).padStart(2,"0")}`;
    // Check movements matching this unit serial
    const active = movements.filter(m => {
      if (m.serialNo !== unit.serialNo && !m.serialNo?.includes(unit.serialNo)) return false;
      if (getDir(m.type) !== "GI") return false;
      if (!m.date) return false;
      const start = m.date;
      const end = m.actualReturn || m.expectedReturn || "";
      if (dateStr < start) return false;
      if (end && dateStr > end) return false;
      if (m.status === "Returned" && m.actualReturn && dateStr > m.actualReturn) return false;
      return true;
    });
    if (active.length === 0) return null;
    const m = active[0];
    const isExpired = m.expectedReturn && dateStr === m.expectedReturn && m.status !== "Returned";
    return {
      status: isExpired ? "reserved" : "demo",
      label: m.remark || m.customerName || m.eventName || m.custodian || "",
    };
  }

  // Manual overrides stored per unit
  const [overrides, setOverrides] = useState({});
  function toggleOverride(unitId, day) {
    const key = `${unitId}_${day}`;
    setOverrides(prev => {
      const cur = prev[key];
      const next = cur === "demo" ? "reserved" : cur === "reserved" ? "available" : "demo";
      return { ...prev, [key]: next };
    });
  }

  const STATUS_COLORS = {
    demo:      { bg: "#dc2626", text: "#fff", label: "Demo at customer site" },
    reserved:  { bg: "#d97706", text: "#fff", label: "Reserved for next Demo" },
    available: { bg: "#16a34a", text: "#fff", label: "Available" },
  };

  // Today highlight
  const todayDay = now.getFullYear() === yr && now.getMonth()+1 === mo ? now.getDate() : -1;

  async function saveUnit(unit) {
    const updated = fleetUnits.map(u => u.id === unit.id ? unit : u);
    setFleetUnits(updated); await save(SK.FLEET, updated); setEditUnit(null);
  }
  async function deleteUnit(id) {
    if (!appConfirm("Remove this unit from fleet?")) return;
    const updated = fleetUnits.filter(u => u.id !== id);
    setFleetUnits(updated); await save(SK.FLEET, updated);
  }
  async function addUnit() {
    if (!newUnit.model || !newUnit.entity) { alert("Model and Entity required."); return; }
    const updated = [...fleetUnits, { ...newUnit, id: Date.now() }];
    setFleetUnits(updated); await save(SK.FLEET, updated);
    setNewUnit({ entity:ENTITIES[0], model:"", variant:"", serialNo:"", location:"", remark:"" });
    setShowAddUnit(false);
  }

  const entityColors = { "Entity A": "#0D1B3E", "Entity B": "#7c3aed", ALL: "#4AACCC" };
  const INP = (label, key, obj, setObj, ph="") => (
    <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
      <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>{label}</label>
      <input value={obj[key]} onChange={e => setObj(o=>({...o,[key]:e.target.value}))}
        placeholder={ph} style={{ border:"1px solid #d1d5db", borderRadius:5, padding:"5px 8px", fontSize:12 }} />
    </div>
  );

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      {/* Header controls */}
      <div style={{ background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)", borderRadius:10, padding:"14px 18px", display:"flex", alignItems:"center", gap:14, flexWrap:"wrap" }}>
        <div>
          <div style={{ color:"#fff", fontWeight:700, fontSize:15 }}>Fleet Calendar</div>
          <div style={{ color:"#C8EEF5", fontSize:11, marginTop:2 }}>Day-by-day unit availability across all demo fleets</div>
        </div>
        <div style={{ display:"flex", gap:8, marginLeft:"auto", flexWrap:"wrap", alignItems:"center" }}>
          {/* Entity filter */}
          <div style={{ display:"flex", gap:4 }}>
            {["ALL",...ENTITIES].map(e => (
              <button key={e} onClick={() => setSelEntity(e)} style={{
                padding:"6px 14px", borderRadius:6, border:"none", cursor:"pointer", fontWeight:700, fontSize:12,
                background: selEntity===e ? (entityColors[e]||"#4AACCC") : "rgba(255,255,255,0.15)",
                color: selEntity===e ? "#fff" : "#C8EEF5",
              }}>{e}</button>
            ))}
          </div>
          {/* Month picker */}
          <input type="month" value={selMonth} onChange={e => setSelMonth(e.target.value)}
            style={{ border:"1px solid rgba(255,255,255,0.3)", borderRadius:6, padding:"6px 10px", fontSize:12, background:"rgba(255,255,255,0.1)", color:"#fff" }} />
          <button onClick={() => setShowAddUnit(s=>!s)}
            style={{ background:showAddUnit?"#374151":"#4AACCC", color:"#fff", border:"none", borderRadius:6, padding:"7px 14px", fontWeight:600, fontSize:12, cursor:"pointer" }}>
            {showAddUnit ? "Cancel" : "+ Add Unit"}
          </button>
        </div>
      </div>

      {/* CEO Highlights */}
      {(() => {
        const now2 = new Date();
        const todayStr2 = today();
        const d7 = new Date(now2); d7.setDate(now2.getDate() + 7);
        const d7str = d7.toISOString().split("T")[0];
        const allActive = fleetUnits.filter(u => selEntity === "ALL" || u.entity === selEntity);
        const totalUnitsCount = allActive.length;
        const outToday = allActive.filter(u =>
          movements.some(m =>
            m.serialNo === u.serialNo && getDir(m.type) === "GI" && m.status !== "Returned" &&
            m.date <= todayStr2 && (!m.expectedReturn || m.expectedReturn >= todayStr2)
          )
        );
        const overdueUnits = allActive.filter(u =>
          movements.some(m => m.serialNo === u.serialNo && isOverdue(m))
        );
        const returningThisWeek = allActive.filter(u =>
          movements.some(m =>
            m.serialNo === u.serialNo && getDir(m.type) === "GI" && m.status !== "Returned" &&
            m.expectedReturn && m.expectedReturn >= todayStr2 && m.expectedReturn <= d7str
          )
        );
        const utilRate = totalUnitsCount > 0 ? Math.round((outToday.length / totalUnitsCount) * 100) : 0;
        const tmkTotal = fleetUnits.filter(u => u.entity === ENTITIES[0]).length;
        const tmsTotal = fleetUnits.filter(u => u.entity === ENTITIES[1]).length;
        const tmkOut = outToday.filter(u => u.entity === ENTITIES[0]).length;
        const tmsOut = outToday.filter(u => u.entity === ENTITIES[1]).length;
        if (totalUnitsCount === 0) return null;
        return (
          <div style={{ background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)", borderRadius:12, padding:"16px 20px", marginBottom:4 }}>
            <div style={{ color:"#C8EEF5", fontSize:11, fontWeight:600, letterSpacing:1, textTransform:"uppercase", marginBottom:12 }}>
              Fleet Management Highlights — {selEntity === "ALL" ? "All Entities" : selEntity} — {selMonth}
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))", gap:10, marginBottom:12 }}>
              {[
                ["Fleet Size", totalUnitsCount, "Total registered units", "#C8EEF5"],
                ["Out Today", outToday.length, "Units at customer / demo", "#ef4444"],
                ["Available", totalUnitsCount - outToday.length, "Units in office", "#4ade80"],
                ["Utilisation", utilRate + "%", "Of fleet deployed today", utilRate >= 70 ? "#4ade80" : utilRate >= 40 ? "#fbbf24" : "#ef4444"],
                ["Returning Soon", returningThisWeek.length, "Expected back in 7 days", "#fbbf24"],
                ["Overdue", overdueUnits.length, "Past expected return", overdueUnits.length > 0 ? "#ef4444" : "#4ade80"],
              ].map(([label, value, sub, col]) => (
                <div key={label} style={{ background:"rgba(255,255,255,0.08)", borderRadius:8, padding:"10px 12px" }}>
                  <div style={{ fontSize:22, fontWeight:800, color:col }}>{value}</div>
                  <div style={{ fontSize:11, fontWeight:700, color:"#fff", marginTop:1 }}>{label}</div>
                  <div style={{ fontSize:9, color:"#9ca3af", marginTop:1 }}>{sub}</div>
                </div>
              ))}
            </div>
            {selEntity === "ALL" && (tmkTotal > 0 || tmsTotal > 0) && (
              <div style={{ display:"flex", gap:10, flexWrap:"wrap", marginBottom:8 }}>
                {[[ENTITIES[0], tmkTotal, tmkOut, "#4AACCC"], [ENTITIES[1], tmsTotal, tmsOut, "#C8EEF5"]].map(([ent, total, out, col]) => total > 0 && (
                  <div key={ent} style={{ background:"rgba(255,255,255,0.06)", borderRadius:8, padding:"8px 14px", display:"flex", gap:12, alignItems:"center" }}>
                    <span style={{ fontWeight:700, fontSize:12, color:col }}>{ent}</span>
                    <span style={{ color:"#ef4444", fontSize:11 }}>⬆ {out} out</span>
                    <span style={{ color:"#4ade80", fontSize:11 }}>⬇ {total - out} in</span>
                    <div style={{ background:"rgba(255,255,255,0.1)", borderRadius:20, padding:"2px 8px", fontSize:10, color:"#fff" }}>
                      {total > 0 ? Math.round((out/total)*100) : 0}% utilised
                    </div>
                  </div>
                ))}
              </div>
            )}
            {overdueUnits.length > 0 && (
              <div style={{ background:"rgba(220,38,38,0.2)", border:"1px solid rgba(220,38,38,0.4)", borderRadius:8, padding:"8px 12px" }}>
                <span style={{ color:"#fca5a5", fontWeight:700, fontSize:12 }}>⚠ Overdue: </span>
                <span style={{ color:"#fecaca", fontSize:12 }}>{overdueUnits.map(u => u.model || u.serialNo).join(", ")}</span>
              </div>
            )}
          </div>
        );
      })()}

      {/* Legend */}
      <div style={{ display:"flex", gap:12, flexWrap:"wrap", alignItems:"center" }}>
        {Object.entries(STATUS_COLORS).map(([k,v]) => (
          <div key={k} style={{ display:"flex", alignItems:"center", gap:5 }}>
            <div style={{ width:14, height:14, borderRadius:3, background:v.bg }} />
            <span style={{ fontSize:12, color:"#374151" }}>{v.label}</span>
          </div>
        ))}
        <span style={{ fontSize:11, color:"#9ca3af", marginLeft:8 }}>Click any cell to cycle status: Demo → Reserved → Available</span>
      </div>

      {/* Add unit form */}
      {showAddUnit && (
        <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:14 }}>
          <SectionHead title="Add Unit to Fleet" />
          <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:10 }}>
            <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
              <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>Entity *</label>
              <select value={newUnit.entity} onChange={e => setNewUnit(o=>({...o,entity:e.target.value}))}
                style={{ border:"1px solid #d1d5db", borderRadius:5, padding:"5px 8px", fontSize:12 }}>
                {ENTITIES.map(en => <option key={en} value={en}>{en}</option>)}
              </select>
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
              <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>Entity *</label>
              <select value={newUnit.entity} onChange={e => setNewUnit(u => ({ ...u, entity:e.target.value }))}
                style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"6px 10px", fontSize:13 }}>
                <option value="">— Select —</option>
                {ENTITIES.map(e => <option key={e} value={e}>{e}</option>)}
              </select>
            </div>
            {INP("Model *","model",newUnit,setNewUnit,"e.g. PUDU T300")}
            {INP("Variant / Full Description","variant",newUnit,setNewUnit,"e.g. WTIDL1 or PUDU T300 (BLACK) WTIDL1")}
            {INP("Serial No","serialNo",newUnit,setNewUnit,"e.g. 8260B5318060021")}
            {INP("Location","location",newUnit,setNewUnit,"e.g. Head Office")}
            {INP("Remark","remark",newUnit,setNewUnit,"e.g. KL Demo")}
          </div>
          <button onClick={addUnit} style={{ marginTop:10, background:"#0D1B3E", color:"#fff", border:"none", borderRadius:6, padding:"8px 18px", fontWeight:600, fontSize:13, cursor:"pointer" }}>
            Add Unit
          </button>
        </div>
      )}

      {/* Calendar tables per entity */}
      {Object.entries(grouped).map(([entity, units]) => (
        <div key={entity} style={{ display:"flex", flexDirection:"column", gap:0 }}>
          {/* Entity header */}
          <div style={{ background: entityColors[entity]||"#0D1B3E", color:"#fff", padding:"8px 14px", borderRadius:"8px 8px 0 0", display:"flex", alignItems:"center", gap:10 }}>
            <span style={{ fontWeight:700, fontSize:13 }}>Demo Fleet — {entity}</span>
            <span style={{ fontSize:11, color:"rgba(255,255,255,0.7)", marginLeft:"auto" }}>{units.length} units · {selMonth}</span>
          </div>

          {/* Scrollable calendar */}
          <div style={{ overflowX:"auto", border:"1px solid #e5e7eb", borderTop:"none", borderRadius:"0 0 8px 8px", background:"#fff" }}>
            <table style={{ borderCollapse:"collapse", fontSize:10, minWidth:"100%" }}>
              <thead>
                {/* Month header row */}
                <tr style={{ background:"#4AACCC" }}>
                  <th style={{ padding:"5px 8px", textAlign:"left", fontWeight:700, color:"#fff", fontSize:10, minWidth:24, borderRight:"1px solid #e5e7eb" }}>No</th>
                  <th style={{ padding:"5px 8px", textAlign:"left", fontWeight:700, color:"#fff", fontSize:10, minWidth:140, borderRight:"1px solid #e5e7eb" }}>Item</th>
                  <th style={{ padding:"5px 8px", textAlign:"left", fontWeight:700, color:"#fff", fontSize:10, minWidth:40, borderRight:"1px solid #e5e7eb" }}>Qty</th>
                  <th style={{ padding:"5px 8px", textAlign:"left", fontWeight:700, color:"#fff", fontSize:10, minWidth:130, borderRight:"1px solid #e5e7eb" }}>Serial No</th>
                  <th style={{ padding:"5px 8px", textAlign:"left", fontWeight:700, color:"#fff", fontSize:10, minWidth:110, borderRight:"1px solid #e5e7eb" }}>Location</th>
                  <th style={{ padding:"5px 8px", textAlign:"left", fontWeight:700, color:"#fff", fontSize:10, minWidth:70, borderRight:"2px solid #0D1B3E" }}>Remark</th>
                  <th colSpan={daysInMonth} style={{ padding:"5px 8px", textAlign:"center", fontWeight:700, color:"#fff", fontSize:11 }}>
                    {new Date(yr,mo-1,1).toLocaleString("en-MY",{month:"long"})} {yr}
                  </th>
                  <th style={{ padding:"5px 8px", borderLeft:"2px solid #0D1B3E", fontWeight:700, color:"#fff", fontSize:10, minWidth:60 }}>Actions</th>
                </tr>
                {/* Day numbers row */}
                <tr style={{ background:"#f9fafb" }}>
                  {["","","","","",""].map((_, i) => <th key={i} style={{ padding:"3px 0", borderRight: i===5?"2px solid #0D1B3E":"1px solid #e5e7eb" }} />)}
                  {days.map(d => {
                    const dateObj = new Date(yr, mo-1, d);
                    const isWknd = dateObj.getDay()===0||dateObj.getDay()===6;
                    const isToday = d === todayDay;
                    return (
                      <th key={d} style={{
                        padding:"2px 0", width:22, minWidth:22, textAlign:"center",
                        fontSize:9, fontWeight:isToday?700:400,
                        color:isToday?"#0D1B3E":isWknd?"#9ca3af":"#374151",
                        background:isToday?"#C8EEF5":isWknd?"#f3f4f6":"#f9fafb",
                        borderRight:"1px solid #e5e7eb",
                      }}>{d}</th>
                    );
                  })}
                  <th style={{ borderLeft:"2px solid #0D1B3E" }} />
                </tr>
              </thead>
              <tbody>
                {units.map((unit, idx) => (
                  <tr key={unit.id} style={{ borderTop:"1px solid #f3f4f6", background: idx%2===0?"#fff":"#fafafa" }}>
                    <td style={{ padding:"5px 8px", color:"#6b7280", borderRight:"1px solid #e5e7eb", textAlign:"center", fontSize:10 }}>{idx+1}</td>
                    <td style={{ padding:"5px 8px", fontWeight:600, color:"#0D1B3E", borderRight:"1px solid #e5e7eb", whiteSpace:"nowrap", fontSize:10 }}>
                      {editUnit?.id===unit.id
                        ? <input value={editUnit.model} onChange={e=>setEditUnit(u=>({...u,model:e.target.value}))} style={{ width:120, fontSize:10, border:"1px solid #d1d5db", borderRadius:3, padding:"2px 4px" }} />
                        : <>{unit.model}{unit.variant ? <span style={{ color:"#9ca3af", fontWeight:400 }}> {unit.variant}</span> : null}</>}
                    </td>
                    <td style={{ padding:"5px 8px", textAlign:"center", borderRight:"1px solid #e5e7eb", color:"#374151" }}>1</td>
                    <td style={{ padding:"5px 8px", borderRight:"1px solid #e5e7eb", color:"#6b7280", fontFamily:"monospace", fontSize:9, whiteSpace:"nowrap" }}>
                      {editUnit?.id===unit.id
                        ? <input value={editUnit.serialNo} onChange={e=>setEditUnit(u=>({...u,serialNo:e.target.value}))} style={{ width:120, fontSize:9, border:"1px solid #d1d5db", borderRadius:3, padding:"2px 4px", fontFamily:"monospace" }} />
                        : unit.serialNo}
                    </td>
                    <td style={{ padding:"5px 8px", borderRight:"1px solid #e5e7eb", color:"#374151", whiteSpace:"nowrap", fontSize:10 }}>
                      {editUnit?.id===unit.id
                        ? <input value={editUnit.location} onChange={e=>setEditUnit(u=>({...u,location:e.target.value}))} style={{ width:90, fontSize:10, border:"1px solid #d1d5db", borderRadius:3, padding:"2px 4px" }} />
                        : unit.location}
                    </td>
                    <td style={{ padding:"5px 8px", borderRight:"2px solid #0D1B3E", fontSize:10,
                      color: unit.remark?.includes("Stock")?"#16a34a": unit.remark?.includes("Loan")?"#d97706":"#374151",
                      fontWeight:600, whiteSpace:"nowrap" }}>
                      {editUnit?.id===unit.id
                        ? <input value={editUnit.remark} onChange={e=>setEditUnit(u=>({...u,remark:e.target.value}))} style={{ width:70, fontSize:10, border:"1px solid #d1d5db", borderRadius:3, padding:"2px 4px" }} />
                        : unit.remark}
                    </td>
                    {/* Day cells */}
                    {days.map(d => {
                      const overrideKey = `${unit.id}_${d}`;
                      const override = overrides[overrideKey];
                      const autoStatus = getDayStatus(unit, d);
                      const finalStatus = override || (autoStatus ? autoStatus.status : "available");
                      const label = autoStatus?.label || "";
                      const col = STATUS_COLORS[finalStatus];
                      const dateObj = new Date(yr, mo-1, d);
                      const isWknd = dateObj.getDay()===0||dateObj.getDay()===6;
                      const isToday = d === todayDay;
                      return (
                        <td key={d} onClick={() => toggleOverride(unit.id, d)} title={label || col.label}
                          style={{
                            width:22, minWidth:22, height:20,
                            background: finalStatus==="available" ? (isWknd?"#dcfce7":isToday?"#bbf7d0":"#dcfce7") : col.bg,
                            borderRight:"1px solid rgba(255,255,255,0.3)",
                            cursor:"pointer", position:"relative",
                            borderLeft: isToday?"2px solid #0D1B3E":"none",
                          }}>
                          {label && finalStatus!=="available" && (
                            <div style={{ fontSize:6, color:"#fff", overflow:"hidden", whiteSpace:"nowrap", textOverflow:"ellipsis", padding:"1px 2px", lineHeight:1.2 }}>
                              {label.slice(0,8)}
                            </div>
                          )}
                        </td>
                      );
                    })}
                    {/* Actions */}
                    <td style={{ padding:"4px 6px", borderLeft:"2px solid #0D1B3E", whiteSpace:"nowrap" }}>
                      {editUnit?.id===unit.id ? (
                        <>
                          <button onClick={() => saveUnit(editUnit)} style={{ background:"#dcfce7", border:"1px solid #86efac", color:"#166534", borderRadius:3, padding:"2px 6px", fontSize:9, cursor:"pointer", marginRight:2 }}>Save</button>
                          <button onClick={() => setEditUnit(null)} style={{ background:"#f3f4f6", border:"1px solid #e5e7eb", color:"#374151", borderRadius:3, padding:"2px 6px", fontSize:9, cursor:"pointer" }}>X</button>
                        </>
                      ) : (
                        <>
                          <button onClick={() => setEditUnit({...unit})} style={{ background:"#e0f2fe", border:"1px solid #7dd3fc", color:"#0369a1", borderRadius:3, padding:"2px 6px", fontSize:9, cursor:"pointer", marginRight:2 }}>Edit</button>
                          <button onClick={() => deleteUnit(unit.id)} style={{ background:"#fee2e2", border:"1px solid #fca5a5", color:"#991b1b", borderRadius:3, padding:"2px 6px", fontSize:9, cursor:"pointer" }}>Del</button>
                        </>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Summary strip */}
          <div style={{ display:"flex", gap:10, padding:"8px 0 2px", flexWrap:"wrap" }}>
            {(() => {
              const avail = units.filter(u => !movements.some(m => m.serialNo===u.serialNo && getDir(m.type)==="GI" && m.status!=="Returned")).length;
              return [
                { label:"Total Units", val:units.length, col:"#0D1B3E" },
                { label:"Available",   val:avail,         col:"#16a34a" },
                { label:"In Field",    val:units.length-avail, col:"#dc2626" },
              ].map(({label,val,col}) => (
                <div key={label} style={{ background:"#fff", border:`1px solid ${col}40`, borderLeft:`3px solid ${col}`, borderRadius:6, padding:"5px 12px" }}>
                  <span style={{ fontSize:16, fontWeight:700, color:col }}>{val}</span>
                  <span style={{ fontSize:10, color:"#6b7280", marginLeft:6 }}>{label}</span>
                </div>
              ));
            })()}
          </div>
        </div>
      ))}

      {filteredUnits.length === 0 && (
        <div style={{ padding:32, textAlign:"center", color:"#9ca3af", fontSize:13 }}>
          <div style={{ fontSize:24, marginBottom:8 }}>📅</div>
          <div style={{ fontWeight:600, fontSize:14, color:"#374151", marginBottom:6 }}>No units yet for {selEntity==="ALL"?"any entity":selEntity}</div>
          <div>Log a movement via ✏️ Log Movement tab — select GI (OUT), fill in Serial Number and Entity. Units appear here automatically.</div>
        </div>
      )}
    </div>
  );
}


// ─── FLEET CALENDAR ──────────────────────────────────────────────────────────
// ─── BACKUP & RESTORE ────────────────────────────────────────────────────────
function ItemMasterSetup({ itemMaster, setItemMaster }) {
  const [newItem, setNewItem] = useState({ sapCode:"", name:"", category:"Robot" });
  const [search, setSearch] = useState("");
  const [editId, setEditId] = useState(null);
  const [editData, setEditData] = useState({});

  const filtered = itemMaster.filter(i =>
    !search || i.name.toLowerCase().includes(search.toLowerCase()) || i.sapCode.toLowerCase().includes(search.toLowerCase())
  );

  async function addItem() {
    if (!newItem.sapCode.trim() || !newItem.name.trim()) { alert("SAP Code and Item Name are required."); return; }
    if (itemMaster.find(i => i.sapCode === newItem.sapCode.trim())) { alert("SAP Code already exists."); return; }
    const item = { ...newItem, id: "ITM" + Date.now(), active: true, sapCode: newItem.sapCode.trim(), name: newItem.name.trim() };
    const updated = [...itemMaster, item];
    setItemMaster(updated);
    await save(SK.ITEMS, updated);
    setNewItem({ sapCode:"", name:"", category:"Robot" });
  }

  async function toggleActive(id) {
    const updated = itemMaster.map(i => i.id === id ? { ...i, active: !i.active } : i);
    setItemMaster(updated); await save(SK.ITEMS, updated);
  }

  async function saveEdit(id) {
    const updated = itemMaster.map(i => i.id === id ? { ...i, ...editData } : i);
    setItemMaster(updated); await save(SK.ITEMS, updated);
    setEditId(null); setEditData({});
  }

  const inp = (label, key, val, setter, ph="") => (
    <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
      <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>{label}</label>
      <input value={val} onChange={e => setter(v => ({ ...v, [key]: e.target.value }))} placeholder={ph}
        style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13 }} />
    </div>
  );

  const catSel = (val, setter) => (
    <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
      <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>Category</label>
      <select value={val} onChange={e => setter(v => ({ ...v, category: e.target.value }))}
        style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13 }}>
        <option value="Robot">Robot</option>
        <option value="Accessory">Accessory</option>
        <option value="Spare Part">Spare Part</option>
        <option value="Other">Other</option>
      </select>
    </div>
  );

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <div style={{ background:"#f0f9ff", border:"1px solid #bae6fd", borderRadius:8, padding:12, fontSize:12, color:"#0369a1" }}>
        <strong>Item Master</strong> — All items here appear as dropdown options in Log Movement, the Public Form, and the Fleet Calendar.
        SAP Code must match exactly what is in your SAP system. Items can be deactivated without deleting.
      </div>

      {/* Add new */}
      <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:16 }}>
        <SectionHead title="Add New Item" />
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr auto", gap:10, alignItems:"flex-end" }}>
          {inp("SAP Item Code *", "sapCode", newItem.sapCode, setNewItem, "e.g. T300-BLACK")}
          {inp("Item Description *", "name", newItem.name, setNewItem, "e.g. PUDU T300 (BLACK) WTIDL1")}
          {catSel(newItem.category, setNewItem)}
          <button onClick={addItem} style={{ background:"#0D1B3E", color:"#fff", border:"none", borderRadius:6, padding:"9px 16px", fontWeight:600, fontSize:13, cursor:"pointer", whiteSpace:"nowrap" }}>
            + Add Item
          </button>
        </div>
      </div>

      {/* Search + list */}
      <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, overflow:"hidden" }}>
        <div style={{ padding:"12px 16px", borderBottom:"1px solid #f3f4f6", display:"flex", gap:10, alignItems:"center" }}>
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search by SAP code or description..."
            style={{ flex:1, border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13 }} />
          <span style={{ fontSize:12, color:"#9ca3af", whiteSpace:"nowrap" }}>{filtered.length} of {itemMaster.length} items</span>
        </div>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
          <thead style={{ background:"#f9fafb" }}>
            <tr>{["SAP Code","Item Description","Category","Status",""].map(h => (
              <th key={h} style={{ padding:"8px 12px", textAlign:"left", fontWeight:600, color:"#374151", borderBottom:"1px solid #e5e7eb", whiteSpace:"nowrap" }}>{h}</th>
            ))}</tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={5} style={{ padding:24, textAlign:"center", color:"#9ca3af" }}>No items found.</td></tr>
            )}
            {filtered.map(item => (
              <tr key={item.id} style={{ borderTop:"1px solid #f3f4f6", background: item.active ? "transparent" : "#fafafa" }}>
                {editId === item.id ? (
                  <>
                    <td style={{ padding:"6px 10px" }}>
                      <input value={editData.sapCode||item.sapCode} onChange={e => setEditData(d=>({...d,sapCode:e.target.value}))}
                        style={{ border:"1px solid #d1d5db", borderRadius:4, padding:"5px 8px", fontSize:12, width:120 }} />
                    </td>
                    <td style={{ padding:"6px 10px" }}>
                      <input value={editData.name||item.name} onChange={e => setEditData(d=>({...d,name:e.target.value}))}
                        style={{ border:"1px solid #d1d5db", borderRadius:4, padding:"5px 8px", fontSize:12, width:280 }} />
                    </td>
                    <td style={{ padding:"6px 10px" }}>
                      <select value={editData.category||item.category} onChange={e => setEditData(d=>({...d,category:e.target.value}))}
                        style={{ border:"1px solid #d1d5db", borderRadius:4, padding:"5px 8px", fontSize:12 }}>
                        <option>Robot</option><option>Accessory</option><option>Spare Part</option><option>Other</option>
                      </select>
                    </td>
                    <td style={{ padding:"6px 10px" }}><Badge label={item.active?"Active":"Inactive"} bg={item.active?"#16a34a":"#9ca3af"} /></td>
                    <td style={{ padding:"6px 10px", display:"flex", gap:6 }}>
                      <button onClick={() => saveEdit(item.id)} style={{ background:"#16a34a", color:"#fff", border:"none", borderRadius:4, padding:"4px 10px", fontSize:11, fontWeight:600, cursor:"pointer" }}>Save</button>
                      <button onClick={() => { setEditId(null); setEditData({}); }} style={{ background:"#f3f4f6", color:"#374151", border:"none", borderRadius:4, padding:"4px 10px", fontSize:11, cursor:"pointer" }}>Cancel</button>
                    </td>
                  </>
                ) : (
                  <>
                    <td style={{ padding:"8px 12px", fontFamily:"monospace", fontSize:12, color:"#0369a1", fontWeight:600 }}>{item.sapCode}</td>
                    <td style={{ padding:"8px 12px", color: item.active?"#374151":"#9ca3af" }}>{item.name}</td>
                    <td style={{ padding:"8px 12px" }}><Badge label={item.category} bg={item.category==="Robot"?"#0D1B3E":item.category==="Accessory"?"#4AACCC":"#9ca3af"} /></td>
                    <td style={{ padding:"8px 12px" }}><Badge label={item.active?"Active":"Inactive"} bg={item.active?"#16a34a":"#9ca3af"} /></td>
                    <td style={{ padding:"8px 12px", display:"flex", gap:6, whiteSpace:"nowrap" }}>
                      <button onClick={() => { setEditId(item.id); setEditData({ sapCode:item.sapCode, name:item.name, category:item.category }); }}
                        style={{ background:"#f0f9ff", color:"#0369a1", border:"1px solid #bae6fd", borderRadius:4, padding:"3px 8px", fontSize:11, cursor:"pointer" }}>Edit</button>
                      <button onClick={() => toggleActive(item.id)}
                        style={{ background: item.active?"#fee2e2":"#dcfce7", color:item.active?"#dc2626":"#16a34a", border:"none", borderRadius:4, padding:"3px 8px", fontSize:11, cursor:"pointer" }}>
                        {item.active?"Deactivate":"Activate"}
                      </button>
                    </td>
                  </>
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── CUSTOMER MASTER SETUP ───────────────────────────────────────────────────
function CustomerMasterSetup({ customers, setCustomers }) {
  const [newCust, setNewCust] = useState({ name:"", country:"Malaysia", contactPerson:"", email:"", phone:"" });
  const [search, setSearch] = useState("");
  const [editId, setEditId] = useState(null);
  const [editData, setEditData] = useState({});

  const filtered = customers.filter(c =>
    !search || c.name.toLowerCase().includes(search.toLowerCase()) || (c.contactPerson||"").toLowerCase().includes(search.toLowerCase())
  );

  async function addCust() {
    if (!newCust.name.trim()) { alert("Customer Name is required."); return; }
    if (customers.find(c => c.name.toLowerCase() === newCust.name.trim().toLowerCase())) { alert("Customer already exists."); return; }
    const c = { ...newCust, id:"CUS"+Date.now(), active:true, name:newCust.name.trim() };
    const updated = [...customers, c];
    setCustomers(updated); await save(SK.CUSTOMERS, updated);
    setNewCust({ name:"", country:"Malaysia", contactPerson:"", email:"", phone:"" });
  }

  async function toggleActive(id) {
    const updated = customers.map(c => c.id===id ? { ...c, active:!c.active } : c);
    setCustomers(updated); await save(SK.CUSTOMERS, updated);
  }

  async function saveEdit(id) {
    const updated = customers.map(c => c.id===id ? { ...c, ...editData } : c);
    setCustomers(updated); await save(SK.CUSTOMERS, updated);
    setEditId(null); setEditData({});
  }

  const inpF = (ph, key, wide=false) => (
    <input value={newCust[key]} onChange={e => setNewCust(v=>({...v,[key]:e.target.value}))} placeholder={ph}
      style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13, width:wide?"100%":"auto" }} />
  );

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
      <div style={{ background:"#f0fdf4", border:"1px solid #86efac", borderRadius:8, padding:12, fontSize:12, color:"#166534" }}>
        <strong>Customer Master</strong> — All customers here appear in the Customer Name dropdown in Log Movement and the Public Form.
        Add customers before logging their demo visits so the correct name is selected, not typed freehand.
      </div>

      {/* Add new */}
      <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:16 }}>
        <SectionHead title="Add New Customer" />
        <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr 1fr 1fr auto", gap:10, alignItems:"flex-end" }}>
          <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
            <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>Customer Name *</label>
            {inpF("e.g. Agilent Technologies", "name", true)}
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
            <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>Country</label>
            {inpF("e.g. Malaysia", "country")}
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
            <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>Contact Person</label>
            {inpF("e.g. John Lim", "contactPerson")}
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:3 }}>
            <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>Phone / Email</label>
            {inpF("e.g. +60 12-345 6789", "phone")}
          </div>
          <button onClick={addCust} style={{ background:"#0D1B3E", color:"#fff", border:"none", borderRadius:6, padding:"9px 16px", fontWeight:600, fontSize:13, cursor:"pointer", whiteSpace:"nowrap" }}>
            + Add
          </button>
        </div>
      </div>

      {/* List */}
      <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, overflow:"hidden" }}>
        <div style={{ padding:"12px 16px", borderBottom:"1px solid #f3f4f6", display:"flex", gap:10, alignItems:"center" }}>
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search customers..."
            style={{ flex:1, border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13 }} />
          <span style={{ fontSize:12, color:"#9ca3af", whiteSpace:"nowrap" }}>{filtered.length} of {customers.length} customers</span>
        </div>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
          <thead style={{ background:"#f9fafb" }}>
            <tr>{["Customer Name","Country","Contact Person","Phone / Email","Status",""].map(h => (
              <th key={h} style={{ padding:"8px 12px", textAlign:"left", fontWeight:600, color:"#374151", borderBottom:"1px solid #e5e7eb", whiteSpace:"nowrap" }}>{h}</th>
            ))}</tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={6} style={{ padding:24, textAlign:"center", color:"#9ca3af" }}>No customers found.</td></tr>
            )}
            {filtered.map(c => (
              <tr key={c.id} style={{ borderTop:"1px solid #f3f4f6", background:c.active?"transparent":"#fafafa" }}>
                {editId===c.id ? (
                  <>
                    <td style={{ padding:"6px 10px" }}><input value={editData.name||c.name} onChange={e=>setEditData(d=>({...d,name:e.target.value}))} style={{ border:"1px solid #d1d5db",borderRadius:4,padding:"5px 8px",fontSize:12,width:200 }} /></td>
                    <td style={{ padding:"6px 10px" }}><input value={editData.country||c.country||""} onChange={e=>setEditData(d=>({...d,country:e.target.value}))} style={{ border:"1px solid #d1d5db",borderRadius:4,padding:"5px 8px",fontSize:12,width:100 }} /></td>
                    <td style={{ padding:"6px 10px" }}><input value={editData.contactPerson||c.contactPerson||""} onChange={e=>setEditData(d=>({...d,contactPerson:e.target.value}))} style={{ border:"1px solid #d1d5db",borderRadius:4,padding:"5px 8px",fontSize:12,width:140 }} /></td>
                    <td style={{ padding:"6px 10px" }}><input value={editData.phone||c.phone||""} onChange={e=>setEditData(d=>({...d,phone:e.target.value}))} style={{ border:"1px solid #d1d5db",borderRadius:4,padding:"5px 8px",fontSize:12,width:140 }} /></td>
                    <td style={{ padding:"6px 10px" }}><Badge label={c.active?"Active":"Inactive"} bg={c.active?"#16a34a":"#9ca3af"} /></td>
                    <td style={{ padding:"6px 10px", display:"flex", gap:6 }}>
                      <button onClick={()=>saveEdit(c.id)} style={{ background:"#16a34a",color:"#fff",border:"none",borderRadius:4,padding:"4px 10px",fontSize:11,fontWeight:600,cursor:"pointer" }}>Save</button>
                      <button onClick={()=>{setEditId(null);setEditData({});}} style={{ background:"#f3f4f6",color:"#374151",border:"none",borderRadius:4,padding:"4px 10px",fontSize:11,cursor:"pointer" }}>Cancel</button>
                    </td>
                  </>
                ) : (
                  <>
                    <td style={{ padding:"8px 12px", fontWeight:600, color:c.active?"#0D1B3E":"#9ca3af" }}>{c.name}</td>
                    <td style={{ padding:"8px 12px", color:"#6b7280" }}>{c.country||"—"}</td>
                    <td style={{ padding:"8px 12px", color:"#6b7280" }}>{c.contactPerson||"—"}</td>
                    <td style={{ padding:"8px 12px", color:"#6b7280" }}>{c.phone||"—"}</td>
                    <td style={{ padding:"8px 12px" }}><Badge label={c.active?"Active":"Inactive"} bg={c.active?"#16a34a":"#9ca3af"} /></td>
                    <td style={{ padding:"8px 12px", display:"flex", gap:6 }}>
                      <button onClick={()=>{setEditId(c.id);setEditData({name:c.name,country:c.country||"",contactPerson:c.contactPerson||"",phone:c.phone||""});}}
                        style={{ background:"#f0fdf4",color:"#166534",border:"1px solid #86efac",borderRadius:4,padding:"3px 8px",fontSize:11,cursor:"pointer" }}>Edit</button>
                      <button onClick={()=>toggleActive(c.id)}
                        style={{ background:c.active?"#fee2e2":"#dcfce7",color:c.active?"#dc2626":"#16a34a",border:"none",borderRadius:4,padding:"3px 8px",fontSize:11,cursor:"pointer" }}>
                        {c.active?"Deactivate":"Activate"}
                      </button>
                    </td>
                  </>
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}


// ─── SETUP ────────────────────────────────────────────────────────────────────
function Setup({ setup, setSetup, appUsers, setAppUsers, itemMaster, setItemMaster, customers, setCustomers }) {
  const [tab, setTab] = useState("general");
  const [saved, setSaved] = useState("");

  async function saveSetup(newSetup) {
    setSetup(newSetup);
    await save(SK.SETUP, newSetup);
    setSaved("Settings saved.");
    setTimeout(() => setSaved(""), 2500);
  }

  async function saveUsers(newUsers) {
    setAppUsers(newUsers);
    await save(SK.USERS, newUsers);
    setSaved("Users saved.");
    setTimeout(() => setSaved(""), 2500);
  }

  const s = setup || DEFAULT_SETUP;

  // General tab
  function GeneralTab() {
    const [appName, setAppName] = useState(s.appName || "Robotics Fleet & Demo Tracker");
    const [companyName, setCompanyName] = useState(s.companyName || "");
    const [entitiesText, setEntitiesText] = useState((s.entities || DEFAULT_SETUP.entities).join("\n"));
    const [regionsText, setRegionsText] = useState((s.regions || DEFAULT_SETUP.regions).join("\n"));
    return (
      <div style={{ display:"flex", flexDirection:"column", gap:14, maxWidth:600 }}>
        {[["App Name", appName, setAppName], ["Company Name", companyName, setCompanyName]].map(([label, val, setter]) => (
          <div key={label} style={{ display:"flex", flexDirection:"column", gap:4 }}>
            <label style={{ fontSize:12, fontWeight:600, color:"#374151" }}>{label}</label>
            <input value={val} onChange={e => setter(e.target.value)} style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"8px 10px", fontSize:13 }} />
          </div>
        ))}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
          {[["Entities (one per line)", entitiesText, setEntitiesText, "Entity A\nEntity B"], ["Regions / Locations (one per line)", regionsText, setRegionsText, "Head Office\nWarehouse"]].map(([label, val, setter, ph]) => (
            <div key={label} style={{ display:"flex", flexDirection:"column", gap:4 }}>
              <label style={{ fontSize:12, fontWeight:600, color:"#374151" }}>{label}</label>
              <textarea value={val} onChange={e => setter(e.target.value)} rows={6} placeholder={ph}
                style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"8px 10px", fontSize:13, resize:"vertical" }} />
            </div>
          ))}
        </div>
        <button onClick={() => saveSetup({ ...s, appName, companyName, entities: entitiesText.split("\n").map(x=>x.trim()).filter(Boolean), regions: regionsText.split("\n").map(x=>x.trim()).filter(Boolean) })}
          style={{ background:"#0D1B3E", color:"#fff", border:"none", borderRadius:8, padding:"10px 24px", fontWeight:700, fontSize:13, cursor:"pointer", alignSelf:"flex-start" }}>
          Save General Settings
        </button>
      </div>
    );
  }

  // Movement types tab (GR or GI)
  function MovTypeTab({ direction }) {
    const types = direction === "GR" ? (s.grTypes || DEFAULT_SETUP.grTypes) : (s.giTypes || DEFAULT_SETUP.giTypes);
    const [rows, setRows] = useState(types.map(t => ({ ...t })));
    const [newLabel, setNewLabel] = useState("");
    const [newSapCat, setNewSapCat] = useState("DEMO");

    function toggle(idx) { const r = [...rows]; r[idx] = { ...r[idx], active: !r[idx].active }; setRows(r); }
    function saveTypes() {
      const updated = direction === "GR" ? { ...s, grTypes: rows } : { ...s, giTypes: rows };
      saveSetup(updated);
    }
    function addType() {
      if (!newLabel.trim()) return;
      const val = direction + "_CUSTOM_" + Date.now();
      setRows([...rows, { value: val, label: newLabel.trim(), sapCat: newSapCat, active: true }]);
      setNewLabel(""); 
    }
    const bg = direction === "GR" ? "#dcfce7" : "#fee2e2";
    const col = direction === "GR" ? "#166534" : "#991b1b";
    return (
      <div style={{ display:"flex", flexDirection:"column", gap:12, maxWidth:700 }}>
        <div style={{ background:bg, border:`1px solid ${col}30`, borderRadius:8, padding:12, fontSize:12, color:col }}>
          <strong>{direction === "GR" ? "GR = Goods Receipt (stock coming IN)" : "GI = Goods Issue (stock going OUT)"}</strong> — Toggle to enable or disable movement types. Add custom types for your business. Changes take effect immediately in Log Movement form.
        </div>
        {rows.map((t, idx) => (
          <div key={t.value} style={{ display:"flex", alignItems:"center", gap:12, padding:"10px 14px", background:"#fff", border:"1px solid #e5e7eb", borderRadius:8 }}>
            <button onClick={() => toggle(idx)} style={{ width:44, height:24, borderRadius:12, border:"none", cursor:"pointer", background:t.active ? "#16a34a" : "#d1d5db", position:"relative", flexShrink:0, transition:"background 0.2s" }}>
              <span style={{ position:"absolute", top:2, left:t.active?22:2, width:20, height:20, background:"#fff", borderRadius:"50%", transition:"left 0.2s" }} />
            </button>
            <div style={{ flex:1 }}>
              <div style={{ fontWeight:600, fontSize:13, color: t.active ? "#374151" : "#9ca3af" }}>{t.label}</div>
              <div style={{ fontSize:11, color:"#9ca3af" }}>SAP Category: {t.sapCat} · Code: {t.value}</div>
            </div>
            <Badge label={t.sapCat} bg={t.sapCat==="STOCK"?"#16a34a":t.sapCat==="GIT"?"#d97706":"#7c3aed"} />
          </div>
        ))}
        <div style={{ display:"flex", gap:10, padding:"12px 14px", background:"#f9fafb", border:"1px dashed #d1d5db", borderRadius:8, alignItems:"flex-end" }}>
          <div style={{ flex:1, display:"flex", flexDirection:"column", gap:4 }}>
            <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>New movement type label</label>
            <input value={newLabel} onChange={e => setNewLabel(e.target.value)} placeholder="e.g. Internal Transfer" style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13 }} />
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
            <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>SAP Category</label>
            <select value={newSapCat} onChange={e => setNewSapCat(e.target.value)} style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13 }}>
              {SAP_STOCK_CATEGORIES.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}
            </select>
          </div>
          <button onClick={addType} style={{ background:"#0D1B3E", color:"#fff", border:"none", borderRadius:6, padding:"9px 16px", fontWeight:600, fontSize:13, cursor:"pointer" }}>+ Add</button>
        </div>
        <button onClick={saveTypes} style={{ background:"#0D1B3E", color:"#fff", border:"none", borderRadius:8, padding:"10px 24px", fontWeight:700, fontSize:13, cursor:"pointer", alignSelf:"flex-start" }}>
          Save {direction} Types
        </button>
      </div>
    );
  }

  // Users tab
  function UsersTab() {
    const [users, setUsers] = useState(appUsers.map(u => ({ ...u, permissions: { ...u.permissions } })));
    const [newUser, setNewUser] = useState({ name:"", pin:"", role:"Sales Personnel", entity:"", region:"", active:true });
    const allPerms = TAB_LIST.reduce((o,t) => ({ ...o, [t]:"view" }), {});

    function addUser() {
      if (!newUser.name || !newUser.pin) { alert("Name and PIN are required."); return; }
      const rolePerms = newUser.role === "Logistics Admin"
        ? TAB_LIST.reduce((o,t) => ({ ...o, [t]: ["User Guide","Version History"].includes(t) ? "view" : "edit" }), {})
        : { ...allPerms, "Log Movement":"none" }; // Sales/marketing use the form only
      const u = { ...newUser, id: Date.now(), permissions: rolePerms };
      const updated = [...users, u];
      setUsers(updated);
      saveUsers(updated);
      setNewUser({ name:"", pin:"", role:"Sales Personnel", entity:"", region:"", active:true });
    }

    function toggleActive(id) {
      const updated = users.map(u => u.id === id ? { ...u, active: !u.active } : u);
      setUsers(updated); saveUsers(updated);
    }

    function setPermission(id, tab, val) {
      const updated = users.map(u => u.id === id ? { ...u, permissions: { ...u.permissions, [tab]: val } } : u);
      setUsers(updated);
    }

    const [expandedUser, setExpandedUser] = useState(null);

    return (
      <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
        <div style={{ background:"#f0f9ff", border:"1px solid #bae6fd", borderRadius:8, padding:12, fontSize:12, color:"#0369a1" }}>
          <strong>Security model:</strong> Admin users have full access via the main app (PIN login). Sales, Engineers, and Marketing staff access the app only through the <strong>Public Form</strong> — they never see the main tracker. Set PIN to a 4-digit number that each user knows.
        </div>
        {/* Add user */}
        <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:14 }}>
          <SectionHead title="Add User" />
          <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"flex-end" }}>
            {[["Full Name","name","text"],["PIN (4 digits)","pin","password"]].map(([label,key,type]) => (
              <div key={key} style={{ display:"flex", flexDirection:"column", gap:4 }}>
                <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>{label}</label>
                <input type={type} value={newUser[key]} onChange={e => setNewUser(u => ({ ...u, [key]:e.target.value }))} maxLength={key==="pin"?4:50} style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13, width:key==="pin"?80:160 }} />
              </div>
            ))}
            {[["Role","role",ROLES],["Entity","entity",["(All)",...(s.entities||DEFAULT_SETUP.entities)]],["Region","region",["(All)",...(s.regions||DEFAULT_SETUP.regions)]]].map(([label,key,opts]) => (
              <div key={key} style={{ display:"flex", flexDirection:"column", gap:4 }}>
                <label style={{ fontSize:11, fontWeight:600, color:"#374151" }}>{label}</label>
                <select value={newUser[key]} onChange={e => setNewUser(u => ({ ...u, [key]:e.target.value }))} style={{ border:"1px solid #d1d5db", borderRadius:6, padding:"7px 10px", fontSize:13 }}>
                  {opts.map(o => <option key={o} value={o}>{o}</option>)}
                </select>
              </div>
            ))}
            <button onClick={addUser} style={{ background:"#0D1B3E", color:"#fff", border:"none", borderRadius:6, padding:"9px 16px", fontWeight:600, fontSize:13, cursor:"pointer" }}>+ Add</button>
          </div>
        </div>
        {/* User list */}
        {users.map(u => (
          <div key={u.id} style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, overflow:"hidden" }}>
            <div style={{ display:"flex", alignItems:"center", gap:12, padding:"12px 16px", background: u.active ? "#f9fafb" : "#fee2e2" }}>
              <div style={{ flex:1 }}>
                <span style={{ fontWeight:700, fontSize:14, color:"#0D1B3E" }}>{u.name}</span>
                <span style={{ marginLeft:10 }}><Badge label={u.role} bg="#4AACCC" /></span>
                {u.entity && <span style={{ marginLeft:6 }}><Badge label={u.entity} bg="#0D1B3E" /></span>}
                {!u.active && <span style={{ marginLeft:6 }}><Badge label="Inactive" bg="#dc2626" /></span>}
              </div>
              <button onClick={() => setExpandedUser(expandedUser === u.id ? null : u.id)} style={{ background:"#f3f4f6", border:"1px solid #e5e7eb", borderRadius:6, padding:"4px 10px", fontSize:11, cursor:"pointer", color:"#374151" }}>
                {expandedUser === u.id ? "Hide Permissions" : "Edit Permissions"}
              </button>
              {u.id !== 1 && (
                <button onClick={() => toggleActive(u.id)} style={{ background: u.active ? "#fee2e2" : "#dcfce7", border:"none", borderRadius:6, padding:"4px 10px", fontSize:11, cursor:"pointer", color: u.active ? "#991b1b" : "#166534" }}>
                  {u.active ? "Deactivate" : "Activate"}
                </button>
              )}
            </div>
            {expandedUser === u.id && (
              <div style={{ padding:14, borderTop:"1px solid #e5e7eb" }}>
                <div style={{ fontSize:11, fontWeight:600, color:"#6b7280", marginBottom:8 }}>TAB PERMISSIONS — none = hidden, view = read only, edit = full access</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))", gap:8 }}>
                  {TAB_LIST.map(tab => (
                    <div key={tab} style={{ display:"flex", alignItems:"center", gap:8, padding:"6px 10px", background:"#f9fafb", borderRadius:6 }}>
                      <span style={{ fontSize:12, flex:1, color:"#374151" }}>{tab}</span>
                      <select value={(u.permissions||{})[tab]||"none"} onChange={e => setPermission(u.id, tab, e.target.value)}
                        style={{ border:"1px solid #d1d5db", borderRadius:4, padding:"3px 6px", fontSize:11 }}>
                        <option value="none">Hidden</option>
                        <option value="view">View Only</option>
                        <option value="edit">Full Edit</option>
                      </select>
                    </div>
                  ))}
                </div>
                <button onClick={() => saveUsers(users)} style={{ marginTop:10, background:"#0D1B3E", color:"#fff", border:"none", borderRadius:6, padding:"8px 18px", fontWeight:600, fontSize:13, cursor:"pointer" }}>
                  Save Permissions
                </button>
              </div>
            )}
          </div>
        ))}
      </div>
    );
  }

  const subTabs = [["general","⚙️ General"],["items","📦 Item Master"],["customers","🏢 Customer Master"],["gr","📥 GR Types"],["gi","📤 GI Types"],["users","👥 Users & Access"]];

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      {saved && <div style={{ background:"#dcfce7", border:"1px solid #86efac", borderRadius:8, padding:"10px 14px", color:"#166534", fontWeight:600, fontSize:13 }}>{saved}</div>}
      <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
        {subTabs.map(([t,l]) => (
          <button key={t} onClick={() => setTab(t)} style={{ padding:"8px 16px", borderRadius:6, border:"1px solid", borderColor:tab===t?"#0D1B3E":"#d1d5db", background:tab===t?"#0D1B3E":"#fff", color:tab===t?"#fff":"#374151", fontWeight:600, fontSize:13, cursor:"pointer" }}>{l}</button>
        ))}
      </div>
      {tab === "general"   && <GeneralTab />}
      {tab === "items"     && <ItemMasterSetup itemMaster={itemMaster} setItemMaster={setItemMaster} />}
      {tab === "customers" && <CustomerMasterSetup customers={customers} setCustomers={setCustomers} />}
      {tab === "gr"        && <MovTypeTab direction="GR" />}
      {tab === "gi"        && <MovTypeTab direction="GI" />}
      {tab === "users"     && <UsersTab />}
    </div>
  );
}


// ─── AI INSIGHTS CHAT (GROQ) ─────────────────────────────────────────────────
function AIChatBox({ movements, personnel, fleetUnits, lockedMonths }) {
  const [messages, setMessages] = useState([
    { role:"assistant", content:"Hi! I'm your Fleet AI. Ask me anything about your demo fleet — utilisation, overdue units, movement trends, or recommendations. Try: *'Which units are overdue?'* or *'Summarise this month's fleet activity.'*" }
  ]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [open, setOpen] = useState(false);
  const bottomRef = React.useRef(null);

  React.useEffect(() => {
    if (open && bottomRef.current) bottomRef.current.scrollIntoView({ behavior:"smooth" });
  }, [messages, open]);

  function buildContext() {
    const now = new Date();
    const thisMonth = now.getFullYear()+"-"+String(now.getMonth()+1).padStart(2,"0");
    const inField = movements.filter(m => getDir(m.type)==="GI" && m.status!=="Returned");
    const overdue = movements.filter(m => isOverdue(m));
    const thisMonthGR = movements.filter(m => getDir(m.type)==="GR" && m.date?.startsWith(thisMonth));
    const thisMonthGI = movements.filter(m => getDir(m.type)==="GI" && m.date?.startsWith(thisMonth));
    return `You are a fleet management AI assistant for a robotics demo fleet tracker.
Current date: ${now.toDateString()}
Fleet summary:
- Total fleet units registered: ${fleetUnits.length}
- Units currently out in field: ${inField.length}
- Overdue returns: ${overdue.length}
- GR (received) this month: ${thisMonthGR.length}
- GI (issued) this month: ${thisMonthGI.length}
- Locked months: ${lockedMonths.join(", ") || "None"}
- Total personnel: ${personnel.length}
Overdue units: ${overdue.map(m=>`${getItemName(m.itemId)} SN:${m.serialNo||"?"} (custodian:${m.custodian}, due:${m.expectedReturn})`).join("; ")||"None"}
Units in field: ${inField.slice(0,10).map(m=>`${getItemName(m.itemId)} (${m.custodian}, ${m.entity||"?"}, since ${m.date})`).join("; ")||"None"}
Answer concisely and helpfully. Use bullet points when listing items. Focus on actionable insights.`;
  }

  async function sendMessage() {
    if (!input.trim() || loading) return;
    const userMsg = { role:"user", content:input.trim() };
    const newMessages = [...messages, userMsg];
    setMessages(newMessages);
    setInput("");
    setLoading(true);
    try {
      const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
        method:"POST",
        headers:{
          "Content-Type":"application/json",
          "Authorization":"Bearer " + (window.__GROQ_KEY__ || "")
        },
        body:JSON.stringify({
          model:"llama3-8b-8192",
          messages:[
            { role:"system", content:buildContext() },
            ...newMessages.map(m=>({ role:m.role, content:m.content }))
          ],
          max_tokens:500,
          temperature:0.7,
        })
      });
      const data = await res.json();
      const reply = data?.choices?.[0]?.message?.content || "Sorry, I couldn't get a response. Check your GROQ API key in Vercel environment variables.";
      setMessages(prev => [...prev, { role:"assistant", content:reply }]);
    } catch(err) {
      setMessages(prev => [...prev, { role:"assistant", content:"Connection error. Ensure GROQ_API_KEY is set in your Vercel environment variables and the app is deployed." }]);
    }
    setLoading(false);
  }

  const QUICK = ["Which units are overdue?","Summarise this month's activity","Which entity has highest utilisation?","List all units currently in field","Any units returning this week?"];

  return (
    <div style={{ position:"fixed", bottom:20, right:20, zIndex:200 }}>
      {/* Toggle button */}
      {!open && (
        <button onClick={()=>setOpen(true)}
          style={{ background:"linear-gradient(135deg,#0D1B3E,#4AACCC)", color:"#fff", border:"none", borderRadius:50, width:56, height:56, fontSize:22, cursor:"pointer", boxShadow:"0 4px 20px rgba(13,27,62,0.4)", display:"flex", alignItems:"center", justifyContent:"center" }}>
          🤖
        </button>
      )}

      {/* Chat panel */}
      {open && (
        <div style={{ width:360, background:"#fff", borderRadius:16, boxShadow:"0 8px 40px rgba(0,0,0,0.2)", display:"flex", flexDirection:"column", maxHeight:"70vh", overflow:"hidden" }}>
          {/* Header */}
          <div style={{ background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)", padding:"14px 16px", display:"flex", alignItems:"center", gap:10, borderRadius:"16px 16px 0 0" }}>
            <span style={{ fontSize:20 }}>🤖</span>
            <div style={{ flex:1 }}>
              <div style={{ color:"#fff", fontWeight:700, fontSize:14 }}>Fleet AI</div>
              <div style={{ color:"#C8EEF5", fontSize:11 }}>Powered by GROQ — ask anything about your fleet</div>
            </div>
            <button onClick={()=>setOpen(false)} style={{ background:"transparent", border:"none", color:"#C8EEF5", fontSize:18, cursor:"pointer" }}>✕</button>
          </div>

          {/* Messages */}
          <div style={{ flex:1, overflowY:"auto", padding:"12px 14px", display:"flex", flexDirection:"column", gap:10, minHeight:200, maxHeight:380 }}>
            {messages.map((msg, i) => (
              <div key={i} style={{ display:"flex", justifyContent:msg.role==="user"?"flex-end":"flex-start" }}>
                <div style={{
                  maxWidth:"85%", padding:"9px 13px", borderRadius:msg.role==="user"?"14px 14px 4px 14px":"14px 14px 14px 4px",
                  background:msg.role==="user"?"#0D1B3E":"#f1f5f9",
                  color:msg.role==="user"?"#fff":"#374151",
                  fontSize:13, lineHeight:1.5, whiteSpace:"pre-wrap"
                }}>
                  {msg.content}
                </div>
              </div>
            ))}
            {loading && (
              <div style={{ display:"flex", justifyContent:"flex-start" }}>
                <div style={{ background:"#f1f5f9", borderRadius:"14px 14px 14px 4px", padding:"9px 13px", fontSize:13, color:"#9ca3af" }}>
                  Thinking...
                </div>
              </div>
            )}
            <div ref={bottomRef} />
          </div>

          {/* Quick prompts */}
          <div style={{ padding:"8px 14px", borderTop:"1px solid #f3f4f6", display:"flex", gap:6, flexWrap:"wrap" }}>
            {QUICK.map(q => (
              <button key={q} onClick={()=>{ setInput(q); }}
                style={{ background:"#f0f9ff", color:"#0369a1", border:"1px solid #bae6fd", borderRadius:20, padding:"4px 10px", fontSize:10, cursor:"pointer", whiteSpace:"nowrap" }}>
                {q}
              </button>
            ))}
          </div>

          {/* Input */}
          <div style={{ padding:"10px 14px", borderTop:"1px solid #f3f4f6", display:"flex", gap:8 }}>
            <input value={input} onChange={e=>setInput(e.target.value)}
              onKeyDown={e=>e.key==="Enter"&&!e.shiftKey&&sendMessage()}
              placeholder="Ask about your fleet..."
              style={{ flex:1, border:"1px solid #d1d5db", borderRadius:20, padding:"8px 14px", fontSize:13, outline:"none" }} />
            <button onClick={sendMessage} disabled={loading||!input.trim()}
              style={{ background:loading||!input.trim()?"#e5e7eb":"#0D1B3E", color:loading||!input.trim()?"#9ca3af":"#fff", border:"none", borderRadius:20, padding:"8px 16px", fontWeight:700, fontSize:13, cursor:loading||!input.trim()?"not-allowed":"pointer" }}>
              Send
            </button>
          </div>
        </div>
      )}
    </div>
  );
}


// ─── VERSION HISTORY ──────────────────────────────────────────────────────────
const INITIAL_VERSIONS = [
];

const VERSION_KEY = "pudu_v2_versions";

function VersionHistory() {
  const [versions, setVersions] = useState([]);
  const [showForm, setShowForm] = useState(false);
  const [saved, setSaved] = useState(false);
  const blank = { version: "", date: today(), type: "Enhancement", changedBy: "", summary: "", changeText: "" };
  const [form, setForm] = useState(blank);

  useEffect(() => {
    load(VERSION_KEY, INITIAL_VERSIONS).then(v => setVersions(v));
  }, []);

  function setF(k, v) { setForm(f => ({ ...f, [k]: v })); }

  async function handleSave() {
    if (!form.version || !form.summary || !form.changedBy) {
      alert("Please fill in Version, Changed By, and Summary.");
      return;
    }
    const newEntry = {
      ...form,
      id: Date.now(),
      changes: form.changeText.split("\n").map(s => s.trim()).filter(Boolean),
    };
    const updated = [newEntry, ...versions];
    setVersions(updated);
    await save(VERSION_KEY, updated);
    setForm(blank);
    setShowForm(false);
    setSaved(true);
    setTimeout(() => setSaved(false), 3000);
  }

  async function delVersion(id) {
    if (!appConfirm("Delete this version entry?")) return;
    const updated = versions.filter(v => v.id !== id);
    setVersions(updated);
    await save(VERSION_KEY, updated);
  }

  const typeColors = {
    "Initial Release": "#0D1B3E",
    "Enhancement":     "#4AACCC",
    "Bug Fix":         "#16a34a",
    "New Feature":     "#7c3aed",
    "Remove Feature":  "#dc2626",
    "Configuration":   "#d97706",
    "Security Update": "#991b1b",
  };

  return (
    <div style={{ maxWidth: 860, display: "flex", flexDirection: "column", gap: 16 }}>
      {/* Header */}
      <div style={{ background: "linear-gradient(135deg,#0D1B3E,#1a3a6b)", borderRadius: 10, padding: "16px 20px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div>
          <div style={{ color: "#fff", fontWeight: 700, fontSize: 16 }}>Version History</div>
          <div style={{ color: "#C8EEF5", fontSize: 12, marginTop: 2 }}>Track every change made to this app after deployment</div>
        </div>
        <button onClick={() => setShowForm(s => !s)}
          style={{ background: showForm ? "#374151" : "#4AACCC", color: "#fff", border: "none", borderRadius: 8, padding: "9px 18px", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
          {showForm ? "Cancel" : "+ Log New Version"}
        </button>
      </div>

      {saved && (
        <div style={{ background: "#dcfce7", border: "1px solid #86efac", borderRadius: 8, padding: "10px 14px", color: "#166534", fontWeight: 600, fontSize: 13 }}>
          Version entry saved successfully.
        </div>
      )}

      {/* Add Version Form */}
      {showForm && (
        <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 18 }}>
          <SectionHead title="Log New Version / Change" />
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
            {[["Version Number *", "version", "text", "e.g. 1.1.0"],
              ["Date *", "date", "date", ""],
              ["Changed By *", "changedBy", "text", "Your name"]].map(([label, key, type, ph]) => (
              <div key={key} style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>{label}</label>
                <input type={type} value={form[key]} onChange={e => setF(key, e.target.value)}
                  placeholder={ph} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
              </div>
            ))}
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Change Type</label>
              <select value={form.type} onChange={e => setF("type", e.target.value)}
                style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }}>
                {Object.keys(typeColors).map(t => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
            <div style={{ gridColumn: "2/-1", display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Summary *</label>
              <input value={form.summary} onChange={e => setF("summary", e.target.value)}
                placeholder="Brief one-line description of this version"
                style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
            </div>
            <div style={{ gridColumn: "1/-1", display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>
                Changes Made (one per line)
              </label>
              <textarea value={form.changeText} onChange={e => setF("changeText", e.target.value)}
                rows={5} placeholder={"Added customer contact field to Log Movement form\nRemoved 6-month trend chart from Dashboard\nFixed overdue calculation for loans"}
                style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 10px", fontSize: 13, resize: "vertical", fontFamily: "inherit" }} />
              <span style={{ fontSize: 11, color: "#9ca3af" }}>Enter each change on a new line — they will appear as a bullet list</span>
            </div>
          </div>
          <button onClick={handleSave}
            style={{ background: "linear-gradient(135deg,#0D1B3E,#4AACCC)", color: "#fff", border: "none", borderRadius: 8, padding: "10px 24px", fontWeight: 700, fontSize: 14, cursor: "pointer" }}>
            Save Version Entry
          </button>
        </div>
      )}

      {/* How to use this tab */}
      <div style={{ background: "#f0f9ff", border: "1px solid #bae6fd", borderRadius: 8, padding: "12px 16px", fontSize: 12, color: "#0369a1" }}>
        <strong>How to use:</strong> Every time you update the app via Claude and re-upload to GitHub, come here and click "+ Log New Version" to record what changed, when, and who requested it. This creates a permanent audit trail for IT and management.
      </div>

      {/* Version Cards */}
      {versions.length === 0 && (
        <div style={{ padding: 24, textAlign: "center", color: "#9ca3af", fontSize: 13 }}>No version entries yet.</div>
      )}

      {versions.map((v, idx) => (
        <div key={v.id} style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, overflow: "hidden" }}>
          {/* Version header */}
          <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "12px 16px", background: "#f9fafb", borderBottom: "1px solid #e5e7eb" }}>
            <div style={{ background: "#0D1B3E", color: "#fff", borderRadius: 6, padding: "4px 12px", fontWeight: 700, fontSize: 13, fontFamily: "monospace" }}>
              v{v.version}
            </div>
            <div style={{ background: typeColors[v.type] || "#4AACCC", color: "#fff", borderRadius: 20, padding: "2px 10px", fontSize: 11, fontWeight: 600 }}>
              {v.type}
            </div>
            {idx === 0 && (
              <div style={{ background: "#dcfce7", color: "#166534", borderRadius: 20, padding: "2px 10px", fontSize: 11, fontWeight: 600, border: "1px solid #86efac" }}>
                Latest
              </div>
            )}
            <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 16 }}>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: "#374151" }}>{fmt(v.date)}</div>
                <div style={{ fontSize: 11, color: "#9ca3af" }}>by {v.changedBy}</div>
              </div>
              {v.id !== 1 && (
                <button onClick={() => delVersion(v.id)}
                  style={{ background: "#fee2e2", border: "1px solid #fca5a5", color: "#991b1b", borderRadius: 4, padding: "3px 8px", fontSize: 11, cursor: "pointer" }}>
                  Del
                </button>
              )}
            </div>
          </div>

          {/* Summary and changes */}
          <div style={{ padding: "12px 16px" }}>
            <div style={{ fontWeight: 600, fontSize: 13, color: "#0D1B3E", marginBottom: 8 }}>{v.summary}</div>
            {v.changes && v.changes.length > 0 && (
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                {v.changes.map((c, i) => (
                  <div key={i} style={{ display: "flex", gap: 8, alignItems: "flex-start" }}>
                    <div style={{ minWidth: 6, height: 6, borderRadius: "50%", background: typeColors[v.type] || "#4AACCC", marginTop: 5, flexShrink: 0 }} />
                    <span style={{ fontSize: 13, color: "#374151" }}>{c}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      ))}
    </div>
  );
}


// APP ROOT
export default function App(){
  const [activeTab,setActiveTab]   = useState("Stock Position");
  const [movements,setMovements]   = useState([]);
  const [personnel,setPersonnel]   = useState([]);
  const [sapSnapshots,setSapSnapshots] = useState([]);
  const [gitItems,setGitItems]     = useState([]);
  const [fleetUnits,setFleetUnits] = useState([]);
  const [versions,setVersions]     = useState([]);
  const [lockedMonths,setLockedMonths] = useState([]);
  const [setup,setSetup]           = useState(null);
  const [appUsers,setAppUsers]     = useState(DEFAULT_USERS);
  const [pendingItems,setPendingItems] = useState([]);
  const [itemMaster,setItemMaster]     = useState(DEFAULT_ITEMS);
  const [customers,setCustomers]       = useState(DEFAULT_CUSTOMERS);
  const [loading,setLoading]       = useState(true);
  const [showClearConfirm,setShowClearConfirm] = useState(false);
  const [sidebarOpen,setSidebarOpen] = useState(true);

  // Auth
  const [currentUser,setCurrentUser] = useState(null);
  const [pinInput,setPinInput]       = useState("");
  const [pinError,setPinError]       = useState("");
  const [showForm,setShowForm]       = useState(false); // public form mode

  useEffect(()=>{
    Promise.all([
      load(SK.MOV,SAMPLE_MOVEMENTS),
      load(SK.PERS,SAMPLE_PERSONNEL),
      load(SK.SAP,SAMPLE_SAP),
      load(SK.GIT,SAMPLE_GIT),
      load(SK.FLEET,SAMPLE_FLEET),
      load(VERSION_KEY,INITIAL_VERSIONS),
      load(SK.LOCK,[]),
      load(SK.SETUP,DEFAULT_SETUP),
      load(SK.USERS,DEFAULT_USERS).then(us => { if(!us[0]?.permissions?.["Pending Approvals"]) { localStorage.removeItem(SK.USERS); return DEFAULT_USERS; } return us; }),
      load(SK.PENDING,[]),
      load(SK.ITEMS, DEFAULT_ITEMS),
      load(SK.CUSTOMERS, DEFAULT_CUSTOMERS),
    ]).then(([m,p,s,g,fl,v,lk,st,us,pnd,itm,cust])=>{
      setMovements(m);setPersonnel(p);setSapSnapshots(s);setGitItems(g);
      setFleetUnits(fl);setVersions(v);setLockedMonths(lk);
      setSetup(st);setPendingItems(pnd);setItemMaster(itm);setCustomers(cust);
      // Always merge loaded user permissions with latest ROLE_PERMS
      // so new tabs added in updates are automatically included
      const mergedUsers = us.map(u => ({
        ...u,
        permissions: {
          ...(ROLE_PERMS[u.role] || {}),
          ...(u.permissions || {}),
          "Pending Approvals": ROLE_PERMS[u.role]?.["Pending Approvals"] || "none",
        }
      }));
      setAppUsers(mergedUsers);
      setLoading(false);
    });
  },[]);

  if(loading) return(
    <div style={{ display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",fontFamily:"Inter,sans-serif",flexDirection:"column",gap:16,background:"#0D1B3E" }}>
      <div style={{ fontSize:48 }}>🤖</div>
      <div style={{ color:"#C8EEF5",fontSize:16,fontWeight:600 }}>Loading Fleet Tracker...</div>
    </div>
  );

  const cfg = setup || DEFAULT_SETUP;
  const allTabs = TAB_LIST;
  const icons={"Stock Position":"📦","Fleet Calendar":"📅","Dashboard":"📊","Log Movement":"✏️","Movement Log":"📋","GIT Register":"🚢","Pending Approvals":"⏳","Admin":"🔒","Setup":"⚙️","Import SAP":"📥","Backup & Restore":"💾","Export Excel":"⬇","User Guide":"📖","Version History":"🕐"};

  // Permission helper
  function canAccess(tab) {
    if (!currentUser) return false;
    const perm = (currentUser.permissions||{})[tab];
    return perm === "edit" || perm === "view";
  }
  function canEdit(tab) {
    if (!currentUser) return false;
    return (currentUser.permissions||{})[tab] === "edit";
  }
  const visibleTabs = allTabs.filter(t => canAccess(t));

  // PIN Login screen
  if (!currentUser && !showForm) return (
    <div style={{ minHeight:"100vh",background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)",display:"flex",alignItems:"center",justifyContent:"center",padding:20,fontFamily:"'Inter','Segoe UI',sans-serif" }}>
      <div style={{ background:"#fff",borderRadius:16,padding:40,maxWidth:420,width:"100%",textAlign:"center",boxShadow:"0 20px 60px rgba(0,0,0,0.3)" }}>
        <div style={{ fontSize:48,marginBottom:8 }}>🤖</div>
        <div style={{ fontSize:22,fontWeight:800,color:"#0D1B3E",marginBottom:4 }}>{cfg.appName}</div>
        <div style={{ fontSize:13,color:"#6b7280",marginBottom:28 }}>Admin Access — Enter your PIN</div>
        <div style={{ display:"flex",gap:10,justifyContent:"center",marginBottom:16 }}>
          {[1,2,3,4].map(i=>(
            <div key={i} style={{ width:48,height:56,border:"2px solid",borderColor:pinInput.length>=i?"#0D1B3E":"#e5e7eb",borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,fontWeight:700,color:"#0D1B3E" }}>
              {pinInput.length>=i?"●":""}
            </div>
          ))}
        </div>
        {pinError && <div style={{ color:"#dc2626",fontSize:13,marginBottom:12 }}>{pinError}</div>}
        <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10 }}>
          {[1,2,3,4,5,6,7,8,9,"✕",0,"✓"].map((k,i)=>(
            <button key={i} onClick={()=>{
              if(k==="✕"){ setPinInput(p=>p.slice(0,-1)); setPinError(""); return; }
              if(k==="✓"){
                const u = appUsers.find(u=>u.pin===pinInput && u.active);
                if(u){ setCurrentUser(u); setPinInput(""); setPinError(""); }
                else { setPinError("Incorrect PIN. Try again."); setPinInput(""); }
                return;
              }
              if(pinInput.length<4) setPinInput(p=>p+k);
            }}
            style={{ padding:"16px",borderRadius:10,border:"1px solid #e5e7eb",background:k==="✓"?"#0D1B3E":k==="✕"?"#fee2e2":"#f9fafb",color:k==="✓"?"#fff":k==="✕"?"#991b1b":"#374151",fontSize:20,fontWeight:700,cursor:"pointer" }}>
              {k}
            </button>
          ))}
        </div>
        <div style={{ borderTop:"1px solid #e5e7eb",paddingTop:16,marginTop:8 }}>
          <div style={{ fontSize:12,color:"#9ca3af",marginBottom:10 }}>Not an admin? Submit a movement:</div>
          <button onClick={()=>setShowForm(true)} style={{ width:"100%",background:"#4AACCC",color:"#0D1B3E",border:"none",borderRadius:8,padding:"11px",fontWeight:700,fontSize:14,cursor:"pointer" }}>
            📝 Open Movement Form
          </button>
        </div>
      </div>
    </div>
  );

  // Public form mode (for sales/marketing/AE)
  if (showForm) return (
    <div>
      <div style={{ position:"fixed",top:0,right:0,zIndex:100 }}>
        <button onClick={()=>setShowForm(false)} style={{ margin:12,background:"rgba(0,0,0,0.5)",color:"#fff",border:"none",borderRadius:6,padding:"6px 12px",fontSize:12,cursor:"pointer" }}>
          ✕ Close Form
        </button>
      </div>
      <PublicForm movements={movements} setMovements={setMovements} personnel={personnel} setup={cfg} pendingItems={pendingItems} setPendingItems={setPendingItems} itemMaster={itemMaster} customers={customers} />
    </div>
  );

  // Main app (authenticated)
  const SIDEBAR_W = sidebarOpen ? 220 : 56;

  return(
    <div style={{ fontFamily:"'Inter','Segoe UI',sans-serif",minHeight:"100vh",background:"#f1f5f9",display:"flex",color:"#111827" }}>

      {/* SIDEBAR */}
      <div style={{ width:SIDEBAR_W,minWidth:SIDEBAR_W,background:"#0D1B3E",display:"flex",flexDirection:"column",minHeight:"100vh",position:"fixed",left:0,top:0,bottom:0,zIndex:50,transition:"width 0.2s" }}>
        {/* Logo */}
        <div style={{ padding:"16px 12px",borderBottom:"1px solid rgba(255,255,255,0.1)",display:"flex",alignItems:"center",gap:10 }}>
          <div style={{ width:32,height:32,background:"linear-gradient(135deg,#C8EEF5,#4AACCC)",borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0 }}>🤖</div>
          {sidebarOpen && <div style={{ color:"#fff",fontWeight:700,fontSize:13,lineHeight:1.2,overflow:"hidden" }}>{cfg.appName}</div>}
          <button onClick={()=>setSidebarOpen(o=>!o)} style={{ marginLeft:"auto",background:"transparent",border:"none",color:"#9ca3af",cursor:"pointer",fontSize:16,flexShrink:0 }}>
            {sidebarOpen?"◀":"▶"}
          </button>
        </div>

        {/* Nav items */}
        <div style={{ flex:1,overflowY:"auto",padding:"8px 0" }}>
          {visibleTabs.map(tab=>{
            const pc = tab==="Pending Approvals" ? pendingItems.filter(p=>(p.approvalStatus||"pending")==="pending").length : 0;
            const isPA = tab==="Pending Approvals";
            const isActive = activeTab===tab;
            return (
              <button key={tab} onClick={()=>setActiveTab(tab)}
                style={{ width:"100%",display:"flex",alignItems:"center",gap:8,
                  padding:sidebarOpen?"10px 16px":"10px 0",
                  border:"none",
                  background: isActive?"rgba(74,172,204,0.2)": isPA&&pc>0?"rgba(220,38,38,0.15)":"transparent",
                  cursor:"pointer",textAlign:"left",
                  borderLeft: isActive?"3px solid #4AACCC": isPA&&pc>0?"3px solid #dc2626":"3px solid transparent" }}>
                <span style={{ fontSize:16,flexShrink:0,marginLeft:sidebarOpen?0:"auto",marginRight:sidebarOpen?0:"auto" }}>{icons[tab]||"•"}</span>
                {sidebarOpen && <span style={{ color:isActive?"#C8EEF5":isPA&&pc>0?"#fca5a5":"#9ca3af",fontSize:13,fontWeight:isActive||isPA?600:400,whiteSpace:"nowrap",overflow:"hidden",flex:1 }}>{tab}</span>}
                {pc>0 && <span style={{ background:"#dc2626",color:"#fff",borderRadius:10,padding:"1px 7px",fontSize:10,fontWeight:700,flexShrink:0 }}>{pc}</span>}
              </button>
            );
          })}
        </div>

        {/* User info + actions */}
        <div style={{ padding:"12px",borderTop:"1px solid rgba(255,255,255,0.1)" }}>
          {sidebarOpen && (
            <div style={{ marginBottom:8,padding:"8px 10px",background:"rgba(255,255,255,0.06)",borderRadius:6 }}>
              <div style={{ color:"#fff",fontSize:12,fontWeight:600 }}>{currentUser?.name}</div>
              <div style={{ color:"#9ca3af",fontSize:10 }}>{currentUser?.role}</div>
            </div>
          )}
          <div style={{ display:"flex",gap:6,flexDirection:sidebarOpen?"row":"column" }}>
            <button onClick={()=>setShowForm(true)} title="Open Form" style={{ flex:1,background:"#4AACCC",color:"#0D1B3E",border:"none",borderRadius:6,padding:"7px 4px",fontSize:sidebarOpen?11:14,fontWeight:700,cursor:"pointer" }}>
              {sidebarOpen?"📝 Form":"📝"}
            </button>
            <button onClick={()=>setCurrentUser(null)} title="Log Out" style={{ flex:1,background:"rgba(220,38,38,0.2)",color:"#fca5a5",border:"none",borderRadius:6,padding:"7px 4px",fontSize:sidebarOpen?11:14,fontWeight:700,cursor:"pointer" }}>
              {sidebarOpen?"🚪 Logout":"🚪"}
            </button>
          </div>
          {sidebarOpen && (
            <button onClick={()=>{
              const url=window.location.href;
              const ta=document.createElement("textarea");ta.value=url;ta.style.position="fixed";ta.style.opacity="0";document.body.appendChild(ta);ta.focus();ta.select();
              try{document.execCommand("copy");alert("Link copied! Share with your team.");}catch(e){alert("App link: "+url);}
              document.body.removeChild(ta);
            }} style={{ width:"100%",marginTop:6,background:"transparent",color:"#6b7280",border:"1px solid rgba(255,255,255,0.1)",borderRadius:6,padding:"6px",fontSize:11,cursor:"pointer" }}>
              🔗 Share Link
            </button>
          )}
          {sidebarOpen && canEdit("Backup & Restore") && (
            <button onClick={()=>setShowClearConfirm(true)} style={{ width:"100%",marginTop:4,background:"transparent",color:"#6b7280",border:"1px solid rgba(255,255,255,0.1)",borderRadius:6,padding:"6px",fontSize:11,cursor:"pointer" }}>
              🗑 Clear Data
            </button>
          )}
          {sidebarOpen && (
            <div style={{ marginTop:8,fontSize:9,color:"rgba(255,255,255,0.2)",textAlign:"center",lineHeight:1.5,padding:"0 4px" }}>
              © 2026 FleetTrack Pro<br/>Proprietary Software<br/>Unauthorised duplication<br/>is prohibited
            </div>
          )}
        </div>
      </div>

      {/* MAIN CONTENT */}
      <div style={{ marginLeft:SIDEBAR_W,flex:1,display:"flex",flexDirection:"column",transition:"margin-left 0.2s" }}>
        {/* Top bar */}
        <div style={{ background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"12px 24px",display:"flex",alignItems:"center",gap:12,position:"sticky",top:0,zIndex:40 }}>
          <div style={{ fontWeight:700,fontSize:16,color:"#0D1B3E",flex:1 }}>{activeTab}</div>
          {!canEdit(activeTab) && canAccess(activeTab) && (
            <div style={{ background:"#fef3c7",border:"1px solid #fbbf24",borderRadius:6,padding:"4px 10px",fontSize:11,fontWeight:600,color:"#92400e" }}>👁 View Only</div>
          )}
          <div style={{ fontSize:12,color:"#9ca3af" }}>{new Date().toLocaleDateString("en-MY",{weekday:"short",day:"2-digit",month:"short",year:"numeric"})}</div>
        </div>

        {/* Tab content */}
        <div style={{ padding:"24px",flex:1 }}>
          {activeTab==="Stock Position"    && <StockPosition movements={movements} sapSnapshots={sapSnapshots} gitItems={gitItems} />}
          {activeTab==="Fleet Calendar"    && <FleetCalendar movements={movements} fleetUnits={fleetUnits} setFleetUnits={setFleetUnits} />}
          {activeTab==="Dashboard"         && <Dashboard movements={movements} personnel={personnel} />}
          {activeTab==="Log Movement"      && <LogMovement movements={movements} setMovements={setMovements} personnel={personnel} lockedMonths={lockedMonths} setup={cfg} fleetUnits={fleetUnits} setFleetUnits={setFleetUnits} itemMaster={itemMaster} customers={customers} readOnly={!canEdit("Log Movement")} />}
          {activeTab==="Movement Log"      && <MovementLog movements={movements} setMovements={setMovements} lockedMonths={lockedMonths} readOnly={!canEdit("Movement Log")} />}
          {activeTab==="GIT Register"      && <GITRegister gitItems={gitItems} setGitItems={setGitItems} readOnly={!canEdit("GIT Register")} />}
          {activeTab==="Pending Approvals" && <PendingApprovals pendingItems={pendingItems} setPendingItems={setPendingItems} movements={movements} setMovements={setMovements} />}
          {activeTab==="Admin"             && <Admin movements={movements} setMovements={setMovements} personnel={personnel} setPersonnel={setPersonnel} sapSnapshots={sapSnapshots} setSapSnapshots={setSapSnapshots} lockedMonths={lockedMonths} setLockedMonths={setLockedMonths} setup={cfg} setSetup={setSetup} currentUser={currentUser} itemMaster={itemMaster} setItemMaster={setItemMaster} customers={customers} setCustomers={setCustomers} fleetUnits={fleetUnits} setFleetUnits={setFleetUnits} versions={versions} setVersions={setVersions} gitItems={gitItems} setGitItems={setGitItems} />}
          {activeTab==="Setup"             && <Setup setup={cfg} setSetup={setSetup} appUsers={appUsers} setAppUsers={setAppUsers} itemMaster={itemMaster} setItemMaster={setItemMaster} customers={customers} setCustomers={setCustomers} />}
          {activeTab==="Import SAP"        && <SapImport sapSnapshots={sapSnapshots} setSapSnapshots={setSapSnapshots} gitItems={gitItems} setGitItems={setGitItems} />}
          {activeTab==="Backup & Restore"  && <Admin movements={movements} setMovements={setMovements} personnel={personnel} setPersonnel={setPersonnel} sapSnapshots={sapSnapshots} setSapSnapshots={setSapSnapshots} lockedMonths={lockedMonths} setLockedMonths={setLockedMonths} setup={cfg} setSetup={setSetup} currentUser={currentUser} itemMaster={itemMaster} setItemMaster={setItemMaster} customers={customers} setCustomers={setCustomers} fleetUnits={fleetUnits} setFleetUnits={setFleetUnits} versions={versions} setVersions={setVersions} gitItems={gitItems} setGitItems={setGitItems} />}
          {activeTab==="Export Excel"      && <ExportExcel movements={movements} personnel={personnel} sapSnapshots={sapSnapshots} gitItems={gitItems} />}
          {activeTab==="User Guide"        && <UserGuide />}
          {activeTab==="Version History"   && <VersionHistory />}
        </div>
      </div>

      {/* AI CHAT BOX */}
      <AIChatBox movements={movements} personnel={personnel} fleetUnits={fleetUnits} lockedMonths={lockedMonths} />

      {/* CLEAR DATA MODAL */}
      {showClearConfirm && (
        <div style={{ position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.6)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:20 }}>
          <div style={{ background:"#fff",borderRadius:12,padding:28,maxWidth:360,width:"100%",boxShadow:"0 20px 60px rgba(0,0,0,0.3)" }}>
            <div style={{ fontSize:32,textAlign:"center",marginBottom:12 }}>⚠️</div>
            <div style={{ fontWeight:700,fontSize:16,color:"#0D1B3E",textAlign:"center",marginBottom:8 }}>Clear All Data?</div>
            <div style={{ fontSize:13,color:"#6b7280",textAlign:"center",marginBottom:20 }}>This will delete all movements, personnel, SAP snapshots, GIT entries, fleet units and version history. This cannot be undone.</div>
            <div style={{ display:"flex",gap:10 }}>
              <button onClick={()=>setShowClearConfirm(false)} style={{ flex:1,background:"#f3f4f6",color:"#374151",border:"none",borderRadius:8,padding:"11px 0",fontWeight:600,fontSize:14,cursor:"pointer" }}>Cancel</button>
              <button onClick={()=>{
                [SK.MOV,SK.PERS,SK.SAP,SK.GIT,SK.FLEET,SK.LOCK,SK.PENDING,VERSION_KEY].forEach(k=>localStorage.removeItem(k));
                setMovements([]);setPersonnel([]);setSapSnapshots([]);setGitItems([]);setFleetUnits([]);setVersions([]);setLockedMonths([]);setPendingItems([]);
                setShowClearConfirm(false);setActiveTab("Stock Position");
              }} style={{ flex:1,background:"#dc2626",color:"#fff",border:"none",borderRadius:8,padding:"11px 0",fontWeight:700,fontSize:14,cursor:"pointer" }}>
                Yes, Clear All
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
