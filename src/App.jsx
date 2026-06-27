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
const ENTITIES = ["TMK", "TMS"];


// ─── STORAGE KEYS ─────────────────────────────────────────────────────────────
const SK = { MOV: "pudu_v2_movements", PERS: "pudu_v2_personnel", SAP: "pudu_v2_sap_snapshots", GIT: "pudu_v2_git", FLEET: "pudu_v2_fleet" };

// ─── SAMPLE DATA ─────────────────────────────────────────────────────────────
const SAMPLE_PERSONNEL = [
  { id: 1, name: "Ahmad Farid", role: "Sales Personnel", region: "Kuala Lumpur", entity: "TMK" },
  { id: 2, name: "Lim Wei Ling", role: "Engineer", region: "Selangor", entity: "TMK" },
  { id: 3, name: "Priya Nair", role: "Marketing", region: "Penang", entity: "TMK" },
  { id: 4, name: "Jason Tan", role: "Sales Personnel", region: "Johor Bahru", entity: "TMS" },
  { id: 5, name: "Siti Rahimah", role: "Sales Personnel", region: "Sabah", entity: "TMS" },
  { id: 6, name: "David Chong", role: "Engineer", region: "Selangor", entity: "TMS" },
  { id: 7, name: "Logistics Admin", role: "Logistics Admin", region: "Selangor", entity: "TMK" },
];

const SAMPLE_SAP = [
  { id: 2001, month: "2026-04", itemId: "BELLABOT", sapStock: 5, sapGit: 2, sapDemo: 2, notes: "April opening" },
  { id: 2002, month: "2026-04", itemId: "CC1", sapStock: 3, sapGit: 0, sapDemo: 1, notes: "" },
  { id: 2003, month: "2026-04", itemId: "KETTYBOT", sapStock: 4, sapGit: 1, sapDemo: 1, notes: "" },
  { id: 2004, month: "2026-05", itemId: "BELLABOT", sapStock: 4, sapGit: 0, sapDemo: 2, notes: "After May sales" },
  { id: 2005, month: "2026-05", itemId: "CC1", sapStock: 3, sapGit: 0, sapDemo: 1, notes: "" },
  { id: 2006, month: "2026-06", itemId: "BELLABOT", sapStock: 3, sapGit: 0, sapDemo: 3, notes: "Jun snapshot" },
  { id: 2007, month: "2026-06", itemId: "HOLABOT", sapStock: 2, sapGit: 1, sapDemo: 1, notes: "New June batch" },
  { id: 2008, month: "2026-06", itemId: "MATRADEE", sapStock: 1, sapGit: 0, sapDemo: 1, notes: "" },
];

const SAMPLE_GIT = [
  { id: 3001, poNo: "PO-2026-0041", itemId: "BELLABOT-PRO", qty: 3, supplier: "Pudu Robotics HQ", shipDate: "2026-06-10", etaDate: "2026-07-05", vesselRef: "MSCTIANJIN-026W", status: "Shipped", notes: "ETA Klang Port" },
  { id: 3002, poNo: "PO-2026-0042", itemId: "SWIFTBOT", qty: 2, supplier: "Pudu Robotics HQ", shipDate: "2026-06-18", etaDate: "2026-07-12", vesselRef: "COSCO-EVER-088", status: "Shipped", notes: "Awaiting customs clearance" },
  { id: 3003, poNo: "PO-2026-0038", itemId: "ACC-DOCK", qty: 5, supplier: "Pudu Accessories", shipDate: "2026-05-30", etaDate: "2026-06-20", vesselRef: "", status: "Received", notes: "GR posted in SAP 21 Jun" },
  { id: 3004, poNo: "PO-2026-0045", itemId: "KETTYBOT-PRO", qty: 2, supplier: "Pudu Robotics HQ", shipDate: "", etaDate: "2026-07-25", vesselRef: "", status: "On Order", notes: "PO confirmed, not yet shipped" },
];

const SAMPLE_MOVEMENTS = [
  // ── TMK entity movements ──────────────────────────────────────────────────
  { id:1001,type:"GR_SUPPLIER",itemId:"T300-BLACK",sapCat:"STOCK",qty:1,serialNo:"8260B5318060021",custodian:"Logistics Admin",entity:"TMK",region:"Kuala Lumpur",location:"KL office",date:"2026-04-02",expectedReturn:"",customerName:"",eventName:"",notes:"",status:"Active",createdAt:"2026-04-02T09:00:00Z" },
  { id:1002,type:"GR_SUPPLIER",itemId:"T300-BLACK",sapCat:"STOCK",qty:1,serialNo:"8260B5318060016",custodian:"Logistics Admin",entity:"TMK",region:"Kuala Lumpur",location:"KL office",date:"2026-04-02",expectedReturn:"",customerName:"",eventName:"",notes:"",status:"Active",createdAt:"2026-04-02T09:05:00Z" },
  { id:1003,type:"GR_SUPPLIER",itemId:"T300-BLACK",sapCat:"STOCK",qty:1,serialNo:"8260B5318060032",custodian:"Logistics Admin",entity:"TMK",region:"Penang",location:"PEN office",date:"2026-04-03",expectedReturn:"",customerName:"",eventName:"",notes:"",status:"Active",createdAt:"2026-04-03T09:00:00Z" },
  { id:1004,type:"GI_LOAN_CUSTOMER",itemId:"T300-BLACK",sapCat:"DEMO",qty:1,serialNo:"8260B5318060045",custodian:"Ahmad Farid",entity:"TMK",region:"Kuala Lumpur",location:"SHRC",date:"2026-04-10",expectedReturn:"2026-11-28",customerName:"SHRC",eventName:"",notes:"Loan to SHRC since 28/11/2025",status:"Active",createdAt:"2026-04-10T09:00:00Z" },
  { id:1005,type:"GI_DEMO_SALES",itemId:"T300-BLACK",sapCat:"DEMO",qty:1,serialNo:"8260B5318060051",custodian:"Ahmad Farid",entity:"TMK",region:"Kuala Lumpur",location:"KL office",date:"2026-05-01",expectedReturn:"2026-07-10",customerName:"",eventName:"",notes:"KL Demo",status:"Active",createdAt:"2026-05-01T09:00:00Z" },
  { id:1006,type:"GI_DEMO_SALES",itemId:"T300-BLACK",sapCat:"DEMO",qty:1,serialNo:"8260B5210060001",custodian:"Priya Nair",entity:"TMK",region:"Penang",location:"Flex P8 Penang",date:"2026-05-15",expectedReturn:"2026-07-20",customerName:"Flex P8",eventName:"",notes:"KL Demo — Flex P8 Penang",status:"Active",createdAt:"2026-05-15T09:00:00Z" },
  { id:1007,type:"GR_SUPPLIER",itemId:"T600-UNDERRIDE",sapCat:"STOCK",qty:1,serialNo:"83A045919070003",custodian:"Logistics Admin",entity:"TMK",region:"Penang",location:"PEN office",date:"2026-04-05",expectedReturn:"",customerName:"",eventName:"",notes:"PEN DEMO",status:"Active",createdAt:"2026-04-05T09:00:00Z" },
  { id:1008,type:"GI_EXHIBITION",itemId:"T600-UNDERRIDE",sapCat:"DEMO",qty:1,serialNo:"83A045919070001",custodian:"Ahmad Farid",entity:"TMK",region:"Kuala Lumpur",location:"Automex",date:"2026-05-22",expectedReturn:"2026-06-05",customerName:"",eventName:"AUTOMEX",notes:"AUTOMEX exhibition",status:"Active",createdAt:"2026-05-22T09:00:00Z" },
  { id:1009,type:"GI_EXHIBITION",itemId:"T600-UNDERRIDE",sapCat:"DEMO",qty:1,serialNo:"83A045919070004",custodian:"Ahmad Farid",entity:"TMK",region:"Kuala Lumpur",location:"KL office",date:"2026-05-22",expectedReturn:"2026-06-05",customerName:"",eventName:"AUTOMEX",notes:"AUTOMEX exhibition",status:"Active",createdAt:"2026-05-22T09:05:00Z" },
  { id:1010,type:"GI_DEMO_SALES",itemId:"T600-WTIDL2",sapCat:"DEMO",qty:1,serialNo:"8.31046E+14",custodian:"Ahmad Farid",entity:"TMK",region:"Kuala Lumpur",location:"KL office",date:"2026-06-01",expectedReturn:"2026-07-15",customerName:"",eventName:"",notes:"KL Demo",status:"Active",createdAt:"2026-06-01T09:00:00Z" },
  { id:1011,type:"GI_DEMO_SALES",itemId:"MT1-MAX",sapCat:"DEMO",qty:1,serialNo:"8.68066E+14",custodian:"Lim Wei Ling",entity:"TMK",region:"Kuala Lumpur",location:"SFP Automation",date:"2026-06-03",expectedReturn:"2026-07-03",customerName:"SFP Automation",eventName:"",notes:"KL Demo",status:"Active",createdAt:"2026-06-03T09:00:00Z" },
  { id:1012,type:"GI_DEMO_SALES",itemId:"MT1-VAC",sapCat:"DEMO",qty:1,serialNo:"8.99066E+14",custodian:"Lim Wei Ling",entity:"TMK",region:"Kuala Lumpur",location:"Automex",date:"2026-06-03",expectedReturn:"2026-06-28",customerName:"",eventName:"AUTOMEX",notes:"Loan to Minke",status:"Active",createdAt:"2026-06-03T09:05:00Z" },
  { id:1013,type:"GI_DEMO_SALES",itemId:"MT1-VAC-CHARGING",sapCat:"DEMO",qty:1,serialNo:"5.67066E+14",custodian:"Lim Wei Ling",entity:"TMK",region:"Kuala Lumpur",location:"Automex",date:"2026-06-03",expectedReturn:"2026-06-28",customerName:"",eventName:"AUTOMEX",notes:"Loan to Minke",status:"Active",createdAt:"2026-06-03T09:10:00Z" },
  { id:1014,type:"GI_DEMO_SALES",itemId:"BELLABOT",sapCat:"DEMO",qty:1,serialNo:"SV1114224449002",custodian:"Ahmad Farid",entity:"TMK",region:"Kuala Lumpur",location:"KL office",date:"2026-06-05",expectedReturn:"2026-07-05",customerName:"",eventName:"",notes:"KL Demo",status:"Active",createdAt:"2026-06-05T09:00:00Z" },
  { id:1015,type:"GR_SUPPLIER",itemId:"CC1",sapCat:"STOCK",qty:1,serialNo:"81100492405001",custodian:"Logistics Admin",entity:"TMK",region:"Kuala Lumpur",location:"KL office",date:"2026-05-01",expectedReturn:"",customerName:"",eventName:"",notes:"",status:"Active",createdAt:"2026-05-01T09:00:00Z" },
  { id:1016,type:"GI_POC_ENGINEER",itemId:"CC1-PRO",sapCat:"DEMO",qty:1,serialNo:"8880B5805070001",custodian:"Lim Wei Ling",entity:"TMK",region:"Penang",location:"PEN office",date:"2026-06-10",expectedReturn:"2026-07-10",customerName:"",eventName:"",notes:"POC at Edz Titan",status:"Active",createdAt:"2026-06-10T09:00:00Z" },
  { id:1017,type:"GR_SUPPLIER",itemId:"HOLABOT",sapCat:"STOCK",qty:1,serialNo:"66076318060002",custodian:"Logistics Admin",entity:"TMK",region:"Kuala Lumpur",location:"UDA warehouse",date:"2026-05-01",expectedReturn:"",customerName:"",eventName:"",notes:"KL Stock",status:"Active",createdAt:"2026-05-01T09:00:00Z" },
  { id:1018,type:"GR_SUPPLIER",itemId:"MT1",sapCat:"STOCK",qty:1,serialNo:"86607631806002",custodian:"Logistics Admin",entity:"TMK",region:"Kuala Lumpur",location:"UDA warehouse",date:"2026-05-01",expectedReturn:"",customerName:"",eventName:"",notes:"KL Stock",status:"Active",createdAt:"2026-05-01T09:10:00Z" },
  { id:1019,type:"GI_POC_ENGINEER",itemId:"MT1",sapCat:"DEMO",qty:1,serialNo:"86607631806002",custodian:"Lim Wei Ling",entity:"TMK",region:"Kuala Lumpur",location:"IOI",date:"2026-06-15",expectedReturn:"2026-07-15",customerName:"IOI",eventName:"",notes:"POC at IOI",status:"Active",createdAt:"2026-06-15T09:00:00Z" },
  { id:1020,type:"GI_DEMO_SALES",itemId:"PUDUBOT2",sapCat:"DEMO",qty:1,serialNo:"85809641308043",custodian:"Priya Nair",entity:"TMK",region:"Kuala Lumpur",location:"UDA warehouse",date:"2026-06-01",expectedReturn:"2026-07-15",customerName:"",eventName:"",notes:"KL Demo",status:"Active",createdAt:"2026-06-01T09:00:00Z" },
  { id:1021,type:"GI_DEMO_SALES",itemId:"T150",sapCat:"DEMO",qty:1,serialNo:"82303631906003",custodian:"Ahmad Farid",entity:"TMK",region:"Kuala Lumpur",location:"UDA warehouse",date:"2026-06-01",expectedReturn:"2026-07-10",customerName:"",eventName:"",notes:"KL Demo",status:"Active",createdAt:"2026-06-01T09:05:00Z" },
  // ── TMS entity movements ──────────────────────────────────────────────────
  { id:2001,type:"GR_SUPPLIER",itemId:"T300-LIFTING",sapCat:"STOCK",qty:1,serialNo:"8260B410050001",custodian:"Logistics Admin",entity:"TMS",region:"Selangor",location:"TMS office",date:"2026-04-01",expectedReturn:"",customerName:"",eventName:"",notes:"",status:"Active",createdAt:"2026-04-01T09:00:00Z" },
  { id:2002,type:"GI_POC_ENGINEER",itemId:"T300-TOWING",sapCat:"DEMO",qty:1,serialNo:"8260B5318060001",custodian:"Jason Tan",entity:"TMS",region:"Selangor",location:"Agilent",date:"2026-06-01",expectedReturn:"2026-06-30",customerName:"Agilent",eventName:"",notes:"POC at Agilent",status:"Active",createdAt:"2026-06-01T09:00:00Z" },
  { id:2003,type:"GR_SUPPLIER",itemId:"T300-CONVEYOR",sapCat:"STOCK",qty:1,serialNo:"8.26095E+14",custodian:"Logistics Admin",entity:"TMS",region:"Selangor",location:"TMS office",date:"2026-04-01",expectedReturn:"",customerName:"",eventName:"",notes:"",status:"Active",createdAt:"2026-04-01T09:00:00Z" },
  { id:2004,type:"GR_SUPPLIER",itemId:"T300",sapCat:"STOCK",qty:1,serialNo:"8260B5318060040",custodian:"Logistics Admin",entity:"TMS",region:"Selangor",location:"TMS office",date:"2026-04-01",expectedReturn:"",customerName:"",eventName:"",notes:"",status:"Active",createdAt:"2026-04-01T09:05:00Z" },
  { id:2005,type:"GR_SUPPLIER",itemId:"PUDUBOT2",sapCat:"STOCK",qty:1,serialNo:"8170D4625050030",custodian:"Logistics Admin",entity:"TMS",region:"Selangor",location:"TMS office",date:"2026-04-01",expectedReturn:"",customerName:"",eventName:"",notes:"",status:"Active",createdAt:"2026-04-01T09:10:00Z" },
  { id:2006,type:"GR_SUPPLIER",itemId:"MT1",sapCat:"STOCK",qty:1,serialNo:"8.66075E+14",custodian:"Logistics Admin",entity:"TMS",region:"Selangor",location:"TMS office",date:"2026-04-01",expectedReturn:"",customerName:"",eventName:"",notes:"",status:"Active",createdAt:"2026-04-01T09:15:00Z" },
  { id:2007,type:"GR_SUPPLIER",itemId:"FLASHBOT",sapCat:"STOCK",qty:1,serialNo:"8FF0D5627050002",custodian:"Logistics Admin",entity:"TMS",region:"Selangor",location:"TMS office",date:"2026-04-01",expectedReturn:"",customerName:"",eventName:"",notes:"",status:"Active",createdAt:"2026-04-01T09:20:00Z" },
  { id:2008,type:"GI_POC_ENGINEER",itemId:"CC1-PRO",sapCat:"DEMO",qty:1,serialNo:"8880B5805070004",custodian:"David Chong",entity:"TMS",region:"Selangor",location:"PT Flex Batam",date:"2026-06-10",expectedReturn:"2026-06-25",customerName:"PT Flex Batam",eventName:"",notes:"POC at PT Flex Batam",status:"Active",createdAt:"2026-06-10T09:00:00Z" },
  { id:2009,type:"GI_POC_ENGINEER",itemId:"T600-STANDARD",sapCat:"DEMO",qty:1,serialNo:"8.31046E+14",custodian:"David Chong",entity:"TMS",region:"Selangor",location:"TMS Office",date:"2026-06-15",expectedReturn:"2026-07-15",customerName:"",eventName:"",notes:"",status:"Active",createdAt:"2026-06-15T09:00:00Z" },
  { id:2010,type:"GI_POC_ENGINEER",itemId:"T600-UNDERRIDE",sapCat:"DEMO",qty:1,serialNo:"83A045919070002",custodian:"Siti Rahimah",entity:"TMS",region:"Selangor",location:"Agilent",date:"2026-06-10",expectedReturn:"2026-06-28",customerName:"Agilent",eventName:"",notes:"POC at Agilent",status:"Active",createdAt:"2026-06-10T09:00:00Z" },
];

// ─── SAMPLE FLEET UNITS ──────────────────────────────────────────────────────
const SAMPLE_FLEET = [
  // TMK units
  { id: 4001, entity:"TMK", model:"PUDU T300 (BLACK)", variant:"WTIDL1", serialNo:"8260B5318060021", location:"KL office", remark:"KL Demo" },
  { id: 4002, entity:"TMK", model:"PUDU T300 (BLACK)", variant:"WTIDL1", serialNo:"8260B5318060016", location:"KL office", remark:"KL Demo" },
  { id: 4003, entity:"TMK", model:"PUDU T300 (BLACK)", variant:"WTIDL1", serialNo:"8260B5318060032", location:"PEN office", remark:"PEN DEMO" },
  { id: 4004, entity:"TMK", model:"PUDU T300 (BLACK)", variant:"WTIDL1", serialNo:"8260B5318060045", location:"Loan to SHRC", remark:"Loan" },
  { id: 4005, entity:"TMK", model:"PUDU T300 (BLACK)", variant:"WTIDL1", serialNo:"8260B5318060051", location:"KL office", remark:"KL Demo" },
  { id: 4006, entity:"TMK", model:"PUDU T300 (BLACK)", variant:"WTIDL1", serialNo:"8260B5210060001", location:"Flex P8 Penang", remark:"KL Demo" },
  { id: 4007, entity:"TMK", model:"PUDU T300 (BLACK)", variant:"WTIDU01", serialNo:"826045318060001", location:"KL office", remark:"KL Demo" },
  { id: 4008, entity:"TMK", model:"PUDU T600 UNDERRIDE", variant:"WTIDU2", serialNo:"83A045919070003", location:"PEN office", remark:"PEN DEMO" },
  { id: 4009, entity:"TMK", model:"PUDU T600 UNDERRIDE", variant:"WTIDU2", serialNo:"83A045919070001", location:"in Automex", remark:"KL Demo" },
  { id: 4010, entity:"TMK", model:"PUDU T600 UNDERRIDE", variant:"WTIDU2", serialNo:"83A045919070004", location:"KL office", remark:"KL Demo" },
  { id: 4011, entity:"TMK", model:"PUDU T600 UNDERRIDE", variant:"WTIDU2", serialNo:"83A045919070001", location:"in Automex", remark:"KL Demo" },
  { id: 4012, entity:"TMK", model:"PUDU T600", variant:"WTIDL2", serialNo:"8.31046E+14", location:"in Automex", remark:"KL Demo" },
  { id: 4013, entity:"TMK", model:"PUDU T600", variant:"WTIDL2", serialNo:"8.31046E+14", location:"KL office", remark:"KL Demo" },
  { id: 4014, entity:"TMK", model:"PUDU MT1 MAX", variant:"MTBC03", serialNo:"8.68066E+14", location:"SFP Automation", remark:"KL Demo" },
  { id: 4015, entity:"TMK", model:"PUDU MT1 VAC", variant:"MTBC02", serialNo:"8.99066E+14", location:"in Automex", remark:"KL Demo" },
  { id: 4016, entity:"TMK", model:"PUDU MT1 VAC Charging Station", variant:"", serialNo:"5.67066E+14", location:"in Automex", remark:"KL Demo" },
  { id: 4017, entity:"TMK", model:"BellaBot (Used Unit)", variant:"", serialNo:"SV111422444P002", location:"KL office", remark:"KL Demo" },
  { id: 4018, entity:"TMK", model:"PUDU CC1", variant:"", serialNo:"811004924050001", location:"KL office", remark:"KL Demo" },
  { id: 4019, entity:"TMK", model:"PUDU CC1 PRO", variant:"", serialNo:"8880B5805070001", location:"PEN office", remark:"KL Demo" },
  { id: 4020, entity:"TMK", model:"PUDU CC1 Docking Station", variant:"", serialNo:"88SD01462B050015", location:"PEN office", remark:"KL Demo" },
  { id: 4021, entity:"TMK", model:"PUDU Mobile Water Station", variant:"", serialNo:"5MS01512060042", location:"PEN office", remark:"KL Demo" },
  { id: 4022, entity:"TMK", model:"PUDU CC1 PRO", variant:"", serialNo:"888126318060001", location:"UDA warehouse", remark:"KL Stock" },
  { id: 4023, entity:"TMK", model:"PUDU MT1", variant:"", serialNo:"866076318060002", location:"UDA warehouse", remark:"KL Stock" },
  { id: 4024, entity:"TMK", model:"Pudu BG1", variant:"", serialNo:"858096413080043", location:"UDA warehouse", remark:"KL Demo" },
  { id: 4025, entity:"TMK", model:"PUDU T150", variant:"", serialNo:"823036319060003", location:"UDA warehouse", remark:"KL Demo" },
  // TMS units
  { id: 4101, entity:"TMS", model:"T300 - Lifting", variant:"", serialNo:"8260B410050001", location:"TMS office", remark:"TMS Demo" },
  { id: 4102, entity:"TMS", model:"T300 - Lifting Charging Station", variant:"", serialNo:"6b004851510029", location:"TMS office", remark:"TMS Demo" },
  { id: 4103, entity:"TMS", model:"T300 - Towing", variant:"", serialNo:"8260B5318060001", location:"Agilent", remark:"POC" },
  { id: 4104, entity:"TMS", model:"T300 - Towing Charging Station", variant:"", serialNo:"5CM005610060254", location:"Agilent", remark:"POC" },
  { id: 4105, entity:"TMS", model:"T300 - Conveyor", variant:"", serialNo:"8.26095E+14", location:"TMS office", remark:"TMS Demo" },
  { id: 4106, entity:"TMS", model:"T300 - Conveyor Charging Station", variant:"", serialNo:"6b004851510026", location:"TMS office", remark:"TMS Demo" },
  { id: 4107, entity:"TMS", model:"T300", variant:"", serialNo:"8260B5318060040", location:"TMS office", remark:"TMS Demo" },
  { id: 4108, entity:"TMS", model:"Pudubot 2", variant:"", serialNo:"8170D4625050030", location:"TMS office", remark:"TMS Demo" },
  { id: 4109, entity:"TMS", model:"MT1", variant:"", serialNo:"8.66075E+14", location:"TMS office", remark:"TMS Demo" },
  { id: 4110, entity:"TMS", model:"MT1 Charging Station", variant:"", serialNo:"567004C20050046", location:"TMS office", remark:"TMS Demo" },
  { id: 4111, entity:"TMS", model:"FLASHBOT", variant:"", serialNo:"8FF0D5627050002", location:"TMS office", remark:"TMS Demo" },
  { id: 4112, entity:"TMS", model:"CC1-Pro", variant:"", serialNo:"8880B505070004", location:"PT Flex @ Batam", remark:"POC" },
  { id: 4113, entity:"TMS", model:"T600 Standard", variant:"", serialNo:"8.31046E+14", location:"TMS Office", remark:"TMS Demo" },
  { id: 4114, entity:"TMS", model:"T600 UNDERRIDE", variant:"", serialNo:"83A045919070002", location:"Agilent", remark:"POC" },
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
const today = () => new Date().toISOString().split("T")[0];
const fmt = (d) => d ? new Date(d).toLocaleDateString("en-MY", { day: "2-digit", month: "short", year: "numeric" }) : "—";
const isOverdue = (m) => m.status !== "Returned" && m.expectedReturn && new Date(m.expectedReturn) < new Date();
const getDir = (t) => t?.startsWith("GR") ? "GR" : "GI";
const getMovLabel = (t) => ALL_MOVEMENT_TYPES.find(m => m.value === t)?.label || t;
const getItemName = (id) => PUDU_ITEMS.find(i => i.id === id)?.name || id;
const getItemCat = (id) => PUDU_ITEMS.find(i => i.id === id)?.category || "";
const getSapCatInfo = (id) => SAP_STOCK_CATEGORIES.find(c => c.id === id) || {};

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
      {/* Legend */}
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

      {/* Model Breakdown this month */}
      <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
        <SectionHead title={"Model Breakdown — " + thisMonth} />
        {(() => {
          const byModel2 = {};
          monthMov.forEach(m=>{
            if(!byModel2[m.itemId]) byModel2[m.itemId]={gr:0,gi:0};
            if(getDir(m.type)==="GR") byModel2[m.itemId].gr+=m.qty;
            else byModel2[m.itemId].gi+=m.qty;
          });
          const rows = Object.entries(byModel2).sort((a,b)=>(b[1].gr+b[1].gi)-(a[1].gr+a[1].gi));
          if(rows.length===0) return <p style={{ color:"#9ca3af",fontSize:13 }}>No movements this month.</p>;
          return (
            <table style={{ width:"100%",borderCollapse:"collapse",fontSize:13 }}>
              <thead style={{ background:"#f9fafb" }}>
                <tr>
                  {["Model","GR (In)","GI (Out)","Net Movement"].map(h=>(
                    <th key={h} style={{ padding:"8px 10px",textAlign:h==="Model"?"left":"center",fontWeight:600,color:"#374151",borderBottom:"1px solid #e5e7eb" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map(([id,{gr,gi}])=>{
                  const net=gr-gi;
                  return(
                    <tr key={id} style={{ borderTop:"1px solid #f3f4f6" }}>
                      <td style={{ padding:"8px 10px",fontWeight:600,color:"#0D1B3E" }}>{getItemName(id)}</td>
                      <td style={{ padding:"8px 10px",textAlign:"center",color:"#16a34a",fontWeight:700 }}>+{gr}</td>
                      <td style={{ padding:"8px 10px",textAlign:"center",color:"#dc2626",fontWeight:700 }}>-{gi}</td>
                      <td style={{ padding:"8px 10px",textAlign:"center",fontWeight:700,color:net>0?"#16a34a":net<0?"#dc2626":"#6b7280" }}>
                        {net>0?"+":""}{net}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot style={{ background:"#f9fafb" }}>
                <tr style={{ borderTop:"2px solid #e5e7eb" }}>
                  <td style={{ padding:"8px 10px",fontWeight:700 }}>TOTAL</td>
                  <td style={{ padding:"8px 10px",textAlign:"center",fontWeight:700,color:"#16a34a" }}>+{rows.reduce((s,[,v])=>s+v.gr,0)}</td>
                  <td style={{ padding:"8px 10px",textAlign:"center",fontWeight:700,color:"#dc2626" }}>-{rows.reduce((s,[,v])=>s+v.gi,0)}</td>
                  <td style={{ padding:"8px 10px",textAlign:"center",fontWeight:700,color:"#0D1B3E" }}>
                    {(()=>{const n=rows.reduce((s,[,v])=>s+v.gr-v.gi,0);return(n>0?"+":"")+n;})()}
                  </td>
                </tr>
              </tfoot>
            </table>
          );
        })()}
      </div>

      {/* 6-month trend */}
      {(() => {
        const months = [];
        for(let i=5;i>=0;i--){
          const d=new Date(); d.setMonth(d.getMonth()-i);
          const m=d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0");
          const mm=movements.filter(x=>x.date?.startsWith(m));
          months.push({ m, gr:mm.filter(x=>getDir(x.type)==="GR").reduce((s,x)=>s+x.qty,0), gi:mm.filter(x=>getDir(x.type)==="GI").reduce((s,x)=>s+x.qty,0) });
        }
        const maxVal=Math.max(...months.flatMap(x=>[x.gr,x.gi]),1);
        return(
          <div style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:16 }}>
            <SectionHead title="6-Month GR vs GI Trend (by units)" />
            <div style={{ display:"flex",gap:8,alignItems:"flex-end",height:120,padding:"0 4px" }}>
              {months.map(({m,gr,gi})=>(
                <div key={m} style={{ flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:3 }}>
                  <div style={{ width:"100%",display:"flex",gap:2,alignItems:"flex-end",height:90 }}>
                    <div title={"GR: "+gr} style={{ flex:1,background:"#16a34a",borderRadius:"3px 3px 0 0",height:gr===0?2:Math.max(4,(gr/maxVal)*88),transition:"height 0.3s" }} />
                    <div title={"GI: "+gi} style={{ flex:1,background:"#dc2626",borderRadius:"3px 3px 0 0",height:gi===0?2:Math.max(4,(gi/maxVal)*88),transition:"height 0.3s" }} />
                  </div>
                  <div style={{ fontSize:9,color:"#6b7280",textAlign:"center",whiteSpace:"nowrap" }}>{m.slice(5)}/{m.slice(2,4)}</div>
                  <div style={{ fontSize:9,color:"#374151",textAlign:"center" }}>
                    <span style={{ color:"#16a34a",fontWeight:600 }}>{gr}</span>/<span style={{ color:"#dc2626",fontWeight:600 }}>{gi}</span>
                  </div>
                </div>
              ))}
            </div>
            <div style={{ display:"flex",gap:16,marginTop:8,justifyContent:"center" }}>
              <div style={{ display:"flex",alignItems:"center",gap:4,fontSize:11,color:"#374151" }}><div style={{ width:12,height:12,background:"#16a34a",borderRadius:2 }} />GR (In)</div>
              <div style={{ display:"flex",alignItems:"center",gap:4,fontSize:11,color:"#374151" }}><div style={{ width:12,height:12,background:"#dc2626",borderRadius:2 }} />GI (Out)</div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}

// ─── LOG MOVEMENT ─────────────────────────────────────────────────────────────
function LogMovement({ movements, setMovements, personnel }) {
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
    const newMov = { ...form, id: Date.now(), qty: Number(form.qty) || 1, createdAt: new Date().toISOString() };
    const updated = [newMov, ...movements];
    setMovements(updated);
    await save(SK.MOV, updated);
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
            <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Item / Model *</label>
            <select value={form.itemId} onChange={e => set("itemId", e.target.value)}
              style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "8px 10px", fontSize: 13 }}>
              <option value="">— Select Item —</option>
              <optgroup label="Robots">{PUDU_ITEMS.filter(i => i.category === "Robot").map(i => <option key={i.id} value={i.id}>{i.name}</option>)}</optgroup>
              <optgroup label="Accessories">{PUDU_ITEMS.filter(i => i.category === "Accessory").map(i => <option key={i.id} value={i.id}>{i.name}</option>)}</optgroup>
            </select>
          </div>
          {inp("Serial No. / Asset Tag", "serialNo", "text", { placeholder: "e.g. PD-BB-2026-001" })}
          {inp("Quantity *", "qty", "number", { min: 1 })}
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
          {inp("Location / Site", "location", "text", { placeholder: "e.g. KL office, UDA warehouse, Automex" })}
          {needsCustomer && inp("Customer Name", "customerName", "text", { placeholder: "e.g. Restaurant Maju Sdn Bhd" })}
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
function MovementLog({ movements, setMovements }) {
  const [fDir, setFDir] = useState("ALL"), [fType, setFType] = useState("ALL"), [fMonth, setFMonth] = useState(""), [fSapCat, setFSapCat] = useState("ALL"), [search, setSearch] = useState("");

  const filtered = useMemo(() => movements.filter(m => {
    if (fDir !== "ALL" && getDir(m.type) !== fDir) return false;
    if (fType !== "ALL" && m.type !== fType) return false;
    if (fSapCat !== "ALL" && m.sapCat !== fSapCat) return false;
    if (fMonth && !m.date?.startsWith(fMonth)) return false;
    if (search && !getItemName(m.itemId).toLowerCase().includes(search.toLowerCase()) && !m.custodian?.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  }), [movements, fDir, fType, fSapCat, fMonth, search]);

  async function markReturned(id) {
    const updated = movements.map(m => m.id === id ? { ...m, status: "Returned", actualReturn: today() } : m);
    setMovements(updated); await save(SK.MOV, updated);
  }
  async function del(id) {
    if (!window.confirm("Delete this record?")) return;
    const updated = movements.filter(m => m.id !== id); setMovements(updated); await save(SK.MOV, updated);
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
            <tr>{["Dir", "SAP Cat", "Movement Type", "Item", "S/N", "Qty", "Custodian", "Region", "Date", "Exp.Return", "Status", "Actions"].map(h => (
              <th key={h} style={{ padding: "9px 8px", textAlign: "left", fontWeight: 600, color: "#374151", borderBottom: "1px solid #e5e7eb", whiteSpace: "nowrap" }}>{h}</th>
            ))}</tr>
          </thead>
          <tbody>
            {filtered.length === 0 && <tr><td colSpan={12} style={{ padding: 24, textAlign: "center", color: "#9ca3af" }}>No records.</td></tr>}
            {filtered.map(m => {
              const over = isOverdue(m);
              const cat = getSapCatInfo(m.sapCat);
              return (
                <tr key={m.id} style={{ borderTop: "1px solid #f3f4f6", background: over ? "#fff7ed" : "transparent" }}>
                  <td style={{ padding: "7px 8px" }}><Badge label={getDir(m.type)} bg={getDir(m.type) === "GR" ? "#16a34a" : "#dc2626"} /></td>
                  <td style={{ padding: "7px 8px" }}><Badge label={m.sapCat || "—"} bg={cat.color || "#9ca3af"} /></td>
                  <td style={{ padding: "7px 8px", fontSize: 11, color: "#374151", maxWidth: 150 }}>{getMovLabel(m.type)}</td>
                  <td style={{ padding: "7px 8px", fontWeight: 600, color: "#0D1B3E", whiteSpace: "nowrap" }}>{getItemName(m.itemId)}</td>
                  <td style={{ padding: "7px 8px", color: "#6b7280", fontSize: 11 }}>{m.serialNo || "—"}</td>
                  <td style={{ padding: "7px 8px", textAlign: "center", fontWeight: 700 }}>{m.qty}</td>
                  <td style={{ padding: "7px 8px", whiteSpace: "nowrap" }}>{m.custodian}</td>
                  <td style={{ padding: "7px 8px", whiteSpace: "nowrap" }}>{m.region}</td>
                  <td style={{ padding: "7px 8px", whiteSpace: "nowrap" }}>{fmt(m.date)}</td>
                  <td style={{ padding: "7px 8px", whiteSpace: "nowrap", color: over ? "#dc2626" : "#374151" }}>{over ? "⚠ " : ""}{fmt(m.expectedReturn)}</td>
                  <td style={{ padding: "7px 8px" }}><Badge label={m.status === "Returned" ? "Returned" : over ? "Overdue" : "Active"} bg={m.status === "Returned" ? "#16a34a" : over ? "#E07820" : "#4AACCC"} /></td>
                  <td style={{ padding: "7px 8px", whiteSpace: "nowrap" }}>
                    {m.status !== "Returned" && getDir(m.type) === "GI" && (
                      <button onClick={() => markReturned(m.id)} style={{ background: "#dcfce7", border: "1px solid #86efac", color: "#166534", borderRadius: 4, padding: "3px 7px", fontSize: 11, cursor: "pointer", marginRight: 4 }}>Return</button>
                    )}
                    <button onClick={() => del(m.id)} style={{ background: "#fee2e2", border: "1px solid #fca5a5", color: "#991b1b", borderRadius: 4, padding: "3px 7px", fontSize: 11, cursor: "pointer" }}>Del</button>
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

// ─── GIT REGISTER ─────────────────────────────────────────────────────────────
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
    if (!window.confirm("Delete GIT record?")) return;
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
function Admin({ movements, personnel, setPersonnel, sapSnapshots, setSapSnapshots }) {
  const [tab, setTab] = useState("sap");
  const [newP, setNewP] = useState({ name: "", role: "Sales Personnel", region: "" });
  const [snapForm, setSnapForm] = useState({ month: new Date().toISOString().slice(0, 7), itemId: "", sapStock: 0, sapGit: 0, sapDemo: 0, notes: "" });

  async function addPersonnel() {
    if (!newP.name) return;
    const updated = [...personnel, { ...newP, id: Date.now() }];
    setPersonnel(updated); await save(SK.PERS, updated);
    setNewP({ name: "", role: "Sales Personnel", region: "" });
  }
  async function removePersonnel(id) {
    const updated = personnel.filter(p => p.id !== id); setPersonnel(updated); await save(SK.PERS, updated);
  }
  async function addSnapshot() {
    if (!snapForm.itemId || !snapForm.month) { alert("Month and Item required."); return; }
    const existing = sapSnapshots.find(s => s.month === snapForm.month && s.itemId === snapForm.itemId);
    let updated;
    if (existing) updated = sapSnapshots.map(s => s.month === snapForm.month && s.itemId === snapForm.itemId ? { ...s, ...snapForm, id: s.id } : s);
    else updated = [{ ...snapForm, id: Date.now(), sapStock: Number(snapForm.sapStock), sapGit: Number(snapForm.sapGit), sapDemo: Number(snapForm.sapDemo) }, ...sapSnapshots];
    setSapSnapshots(updated); await save(SK.SAP, updated);
    setSnapForm(f => ({ ...f, itemId: "", sapStock: 0, sapGit: 0, sapDemo: 0, notes: "" }));
  }
  async function delSnap(id) {
    const updated = sapSnapshots.filter(s => s.id !== id); setSapSnapshots(updated); await save(SK.SAP, updated);
  }

  const byPerson = {};
  movements.forEach(m => {
    if (!byPerson[m.custodian]) byPerson[m.custodian] = { total: 0, inField: 0, overdue: 0 };
    byPerson[m.custodian].total++;
    if (getDir(m.type) === "GI" && m.status !== "Returned") byPerson[m.custodian].inField++;
    if (isOverdue(m)) byPerson[m.custodian].overdue++;
  });

  const month = snapForm.month;
  const monthSnaps = sapSnapshots.filter(s => s.month === month);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        {[["sap","SAP Snapshots"],["monthly","Monthly Summary"],["personnel","Personnel"],["team","Team Activity"]].map(([t,l]) => (
          <button key={t} onClick={() => setTab(t)} style={{ padding: "8px 16px", borderRadius: 6, border: "1px solid", borderColor: tab === t ? "#0D1B3E" : "#d1d5db", background: tab === t ? "#0D1B3E" : "#fff", color: tab === t ? "#fff" : "#374151", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>{l}</button>
        ))}
      </div>

      {tab === "sap" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
            <SectionHead title="Enter SAP Month-End Stock Snapshot" />
            <p style={{ fontSize: 12, color: "#6b7280", margin: "0 0 12px" }}>After month-end close in SAP, enter the closing stock for each item. This feeds the Stock Position tab.</p>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Month *</label>
                <input type="month" value={snapForm.month} onChange={e => setSnapForm(f => ({ ...f, month: e.target.value }))} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
              </div>
              <div style={{ gridColumn: "2 / -1", display: "flex", flexDirection: "column", gap: 4 }}>
                <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Item *</label>
                <select value={snapForm.itemId} onChange={e => setSnapForm(f => ({ ...f, itemId: e.target.value }))} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }}>
                  <option value="">— Select Item —</option>
                  <optgroup label="Robots">{PUDU_ITEMS.filter(i => i.category === "Robot").map(i => <option key={i.id} value={i.id}>{i.name}</option>)}</optgroup>
                  <optgroup label="Accessories">{PUDU_ITEMS.filter(i => i.category === "Accessory").map(i => <option key={i.id} value={i.id}>{i.name}</option>)}</optgroup>
                </select>
              </div>
              {[["SAP Stock (unrestricted)", "sapStock", "#16a34a"], ["SAP GIT", "sapGit", "#d97706"], ["SAP Demo", "sapDemo", "#7c3aed"]].map(([l, k, c]) => (
                <div key={k} style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  <label style={{ fontSize: 12, fontWeight: 600, color: c }}>{l}</label>
                  <input type="number" min={0} value={snapForm[k]} onChange={e => setSnapForm(f => ({ ...f, [k]: e.target.value }))} style={{ border: `1px solid ${c}60`, borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
                </div>
              ))}
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                <label style={{ fontSize: 12, fontWeight: 600, color: "#374151" }}>Notes</label>
                <input value={snapForm.notes} onChange={e => setSnapForm(f => ({ ...f, notes: e.target.value }))} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
              </div>
              <div style={{ alignSelf: "flex-end" }}>
                <button onClick={addSnapshot} style={{ background: "#0D1B3E", color: "#fff", border: "none", borderRadius: 6, padding: "9px 18px", fontWeight: 600, fontSize: 13, cursor: "pointer", width: "100%" }}>Save Snapshot</button>
              </div>
            </div>
          </div>
          <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 16 }}>
            <SectionHead title={`SAP Snapshots — ${month}`} />
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead style={{ background: "#f9fafb" }}>
                <tr>{["Item", "SAP Stock", "SAP GIT", "SAP Demo", "Notes", ""].map(h => <th key={h} style={{ padding: "8px 10px", textAlign: h === "Item" || h === "Notes" || h === "" ? "left" : "center", fontWeight: 600, color: "#374151", borderBottom: "1px solid #e5e7eb" }}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {monthSnaps.length === 0 && <tr><td colSpan={6} style={{ padding: 16, textAlign: "center", color: "#9ca3af" }}>No snapshots for {month} yet.</td></tr>}
                {monthSnaps.map(s => (
                  <tr key={s.id} style={{ borderTop: "1px solid #f3f4f6" }}>
                    <td style={{ padding: "8px 10px", fontWeight: 600 }}>{getItemName(s.itemId)}</td>
                    <td style={{ padding: "8px 10px", textAlign: "center" }}><Badge label={s.sapStock} bg="#16a34a" /></td>
                    <td style={{ padding: "8px 10px", textAlign: "center" }}><Badge label={s.sapGit} bg={s.sapGit > 0 ? "#d97706" : "#9ca3af"} /></td>
                    <td style={{ padding: "8px 10px", textAlign: "center" }}><Badge label={s.sapDemo} bg={s.sapDemo > 0 ? "#7c3aed" : "#9ca3af"} /></td>
                    <td style={{ padding: "8px 10px", fontSize: 12, color: "#6b7280" }}>{s.notes}</td>
                    <td style={{ padding: "8px 10px" }}><button onClick={() => delSnap(s.id)} style={{ background: "#fee2e2", border: "1px solid #fca5a5", color: "#991b1b", borderRadius: 4, padding: "3px 7px", fontSize: 11, cursor: "pointer" }}>Del</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tab === "monthly" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <label style={{ fontWeight: 600, fontSize: 13 }}>Month:</label>
            <input type="month" value={snapForm.month} onChange={e => setSnapForm(f => ({ ...f, month: e.target.value }))} style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "7px 10px", fontSize: 13 }} />
          </div>
          {(() => {
            const mm = movements.filter(m => m.date?.startsWith(snapForm.month));
            const grByType = {}, giByType = {}, byModel = {};
            mm.forEach(m => {
              if (getDir(m.type) === "GR") grByType[m.type] = (grByType[m.type] || 0) + m.qty;
              else giByType[m.type] = (giByType[m.type] || 0) + m.qty;
              if (!byModel[m.itemId]) byModel[m.itemId] = { gr: 0, gi: 0 };
              if (getDir(m.type) === "GR") byModel[m.itemId].gr += m.qty;
              else byModel[m.itemId].gi += m.qty;
            });
            return (
              <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                  {[["GR Summary", grByType, MOVEMENT_TYPES.GR, "#16a34a"], ["GI Summary", giByType, MOVEMENT_TYPES.GI, "#dc2626"]].map(([title, data, types, col]) => (
                    <div key={title}                     style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 14 }}>
                      <SectionHead title={title + " — Total: " + Object.values(data).reduce((s,v)=>s+v,0) + " units"} />
                      {Object.keys(data).length === 0 ? <p style={{ color: "#9ca3af", fontSize: 13 }}>None this month.</p> :
                        Object.entries(data).map(([t, qty]) => (
                          <div key={t} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: "1px solid #f3f4f6" }}>
                            <span style={{ fontSize: 12 }}>{types.find(x=>x.value===t)?.label||t}</span>
                            <Badge label={qty} bg={col} />
                          </div>
                        ))}
                    </div>
                  ))}
                </div>
                <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 14 }}>
                  <SectionHead title="Model Breakdown" />
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead style={{ background: "#f9fafb" }}><tr>
                      {["Model","GR (In)","GI (Out)","Net"].map(h=><th key={h} style={{ padding:"7px 10px",textAlign:h==="Model"?"left":"center",fontWeight:600,color:"#374151",borderBottom:"1px solid #e5e7eb" }}>{h}</th>)}
                    </tr></thead>
                    <tbody>
                      {Object.keys(byModel).length===0 && <tr><td colSpan={4} style={{ padding:14,textAlign:"center",color:"#9ca3af" }}>No data.</td></tr>}
                      {Object.entries(byModel).map(([id,{gr,gi}])=>(
                        <tr key={id} style={{ borderTop:"1px solid #f3f4f6" }}>
                          <td style={{ padding:"7px 10px",fontWeight:600 }}>{getItemName(id)}</td>
                          <td style={{ padding:"7px 10px",textAlign:"center",color:"#16a34a",fontWeight:700 }}>+{gr}</td>
                          <td style={{ padding:"7px 10px",textAlign:"center",color:"#dc2626",fontWeight:700 }}>-{gi}</td>
                          <td style={{ padding:"7px 10px",textAlign:"center",fontWeight:700,color:gr-gi>=0?"#0D1B3E":"#dc2626" }}>{gr-gi>=0?"+":""}{gr-gi}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })()}
        </div>
      )}

      {tab === "personnel" && (
        <div style={{ display:"flex",flexDirection:"column",gap:14 }}>
          <div style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:14 }}>
            <SectionHead title="Add Personnel" />
            <div style={{ display:"flex",gap:10,flexWrap:"wrap" }}>
              <input value={newP.name} onChange={e=>setNewP(p=>({...p,name:e.target.value}))} placeholder="Full name" style={{ border:"1px solid #d1d5db",borderRadius:6,padding:"7px 10px",fontSize:13 }} />
              <select value={newP.role} onChange={e=>setNewP(p=>({...p,role:e.target.value}))} style={{ border:"1px solid #d1d5db",borderRadius:6,padding:"7px 10px",fontSize:13 }}>
                {ROLES.map(r=><option key={r} value={r}>{r}</option>)}
              </select>
              <select value={newP.entity||""} onChange={e=>setNewP(p=>({...p,entity:e.target.value}))} style={{ border:"1px solid #d1d5db",borderRadius:6,padding:"7px 10px",fontSize:13 }}>
                <option value="">Entity</option>{ENTITIES.map(e=><option key={e} value={e}>{e}</option>)}
              </select>
              <select value={newP.region} onChange={e=>setNewP(p=>({...p,region:e.target.value}))} style={{ border:"1px solid #d1d5db",borderRadius:6,padding:"7px 10px",fontSize:13 }}>
                <option value="">Region</option>{REGIONS.map(r=><option key={r} value={r}>{r}</option>)}
              </select>
              <button onClick={addPersonnel} style={{ background:"#0D1B3E",color:"#fff",border:"none",borderRadius:6,padding:"9px 16px",fontWeight:600,fontSize:13,cursor:"pointer" }}>+ Add</button>
            </div>
          </div>
          <div style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:14 }}>
            <SectionHead title="Personnel List" />
            <table style={{ width:"100%",borderCollapse:"collapse",fontSize:13 }}>
              <thead style={{ background:"#f9fafb" }}><tr>{["Name","Role","Entity","Region",""].map(h=><th key={h} style={{ padding:"8px 10px",textAlign:"left",fontWeight:600,borderBottom:"1px solid #e5e7eb" }}>{h}</th>)}</tr></thead>
              <tbody>
                {personnel.map(p=>(
                  <tr key={p.id} style={{ borderTop:"1px solid #f3f4f6" }}>
                    <td style={{ padding:"8px 10px",fontWeight:600 }}>{p.name}</td>
                    <td style={{ padding:"8px 10px" }}><Badge label={p.role} bg="#4AACCC" /></td>
                    <td style={{ padding:"8px 10px" }}>{p.entity ? <Badge label={p.entity} bg={p.entity==="TMK"?"#0D1B3E":"#4AACCC"} /> : "—"}</td>
                    <td style={{ padding:"8px 10px",color:"#6b7280" }}>{p.region||"—"}</td>
                    <td style={{ padding:"8px 10px" }}><button onClick={()=>removePersonnel(p.id)} style={{ background:"#fee2e2",border:"1px solid #fca5a5",color:"#991b1b",borderRadius:4,padding:"3px 8px",fontSize:11,cursor:"pointer" }}>Remove</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tab === "team" && (
        <div style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:14 }}>
          <SectionHead title="Team Activity (All Time)" />
          <table style={{ width:"100%",borderCollapse:"collapse",fontSize:13 }}>
            <thead style={{ background:"#f9fafb" }}><tr>{["Name","Role","Region","Total Movements","Units In Field","Overdue"].map(h=><th key={h} style={{ padding:"8px 10px",textAlign:"left",fontWeight:600,borderBottom:"1px solid #e5e7eb" }}>{h}</th>)}</tr></thead>
            <tbody>
              {personnel.map(p=>{
                const s=byPerson[p.name]||{total:0,inField:0,overdue:0};
                return(
                  <tr key={p.id} style={{ borderTop:"1px solid #f3f4f6" }}>
                    <td style={{ padding:"8px 10px",fontWeight:600 }}>{p.name}</td>
                    <td style={{ padding:"8px 10px" }}><Badge label={p.role} bg="#4AACCC" /></td>
                    <td style={{ padding:"8px 10px",color:"#6b7280" }}>{p.region}</td>
                    <td style={{ padding:"8px 10px",textAlign:"center",fontWeight:700 }}>{s.total}</td>
                    <td style={{ padding:"8px 10px",textAlign:"center" }}><Badge label={s.inField} bg="#0D1B3E" /></td>
                    <td style={{ padding:"8px 10px",textAlign:"center" }}>{s.overdue>0?<Badge label={s.overdue} bg="#E07820" />:"—"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ─── SAP IMPORT ───────────────────────────────────────────────────────────────
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
      "Status":isOverdue(m)?"Overdue":(m.status||"Active"),"Notes":m.notes||"",
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
    {who:"Marketing - Exhibition OUT and Return",color:"#7c3aed",bg:"#f5f3ff",steps:["Click Log Movement tab","Click Goods Issue (OUT)","Select Exhibition Out - Marketing","Fill in Event Name (e.g. MIHAS 2026 KLCC)","Fill Serial No, Qty (can be more than 1 for exhibition), Date, Expected Return","To return after event: Log Movement - Goods Receipt IN - Return Exhibition Marketing"]},
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
  const [newUnit, setNewUnit] = useState({ entity:"TMK", model:"", variant:"", serialNo:"", location:"", remark:"" });

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
      label: m.customerName || m.eventName || m.custodian || "",
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
    if (!window.confirm("Remove this unit from fleet?")) return;
    const updated = fleetUnits.filter(u => u.id !== id);
    setFleetUnits(updated); await save(SK.FLEET, updated);
  }
  async function addUnit() {
    if (!newUnit.model || !newUnit.entity) { alert("Model and Entity required."); return; }
    const updated = [...fleetUnits, { ...newUnit, id: Date.now() }];
    setFleetUnits(updated); await save(SK.FLEET, updated);
    setNewUnit({ entity:"TMK", model:"", variant:"", serialNo:"", location:"", remark:"" });
    setShowAddUnit(false);
  }

  const entityColors = { TMK: "#0D1B3E", TMS: "#7c3aed", ALL: "#4AACCC" };
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
            {["ALL","TMK","TMS"].map(e => (
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
            {INP("Model *","model",newUnit,setNewUnit,"e.g. PUDU T300")}
            {INP("Variant","variant",newUnit,setNewUnit,"e.g. WTIDL1")}
            {INP("Serial No","serialNo",newUnit,setNewUnit,"e.g. 8260B5318060021")}
            {INP("Location","location",newUnit,setNewUnit,"e.g. KL office")}
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
            <span style={{ fontWeight:700, fontSize:13 }}>DEMO of PUDU under {entity}</span>
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
          No units found for {selEntity==="ALL"?"any entity":selEntity}. Click + Add Unit to get started.
        </div>
      )}
    </div>
  );
}


// ─── FLEET CALENDAR ──────────────────────────────────────────────────────────
// ─── BACKUP & RESTORE ────────────────────────────────────────────────────────
function BackupRestore({ movements, setMovements, personnel, setPersonnel,
    sapSnapshots, setSapSnapshots, gitItems, setGitItems,
    fleetUnits, setFleetUnits, versions, setVersions }) {

  const [restoreStatus, setRestoreStatus] = useState("idle");
  const [restorePreview, setRestorePreview] = useState(null);
  const [backupDone, setBackupDone] = useState(false);
  const [errorMsg, setErrorMsg] = useState("");

  // ── BACKUP ────────────────────────────────────────────────────────────────
  function doBackup() {
    const payload = {
      exportedAt: new Date().toISOString(),
      appVersion: "IFACT Pudu Tracker",
      data: {
        movements,
        personnel,
        sapSnapshots,
        gitItems,
        fleetUnits,
        versions,
      }
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href     = url;
    a.download = "PuduTracker_Backup_" + new Date().toISOString().slice(0,10) + ".json";
    a.click();
    URL.revokeObjectURL(url);
    setBackupDone(true);
    setTimeout(() => setBackupDone(false), 4000);
  }

  // ── RESTORE ───────────────────────────────────────────────────────────────
  async function handleRestoreFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setRestoreStatus("reading");
    setErrorMsg("");
    try {
      const text = await file.text();
      const json = JSON.parse(text);
      if (!json.data || !json.exportedAt) throw new Error("Invalid backup file — missing data or exportedAt fields.");
      const d = json.data;
      const preview = {
        exportedAt: json.exportedAt,
        movements:  Array.isArray(d.movements)    ? d.movements.length    : 0,
        personnel:  Array.isArray(d.personnel)    ? d.personnel.length    : 0,
        snapshots:  Array.isArray(d.sapSnapshots) ? d.sapSnapshots.length : 0,
        git:        Array.isArray(d.gitItems)      ? d.gitItems.length     : 0,
        fleet:      Array.isArray(d.fleetUnits)    ? d.fleetUnits.length   : 0,
        versions:   Array.isArray(d.versions)      ? d.versions.length     : 0,
        raw: d,
      };
      setRestorePreview(preview);
      setRestoreStatus("preview");
    } catch(err) {
      setErrorMsg(err.message || "Could not read file.");
      setRestoreStatus("error");
    }
    e.target.value = "";
  }

  async function confirmRestore() {
    const d = restorePreview.raw;
    if (Array.isArray(d.movements))    { setMovements(d.movements);       await save(SK.MOV,  d.movements); }
    if (Array.isArray(d.personnel))    { setPersonnel(d.personnel);       await save(SK.PERS, d.personnel); }
    if (Array.isArray(d.sapSnapshots)) { setSapSnapshots(d.sapSnapshots); await save(SK.SAP,  d.sapSnapshots); }
    if (Array.isArray(d.gitItems))     { setGitItems(d.gitItems);         await save(SK.GIT,  d.gitItems); }
    if (Array.isArray(d.fleetUnits))   { setFleetUnits(d.fleetUnits);     await save(SK.FLEET,d.fleetUnits); }
    if (Array.isArray(d.versions))     { setVersions(d.versions);         await save(VERSION_KEY, d.versions); }
    setRestoreStatus("done");
    setRestorePreview(null);
  }

  function cancelRestore() { setRestoreStatus("idle"); setRestorePreview(null); }

  const counts = [
    ["Movement Records",    movements.length,    "#0D1B3E"],
    ["Personnel",           personnel.length,    "#4AACCC"],
    ["SAP Snapshots",       sapSnapshots.length, "#16a34a"],
    ["GIT Entries",         gitItems.length,     "#d97706"],
    ["Fleet Units",         fleetUnits.length,   "#7c3aed"],
    ["Version History",     versions.length,     "#E07820"],
  ];

  return (
    <div style={{ maxWidth:760, display:"flex", flexDirection:"column", gap:18 }}>

      {/* Header */}
      <div style={{ background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)", borderRadius:10, padding:"16px 20px", color:"#fff" }}>
        <div style={{ fontWeight:700, fontSize:16, marginBottom:4 }}>Backup & Restore</div>
        <div style={{ fontSize:12, color:"#C8EEF5" }}>
          Export all app data to a JSON file for safekeeping. Restore from any previous backup if something goes wrong.
        </div>
      </div>

      {/* Current data snapshot */}
      <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:16 }}>
        <SectionHead title="Current Data — Live Snapshot" />
        <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:10 }}>
          {counts.map(([label, count, color]) => (
            <div key={label} style={{ background:"#f9fafb", border:"1px solid #e5e7eb", borderRadius:8, padding:"10px 14px", textAlign:"center" }}>
              <div style={{ fontSize:24, fontWeight:700, color }}>{count}</div>
              <div style={{ fontSize:11, color:"#374151", marginTop:2 }}>{label}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Backup */}
      <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:16 }}>
        <SectionHead title="Step 1 — Download Backup" />
        <p style={{ fontSize:13, color:"#374151", margin:"0 0 12px" }}>
          Downloads a single JSON file containing all your data. Save it to SharePoint, OneDrive, or email it to yourself. Do this monthly or before any major update.
        </p>
        {backupDone && (
          <div style={{ background:"#dcfce7", border:"1px solid #86efac", borderRadius:8, padding:"10px 14px", color:"#166534", fontWeight:600, fontSize:13, marginBottom:10 }}>
            Backup downloaded — check your Downloads folder.
          </div>
        )}
        <button onClick={doBackup} style={{
          background:"linear-gradient(135deg,#0D1B3E,#1a3a6b)", color:"#fff",
          border:"none", borderRadius:8, padding:"11px 28px",
          fontWeight:700, fontSize:14, cursor:"pointer"
        }}>
          ⬇ Download Backup (.json)
        </button>
      </div>

      {/* Restore */}
      <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:10, padding:16 }}>
        <SectionHead title="Step 2 — Restore from Backup" />
        <div style={{ background:"#fff7ed", border:"1px solid #fed7aa", borderRadius:8, padding:"10px 14px", fontSize:12, color:"#92400e", marginBottom:12 }}>
          <strong>Warning:</strong> Restoring will OVERWRITE all current data with the backup data. This cannot be undone. Download a fresh backup first if you want to preserve current data.
        </div>

        {restoreStatus === "idle" || restoreStatus === "error" ? (
          <div>
            <label style={{ display:"inline-flex", alignItems:"center", gap:8, background:"#f9fafb", border:"2px solid #d1d5db", borderRadius:8, padding:"10px 22px", fontWeight:600, fontSize:14, cursor:"pointer", color:"#374151" }}>
              📂 Choose Backup File (.json)
              <input type="file" accept=".json" onChange={handleRestoreFile} style={{ display:"none" }} />
            </label>
            {restoreStatus === "error" && (
              <div style={{ marginTop:10, background:"#fee2e2", border:"1px solid #fca5a5", borderRadius:8, padding:"10px 14px", color:"#991b1b", fontSize:13 }}>
                <strong>Error:</strong> {errorMsg}
              </div>
            )}
          </div>
        ) : restoreStatus === "reading" ? (
          <div style={{ color:"#6b7280", fontSize:13 }}>Reading backup file...</div>
        ) : restoreStatus === "done" ? (
          <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
            <div style={{ background:"#dcfce7", border:"1px solid #86efac", borderRadius:8, padding:"12px 16px", color:"#166534", fontWeight:600, fontSize:14 }}>
              Restore complete — all data has been updated from the backup.
            </div>
            <button onClick={cancelRestore} style={{ alignSelf:"flex-start", background:"#f9fafb", border:"1px solid #d1d5db", borderRadius:6, padding:"8px 16px", fontWeight:600, fontSize:13, cursor:"pointer" }}>
              Done
            </button>
          </div>
        ) : null}

        {/* Preview before confirm */}
        {restoreStatus === "preview" && restorePreview && (
          <div style={{ display:"flex", flexDirection:"column", gap:12, marginTop:12 }}>
            <div style={{ background:"#f0f9ff", border:"1px solid #bae6fd", borderRadius:8, padding:"12px 16px" }}>
              <div style={{ fontWeight:700, fontSize:13, color:"#0D1B3E", marginBottom:6 }}>
                Backup from: {new Date(restorePreview.exportedAt).toLocaleString("en-MY")}
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:8 }}>
                {[
                  ["Movements",    restorePreview.movements],
                  ["Personnel",    restorePreview.personnel],
                  ["SAP Snapshots",restorePreview.snapshots],
                  ["GIT Entries",  restorePreview.git],
                  ["Fleet Units",  restorePreview.fleet],
                  ["Versions",     restorePreview.versions],
                ].map(([label, count]) => (
                  <div key={label} style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:6, padding:"8px 10px", textAlign:"center" }}>
                    <div style={{ fontSize:18, fontWeight:700, color:"#0D1B3E" }}>{count}</div>
                    <div style={{ fontSize:10, color:"#6b7280" }}>{label}</div>
                  </div>
                ))}
              </div>
            </div>
            <div style={{ display:"flex", gap:10 }}>
              <button onClick={confirmRestore} style={{ background:"linear-gradient(135deg,#dc2626,#991b1b)", color:"#fff", border:"none", borderRadius:8, padding:"10px 24px", fontWeight:700, fontSize:13, cursor:"pointer" }}>
                Confirm Restore (Overwrites Current Data)
              </button>
              <button onClick={cancelRestore} style={{ background:"#f9fafb", border:"1px solid #d1d5db", borderRadius:8, padding:"10px 18px", fontWeight:600, fontSize:13, cursor:"pointer" }}>
                Cancel
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Best practice guide */}
      <div style={{ background:"#f0fdf4", border:"1px solid #86efac", borderRadius:10, padding:16 }}>
        <SectionHead title="Backup Best Practices" />
        {[
          ["When to backup", "Before every app update, at month-end after all data is entered, and whenever IT requests it"],
          ["Where to save", "SharePoint, OneDrive, or email to yourself and the Logistics Admin. Keep at least 2 copies."],
          ["File naming",   "The file auto-names with today's date (e.g. PuduTracker_Backup_2026-06-27.json). Do not rename it."],
          ["How to restore","If the app loses data after an update, go to Backup & Restore tab, choose your latest .json file, preview and confirm."],
          ["After restoring","Check the Movement Log and Fleet Calendar to confirm all data is back. Then download a fresh backup immediately."],
        ].map(([title, desc]) => (
          <div key={title} style={{ display:"flex", gap:12, padding:"7px 0", borderBottom:"1px solid #dcfce7" }}>
            <div style={{ minWidth:130, fontWeight:600, fontSize:12, color:"#166534" }}>{title}</div>
            <div style={{ fontSize:12, color:"#374151" }}>{desc}</div>
          </div>
        ))}
      </div>
    </div>
  );
}


// ─── VERSION HISTORY ──────────────────────────────────────────────────────────
const INITIAL_VERSIONS = [
  {
    id: 1,
    version: "1.0.0",
    date: "2026-06-27",
    type: "Initial Release",
    changedBy: "Esther Bee Kuan",
    summary: "Initial deployment of IFACT Robotics Pudu Stock Movement Tracker",
    changes: [
      "Stock Position tab — combined SAP Stock, GIT, Demo, and Effective Available view",
      "Dashboard — KPIs, overdue alerts, in-field by custodian and region, 6-month GR/GI trend",
      "Log Movement — demo out/return, POC, exhibition, loan, customer delivery, GIT transfer",
      "Movement Log — full filterable history with Return button",
      "GIT Register — PO tracking from On Order to Received",
      "Admin — SAP snapshots, monthly summary, personnel register, team activity",
      "Import SAP — upload SAP Excel template to auto-update Stock Position and GIT Register",
      "Export Excel — 9-sheet report with multiple filter criteria",
      "User Guide tab — role-by-role how-to instructions",
    ]
  }
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
    if (!window.confirm("Delete this version entry?")) return;
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
  const [activeTab,setActiveTab]=useState("Stock Position");
  const [movements,setMovements]=useState([]);
  const [personnel,setPersonnel]=useState([]);
  const [sapSnapshots,setSapSnapshots]=useState([]);
  const [gitItems,setGitItems]=useState([]);
  const [fleetUnits,setFleetUnits]=useState([]);
  const [versions,setVersions]=useState([]);
  const [loading,setLoading]=useState(true);

  useEffect(()=>{
    Promise.all([
      load(SK.MOV,SAMPLE_MOVEMENTS),
      load(SK.PERS,SAMPLE_PERSONNEL),
      load(SK.SAP,SAMPLE_SAP),
      load(SK.GIT,SAMPLE_GIT),
      load(SK.FLEET,SAMPLE_FLEET),
      load(VERSION_KEY,INITIAL_VERSIONS),
    ]).then(([m,p,s,g,fl,v])=>{
      setMovements(m);setPersonnel(p);setSapSnapshots(s);setGitItems(g);setFleetUnits(fl);setVersions(v);setLoading(false);
    });
  },[]);

  if(loading) return <div style={{ display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",fontFamily:"Inter,sans-serif",color:"#0D1B3E",fontSize:15 }}>Loading IFACT Robotics Tracker...</div>;

  const icons={"Stock Position":"📦","Fleet Calendar":"📅","Dashboard":"📊","Log Movement":"✏️","Movement Log":"📋","GIT Register":"🚢","Admin":"🔒","Import SAP":"📥","Export Excel":"⬇","User Guide":"📖","Backup & Restore":"💾","Version History":"🕐"};
  const NAV_TABS=["Stock Position","Fleet Calendar","Dashboard","Log Movement","Movement Log","GIT Register","Admin","Import SAP","Backup & Restore","Export Excel","User Guide","Version History"];

  return(
    <div style={{ fontFamily:"'Inter','Segoe UI',sans-serif",minHeight:"100vh",background:"#f8fafc",color:"#111827" }}>
      <div style={{ background:"linear-gradient(135deg,#0D1B3E 0%,#1a3a6b 100%)",padding:"0 20px" }}>
        <div style={{ maxWidth:1280,margin:"0 auto" }}>
          <div style={{ display:"flex",alignItems:"center",gap:14,padding:"14px 0 0" }}>
            <div style={{ background:"linear-gradient(135deg,#C8EEF5,#4AACCC)",borderRadius:8,width:38,height:38,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18 }}>🤖</div>
            <div>
              <div style={{ color:"#fff",fontWeight:700,fontSize:17 }}>IFACT Robotics — Stock Movement Tracker</div>
              <div style={{ color:"#C8EEF5",fontSize:11 }}>Pudu Robotics | Stock · GIT · Demo · POC · Exhibition</div>
            </div>
            <div style={{ marginLeft:"auto", display:"flex", alignItems:"center", gap:12 }}>
              <div style={{ color:"#C8EEF5", fontSize:11 }}>{new Date().toLocaleDateString("en-MY",{weekday:"short",day:"2-digit",month:"short",year:"numeric"})}</div>
              <button
                onClick={() => {
                  const url = window.location.href;
                  const ta = document.createElement("textarea");
                  ta.value = url;
                  ta.style.position = "fixed";
                  ta.style.opacity = "0";
                  document.body.appendChild(ta);
                  ta.focus();
                  ta.select();
                  try {
                    document.execCommand("copy");
                    alert("App link copied!\n\n" + url + "\n\nPaste and share this link with your team via WhatsApp or email.");
                  } catch(e) {
                    alert("Your app link:\n\n" + url + "\n\nCopy this link and share with your team.");
                  }
                  document.body.removeChild(ta);
                }}
                style={{ background:"#4AACCC", color:"#0D1B3E", border:"none", borderRadius:6, padding:"5px 12px", fontWeight:700, fontSize:11, cursor:"pointer", whiteSpace:"nowrap" }}>
                🔗 Share App
              </button>
            </div>
          </div>
          <div style={{ display:"flex",gap:2,marginTop:10,overflowX:"auto" }}>
            {NAV_TABS.map(tab=>(
              <button key={tab} onClick={()=>setActiveTab(tab)} style={{ padding:"9px 14px",border:"none",cursor:"pointer",borderRadius:"6px 6px 0 0",fontWeight:600,fontSize:12,background:activeTab===tab?"#f8fafc":"transparent",color:activeTab===tab?"#0D1B3E":"#C8EEF5",whiteSpace:"nowrap" }}>
                {icons[tab]} {tab}
              </button>
            ))}
          </div>
        </div>
      </div>
      <div style={{ maxWidth:1280,margin:"0 auto",padding:"22px 20px" }}>
        {activeTab==="Stock Position" && <StockPosition movements={movements} sapSnapshots={sapSnapshots} gitItems={gitItems} />}

        {activeTab==="Fleet Calendar" && <FleetCalendar movements={movements} fleetUnits={fleetUnits} setFleetUnits={setFleetUnits} />}
        {activeTab==="Dashboard" && <Dashboard movements={movements} personnel={personnel} />}
        {activeTab==="Log Movement" && <LogMovement movements={movements} setMovements={setMovements} personnel={personnel} />}
        {activeTab==="Movement Log" && <MovementLog movements={movements} setMovements={setMovements} />}
        {activeTab==="GIT Register" && <GITRegister gitItems={gitItems} setGitItems={setGitItems} />}
        {activeTab==="Admin" && <Admin movements={movements} personnel={personnel} setPersonnel={setPersonnel} sapSnapshots={sapSnapshots} setSapSnapshots={setSapSnapshots} />}
        {activeTab==="Import SAP" && <SapImport sapSnapshots={sapSnapshots} setSapSnapshots={setSapSnapshots} gitItems={gitItems} setGitItems={setGitItems} />}
        {activeTab==="Export Excel" && <ExportExcel movements={movements} personnel={personnel} sapSnapshots={sapSnapshots} gitItems={gitItems} />}
        {activeTab==="Backup & Restore" && <BackupRestore
          movements={movements} setMovements={setMovements}
          personnel={personnel} setPersonnel={setPersonnel}
          sapSnapshots={sapSnapshots} setSapSnapshots={setSapSnapshots}
          gitItems={gitItems} setGitItems={setGitItems}
          fleetUnits={fleetUnits} setFleetUnits={setFleetUnits}
          versions={versions} setVersions={setVersions}
        />}
        {activeTab==="User Guide" && <UserGuide />}
        {activeTab==="Version History" && <VersionHistory />}
      </div>
    </div>
  );
}
