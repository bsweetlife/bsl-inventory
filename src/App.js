import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from './lib/supabase';

// ── TRANSLATIONS ──────────────────────────────────────────────
const T = {
  en: {
    title: 'BSL Inventory', dashboard: 'Dashboard', chat: 'Chat with Claude',
    addProduct: '+ Add Product', bulkImport: '📥 Bulk Import', uploadOrders: '📦 Upload Orders', exportCSV: '⬇ Export CSV',
    totalProducts: 'Total Products', inventoryValue: 'Inventory Value', lowStock: 'Low Stock', critical: 'Critical',
    allProducts: 'All Products', alerts: 'Low Stock Alerts', channels: 'By Channel', log: 'Change Log',
    status: 'Status', product: 'Product', rootSku: 'Root SKU', stock: 'Stock', days: 'Days', cost: 'Cost', price: 'Price', actions: 'Actions',
    ok: 'OK', low: 'Low', crit: 'Critical',
    productName: 'Product name', category: 'Category', monthlyS: 'Monthly sales', reorder: 'Reorder point', supplier: 'Supplier',
    marketplaceSKUs: 'Marketplace SKUs', amazonSku: 'Amazon SKU', walmartSku: 'Walmart SKU', targetSku: 'Target SKU', temuSku: 'Temu SKU', otherSku: 'Other SKU',
    save: 'Save', cancel: 'Cancel', edit: 'Edit', delete: 'Delete',
    selectPlatform: 'Select platform then upload order file',
    confirmColumns: 'Confirm columns — auto-detected from your file',
    skuCol: 'SKU column *', qtyCol: 'Quantity column *', orderIdCol: 'Order ID (dupe check)',
    previewDeductions: 'Preview →', applyDeductions: '✓ Apply deductions', back: '← Back',
    matched: 'matched', notFound: 'not found', current: 'Current', sold: 'Sold', newStock: 'New stock',
    duplicateTitle: 'Possible duplicate upload', duplicateMsg: 'This file looks like it was already uploaded. Processing again will double-deduct quantities.',
    processAnyway: 'Process anyway', goBack: '← Go back',
    dropFile: 'Drop file or click to browse', csvOrXlsx: '.csv or .xlsx',
    downloadTemplate: '⬇ Download template',
    importStep: 'Bulk import — step', of: 'of',
    addUpdate: 'Add / update', replaceAll: 'Replace all',
    addUpdateDesc: 'Adds new products; updates existing by Root SKU',
    replaceDesc: 'Removes all current products and replaces',
    productsReady: 'products ready to import',
    import: 'Import', products: 'products',
    noProducts: 'No products yet — click + Add Product or use Bulk Import',
    noAlerts: 'No low stock items',
    noLog: 'No changes recorded yet',
    saved: 'Saved', sessionOnly: 'Session only',
    chatPlaceholder: 'Ask me anything about your inventory...',
    chatWelcome: "Hi! I'm Claude, your BSL inventory manager. I have full access to your inventory data. Ask me anything — stock levels, sales trends, reorder suggestions, or upload order files and I'll update everything automatically.",
    thinking: 'Thinking...',
    uploadOrdersTitle: 'Upload Orders',
    step: 'Step',
    name: 'Name',
    amazon: 'Amazon', walmart: 'Walmart', target: 'Target', temu: 'Temu', other: 'Other',
  },
  es: {
    title: 'BSL Inventario', dashboard: 'Panel', chat: 'Chat con Claude',
    addProduct: '+ Agregar Producto', bulkImport: '📥 Importar Masivo', uploadOrders: '📦 Subir Órdenes', exportCSV: '⬇ Exportar CSV',
    totalProducts: 'Total Productos', inventoryValue: 'Valor Inventario', lowStock: 'Stock Bajo', critical: 'Crítico',
    allProducts: 'Todos los Productos', alerts: 'Alertas Stock Bajo', channels: 'Por Canal', log: 'Registro de Cambios',
    status: 'Estado', product: 'Producto', rootSku: 'SKU Raíz', stock: 'Stock', days: 'Días', cost: 'Costo', price: 'Precio', actions: 'Acciones',
    ok: 'OK', low: 'Bajo', crit: 'Crítico',
    productName: 'Nombre del producto', category: 'Categoría', monthlyS: 'Ventas mensuales', reorder: 'Punto de reorden', supplier: 'Proveedor',
    marketplaceSKUs: 'SKUs por Marketplace', amazonSku: 'SKU Amazon', walmartSku: 'SKU Walmart', targetSku: 'SKU Target', temuSku: 'SKU Temu', otherSku: 'SKU Otro',
    save: 'Guardar', cancel: 'Cancelar', edit: 'Editar', delete: 'Eliminar',
    selectPlatform: 'Selecciona plataforma y sube el archivo de órdenes',
    confirmColumns: 'Confirma columnas — detectadas automáticamente',
    skuCol: 'Columna SKU *', qtyCol: 'Columna Cantidad *', orderIdCol: 'ID Orden (detección duplicados)',
    previewDeductions: 'Vista previa →', applyDeductions: '✓ Aplicar deducciones', back: '← Atrás',
    matched: 'encontrados', notFound: 'no encontrado', current: 'Actual', sold: 'Vendido', newStock: 'Nuevo stock',
    duplicateTitle: 'Posible archivo duplicado', duplicateMsg: 'Este archivo parece haber sido subido antes. Procesarlo de nuevo deducirá las cantidades dos veces.',
    processAnyway: 'Procesar de todas formas', goBack: '← Regresar',
    dropFile: 'Suelta el archivo o haz clic para buscar', csvOrXlsx: '.csv o .xlsx',
    downloadTemplate: '⬇ Descargar plantilla',
    importStep: 'Importar masivo — paso', of: 'de',
    addUpdate: 'Agregar / actualizar', replaceAll: 'Reemplazar todo',
    addUpdateDesc: 'Agrega nuevos productos; actualiza existentes por SKU Raíz',
    replaceDesc: 'Elimina todos los productos actuales y los reemplaza',
    productsReady: 'productos listos para importar',
    import: 'Importar', products: 'productos',
    noProducts: 'Sin productos — haz clic en + Agregar Producto o usa Importar Masivo',
    noAlerts: 'Sin artículos con stock bajo',
    noLog: 'Sin cambios registrados aún',
    saved: 'Guardado', sessionOnly: 'Solo sesión',
    chatPlaceholder: 'Pregúntame sobre tu inventario...',
    chatWelcome: '¡Hola! Soy Claude, tu gerente de inventario BSL. Tengo acceso completo a tus datos. Pregúntame lo que necesites — niveles de stock, tendencias de ventas, sugerencias de reorden, o sube archivos de órdenes y actualizaré todo automáticamente.',
    thinking: 'Pensando...',
    uploadOrdersTitle: 'Subir Órdenes',
    step: 'Paso',
    name: 'Nombre',
    amazon: 'Amazon', walmart: 'Walmart', target: 'Target', temu: 'Temu', other: 'Otro',
  }
};

// ── PLATFORM CONFIG ───────────────────────────────────────────
const PLAT = {
  amazon:  { l: 'Amazon',      c: '#FF9900', sku: ['sku','seller-sku','asin'],         qty: ['quantity','quantity-purchased','qty','units'], oid: ['amazon-order-id','order-id','order_id'] },
  walmart: { l: 'Walmart',     c: '#0071DC', sku: ['sku','item-sku','seller-sku'],      qty: ['qty','quantity','units-ordered'],              oid: ['order_id','purchase-order-id','order-id'] },
  target:  { l: 'Target Plus', c: '#CC0000', sku: ['sku','tcin','seller-sku'],          qty: ['quantity','qty','quantity-ordered'],           oid: ['order-id','po-number','order_id'] },
  temu:    { l: 'Temu',        c: '#7B2D8B', sku: ['sku','product-sku','seller-sku'],   qty: ['quantity','qty','amount'],                    oid: ['order-id','order_id','order-number'] },
  other:   { l: 'Other',       c: '#888',    sku: ['sku','item-sku','product-sku'],     qty: ['quantity','qty','units'],                     oid: ['order-id','order_id'] },
};

const MPK = ['amz','wmt','tgt','temu','other_sku'];

// ── HELPERS ───────────────────────────────────────────────────
const getStatus = p => { const d=(p.velocity||0)/30; if(!d) return 'ok'; const v=p.stock/d; return v<=15?'crit':v<=30?'low':'ok'; };
const getDays   = p => { const d=(p.velocity||0)/30; return d?Math.round(p.stock/d):null; };
const fmtMoney  = n => (n!=null&&n!=='')?'$'+parseFloat(n).toFixed(2):'—';
const hashStr   = s => { let h=0; for(let i=0;i<Math.min(s.length,500);i++) h=(Math.imul(31,h)+s.charCodeAt(i))|0; return h.toString(); };
const findCol   = (hdrs,cands) => { for(const c of cands){const i=hdrs.findIndex(h=>h.toLowerCase().replace(/[\s_-]+/g,'-')===c); if(i>=0)return i;} for(const c of cands){const i=hdrs.findIndex(h=>h.toLowerCase().includes(c.replace(/-/g,''))); if(i>=0)return i;} return -1; };

function readXLSX(file, cb) {
  const r = new FileReader();
  r.onload = e => { try { const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'}); const ws=wb.Sheets[wb.SheetNames[0]]; cb(null,XLSX.utils.sheet_to_json(ws,{header:1,defval:''})); } catch(err){cb(err);} };
  r.readAsArrayBuffer(file);
}

const COL_MAP = {"name":"name","product name":"name","product":"name","root sku":"sku","sku":"sku","category":"category","cat":"category","stock":"stock","current stock":"stock","qty":"stock","quantity":"stock","monthly sales":"velocity","velocity":"velocity","cost":"cost","cost price":"cost","price":"price","selling price":"price","reorder":"reorder","reorder point":"reorder","supplier":"supplier","supplier name":"supplier","amazon sku":"amz","amz":"amz","walmart sku":"wmt","wmt":"wmt","target sku":"tgt","tgt":"tgt","temu sku":"temu","temu":"temu","other sku":"other_sku","other":"other_sku"};

const emptyProduct = () => ({id:null,name:'',sku:'',category:'',stock:'',velocity:'',cost:'',price:'',reorder:'',supplier:'',amz:'',wmt:'',tgt:'',temu:'',other_sku:''});

// ── STYLES ────────────────────────────────────────────────────
const S = {
  app:   { fontFamily:'system-ui,sans-serif', minHeight:'100vh', background:'#f5f5f5' },
  nav:   { background:'#111', color:'#fff', padding:'0 1.5rem', display:'flex', alignItems:'center', justifyContent:'space-between', height:52, position:'sticky', top:0, zIndex:50 },
  main:  { maxWidth:1100, margin:'0 auto', padding:'1.25rem' },
  card:  { background:'#fff', borderRadius:12, padding:'1.25rem', boxShadow:'0 1px 3px rgba(0,0,0,.08)' },
  btn:   { fontSize:12, padding:'6px 13px', borderRadius:8, border:'0.5px solid #ddd', background:'transparent', cursor:'pointer', display:'inline-flex', alignItems:'center', gap:4, fontFamily:'inherit' },
  btnPrimary: { fontSize:12, padding:'6px 13px', borderRadius:8, border:'none', background:'#111', color:'#fff', cursor:'pointer', display:'inline-flex', alignItems:'center', gap:4, fontFamily:'inherit' },
  input: { fontSize:13, padding:'6px 9px', borderRadius:8, border:'0.5px solid #ddd', width:'100%', boxSizing:'border-box', fontFamily:'inherit' },
  overlay: { position:'fixed', top:0, left:0, right:0, bottom:0, background:'rgba(0,0,0,.45)', zIndex:200, display:'flex', alignItems:'center', justifyContent:'center' },
  sheet:   { background:'#fff', borderRadius:16, padding:'1.5rem', width:560, maxWidth:'95vw', maxHeight:'85vh', overflowY:'auto' },
  th:    { textAlign:'left', fontSize:10, fontWeight:600, color:'#888', padding:'7px 10px', borderBottom:'1px solid #eee', background:'#fafafa', whiteSpace:'nowrap', textTransform:'uppercase', letterSpacing:'.04em', position:'sticky', top:0 },
  td:    { padding:'8px 10px', borderBottom:'1px solid #f5f5f5', fontSize:12, verticalAlign:'middle', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' },
};

// ── STATUS BADGE ──────────────────────────────────────────────
const Badge = ({ status, t }) => {
  const m = { ok:['#D4EDDA','#155724'], low:['#FFF3CD','#856404'], crit:['#F8D7DA','#721C24'] };
  const [bg,c] = m[status]||m.ok;
  return <span style={{background:bg,color:c,padding:'2px 8px',borderRadius:99,fontSize:11,fontWeight:500}}>{t[status]||status}</span>;
};

// ── MAIN APP ──────────────────────────────────────────────────
export default function App() {
  const [lang, setLang]       = useState('en');
  const [page, setPage]       = useState('dashboard');
  const [tab,  setTab]        = useState('all');
  const [prods, setProds]     = useState([]);
  const [logEntries, setLog]  = useState([]);
  const [hashes, setHashes]   = useState([]);
  const [loading, setLoading] = useState(true);
  const [modal, setModal]     = useState(null);
  const [mdata, setMdata]     = useState({});
  const [chatMsgs, setChatMsgs] = useState([]);
  const [chatInput, setChatInput] = useState('');
  const [chatLoading, setChatLoading] = useState(false);
  const chatEndRef = useRef(null);
  const t = T[lang];

  // Load data from Supabase
  useEffect(() => {
    loadAll();
  }, []);

  useEffect(() => {
    if (chatMsgs.length === 0) {
      setChatMsgs([{ role: 'assistant', content: t.chatWelcome }]);
    }
  }, [lang]);

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [chatMsgs]);

  async function loadAll() {
    setLoading(true);
    try {
      const [{ data: p }, { data: l }, { data: h }] = await Promise.all([
        supabase.from('products').select('*').order('name'),
        supabase.from('change_log').select('*').order('created_at', { ascending: false }).limit(200),
        supabase.from('uploaded_files').select('file_hash'),
      ]);
      if (p) setProds(p);
      if (l) setLog(l);
      if (h) setHashes(h.map(x => x.file_hash));
    } catch (e) { console.error(e); }
    setLoading(false);
    setChatMsgs([{ role: 'assistant', content: t.chatWelcome }]);
  }

  // ── PRODUCT CRUD ────────────────────────────────────────────
  async function saveProduct(form) {
    const entry = { ...form, stock: parseFloat(form.stock)||0, velocity: parseFloat(form.velocity)||0, cost: parseFloat(form.cost)||0, price: parseFloat(form.price)||0, reorder: parseFloat(form.reorder)||0 };
    delete entry.id;
    if (form.id) {
      const old = prods.find(p => p.id === form.id);
      await supabase.from('products').update(entry).eq('id', form.id);
      const diff = (parseFloat(form.stock)||0) - old.stock;
      await supabase.from('change_log').insert({ description: diff !== 0 ? `Stock: ${form.name} ${old.stock}→${form.stock}` : `Updated: ${form.name}`, qty_change: diff || null });
    } else {
      await supabase.from('products').insert(entry);
      await supabase.from('change_log').insert({ description: `Added: ${form.name}`, qty_change: parseFloat(form.stock)||0 });
    }
    await loadAll();
    setModal(null);
  }

  async function deleteProduct(id) {
    const p = prods.find(x => x.id === id);
    if (!window.confirm(`Delete "${p.name}"?`)) return;
    await supabase.from('products').delete().eq('id', id);
    await supabase.from('change_log').insert({ description: `Deleted: ${p.name}` });
    await loadAll();
  }

  // ── BULK IMPORT ─────────────────────────────────────────────
  function handleImpFile(e) {
    const f = e.target.files[0]; if (!f) return;
    readXLSX(f, (err, rows) => {
      if (err || rows.length < 2) { alert('Cannot read file'); return; }
      const hdrs = rows[0].map(h => (h||'').toString().trim().toLowerCase());
      const fm = hdrs.map(h => COL_MAP[h]||null);
      const parsed = [], errors = [];
      rows.slice(1).filter(r => r.some(c => c !== '')).forEach((row, i) => {
        const o = {}; fm.forEach((ff, ci) => { if (ff) o[ff] = (row[ci]??'').toString().trim(); });
        if (!o.name) { errors.push(`Row ${i+2}: no name`); return; }
        parsed.push({ name:o.name, sku:o.sku||'', category:o.category||'', stock:parseFloat(o.stock)||0, velocity:parseFloat(o.velocity)||0, cost:parseFloat(o.cost)||0, price:parseFloat(o.price)||0, reorder:parseFloat(o.reorder)||0, supplier:o.supplier||'', amz:o.amz||'', wmt:o.wmt||'', tgt:o.tgt||'', temu:o.temu||'', other_sku:o.other_sku||'' });
      });
      setMdata(prev => ({ ...prev, step:2, parsed, errors }));
    });
    e.target.value = '';
  }

  async function confirmImport() {
    const { parsed, mode } = mdata;
    if (mode === 'replace') {
      await supabase.from('products').delete().neq('id', 0);
      await supabase.from('products').insert(parsed);
    } else {
      for (const p of parsed) {
        const existing = prods.find(x => x.sku && x.sku === p.sku);
        if (existing) await supabase.from('products').update(p).eq('id', existing.id);
        else await supabase.from('products').insert(p);
      }
    }
    await supabase.from('change_log').insert({ description: `Bulk import: ${parsed.length} products (${mode||'add'})`, qty_change: parsed.length });
    await loadAll();
    setModal(null);
  }

  // ── ORDERS UPLOAD ───────────────────────────────────────────
  function handleOrdFile(e) {
    const f = e.target.files[0]; if (!f) return;
    readXLSX(f, (err, rows) => {
      if (err || rows.length < 2) { alert('Cannot read file'); return; }
      const hash = hashStr(rows[0].join(',') + (rows[1]||[]).join(','));
      const headers = rows[0].map(h => (h||'').toString().trim());
      const dataRows = rows.slice(1).filter(r => r.some(c => c !== ''));
      const cfg = PLAT[mdata.platform||'amazon'];
      const hl = headers.map(h => h.toLowerCase());
      setMdata(prev => ({ ...prev, hash, headers, rows: dataRows, skuCol: findCol(hl, cfg.sku), qtyCol: findCol(hl, cfg.qty), oidCol: findCol(hl, cfg.oid), step: 2 }));
    });
    e.target.value = '';
  }

  function buildPreview() {
    if (mdata.skuCol < 0 || mdata.qtyCol < 0) { alert('Map SKU and Quantity columns'); return; }
    const mpKey = { amazon:'amz', walmart:'wmt', target:'tgt', temu:'temu', other:'other_sku' }[mdata.platform]||'other_sku';
    const agg = {};
    mdata.rows.forEach(row => {
      const sku = (row[mdata.skuCol]||'').toString().trim();
      const qty = parseFloat((row[mdata.qtyCol]||'0').toString().replace(/[^0-9.-]/g,''))||0;
      if (!sku || qty <= 0) return;
      if (!agg[sku]) agg[sku] = { sku, qty: 0 };
      agg[sku].qty += qty;
    });
    const warnings = [];
    const preview = Object.values(agg).map(item => {
      let prod = prods.find(p => p[mpKey] === item.sku) || prods.find(p => p.sku === item.sku) || prods.find(p => MPK.some(k => p[k] === item.sku));
      const newStock = prod ? prod.stock - item.qty : null;
      if (prod && newStock < 0) warnings.push(`"${prod.name}" → ${newStock}`);
      if (!prod) warnings.push(`"${item.sku}" not found`);
      return { ...item, prod, newStock };
    });
    setMdata(prev => ({ ...prev, preview, warnings, step: 3 }));
  }

  async function applyOrders() {
    const platLabel = PLAT[mdata.platform]?.l;
    let deducted = 0, skipped = 0;
    for (const item of mdata.preview) {
      if (!item.prod) { skipped++; continue; }
      const oldStock = item.prod.stock;
      const newStock = oldStock - item.qty;
      await supabase.from('products').update({ stock: newStock }).eq('id', item.prod.id);
      await supabase.from('orders_log').insert({ platform: mdata.platform, order_sku: item.sku, product_id: item.prod.id, product_name: item.prod.name, qty_sold: item.qty, stock_before: oldStock, stock_after: newStock });
      await supabase.from('change_log').insert({ description: `${platLabel}: ${item.prod.name} [${item.sku}] ${oldStock}→${newStock}`, qty_change: -item.qty, platform: mdata.platform });
      deducted++;
    }
    await supabase.from('change_log').insert({ description: `${platLabel} upload: ${deducted} updated, ${skipped} skipped` });
    if (mdata.hash) await supabase.from('uploaded_files').insert({ file_hash: mdata.hash });
    await loadAll();
    setModal(null);
  }

  function confirmOrders() {
    if (hashes.includes(mdata.hash)) { setMdata(prev => ({ ...prev, step: 4 })); return; }
    applyOrders();
  }

  // ── EXPORT CSV ──────────────────────────────────────────────
  function exportCSV() {
    const headers = ['Name','Root SKU','Category','Stock','Monthly Sales','Cost','Price','Reorder','Supplier','Amazon SKU','Walmart SKU','Target SKU','Temu SKU','Other SKU','Status','Days Left'];
    const rows = prods.map(p => [p.name,p.sku,p.category,p.stock,p.velocity,p.cost,p.price,p.reorder,p.supplier,p.amz,p.wmt,p.tgt,p.temu,p.other_sku,getStatus(p),getDays(p)??'—'].map(v=>`"${String(v??'').replace(/"/g,'""')}"`).join(','));
    const csv = [headers.join(','), ...rows].join('\n');
    const a = document.createElement('a'); a.href='data:text/csv;charset=utf-8,'+encodeURIComponent(csv); a.download='bsl_inventory_'+new Date().toISOString().slice(0,10)+'.csv'; a.click();
  }

  function downloadTemplate() {
    const h = ['Name','Root SKU','Category','Stock','Monthly Sales','Cost','Price','Reorder Point','Supplier','Amazon SKU','Walmart SKU','Target SKU','Temu SKU','Other SKU'];
    const ex = ['Vanilla Cake Mix','BSL-CAKE-001','Baking Mixes','200','80','4.50','12.99','50','Supplier Co.','AMZ-001','WMT-001','TGT-001','TEMU-001',''];
    const csv = [h,ex].map(r=>r.map(v=>`"${v}"`).join(',')).join('\n');
    const a = document.createElement('a'); a.href='data:text/csv;charset=utf-8,'+encodeURIComponent(csv); a.download='inventory_template.csv'; a.click();
  }

  // ── AI CHAT ─────────────────────────────────────────────────
  async function sendChat(e) {
    e.preventDefault();
    if (!chatInput.trim() || chatLoading) return;
    const userMsg = chatInput.trim();
    setChatInput('');
    const newMsgs = [...chatMsgs, { role: 'user', content: userMsg }];
    setChatMsgs(newMsgs);
    setChatLoading(true);
    try {
      const inventoryContext = `Current inventory data (${prods.length} products):
${prods.map(p => `- ${p.name} | Root SKU: ${p.sku} | Stock: ${p.stock} | Velocity: ${p.velocity}/mo | Cost: $${p.cost} | Price: $${p.price} | Status: ${getStatus(p)} | Days left: ${getDays(p)??'N/A'} | Amazon: ${p.amz||'—'} | Walmart: ${p.wmt||'—'} | Target: ${p.tgt||'—'} | Temu: ${p.temu||'—'}`).join('\n')}

Recent changes:
${logEntries.slice(0,10).map(l => `- ${new Date(l.created_at).toLocaleDateString()}: ${l.description}`).join('\n')}`;

      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-api-key': process.env.REACT_APP_ANTHROPIC_KEY, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-calls': 'true' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 1000,
          system: `You are Claude, the inventory manager for BSL (Blooming Sweet Life Corp), a multi-channel e-commerce business selling on Amazon, Walmart, Target, Temu and other platforms. You have full access to their inventory data. Be concise, helpful, and proactive. Respond in ${lang === 'es' ? 'Spanish' : 'English'}. When asked about stock levels, sales, or reorders, reference the actual data provided. Here is the current inventory:\n\n${inventoryContext}`,
          messages: newMsgs.filter(m => m.role !== 'system').map(m => ({ role: m.role, content: m.content })),
        }),
      });
      const data = await response.json();
      const reply = data.content?.[0]?.text || 'Sorry, I could not process that.';
      setChatMsgs(prev => [...prev, { role: 'assistant', content: reply }]);
    } catch (err) {
      setChatMsgs(prev => [...prev, { role: 'assistant', content: 'Error connecting to Claude API. Please check your API key.' }]);
    }
    setChatLoading(false);
  }

  // ── COMPUTED ─────────────────────────────────────────────────
  const totalVal = prods.reduce((a, p) => a + (p.stock * (p.cost||0)), 0);
  const lowN  = prods.filter(p => getStatus(p) !== 'ok').length;
  const critN = prods.filter(p => getStatus(p) === 'crit').length;
  const alerts = prods.filter(p => getStatus(p) !== 'ok');

  // ── PRODUCT TABLE ────────────────────────────────────────────
  const ProdTable = ({ list }) => {
    if (!list.length) return <div style={{textAlign:'center',padding:'2rem',color:'#aaa',fontSize:13}}>{tab==='alerts'?t.noAlerts:t.noProducts}</div>;
    return (
      <div style={{overflowX:'auto',border:'1px solid #eee',borderRadius:12}}>
        <table style={{width:'100%',borderCollapse:'collapse',fontSize:12}}>
          <thead><tr>{[t.status,t.product,t.rootSku,t.stock,t.days,t.cost,t.price,t.actions].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
          <tbody>
            {list.map(p => {
              const st=getStatus(p), dl=getDays(p);
              return <tr key={p.id}>
                <td style={S.td}><Badge status={st} t={t}/></td>
                <td style={{...S.td,fontWeight:500,maxWidth:160}} title={p.name}>{p.name}</td>
                <td style={S.td}><code style={{fontSize:10,background:'#f5f5f5',padding:'1px 5px',borderRadius:4}}>{p.sku||'—'}</code></td>
                <td style={{...S.td,color:p.stock<0?'#dc3545':undefined}}>{p.stock}</td>
                <td style={{...S.td,color:st==='crit'?'#dc3545':st==='low'?'#856404':undefined}}>{dl!=null?`${dl}d`:'—'}</td>
                <td style={S.td}>{fmtMoney(p.cost)}</td>
                <td style={S.td}>{fmtMoney(p.price)}</td>
                <td style={{...S.td,whiteSpace:'nowrap'}}>
                  <button style={{...S.btn,padding:'3px 7px'}} onClick={()=>{setMdata({form:{...p}});setModal('product')}}>✏️</button>
                  <button style={{...S.btn,padding:'3px 7px',marginLeft:3,color:'#dc3545',borderColor:'#f5c6cb'}} onClick={()=>deleteProduct(p.id)}>🗑</button>
                </td>
              </tr>;
            })}
          </tbody>
        </table>
      </div>
    );
  };

  // ── MODALS ───────────────────────────────────────────────────
  const ProductModal = () => {
    const [form, setForm] = useState(mdata.form || emptyProduct());
    const isEdit = !!form.id;
    return (
      <div style={S.overlay} onClick={e=>e.target===e.currentTarget&&setModal(null)}>
        <div style={S.sheet}>
          <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem'}}>{isEdit?t.edit:t.addProduct}</div>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
            {[{l:t.productName,k:'name',full:true,ph:'Vanilla Cake Mix'},{l:t.rootSku,k:'sku',ph:'BSL-CAKE-001'},{l:t.category,k:'category',ph:'Baking Mixes'},{l:t.stock,k:'stock',t:'number',ph:'0'},{l:t.monthlyS,k:'velocity',t:'number',ph:'0'},{l:`${t.cost} ($)`,k:'cost',t:'number',ph:'0.00'},{l:`${t.price} ($)`,k:'price',t:'number',ph:'0.00'},{l:t.reorder,k:'reorder',t:'number',ph:'0'},{l:t.supplier,k:'supplier',full:true,ph:'Supplier Co.'}].map(f=>(
              <div key={f.k} style={{gridColumn:f.full?'1/-1':undefined,display:'flex',flexDirection:'column',gap:3}}>
                <label style={{fontSize:11,color:'#888',fontWeight:500}}>{f.l}</label>
                <input style={S.input} type={f.t||'text'} placeholder={f.ph} value={form[f.k]||''} onChange={e=>setForm(prev=>({...prev,[f.k]:e.target.value}))}/>
              </div>
            ))}
            <div style={{gridColumn:'1/-1',fontSize:10,fontWeight:600,color:'#888',textTransform:'uppercase',letterSpacing:'.05em',marginTop:6,paddingTop:8,borderTop:'1px solid #eee'}}>{t.marketplaceSKUs}</div>
            {[{l:t.amazonSku,k:'amz'},{l:t.walmartSku,k:'wmt'},{l:t.targetSku,k:'tgt'},{l:t.temuSku,k:'temu'},{l:t.otherSku,k:'other_sku'}].map(f=>(
              <div key={f.k} style={{display:'flex',flexDirection:'column',gap:3}}>
                <label style={{fontSize:11,color:'#888',fontWeight:500}}>{f.l}</label>
                <input style={S.input} placeholder={f.k.toUpperCase()+'-001'} value={form[f.k]||''} onChange={e=>setForm(prev=>({...prev,[f.k]:e.target.value}))}/>
              </div>
            ))}
          </div>
          <div style={{display:'flex',gap:8,justifyContent:'flex-end',marginTop:'1rem'}}>
            <button style={S.btn} onClick={()=>setModal(null)}>{t.cancel}</button>
            <button style={S.btnPrimary} onClick={()=>saveProduct(form)}>{t.save}</button>
          </div>
        </div>
      </div>
    );
  };

  const ImportModal = () => {
    const step = mdata.step||1;
    const fRef = useRef();
    return (
      <div style={S.overlay} onClick={e=>e.target===e.currentTarget&&setModal(null)}>
        <div style={S.sheet}>
          <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem'}}>{t.importStep} {step} {t.of} 2</div>
          {step===1&&<>
            <div style={{border:'2px dashed #ddd',borderRadius:12,padding:'2rem',textAlign:'center',cursor:'pointer',background:'#fafafa',marginBottom:'1rem'}} onClick={()=>fRef.current.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f){const dt=new DataTransfer();dt.items.add(f);fRef.current.files=dt.files;handleImpFile({target:fRef.current})}}}>
              <div style={{fontSize:26,marginBottom:6}}>📂</div>
              <div style={{fontSize:13,fontWeight:600,marginBottom:4}}>{t.dropFile}</div>
              <div style={{fontSize:11,color:'#aaa'}}>{t.csvOrXlsx}</div>
              <input ref={fRef} type="file" accept=".csv,.xlsx,.xls" style={{display:'none'}} onChange={handleImpFile}/>
            </div>
            <div style={{display:'flex',justifyContent:'space-between'}}>
              <button style={{...S.btn,color:'#4a90e2',borderColor:'#4a90e2',fontSize:11}} onClick={downloadTemplate}>{t.downloadTemplate}</button>
              <button style={S.btn} onClick={()=>setModal(null)}>{t.cancel}</button>
            </div>
          </>}
          {step===2&&<>
            {(mdata.errors||[]).length>0&&<div style={{background:'#FFF3CD',border:'1px solid #ffc107',borderRadius:8,padding:'9px 12px',marginBottom:'1rem',fontSize:12,color:'#856404'}}>{mdata.errors.length} row(s) skipped: {mdata.errors.slice(0,3).join('; ')}</div>}
            <div style={{display:'flex',gap:8,marginBottom:'1rem'}}>
              {['add','replace'].map(m=><div key={m} onClick={()=>setMdata(prev=>({...prev,mode:m}))} style={{flex:1,border:`2px solid ${(mdata.mode||'add')===m?'#111':'#eee'}`,borderRadius:10,padding:'9px 12px',cursor:'pointer',background:(mdata.mode||'add')===m?'#f8f8f8':'#fff'}}>
                <div style={{fontSize:13,fontWeight:600,marginBottom:2}}>{m==='add'?t.addUpdate:t.replaceAll}</div>
                <div style={{fontSize:11,color:'#aaa'}}>{m==='add'?t.addUpdateDesc:t.replaceDesc}</div>
              </div>)}
            </div>
            <div style={{fontSize:12,color:'#555',marginBottom:8,fontWeight:500}}>{(mdata.parsed||[]).length} {t.productsReady}</div>
            <div style={{overflowX:'auto',border:'1px solid #eee',borderRadius:10,marginBottom:'1rem',maxHeight:200,overflowY:'auto'}}>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
                <thead><tr>{[t.name,'Root SKU',t.stock,t.amazon,t.walmart].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                <tbody>{(mdata.parsed||[]).map((p,i)=><tr key={i}><td style={S.td}>{p.name}</td><td style={S.td}><code style={{fontSize:10,background:'#f5f5f5',padding:'1px 4px',borderRadius:3}}>{p.sku||'—'}</code></td><td style={S.td}>{p.stock}</td><td style={S.td}>{p.amz||'—'}</td><td style={S.td}>{p.wmt||'—'}</td></tr>)}</tbody>
              </table>
            </div>
            <div style={{display:'flex',gap:8,justifyContent:'space-between'}}>
              <button style={S.btn} onClick={()=>setMdata(prev=>({...prev,step:1}))}>{t.back}</button>
              <div style={{display:'flex',gap:8}}>
                <button style={S.btn} onClick={()=>setModal(null)}>{t.cancel}</button>
                <button style={S.btnPrimary} onClick={confirmImport}>{t.import} {(mdata.parsed||[]).length} {t.products}</button>
              </div>
            </div>
          </>}
        </div>
      </div>
    );
  };

  const OrdersModal = () => {
    const step = mdata.step||1;
    const oRef = useRef();
    const plat = mdata.platform||'amazon';
    return (
      <div style={S.overlay} onClick={e=>e.target===e.currentTarget&&setModal(null)}>
        <div style={S.sheet}>
          <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem'}}>{t.uploadOrdersTitle} — {t.step} {Math.min(step,3)} / 3</div>
          {step===1&&<>
            <div style={{fontSize:12,color:'#555',marginBottom:12}}>{t.selectPlatform}</div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:12}}>
              {Object.entries(PLAT).map(([k,v])=><div key={k} onClick={()=>setMdata(prev=>({...prev,platform:k}))} style={{border:`2px solid ${plat===k?v.c:'#eee'}`,borderRadius:10,padding:'10px 12px',cursor:'pointer',background:plat===k?v.c+'11':'#fff',display:'flex',alignItems:'center',gap:8}}>
                <div style={{width:8,height:8,borderRadius:'50%',background:v.c,flexShrink:0}}/>
                <span style={{fontSize:13,fontWeight:plat===k?600:400}}>{v.l}</span>
              </div>)}
            </div>
            <div style={{border:'2px dashed #ddd',borderRadius:12,padding:'1.5rem',textAlign:'center',cursor:'pointer',background:'#fafafa'}} onClick={()=>oRef.current.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f){const dt=new DataTransfer();dt.items.add(f);oRef.current.files=dt.files;handleOrdFile({target:oRef.current})}}}>
              <div style={{fontSize:24,marginBottom:6}}>📂</div>
              <div style={{fontSize:13,fontWeight:600,marginBottom:3}}>{t.dropFile}</div>
              <div style={{fontSize:11,color:'#aaa'}}>{PLAT[plat].l} — {t.csvOrXlsx}</div>
              <input ref={oRef} type="file" accept=".csv,.xlsx,.xls" style={{display:'none'}} onChange={handleOrdFile}/>
            </div>
            <div style={{display:'flex',justifyContent:'flex-end',marginTop:12}}>
              <button style={S.btn} onClick={()=>setModal(null)}>{t.cancel}</button>
            </div>
          </>}
          {step===2&&<>
            <div style={{fontSize:12,color:'#555',marginBottom:12}}>{t.confirmColumns} {PLAT[plat].l}</div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:10,marginBottom:12}}>
              {[{l:t.skuCol,k:'skuCol'},{l:t.qtyCol,k:'qtyCol'},{l:t.orderIdCol,k:'oidCol'}].map(f=>(
                <div key={f.k} style={{display:'flex',flexDirection:'column',gap:3}}>
                  <label style={{fontSize:11,color:'#888',fontWeight:500}}>{f.l}</label>
                  <select style={{...S.input,width:'auto'}} value={mdata[f.k]??-1} onChange={e=>setMdata(prev=>({...prev,[f.k]:parseInt(e.target.value)}))}>
                    <option value={-1}>— not mapped —</option>
                    {(mdata.headers||[]).map((h,i)=><option key={i} value={i}>{h}</option>)}
                  </select>
                </div>
              ))}
            </div>
            <div style={{display:'flex',gap:8,justifyContent:'space-between'}}>
              <button style={S.btn} onClick={()=>setMdata(prev=>({...prev,step:1}))}>{t.back}</button>
              <div style={{display:'flex',gap:8}}>
                <button style={S.btn} onClick={()=>setModal(null)}>{t.cancel}</button>
                <button style={{...S.btn,background:'#e67e22',color:'#fff',border:'none'}} onClick={buildPreview}>{t.previewDeductions}</button>
              </div>
            </div>
          </>}
          {step===3&&<>
            <div style={{fontSize:12,color:'#555',marginBottom:8}}>{PLAT[plat].l} — {(mdata.preview||[]).filter(p=>p.prod).length} {t.matched}, {(mdata.preview||[]).filter(p=>!p.prod).length} {t.notFound}</div>
            {(mdata.warnings||[]).length>0&&<div style={{background:'#FFF3CD',border:'1px solid #ffc107',borderRadius:8,padding:'9px 12px',marginBottom:'1rem',fontSize:12,color:'#856404'}}>{mdata.warnings.slice(0,4).join(' · ')}</div>}
            <div style={{overflowX:'auto',border:'1px solid #eee',borderRadius:10,marginBottom:'1rem',maxHeight:220,overflowY:'auto'}}>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
                <thead><tr>{['SKU',t.product,t.current,t.sold,t.newStock].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                <tbody>{(mdata.preview||[]).map((item,i)=><tr key={i} style={{background:!item.prod?'#FFF5F5':item.newStock<0?'#FFFBF0':''}}>
                  <td style={S.td}><code style={{fontSize:10,background:'#f5f5f5',padding:'1px 4px',borderRadius:3}}>{item.sku}</code></td>
                  <td style={{...S.td,fontWeight:500,color:item.prod?'#111':'#dc3545'}}>{item.prod?item.prod.name:t.notFound}</td>
                  <td style={S.td}>{item.prod?item.prod.stock:'—'}</td>
                  <td style={{...S.td,color:'#dc3545',fontWeight:600}}>-{item.qty}</td>
                  <td style={{...S.td,fontWeight:600,color:item.newStock<0?'#dc3545':item.newStock<=10?'#856404':'#28a745'}}>{item.prod!=null?item.newStock:'—'}</td>
                </tr>)}</tbody>
              </table>
            </div>
            <div style={{display:'flex',gap:8,justifyContent:'space-between'}}>
              <button style={S.btn} onClick={()=>setMdata(prev=>({...prev,step:2}))}>{t.back}</button>
              <div style={{display:'flex',gap:8}}>
                <button style={S.btn} onClick={()=>setModal(null)}>{t.cancel}</button>
                <button style={{...S.btn,background:'#28a745',color:'#fff',border:'none'}} onClick={confirmOrders}>{t.applyDeductions} ({(mdata.preview||[]).filter(p=>p.prod).length})</button>
              </div>
            </div>
          </>}
          {step===4&&<>
            <div style={{textAlign:'center',padding:'1.5rem 0'}}>
              <div style={{fontSize:36,marginBottom:12}}>⚠️</div>
              <div style={{fontSize:15,fontWeight:600,marginBottom:8}}>{t.duplicateTitle}</div>
              <div style={{fontSize:12,color:'#666',marginBottom:'1.5rem'}}>{t.duplicateMsg}</div>
              <div style={{display:'flex',gap:10,justifyContent:'center'}}>
                <button style={S.btn} onClick={()=>setMdata(prev=>({...prev,step:3}))}>{t.goBack}</button>
                <button style={{...S.btn,background:'#dc3545',color:'#fff',border:'none'}} onClick={applyOrders}>{t.processAnyway}</button>
              </div>
            </div>
          </>}
        </div>
      </div>
    );
  };

  // ── RENDER ───────────────────────────────────────────────────
  return (
    <div style={S.app}>
      {/* NAV */}
      <nav style={S.nav}>
        <div style={{display:'flex',alignItems:'center',gap:'1.5rem'}}>
          <span style={{fontWeight:700,fontSize:16,letterSpacing:'.02em'}}>{t.title}</span>
          <button style={{...S.btn,background:page==='dashboard'?'rgba(255,255,255,.15)':'transparent',color:'#fff',border:'none',fontSize:13}} onClick={()=>setPage('dashboard')}>{t.dashboard}</button>
          <button style={{...S.btn,background:page==='chat'?'rgba(255,255,255,.15)':'transparent',color:'#fff',border:'none',fontSize:13}} onClick={()=>setPage('chat')}>{t.chat}</button>
        </div>
        <button style={{...S.btn,color:'#fff',border:'1px solid rgba(255,255,255,.3)',fontSize:12}} onClick={()=>setLang(l=>l==='en'?'es':'en')}>{lang==='en'?'🇲🇽 Español':'🇺🇸 English'}</button>
      </nav>

      <main style={S.main}>
        {loading && <div style={{textAlign:'center',padding:'3rem',color:'#aaa'}}>Loading...</div>}

        {/* DASHBOARD */}
        {!loading && page==='dashboard' && <>
          {/* Header */}
          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:'1.25rem',flexWrap:'wrap',gap:8}}>
            <div style={{fontSize:20,fontWeight:600}}>{t.dashboard}</div>
            <div style={{display:'flex',gap:7,flexWrap:'wrap'}}>
              <button style={S.btn} onClick={exportCSV}>{t.exportCSV}</button>
              <button style={{...S.btn,color:'#4a90e2',borderColor:'#4a90e2'}} onClick={()=>{setMdata({step:1,mode:'add',parsed:[],errors:[]});setModal('import')}}>{t.bulkImport}</button>
              <button style={{...S.btn,color:'#e67e22',borderColor:'#e67e22',fontWeight:600}} onClick={()=>{setMdata({step:1,platform:'amazon',headers:[],rows:[],skuCol:-1,qtyCol:-1,oidCol:-1,hash:'',preview:[],warnings:[]});setModal('orders')}}>{t.uploadOrders}</button>
              <button style={S.btnPrimary} onClick={()=>{setMdata({form:emptyProduct()});setModal('product')}}>{t.addProduct}</button>
            </div>
          </div>

          {/* Metrics */}
          <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:9,marginBottom:'1.5rem'}}>
            {[{l:t.totalProducts,v:prods.length},{l:t.inventoryValue,v:'$'+Math.round(totalVal).toLocaleString()},{l:t.lowStock,v:lowN,a:lowN>0},{l:t.critical,v:critN,a:critN>0}].map(m=>(
              <div key={m.l} style={{background:m.a?'#FFF3CD':'#fff',borderRadius:12,padding:'14px 16px',boxShadow:'0 1px 3px rgba(0,0,0,.08)'}}>
                <div style={{fontSize:11,color:m.a?'#856404':'#888',marginBottom:4}}>{m.l}</div>
                <div style={{fontSize:22,fontWeight:700,color:m.a?'#856404':'#111'}}>{m.v}</div>
              </div>
            ))}
          </div>

          {/* Tabs */}
          <div style={{display:'flex',gap:6,marginBottom:'1rem',flexWrap:'wrap'}}>
            {[[' all',t.allProducts],['alerts',`${t.alerts} (${alerts.length})`],['channels',t.channels],['log',t.log]].map(([k,l])=>(
              <button key={k} onClick={()=>setTab(k.trim())} style={{...S.btn,background:tab===k.trim()?'#111':'transparent',color:tab===k.trim()?'#fff':'#555',border:tab===k.trim()?'none':'1px solid #ddd'}}>{l}</button>
            ))}
          </div>

          <div style={S.card}>
            {tab==='all'&&<ProdTable list={prods}/>}
            {tab==='alerts'&&<ProdTable list={alerts}/>}
            {tab==='channels'&&(
              <div style={{overflowX:'auto'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:12}}>
                  <thead><tr>{[t.product,'Root SKU',t.amazon,t.walmart,t.target,t.temu,t.other].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                  <tbody>
                    {!prods.length&&<tr><td colSpan={7} style={{textAlign:'center',padding:'2rem',color:'#aaa'}}>{t.noProducts}</td></tr>}
                    {prods.map(p=><tr key={p.id}>
                      <td style={{...S.td,fontWeight:500}}>{p.name}</td>
                      <td style={S.td}><code style={{fontSize:10,background:'#f5f5f5',padding:'1px 5px',borderRadius:4}}>{p.sku||'—'}</code></td>
                      {['amz','wmt','tgt','temu','other_sku'].map(k=>{
                        const platKey = {amz:'amazon',wmt:'walmart',tgt:'target',temu:'temu',other_sku:'other'}[k];
                        return <td key={k} style={S.td}>{p[k]?<span style={{background:PLAT[platKey]?.c+'22',color:PLAT[platKey]?.c,padding:'1px 7px',borderRadius:99,fontSize:11,fontWeight:500}}>{p[k]}</span>:'—'}</td>;
                      })}
                    </tr>)}
                  </tbody>
                </table>
              </div>
            )}
            {tab==='log'&&(
              <div>
                {!logEntries.length&&<div style={{textAlign:'center',padding:'2rem',color:'#aaa',fontSize:13}}>{t.noLog}</div>}
                {logEntries.map((l,i)=>(
                  <div key={l.id} style={{display:'flex',gap:10,alignItems:'baseline',padding:'8px 0',borderBottom:i<logEntries.length-1?'1px solid #f5f5f5':'none',fontSize:12}}>
                    <span style={{fontSize:10,color:'#aaa',minWidth:130,flexShrink:0}}>{new Date(l.created_at).toLocaleString()}</span>
                    {l.platform&&<span style={{fontSize:10,background:PLAT[l.platform]?.c+'22',color:PLAT[l.platform]?.c,padding:'1px 6px',borderRadius:99,fontWeight:600,flexShrink:0}}>{PLAT[l.platform]?.l}</span>}
                    <span style={{flex:1}}>{l.description}</span>
                    {l.qty_change!=null&&<span style={{fontWeight:600,color:l.qty_change<0?'#dc3545':'#28a745',flexShrink:0}}>{l.qty_change>0?'+':''}{l.qty_change}</span>}
                  </div>
                ))}
              </div>
            )}
          </div>
        </>}

        {/* CHAT */}
        {!loading && page==='chat' && (
          <div style={{...S.card,display:'flex',flexDirection:'column',height:'calc(100vh - 120px)'}}>
            <div style={{fontSize:16,fontWeight:600,marginBottom:'1rem',paddingBottom:'0.75rem',borderBottom:'1px solid #eee'}}>{t.chat}</div>
            <div style={{flex:1,overflowY:'auto',display:'flex',flexDirection:'column',gap:12,paddingBottom:'1rem'}}>
              {chatMsgs.map((m,i)=>(
                <div key={i} style={{display:'flex',justifyContent:m.role==='user'?'flex-end':'flex-start'}}>
                  <div style={{maxWidth:'75%',padding:'10px 14px',borderRadius:12,background:m.role==='user'?'#111':'#f0f0f0',color:m.role==='user'?'#fff':'#111',fontSize:13,lineHeight:1.5,whiteSpace:'pre-wrap'}}>
                    {m.content}
                  </div>
                </div>
              ))}
              {chatLoading&&<div style={{display:'flex',justifyContent:'flex-start'}}><div style={{padding:'10px 14px',borderRadius:12,background:'#f0f0f0',fontSize:13,color:'#888'}}>{t.thinking}</div></div>}
              <div ref={chatEndRef}/>
            </div>
            <form onSubmit={sendChat} style={{display:'flex',gap:8,paddingTop:'0.75rem',borderTop:'1px solid #eee'}}>
              <input style={{...S.input,flex:1}} value={chatInput} onChange={e=>setChatInput(e.target.value)} placeholder={t.chatPlaceholder} disabled={chatLoading}/>
              <button type="submit" style={{...S.btnPrimary,padding:'8px 16px'}} disabled={chatLoading || !chatInput.trim()}>→</button>
            </form>
          </div>
        )}
      </main>

      {/* MODALS */}
      {modal==='product'&&<ProductModal/>}
      {modal==='import'&&<ImportModal/>}
      {modal==='orders'&&<OrdersModal/>}
    </div>
  );
}
