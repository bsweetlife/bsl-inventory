// BSL Inventory v4.23 - bulk-update-stock-tool
import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from './lib/supabase';

const T = {
  en: {
    title:'BSL Inventory',dashboard:'Dashboard',chat:'Chat with Claude',customers:'Customers',reports:'Reports',notes:'Notes',
    addProduct:'+ Add Product',bulkImport:'📥 Bulk Import',uploadOrders:'📦 Upload Orders',exportCSV:'⬇ Export CSV',uploadPacking:'📸 Packing List',
    totalProducts:'Total Products',inventoryValue:'Inventory Value',lowStock:'Low Stock',critical:'Critical',
    allProducts:'All Products',alerts:'Low Stock Alerts',channels:'By Channel',log:'Change Log',
    status:'Status',product:'Product',rootSku:'Root SKU',stock:'Stock (Singles)',days:'Days',cost:'Cost',price:'Price',actions:'Actions',
    ok:'OK',low:'Low',crit:'Critical',
    productName:'Product name',category:'Category',monthlyS:'Monthly sales (singles)',reorder:'Reorder point (singles)',supplier:'Supplier',
    marketplaceSKUs:'Marketplace SKUs & Pack Sizes',amazonSku:'Amazon SKU',walmartSku:'Walmart SKU',targetSku:'Target SKU',temuSku:'Temu SKU',otherSku:'Other SKU',
    packSize:'Pack size',
    save:'Save',cancel:'Cancel',edit:'Edit',delete:'Delete',
    noProducts:'No products yet — click + Add Product or use Bulk Import',
    noAlerts:'No low stock items',noLog:'No changes yet',
    chatPlaceholder:'Ask me anything about your inventory...',
    thinking:'Thinking...',
    amazon:'Amazon',walmart:'Walmart',target:'Target',temu:'Temu',other:'Other',direct:'Direct',
    customerName:'Customer name',customerEmail:'Email (optional)',customerPhone:'Phone (optional)',
    addSale:'+ New Sale',saleDate:'Sale date',platform:'Platform',
    addNote:'+ Add Note',noteTitle:'Title',noteContent:'Content',noteCategory:'Category',
    reorderSuggestions:'Reorder Suggestions',velocityReport:'Sales Velocity',
    packingListTitle:'Upload Packing List',processing:'Processing...',
    confirmDeductions:'Confirm Deductions',applyDeductions:'✓ Apply',
    inbound:'Inbound (stock IN)',outbound:'Outbound (stock OUT)',
    name:'Name',
  },
  es: {
    title:'BSL Inventario',dashboard:'Panel',chat:'Chat con Claude',customers:'Clientes',reports:'Reportes',notes:'Notas',
    addProduct:'+ Agregar Producto',bulkImport:'📥 Importar Masivo',uploadOrders:'📦 Subir Órdenes',exportCSV:'⬇ Exportar CSV',uploadPacking:'📸 Lista de Empaque',
    totalProducts:'Total Productos',inventoryValue:'Valor Inventario',lowStock:'Stock Bajo',critical:'Crítico',
    allProducts:'Todos los Productos',alerts:'Alertas Stock Bajo',channels:'Por Canal',log:'Registro',
    status:'Estado',product:'Producto',rootSku:'SKU Raíz',stock:'Stock (Singles)',days:'Días',cost:'Costo',price:'Precio',actions:'Acciones',
    ok:'OK',low:'Bajo',crit:'Crítico',
    productName:'Nombre del producto',category:'Categoría',monthlyS:'Ventas mensuales (singles)',reorder:'Punto de reorden (singles)',supplier:'Proveedor',
    marketplaceSKUs:'SKUs por Marketplace y Tamaño de Pack',amazonSku:'SKU Amazon',walmartSku:'SKU Walmart',targetSku:'SKU Target',temuSku:'SKU Temu',otherSku:'SKU Otro',
    packSize:'Tamaño de pack',
    save:'Guardar',cancel:'Cancelar',edit:'Editar',delete:'Eliminar',
    noProducts:'Sin productos — haz clic en + Agregar Producto',
    noAlerts:'Sin artículos con stock bajo',noLog:'Sin cambios aún',
    chatPlaceholder:'Pregúntame sobre tu inventario...',
    thinking:'Pensando...',
    amazon:'Amazon',walmart:'Walmart',target:'Target',temu:'Temu',other:'Otro',direct:'Directo',
    customerName:'Nombre del cliente',customerEmail:'Email (opcional)',customerPhone:'Teléfono (opcional)',
    addSale:'+ Nueva Venta',saleDate:'Fecha de venta',platform:'Plataforma',
    addNote:'+ Agregar Nota',noteTitle:'Título',noteContent:'Contenido',noteCategory:'Categoría',
    reorderSuggestions:'Sugerencias de Reorden',velocityReport:'Velocidad de Ventas',
    packingListTitle:'Subir Lista de Empaque',processing:'Procesando...',
    confirmDeductions:'Confirmar Deducciones',applyDeductions:'✓ Aplicar',
    inbound:'Entrada (stock IN)',outbound:'Salida (stock OUT)',
    name:'Nombre',
  }
};

const PLAT={amazon:{l:'Amazon',c:'#FF9900',sku:['sku','seller-sku','asin'],qty:['quantity','quantity-purchased','qty','units'],oid:['amazon-order-id','order-id','order_id']},walmart:{l:'Walmart',c:'#0071DC',sku:['sku','item-sku','seller-sku'],qty:['qty','quantity','units-ordered'],oid:['order_id','purchase-order-id','order-id']},target:{l:'Target Plus',c:'#CC0000',sku:['sku','tcin','seller-sku'],qty:['quantity','qty','quantity-ordered'],oid:['order-id','po-number','order_id']},temu:{l:'Temu',c:'#7B2D8B',sku:['sku','product-sku','seller-sku'],qty:['quantity','qty','amount'],oid:['order-id','order_id','order-number']},other:{l:'Other',c:'#888',sku:['sku','item-sku','product-sku'],qty:['quantity','qty','units'],oid:['order-id','order_id']}};
const MPK=['amz','wmt','tgt','temu','other_sku'];
const COL_MAP={"name":"name","product name":"name","product":"name","root sku":"sku","sku":"sku","category":"category","cat":"category","stock":"stock","stock (singles)":"stock","current stock":"stock","qty":"stock","quantity":"stock","monthly sales":"velocity","monthly sales (singles)":"velocity","velocity":"velocity","cost":"cost","cost per single ($)":"cost","cost price":"cost","price":"price","selling price single ($)":"price","selling price":"price","reorder":"reorder","reorder point":"reorder","reorder point (singles)":"reorder","supplier":"supplier","supplier name":"supplier","amazon sku":"amz","amz":"amz","amz pack size":"amz_pack_size","walmart sku":"wmt","wmt":"wmt","wmt pack size":"wmt_pack_size","target sku":"tgt","tgt":"tgt","tgt pack size":"tgt_pack_size","temu sku":"temu","temu":"temu","temu pack size":"temu_pack_size","other sku":"other_sku","other":"other_sku","other pack size":"other_pack_size","product notes":"notes_field","raw material cost per kg":"raw_material_cost_per_kg","raw material price per kg":"raw_material_cost_per_kg","raw_material_cost_per_kg":"raw_material_cost_per_kg","raw material/kg":"raw_material_cost_per_kg","jumbo box cost":"jumbo_box_cost","jumbo_box_cost":"jumbo_box_cost","jumbo box":"jumbo_box_cost","weight oz":"weight_oz","weight (oz)":"weight_oz","packaging cost":"packaging_cost","packaging_cost":"packaging_cost","box cost":"box_cost","box_cost":"box_cost"};

const gs=p=>{const d=(p.velocity||0)/30;if(!d)return'ok';const v=p.stock/d;return v<=15?'crit':v<=30?'low':'ok'};
const gd=p=>{const d=(p.velocity||0)/30;return d?Math.round(p.stock/d):null};
const fm=n=>(n!=null&&n!=='')?'$'+parseFloat(n).toFixed(2):'—';
const fm2=n=>'$'+parseFloat(n||0).toFixed(4);
const hs=s=>{let h=0;for(let i=0;i<Math.min(s.length,500);i++)h=(Math.imul(31,h)+s.charCodeAt(i))|0;return h.toString()};
const fc=(hdrs,cs)=>{for(const c of cs){const i=hdrs.findIndex(h=>h.toLowerCase().replace(/[\s_-]+/g,'-')===c);if(i>=0)return i;}for(const c of cs){const i=hdrs.findIndex(h=>h.toLowerCase().includes(c.replace(/-/g,'')));if(i>=0)return i;}return -1};
const ep=()=>({id:null,name:'',sku:'',category:'',stock:'',velocity:'',cost:'',price:'',reorder:'',supplier:'',amz:'',wmt:'',tgt:'',temu:'',other_sku:'',amz_pack_size:1,wmt_pack_size:1,tgt_pack_size:1,temu_pack_size:1,other_pack_size:1,product_type:'finished',weight_oz:'',raw_material_cost_per_kg:'',packaging_cost:'',box_cost:'',jumbo_box_cost:'',cost_notes:''});

function readXLSX(file,cb){const r=new FileReader();r.onload=e=>{try{const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'});const ws=wb.Sheets[wb.SheetNames[0]];cb(null,XLSX.utils.sheet_to_json(ws,{header:1,defval:''}))}catch(err){cb(err)}};r.readAsArrayBuffer(file)}

const S={
  app:{fontFamily:'system-ui,sans-serif',minHeight:'100vh',background:'#f0f0f0'},
  nav:{background:'#111',color:'#fff',padding:'0 1.5rem',display:'flex',alignItems:'center',justifyContent:'space-between',height:52,position:'sticky',top:0,zIndex:50},
  main:{maxWidth:1200,margin:'0 auto',padding:'1.25rem'},
  card:{background:'#fff',borderRadius:12,padding:'1.25rem',boxShadow:'0 1px 3px rgba(0,0,0,.08)'},
  btn:{fontSize:12,padding:'6px 13px',borderRadius:8,border:'0.5px solid #ddd',background:'transparent',cursor:'pointer',display:'inline-flex',alignItems:'center',gap:4,fontFamily:'inherit'},
  btnP:{fontSize:12,padding:'6px 13px',borderRadius:8,border:'none',background:'#111',color:'#fff',cursor:'pointer',display:'inline-flex',alignItems:'center',gap:4,fontFamily:'inherit'},
  inp:{fontSize:13,padding:'6px 9px',borderRadius:8,border:'0.5px solid #ddd',width:'100%',boxSizing:'border-box',fontFamily:'inherit'},
  overlay:{position:'fixed',top:0,left:0,right:0,bottom:0,background:'rgba(0,0,0,.45)',zIndex:200,display:'flex',alignItems:'center',justifyContent:'center'},
  sheet:{background:'#fff',borderRadius:16,padding:'1.5rem',width:600,maxWidth:'95vw',maxHeight:'88vh',overflowY:'auto'},
  th:{textAlign:'left',fontSize:10,fontWeight:600,color:'#888',padding:'7px 10px',borderBottom:'1px solid #eee',background:'#fafafa',whiteSpace:'nowrap',textTransform:'uppercase',letterSpacing:'.04em',position:'sticky',top:0},
  td:{padding:'8px 10px',borderBottom:'1px solid #f5f5f5',fontSize:12,verticalAlign:'middle',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'},
};

const Badge=({status,t})=>{const m={ok:['#D4EDDA','#155724'],low:['#FFF3CD','#856404'],crit:['#F8D7DA','#721C24']};const[bg,c]=m[status]||m.ok;return<span style={{background:bg,color:c,padding:'2px 8px',borderRadius:99,fontSize:11,fontWeight:500}}>{t[status]||status}</span>};

function LoginScreen(){
  const[email,setEmail]=useState('');
  const[password,setPassword]=useState('');
  const[error,setError]=useState('');
  const[loading,setLoading]=useState(false);
  async function handleLogin(e){
    e.preventDefault();setLoading(true);setError('');
    const{error}=await supabase.auth.signInWithPassword({email,password});
    if(error)setError(error.message);
    setLoading(false);
  }
  return(
    <div style={{display:'flex',alignItems:'center',justifyContent:'center',height:'100vh',background:'#f0f0f0',fontFamily:'system-ui,sans-serif'}}>
      <div style={{background:'#fff',borderRadius:16,padding:'2.5rem',width:380,boxShadow:'0 4px 24px rgba(0,0,0,.1)'}}>
        <div style={{textAlign:'center',marginBottom:'2rem'}}>
          <div style={{fontSize:24,fontWeight:700,marginBottom:4}}>BSL Inventory</div>
          <div style={{fontSize:13,color:'#888'}}>Sign in to continue</div>
        </div>
        <form onSubmit={handleLogin} style={{display:'flex',flexDirection:'column',gap:12}}>
          <div style={{display:'flex',flexDirection:'column',gap:4}}>
            <label style={{fontSize:12,color:'#888',fontWeight:500}}>Email</label>
            <input style={{fontSize:13,padding:'9px 12px',borderRadius:8,border:'1px solid #ddd'}} type="email" placeholder="your@email.com" value={email} onChange={e=>setEmail(e.target.value)} required/>
          </div>
          <div style={{display:'flex',flexDirection:'column',gap:4}}>
            <label style={{fontSize:12,color:'#888',fontWeight:500}}>Password</label>
            <input style={{fontSize:13,padding:'9px 12px',borderRadius:8,border:'1px solid #ddd'}} type="password" placeholder="••••••••" value={password} onChange={e=>setPassword(e.target.value)} required/>
          </div>
          {error&&<div style={{fontSize:12,color:'#dc3545',background:'#F8D7DA',padding:'8px 12px',borderRadius:8}}>{error}</div>}
          <button type="submit" style={{marginTop:4,padding:'10px',borderRadius:8,border:'none',background:'#111',color:'#fff',fontSize:14,fontWeight:600,cursor:'pointer'}} disabled={loading}>{loading?'Signing in...':'Sign in'}</button>
        </form>
      </div>
    </div>
  );
}

export default function App(){
  const[session,setSession]=useState(null);
  const[authChecked,setAuthChecked]=useState(false);
  useEffect(()=>{
    supabase.auth.getSession().then(({data:{session}})=>{setSession(session);setAuthChecked(true);});
    const{data:{subscription}}=supabase.auth.onAuthStateChange((_,session)=>setSession(session));
    return()=>subscription.unsubscribe();
  },[]);
  if(!authChecked)return<div style={{display:'flex',alignItems:'center',justifyContent:'center',height:'100vh',fontFamily:'system-ui',color:'#888',fontSize:14}}>Loading...</div>;
  if(!session)return<LoginScreen/>;
  return<AppMain session={session}/>;
}

function AppMain({session}){
  const userEmail=session?.user?.email||'unknown';
  const[lang,setLang]=useState('en');
  const[page,setPage]=useState('dashboard');
  const[tab,setTab]=useState('all');
  const[prods,setProds]=useState([]);
  const[logEntries,setLog]=useState([]);
  const[hashes,setHashes]=useState([]);
  const[customers,setCustomers]=useState([]);
  const[agentNotes,setAgentNotes]=useState([]);
  const[globalSettings,setGlobalSettings]=useState({raw_material_waste_pct:0.005,packaging_waste_pct:0.005,filling_cost:1.15,packaging_cost_default:0.45,fee_amz:0.15,fee_wmt:0.15,fee_tgt:0.15,fee_temu:0.12,fee_other:0.10});
  const[costCalcOverrides,setCostCalcOverrides]=useState({});
  const[costModal,setCostModal]=useState(null);
  const[locationModal,setLocationModal]=useState(null);
  const[locations,setLocations]=useState([]);
  const[loading,setLoading]=useState(true);
  const[modal,setModal]=useState(null);
  const[mdata,setMdata]=useState({});
  const[chatMsgs,setChatMsgs]=useState([]);
  const[chatInput,setChatInput]=useState('');
  const[chatLoading,setChatLoading]=useState(false);
  const[chatSessions,setChatSessions]=useState([]);
  const[currentSessionId,setCurrentSessionId]=useState(null);
  const[sessionsLoading,setSessionsLoading]=useState(false);
  const[chatFile,setChatFile]=useState(null);
  const[pendingDangerousTool,setPendingDangerousTool]=useState(null);
  const[passwordInput,setPasswordInput]=useState('');
  const[passwordError,setPasswordError]=useState('');
  const chatEndRef=useRef(null);
  const chatFileRef=useRef(null);
  const t=T[lang];

  useEffect(()=>{loadAll(true);},[]);
  useEffect(()=>{if(chatMsgs.length===0)setChatMsgs([{role:'assistant',content:T[lang].chatPlaceholder}]);},[lang]);
  useEffect(()=>{chatEndRef.current?.scrollIntoView({behavior:'smooth'});},[chatMsgs]);

  async function loadAll(isFirstLoad=false){
    setLoading(true);
    try{
      const[{data:p},{data:l},{data:h},{data:c},{data:n},{data:gs},{data:locs}]=await Promise.all([
        supabase.from('products').select('*').order('name'),
        supabase.from('change_log').select('*').order('created_at',{ascending:false}).limit(200),
        supabase.from('uploaded_files').select('file_hash'),
        supabase.from('customer_sales').select('*,customer_sale_items(*)').order('created_at',{ascending:false}).limit(100),
        supabase.from('agent_notes').select('*').order('created_at',{ascending:false}),
        supabase.from('global_settings').select('*'),
        supabase.from('inventory_locations').select('*'),
      ]);
      if(p)setProds(p);
      if(l)setLog(l);
      if(h)setHashes(h.map(x=>x.file_hash));
      if(c)setCustomers(c);
      if(n)setAgentNotes(n);
      if(gs){const s={};gs.forEach(g=>s[g.key]=parseFloat(g.value));setGlobalSettings(prev=>({...prev,...s}));}
      if(locs)setLocations(locs);
      // Load chat sessions separately (not blocking)
      loadChatSessions();
    }catch(e){console.error(e);}
    setLoading(false);
    // Only reset chat on first page load, never during data refreshes mid-conversation
    if(isFirstLoad){
      setChatMsgs([{role:'assistant',content:`Hi! I'm Claude, your BSL inventory manager. I have full access to your inventory, customer sales, and notes. Ask me anything — stock levels, reorder suggestions, sales trends, or upload a packing list and I'll process it automatically.`}]);
    }
  }

  async function saveGlobalSettings(gs){
    for(const[key,value] of Object.entries(gs)){
      await supabase.from('global_settings').update({value,updated_at:new Date().toISOString()}).eq('key',key);
    }
    setGlobalSettings(gs);
  }

  async function saveProduct(form){
    const entry={...form,stock:parseFloat(form.stock)||0,velocity:parseFloat(form.velocity)||0,cost:parseFloat(form.cost)||0,price:parseFloat(form.price)||0,reorder:parseFloat(form.reorder)||0,amz_pack_size:parseFloat(form.amz_pack_size)||1,wmt_pack_size:parseFloat(form.wmt_pack_size)||1,tgt_pack_size:parseFloat(form.tgt_pack_size)||1,temu_pack_size:parseFloat(form.temu_pack_size)||1,other_pack_size:parseFloat(form.other_pack_size)||1,weight_oz:parseFloat(form.weight_oz)||0,raw_material_cost_per_kg:parseFloat(form.raw_material_cost_per_kg)||0,packaging_cost:parseFloat(form.packaging_cost)||0,box_cost:parseFloat(form.box_cost)||0,jumbo_box_cost:parseFloat(form.jumbo_box_cost)||0,product_type:form.product_type||'finished'};
    delete entry.id;
    if(form.id){
      const old=prods.find(p=>p.id===form.id);
      await supabase.from('products').update(entry).eq('id',form.id);
      const diff=(parseFloat(form.stock)||0)-old.stock;
      await supabase.from('change_log').insert({description:diff!==0?`Stock: ${form.name} ${old.stock}→${form.stock}`:`Updated: ${form.name}`,qty_change:diff||null,user_email:userEmail});
    }else{
      await supabase.from('products').insert(entry);
      await supabase.from('change_log').insert({description:`Added: ${form.name}`,qty_change:parseFloat(form.stock)||0});
    }
    await loadAll();setModal(null);
  }

  async function deleteProduct(id){
    const p=prods.find(x=>x.id===id);
    if(!window.confirm(`Delete "${p.name}"?`))return;
    await supabase.from('products').delete().eq('id',id);
    await supabase.from('change_log').insert({description:`Deleted: ${p.name}`,user_email:userEmail});
    await loadAll();
  }

  // ── BULK IMPORT ─────────────────────────────────────────────
  function handleImpFile(e){
    const f=e.target.files[0];if(!f)return;
    readXLSX(f,(err,rows)=>{
      if(err||rows.length<2){alert('Cannot read file');return;}
      const hdrs=rows[0].map(h=>(h||'').toString().trim().toLowerCase());
      const fm2=hdrs.map(h=>COL_MAP[h]||null);
      const parsed=[],errors=[];
      rows.slice(1).filter(r=>r.some(c=>c!=='')).forEach((row,i)=>{
        const o={};fm2.forEach((ff,ci)=>{if(ff)o[ff]=(row[ci]??'').toString().trim()});
        if(!o.name){errors.push(`Row ${i+2}: no name`);return;}
        parsed.push({name:o.name,sku:o.sku||'',category:o.category||'',stock:parseFloat(o.stock)||0,velocity:parseFloat(o.velocity)||0,cost:parseFloat(o.cost)||0,price:parseFloat(o.price)||0,reorder:parseFloat(o.reorder)||0,supplier:o.supplier||'',amz:o.amz||'',wmt:o.wmt||'',tgt:o.tgt||'',temu:o.temu||'',other_sku:o.other_sku||'',amz_pack_size:parseFloat(o.amz_pack_size)||1,wmt_pack_size:parseFloat(o.wmt_pack_size)||1,tgt_pack_size:parseFloat(o.tgt_pack_size)||1,temu_pack_size:parseFloat(o.temu_pack_size)||1,other_pack_size:parseFloat(o.other_pack_size)||1,weight_oz:parseFloat(o.weight_oz)||0,raw_material_cost_per_kg:parseFloat(o.raw_material_cost_per_kg)||0,packaging_cost:parseFloat(o.packaging_cost)||0,box_cost:parseFloat(o.box_cost)||0,jumbo_box_cost:parseFloat(o.jumbo_box_cost)||0});
      });
      setMdata(prev=>({...prev,step:2,parsed,errors}));
    });
    e.target.value='';
  }

  async function confirmImport(){
    const{parsed,mode}=mdata;
    if(mode==='replace'){await supabase.from('products').delete().neq('id',0);await supabase.from('products').insert(parsed);}
    else{for(const p of parsed){const existing=prods.find(x=>x.sku&&x.sku===p.sku);if(existing)await supabase.from('products').update(p).eq('id',existing.id);else await supabase.from('products').insert(p);}}
    await supabase.from('change_log').insert({description:`Bulk import: ${parsed.length} products (${mode||'add'})`,qty_change:parsed.length,user_email:userEmail});
    await loadAll();setModal(null);
  }

  // ── ORDERS UPLOAD ───────────────────────────────────────────
  function handleOrdFile(e){
    const f=e.target.files[0];if(!f)return;
    readXLSX(f,(err,rows)=>{
      if(err||rows.length<2){alert('Cannot read file');return;}
      const hash=hs(rows[0].join(',')+(rows[1]||[]).join(','));
      const headers=rows[0].map(h=>(h||'').toString().trim());
      const dataRows=rows.slice(1).filter(r=>r.some(c=>c!==''));
      const cfg=PLAT[mdata.platform||'amazon'];
      const hl=headers.map(h=>h.toLowerCase());
      setMdata(prev=>({...prev,hash,headers,rows:dataRows,skuCol:fc(hl,cfg.sku),qtyCol:fc(hl,cfg.qty),oidCol:fc(hl,cfg.oid),step:2}));
    });
    e.target.value='';
  }

  function buildPreview(){
    if(mdata.skuCol<0||mdata.qtyCol<0){alert('Map SKU and Quantity columns');return;}
    const mpKey={amazon:'amz',walmart:'wmt',target:'tgt',temu:'temu',other:'other_sku'}[mdata.platform]||'other_sku';
    const packKey={amazon:'amz_pack_size',walmart:'wmt_pack_size',target:'tgt_pack_size',temu:'temu_pack_size',other:'other_pack_size'}[mdata.platform]||'other_pack_size';
    const agg={};
    mdata.rows.forEach(row=>{
      const sku=(row[mdata.skuCol]||'').toString().trim();
      const qty=parseFloat((row[mdata.qtyCol]||'0').toString().replace(/[^0-9.-]/g,''))||0;
      if(!sku||qty<=0)return;
      if(!agg[sku])agg[sku]={sku,packsOrdered:0};
      agg[sku].packsOrdered+=qty;
    });
    const warnings=[];
    const preview=Object.values(agg).map(item=>{
      let prod=prods.find(p=>p[mpKey]===item.sku)||prods.find(p=>p.sku===item.sku)||prods.find(p=>MPK.some(k=>p[k]===item.sku));
      const packSize=prod?((prod[packKey])||1):1;
      const singlesDeducted=item.packsOrdered*packSize;
      const newStock=prod?prod.stock-singlesDeducted:null;
      if(prod&&newStock<0)warnings.push(`"${prod.name}" will go negative (${newStock})`);
      if(!prod)warnings.push(`"${item.sku}" not found`);
      return{...item,prod,packSize,singlesDeducted,newStock};
    });
    setMdata(prev=>({...prev,preview,warnings,step:3}));
  }

  // ── AUTO VELOCITY CALCULATION ───────────────────────────────
  async function recalculateVelocity(productIds=null){
    // Get all orders from last 30 days
    const thirtyDaysAgo=new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate()-30);
    
    let query=supabase.from('orders_log')
      .select('product_id,qty_sold,created_at')
      .gte('created_at',thirtyDaysAgo.toISOString());
    
    if(productIds&&productIds.length>0){
      query=query.in('product_id',productIds);
    }
    
    const{data:orders,error}=await query;
    if(error||!orders||orders.length===0)return;

    // Aggregate sales per product
    const salesByProduct={};
    orders.forEach(o=>{
      if(!salesByProduct[o.product_id])salesByProduct[o.product_id]=0;
      salesByProduct[o.product_id]+=parseFloat(o.qty_sold)||0;
    });

    // Calculate days since first order (up to 30)
    const firstOrderDate=new Date(Math.min(...orders.map(o=>new Date(o.created_at))));
    const daysSinceFirst=Math.max(1,Math.min(30,Math.round((new Date()-firstOrderDate)/(1000*60*60*24))));

    // Update velocity for each product (monthly = sales/days * 30)
    for(const[productId,totalSold] of Object.entries(salesByProduct)){
      const monthlyVelocity=Math.round((totalSold/daysSinceFirst)*30*10)/10;
      await supabase.from('products').update({velocity:monthlyVelocity}).eq('id',parseInt(productId));
    }
  }

  async function applyOrders(){
    const platLabel=PLAT[mdata.platform]?.l;
    for(const item of mdata.preview){
      if(!item.prod)continue;
      const newStock=item.prod.stock-item.singlesDeducted;
      await supabase.from('products').update({stock:newStock}).eq('id',item.prod.id);
      await supabase.from('orders_log').insert({platform:mdata.platform,order_sku:item.sku,product_id:item.prod.id,product_name:item.prod.name,qty_sold:item.singlesDeducted,stock_before:item.prod.stock,stock_after:newStock});
      await supabase.from('change_log').insert({description:`${platLabel}: ${item.prod.name} [${item.sku}] ${item.packsOrdered} packs × ${item.packSize} = ${item.singlesDeducted} singles. ${item.prod.stock}→${newStock}`,qty_change:-item.singlesDeducted,platform:mdata.platform,user_email:userEmail});
    }
    await supabase.from('change_log').insert({description:`${platLabel} upload complete`,user_email:userEmail});
    if(mdata.hash)await supabase.from('uploaded_files').insert({file_hash:mdata.hash});
    // Recalculate velocity for affected products
    const affectedIds=mdata.preview.filter(i=>i.prod).map(i=>i.prod.id);
    await recalculateVelocity(affectedIds);
    await loadAll();setModal(null);
  }

  function confirmOrders(){if(hashes.includes(mdata.hash)){setMdata(prev=>({...prev,step:4}));return;}applyOrders();}

  // ── PACKING LIST ────────────────────────────────────────────
  function handlePackingFile(e){
    const f=e.target.files[0];if(!f)return;
    const reader=new FileReader();
    reader.onload=async(evt)=>{
      const base64=evt.target.result.split(',')[1];
      const mediaType=f.type||'image/jpeg';
      setMdata(prev=>({...prev,processing:true,step:2}));
      try{
        const res=await fetch('/api/parse-packing-list',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({imageBase64:base64,mediaType,products:prods})});
        const data=await res.json();
        if(data.error){alert('Could not read document: '+data.error);setMdata(prev=>({...prev,processing:false,step:1}));return;}
        setMdata(prev=>({...prev,processing:false,parsed:data,step:3}));
      }catch(err){alert('Error processing file');setMdata(prev=>({...prev,processing:false,step:1}));}
    };
    reader.readAsDataURL(f);
    e.target.value='';
  }

  async function applyPackingList(){
    const{parsed,customerName,customerEmail,customerPhone}=mdata;
    const direction=parsed.direction||'outbound';
    // Create packing list record
    await supabase.from('packing_list_uploads').insert({supplier:parsed.supplier,customer_name:parsed.customer||customerName,packing_list_number:parsed.packingListNumber,po_number:parsed.poNumber,direction,total_singles:parsed.totalSingles,processed:true});
    // If outbound — create customer sale and deduct
    if(direction==='outbound'){
      const{data:sale}=await supabase.from('customer_sales').insert({customer_name:parsed.customer||customerName,customer_email:customerEmail||null,customer_phone:customerPhone||null,sale_date:parsed.date||new Date().toISOString().slice(0,10),platform:'direct',total_singles:parsed.totalSingles}).select().single();
      for(const item of(parsed.items||[])){
        if(item.included===false||!item.matchedSku)continue;
        const prod=prods.find(p=>p.sku===item.matchedSku);
        if(!prod)continue;
        const newStock=prod.stock-item.singles;
        await supabase.from('products').update({stock:newStock}).eq('id',prod.id);
        await supabase.from('customer_sale_items').insert({sale_id:sale.id,product_id:prod.id,product_name:prod.name,root_sku:prod.sku,pack_size:item.packSize||1,packs_ordered:item.totalUnits/(item.packSize||1),singles_deducted:item.singles});
        await supabase.from('change_log').insert({description:`Direct sale to ${parsed.customer||customerName}: ${prod.name} -${item.singles} singles`,qty_change:-item.singles,platform:'direct',user_email:userEmail});
      }
    }else{
      // Inbound — add stock
      for(const item of(parsed.items||[])){
        if(item.included===false||!item.matchedSku)continue;
        const prod=prods.find(p=>p.sku===item.matchedSku);
        if(!prod)continue;
        const newStock=prod.stock+item.singles;
        await supabase.from('products').update({stock:newStock}).eq('id',prod.id);
        await supabase.from('change_log').insert({description:`Received from ${parsed.supplier}: ${prod.name} +${item.singles} singles`,qty_change:item.singles,platform:'inbound',user_email:userEmail});
      }
    }
    // Recalculate velocity for affected products
    const packingIds=(parsed.items||[]).filter(i=>i.matchedSku).map(i=>prods.find(p=>p.sku===i.matchedSku)?.id).filter(Boolean);
    if(packingIds.length)await recalculateVelocity(packingIds);
    await loadAll();setModal(null);
  }

  // ── EXPORT CSV ──────────────────────────────────────────────
  function exportCSV(){
    const headers=['Name','Root SKU','Category','Stock (Singles)','Monthly Sales','Cost','Price','Reorder','Supplier','Amazon SKU','AMZ Pack','Walmart SKU','WMT Pack','Target SKU','TGT Pack','Temu SKU','Temu Pack','Other SKU','Other Pack','Status','Days Left','Product Type','Weight (oz)','Raw Material/kg ($)','Packaging Cost ($)','Box Cost ($)','Jumbo Box Cost ($)'];
    const rows=prods.map(p=>[p.name,p.sku,p.category,p.stock,p.velocity,p.cost,p.price,p.reorder,p.supplier,p.amz,p.amz_pack_size||1,p.wmt,p.wmt_pack_size||1,p.tgt,p.tgt_pack_size||1,p.temu,p.temu_pack_size||1,p.other_sku,p.other_pack_size||1,gs(p),gd(p)??'—',p.product_type||'finished',p.weight_oz||'',p.raw_material_cost_per_kg||'',p.packaging_cost||'',p.box_cost||'',p.jumbo_box_cost||''].map(v=>`"${String(v??'').replace(/"/g,'""')}"`).join(','));
    const csv=[headers.join(','),...rows].join('\n');
    const a=document.createElement('a');a.href='data:text/csv;charset=utf-8,'+encodeURIComponent(csv);a.download='bsl_inventory_'+new Date().toISOString().slice(0,10)+'.csv';a.click();
  }

  function downloadTemplate(){
    const h=['Root SKU','Product Name','Category','Supplier','Stock (Singles)','Monthly Sales (Singles)','Cost Per Single ($)','Selling Price Single ($)','Reorder Point (Singles)','Amazon SKU','AMZ Pack Size','AMZ Selling Price ($)','Walmart SKU','WMT Pack Size','WMT Selling Price ($)','Target SKU','TGT Pack Size','TGT Selling Price ($)','Temu SKU','Temu Pack Size','Temu Selling Price ($)','Other SKU','Other Pack Size','Other Selling Price ($)','Product Notes','Weight (oz)','Raw Material Cost Per Kg','Packaging Cost','Box Cost','Jumbo Box Cost'];
    const ex=['BSL-LACT-150','Lactose Free 1.5lb','Lactose Free','EVI Labs','0','0','0.00','0.00','0','AMZ-LF150-001','6','0.00','WMT-LF150-001','6','0.00','','1','0.00','','1','0.00','','1','0.00','EVI Labs product','24','4.24','0.50','0.05','0.00'];
    const csv=[h,ex].map(r=>r.map(v=>`"${v}"`).join(',')).join('\n');
    const a=document.createElement('a');a.href='data:text/csv;charset=utf-8,'+encodeURIComponent(csv);a.download='BSL_SKU_Template.csv';a.click();
  }

  // ── AI CHAT ─────────────────────────────────────────────────
  // ── INVENTORY TOOLS ─────────────────────────────────────────
  const inventoryTools=[
    {
      name:'update_stock',
      description:'Update the stock level of a product. Use this when the user asks to add, remove, set, or adjust stock for any product.',
      input_schema:{
        type:'object',
        properties:{
          product_id:{type:'number',description:'The numeric ID of the product'},
          product_name:{type:'string',description:'Name of the product for confirmation'},
          new_stock:{type:'number',description:'The new stock level in singles after the update'},
          reason:{type:'string',description:'Brief reason for the update e.g. "manual adjustment", "received shipment", "direct sale"'}
        },
        required:['product_id','product_name','new_stock','reason']
      }
    },
    {
      name:'update_product_field',
      description:'Update a specific field of a product like price, cost, velocity, reorder point, or supplier.',
      input_schema:{
        type:'object',
        properties:{
          product_id:{type:'number',description:'The numeric ID of the product'},
          product_name:{type:'string',description:'Name of the product'},
          field:{type:'string',enum:['price','cost','velocity','reorder','supplier','category'],description:'Which field to update'},
          value:{type:'string',description:'New value for the field'}
        },
        required:['product_id','product_name','field','value']
      }
    },
    {
      name:'bulk_update_stock',
      description:'Update stock for multiple products at once. Use this when the user provides a list of products with quantities — e.g. pasted inventory counts, box counts, or shipment lists. Convert boxes×units to singles before calling. Call this ONCE with the full array, do not loop with update_stock.',
      input_schema:{
        type:'object',
        properties:{
          updates:{
            type:'array',
            description:'Array of stock updates to apply',
            items:{
              type:'object',
              properties:{
                product_id:{type:'number',description:'Numeric ID of the product'},
                product_name:{type:'string',description:'Name of the product'},
                new_stock:{type:'number',description:'New stock level in singles'},
                reason:{type:'string',description:'Brief reason e.g. "inventory count", "shipment received"'}
              },
              required:['product_id','product_name','new_stock','reason']
            }
          }
        },
        required:['updates']
      }
    },
    {
      name:'clear_all_stock',
      description:'Set ALL products stock to 0. ONLY use this when the user explicitly asks to clear, reset, or zero out all inventory stock. This is a destructive action that requires a password confirmation from the user.',
      input_schema:{
        type:'object',
        properties:{
          reason:{type:'string',description:'Reason for clearing all stock'}
        },
        required:['reason']
      }
    }
  ];

  async function executeToolCall(toolName, toolInput){
    if(toolName==='update_stock'){
      const{product_id,product_name,new_stock,reason}=toolInput;
      const prod=prods.find(p=>p.id===product_id);
      if(!prod)return{success:false,error:`Product ID ${product_id} not found`};
      const old_stock=prod.stock;
      await supabase.from('products').update({stock:new_stock}).eq('id',product_id);
      await supabase.from('change_log').insert({description:`Chat update: ${product_name} ${old_stock}→${new_stock} (${reason})`,qty_change:new_stock-old_stock,user_email:userEmail});
      await loadAll();
      return{success:true,message:`Updated ${product_name}: ${old_stock} → ${new_stock} singles (${reason})`};
    }
    if(toolName==='update_product_field'){
      const{product_id,product_name,field,value}=toolInput;
      const update={};
      update[field]=isNaN(value)?value:parseFloat(value);
      await supabase.from('products').update(update).eq('id',product_id);
      await supabase.from('change_log').insert({description:`Chat update: ${product_name} ${field} set to ${value}`,user_email:userEmail});
      await loadAll();
      return{success:true,message:`Updated ${product_name}: ${field} = ${value}`};
    }
    if(toolName==='bulk_update_stock'){
      const{updates}=toolInput;
      if(!updates||!updates.length)return{success:false,error:'No updates provided'};
      const results=[];
      for(const u of updates){
        const prod=prods.find(p=>p.id===u.product_id);
        if(!prod){results.push(`❌ ID ${u.product_id} not found`);continue;}
        const old_stock=prod.stock;
        await supabase.from('products').update({stock:u.new_stock}).eq('id',u.product_id);
        await supabase.from('change_log').insert({description:`Chat bulk update: ${u.product_name} ${old_stock}→${u.new_stock} (${u.reason})`,qty_change:u.new_stock-old_stock,user_email:userEmail});
        results.push(`✅ ${u.product_name}: ${old_stock} → ${u.new_stock}`);
      }
      await loadAll();
      return{success:true,message:`Bulk update complete:\n${results.join('\n')}`};
    }
    if(toolName==='clear_all_stock'){
      // This is handled via password prompt — should not reach here directly
      return{success:false,error:'clear_all_stock requires password confirmation'};
    }
    return{success:false,error:'Unknown tool'};
  }

  async function executeClearAllStock(reason){
    for(const p of prods){
      await supabase.from('products').update({stock:0}).eq('id',p.id);
    }
    await supabase.from('change_log').insert({description:`⚠️ CLEAR ALL STOCK: All ${prods.length} products set to 0 (${reason})`,qty_change:0,user_email:userEmail});
    await loadAll();
  }

  const SYSTEM_PASSWORD='BSL2025!';

  // ── CHAT SESSIONS ─────────────────────────────────────────────
  async function loadChatSessions(){
    setSessionsLoading(true);
    const{data}=await supabase.from('chat_sessions').select('id,title,created_at,updated_at').eq('user_email',userEmail).order('updated_at',{ascending:false}).limit(50);
    if(data)setChatSessions(data);
    setSessionsLoading(false);
  }

  async function saveSessionMessages(sessionId,msgs){
    const filtered=msgs.filter(m=>typeof m.content==='string');
    await supabase.from('chat_sessions').update({messages:filtered,updated_at:new Date().toISOString()}).eq('id',sessionId);
  }

  async function createNewSession(firstUserMsg){
    const title=firstUserMsg.slice(0,60)+(firstUserMsg.length>60?'...':'');
    const{data}=await supabase.from('chat_sessions').insert({user_email:userEmail,title,messages:[],created_at:new Date().toISOString(),updated_at:new Date().toISOString()}).select().single();
    if(data){
      setCurrentSessionId(data.id);
      setChatSessions(prev=>[{id:data.id,title,created_at:data.created_at,updated_at:data.updated_at},...prev]);
      return data.id;
    }
    return null;
  }

  async function loadSession(session){
    const{data}=await supabase.from('chat_sessions').select('messages').eq('id',session.id).single();
    if(data&&data.messages&&data.messages.length>0){
      setChatMsgs(data.messages);
    }else{
      setChatMsgs([{role:'assistant',content:`Hi! I'm Claude, your BSL inventory manager. I have full access to your inventory, customer sales, and notes. Ask me anything — stock levels, reorder suggestions, sales trends, or upload a packing list and I'll process it automatically.`}]);
    }
    setCurrentSessionId(session.id);
    setPage('chat');
  }

  async function deleteSession(id,e){
    e.stopPropagation();
    if(!window.confirm('Delete this conversation?'))return;
    await supabase.from('chat_sessions').delete().eq('id',id);
    setChatSessions(prev=>prev.filter(s=>s.id!==id));
    if(currentSessionId===id){
      setCurrentSessionId(null);
      setChatMsgs([{role:'assistant',content:`Hi! I'm Claude, your BSL inventory manager. I have full access to your inventory, customer sales, and notes. Ask me anything — stock levels, reorder suggestions, sales trends, or upload a packing list and I'll process it automatically.`}]);
    }
  }

  function startNewChat(){
    setCurrentSessionId(null);
    setChatMsgs([{role:'assistant',content:`Hi! I'm Claude, your BSL inventory manager. I have full access to your inventory, customer sales, and notes. Ask me anything — stock levels, reorder suggestions, sales trends, or upload a packing list and I'll process it automatically.`}]);
    setChatInput('');
    setChatFile(null);
  }

  async function confirmDangerousAction(){
    if(passwordInput!==SYSTEM_PASSWORD){
      setPasswordError('Incorrect password. Try again.');
      setPasswordInput('');
      return;
    }
    const{toolUseBlock,data,newMsgs,systemPrompt,reason}=pendingDangerousTool;
    setPendingDangerousTool(null);
    setPasswordInput('');
    setPasswordError('');
    setChatMsgs(prev=>[...prev,{role:'assistant',content:`🔧 Clearing all stock...`}]);
    try{
      await executeClearAllStock(reason);
      // Send second API call so Claude can confirm
      const toolResult={success:true,message:`All ${prods.length} products have been set to 0 stock.`};
      const msgs2=[
        ...newMsgs.filter(m=>m.role!=='system').map(m=>({role:m.role,content:m.content})),
        {role:'assistant',content:data.content},
        {role:'user',content:[{type:'tool_result',tool_use_id:toolUseBlock.id,content:JSON.stringify(toolResult)}]}
      ];
      const res2=await fetch('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({
        model:'claude-sonnet-4-20250514',
        max_tokens:500,
        system:systemPrompt,
        tools:inventoryTools,
        messages:msgs2,
      })});
      const data2=await res2.json();
      const finalReply=data2.content?.find(b=>b.type==='text')?.text||'All stock cleared to 0.';
      setChatMsgs(prev=>[...prev.slice(0,-1),{role:'assistant',content:finalReply}]);
    }catch(err){
      setChatMsgs(prev=>[...prev.slice(0,-1),{role:'assistant',content:'Error clearing stock. Please try again.'}]);
    }
  }

  async function sendChat(e){
    e.preventDefault();
    if((!chatInput.trim()&&!chatFile)||chatLoading)return;
    const userMsg=chatInput.trim();
    const fileToSend=chatFile;
    setChatInput('');
    setChatFile(null);
    if(chatFileRef.current)chatFileRef.current.value='';

    // Build user message content — text + optional file
    let userContent=userMsg;
    let filePreview=null;
    if(fileToSend){
      const base64=await new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result.split(',')[1]);r.onerror=rej;r.readAsDataURL(fileToSend);});
      const isPDF=fileToSend.type==='application/pdf';
      const isImage=fileToSend.type.startsWith('image/');
      filePreview={name:fileToSend.name,type:fileToSend.type,base64};
      userContent=[
        ...(isPDF?[{type:'document',source:{type:'base64',media_type:'application/pdf',data:base64}}]:[]),
        ...(isImage?[{type:'image',source:{type:'base64',media_type:fileToSend.type,data:base64}}]:[]),
        {type:'text',text:userMsg||`Please analyze this file: ${fileToSend.name}`}
      ];
    }

    const displayMsg={role:'user',content:userMsg||(fileToSend?`📎 ${fileToSend.name}`:'')};
    const newMsgs=[...chatMsgs,displayMsg];
    setChatMsgs(newMsgs);
    setChatLoading(true);
    // Auto-create session on first real user message
    let sessionId=currentSessionId;
    if(!sessionId&&!chatMsgs.some(m=>m.role==='user')){
      sessionId=await createNewSession(userMsg||(fileToSend?fileToSend.name:'New chat'));
    }
    try{
      const notesContext=agentNotes.length?`\nPermanent notes & instructions:\n${agentNotes.map(n=>`[${n.category}] ${n.title}: ${n.content}`).join('\n')}`:'';
      const inventoryContext=`Current inventory (${prods.length} products, all in SINGLES):\n${prods.map(p=>`- ID:${p.id} | ${p.name} | Root SKU: ${p.sku} | Stock: ${p.stock} | Velocity: ${p.velocity}/mo | Cost: $${p.cost} | Price: $${p.price} | Reorder at: ${p.reorder} | Status: ${gs(p)} | Supplier: ${p.supplier||'—'}`).join('\n')}`;
      const recentLog=`\nRecent changes:\n${logEntries.slice(0,10).map(l=>`- ${new Date(l.created_at).toLocaleDateString()}: ${l.description}`).join('\n')}`;

      const systemPrompt=`You are Claude, the inventory manager for BSL (Blooming Sweet Life Corp). You have tools to directly update inventory. RULES: 1) All stock in SINGLES. 2) When user asks to update stock for ONE product, USE the update_stock tool. 3) When user pastes or provides a LIST of products with quantities, USE the bulk_update_stock tool with ALL products in a single call — never loop one by one. 4) For boxes×units, multiply to get singles (e.g. 60 boxes × 12 units = 720 singles). 5) Always confirm what you did after using a tool. 6) Be concise. Respond in ${lang==='es'?'Spanish':'English'}. 7) IMPORTANT: When the user asks to clear, reset, or zero all stock/inventory — call the clear_all_stock tool IMMEDIATELY with a reason. Do NOT ask for a password in chat — the app handles password confirmation automatically. Just call the tool.${notesContext}\n\n${inventoryContext}${recentLog}`;

      // Build messages array — replace last user message content with rich content if file attached
      const apiMsgs=newMsgs.filter(m=>m.role!=='system').map((m,i)=>{
        if(i===newMsgs.length-1&&filePreview)return{role:'user',content:userContent};
        return{role:m.role,content:m.content};
      });

      // First API call with tools
      const res=await fetch('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({
        model:'claude-sonnet-4-20250514',
        max_tokens:1000,
        system:systemPrompt,
        tools:inventoryTools,
        messages:apiMsgs,
      })});
      const data=await res.json();

      // Check if Claude wants to use a tool
      const toolUseBlock=data.content?.find(b=>b.type==='tool_use');
      const textBlock=data.content?.find(b=>b.type==='text');

      if(toolUseBlock){
        // Intercept dangerous tools — require password first
        if(toolUseBlock.name==='clear_all_stock'){
          setChatMsgs(prev=>[...prev,{role:'assistant',content:`⚠️ This will set ALL ${prods.length} products to 0 stock. Please enter the system password to confirm.`}]);
          setPendingDangerousTool({toolUseBlock,data,newMsgs,systemPrompt,reason:toolUseBlock.input.reason});
          setChatLoading(false);
          return;
        }

        // Show thinking message
        setChatMsgs(prev=>[...prev,{role:'assistant',content:`🔧 Updating inventory...`}]);
        
        // Execute the tool
        const toolResult=await executeToolCall(toolUseBlock.name, toolUseBlock.input);
        
        // Second API call with tool result so Claude can respond
        const msgs2=[
          ...newMsgs.filter(m=>m.role!=='system').map(m=>({role:m.role,content:m.content})),
          {role:'assistant',content:data.content},
          {role:'user',content:[{type:'tool_result',tool_use_id:toolUseBlock.id,content:JSON.stringify(toolResult)}]}
        ];
        const res2=await fetch('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({
          model:'claude-sonnet-4-20250514',
          max_tokens:500,
          system:systemPrompt,
          tools:inventoryTools,
          messages:msgs2,
        })});
        const data2=await res2.json();
        const finalReply=data2.content?.find(b=>b.type==='text')?.text||'Done!';
        // Replace thinking message with final reply + auto-save
        setChatMsgs(prev=>{const updated=[...prev.slice(0,-1),{role:'assistant',content:finalReply}];if(sessionId)saveSessionMessages(sessionId,updated);return updated;});
      } else {
        const reply=textBlock?.text||'Sorry, I could not process that.';
        setChatMsgs(prev=>{const updated=[...prev,{role:'assistant',content:reply}];if(sessionId)saveSessionMessages(sessionId,updated);return updated;});
      }
    }catch(err){
      console.error(err);
      setChatMsgs(prev=>[...prev,{role:'assistant',content:'Error connecting to Claude API.'}]);
    }
    setChatLoading(false);
  }

  // ── COST CALCULATOR ──────────────────────────────────────────
  const OZ_PER_KG=35.274;
  function calcCost(p,gs){
    if(p.product_type!=='packaged')return{total:parseFloat(p.cost)||0,breakdown:null};
    const rmWaste=gs.raw_material_waste_pct||0.005;
    const pkgWaste=gs.packaging_waste_pct||0.005;
    const filling=gs.filling_cost||1.15;
    // Convert price per kg to price per oz
    const pricePerKg=parseFloat(p.raw_material_cost_per_kg)||0;
    const pricePerOz=pricePerKg/OZ_PER_KG;
    const weightOz=parseFloat(p.weight_oz)||0;
    const rawMaterial=pricePerOz*weightOz;
    const rawWithWaste=rawMaterial*(1+rmWaste);
    const pkgCost=parseFloat(p.packaging_cost)||0;
    const pkgWithWaste=pkgCost*(1+pkgWaste);
    const box=parseFloat(p.box_cost)||0;
    const jumboBox=parseFloat(p.jumbo_box_cost)||0;
    const total=rawWithWaste+pkgWithWaste+filling+box+jumboBox;
    return{
      total:Math.round(total*10000)/10000,
      pricePerKg,pricePerOz,rawMaterial,rawWithWaste,
      pkgCost,pkgWithWaste,filling,box,jumboBox,
      rmWaste,pkgWaste,weightOz
    };
  }

  // ── COMPUTED ─────────────────────────────────────────────────
  const totalVal=prods.reduce((a,p)=>a+(p.stock*(calcCost(p,globalSettings).total||0)),0);
  const lowN=prods.filter(p=>gs(p)!=='ok').length;
  const critN=prods.filter(p=>gs(p)==='crit').length;
  const alerts=prods.filter(p=>gs(p)!=='ok');
  const reorderItems=prods.filter(p=>p.stock<=p.reorder&&p.reorder>0);

  // ── FILTERS ──────────────────────────────────────────────────
  const[search,setSearch]=useState('');
  const[filterCat,setFilterCat]=useState('');
  const[filterStatus,setFilterStatus]=useState('');
  const[sortBy,setSortBy]=useState('name');
  const[sortDir,setSortDir]=useState('asc');

  const categories=[...new Set(prods.map(p=>p.category).filter(Boolean))].sort();

  const filteredProds=prods.filter(p=>{
    const q=search.toLowerCase();
    const matchQ=!q||p.name.toLowerCase().includes(q)||(p.sku||'').toLowerCase().includes(q)||(p.amz||'').toLowerCase().includes(q)||(p.wmt||'').toLowerCase().includes(q);
    const matchCat=!filterCat||p.category===filterCat;
    const matchStatus=!filterStatus||gs(p)===filterStatus;
    return matchQ&&matchCat&&matchStatus;
  }).sort((a,b)=>{
    let av,bv;
    if(sortBy==='name'){av=a.name||'';bv=b.name||'';}
    else if(sortBy==='stock'){av=a.stock||0;bv=b.stock||0;}
    else if(sortBy==='price'){av=a.price||0;bv=b.price||0;}
    else if(sortBy==='cost'){av=a.cost||0;bv=b.cost||0;}
    else if(sortBy==='days'){av=gd(a)??9999;bv=gd(b)??9999;}
    else if(sortBy==='status'){av={ok:0,low:1,crit:2}[gs(a)]||0;bv={ok:0,low:1,crit:2}[gs(b)]||0;}
    else{av=a.name||'';bv=b.name||'';}
    if(typeof av==='string') return sortDir==='asc'?av.localeCompare(bv):bv.localeCompare(av);
    return sortDir==='asc'?av-bv:bv-av;
  });

  function toggleSort(col){if(sortBy===col)setSortDir(d=>d==='asc'?'desc':'asc');else{setSortBy(col);setSortDir('asc');}}
  const sortIcon=(col)=>sortBy===col?(sortDir==='asc'?'↑':'↓'):'';

  // ── FILTER BAR ───────────────────────────────────────────────
  const FilterBar=()=>(
    <div style={{display:'flex',gap:8,marginBottom:'1rem',flexWrap:'wrap',alignItems:'center'}}>
      <input style={{...S.inp,flex:1,minWidth:160}} placeholder="Search by name, SKU..." value={search} onChange={e=>setSearch(e.target.value)}/>
      <select style={{...S.inp,width:'auto'}} value={filterCat} onChange={e=>setFilterCat(e.target.value)}>
        <option value="">All categories</option>
        {categories.map(c=><option key={c} value={c}>{c}</option>)}
      </select>
      <select style={{...S.inp,width:'auto'}} value={filterStatus} onChange={e=>setFilterStatus(e.target.value)}>
        <option value="">All statuses</option>
        <option value="ok">OK</option>
        <option value="low">Low stock</option>
        <option value="crit">Critical</option>
      </select>
      {(search||filterCat||filterStatus)&&<button style={{...S.btn,color:'#dc3545',borderColor:'#f5c6cb',fontSize:11}} onClick={()=>{setSearch('');setFilterCat('');setFilterStatus('');}}>✕ Clear</button>}
      <span style={{fontSize:11,color:'#aaa',marginLeft:'auto'}}>{filteredProds.length} of {prods.length} products</span>
    </div>
  );

  // ── PRODUCT TABLE ────────────────────────────────────────────
  const[colWidths,setColWidths]=useState({status:70,name:220,sku:120,stock:90,days:70,cost:90,price:90,actions:80});
  const resizing=useRef(null);
  function onResizeStart(col,e){
    e.preventDefault();
    const startX=e.clientX;
    const startW=colWidths[col];
    resizing.current={col,startX,startW};
    function onMove(ev){
      const delta=ev.clientX-resizing.current.startX;
      const newW=Math.max(50,resizing.current.startW+delta);
      setColWidths(prev=>({...prev,[resizing.current.col]:newW}));
    }
    function onUp(){
      resizing.current=null;
      window.removeEventListener('mousemove',onMove);
      window.removeEventListener('mouseup',onUp);
    }
    window.addEventListener('mousemove',onMove);
    window.addEventListener('mouseup',onUp);
  }
  const ResizeTh=({col,children,onClick,style})=>(
    <th style={{...S.th,width:colWidths[col],minWidth:colWidths[col],maxWidth:colWidths[col],position:'relative',userSelect:'none',borderRight:'2px solid #ddd',...style}} onClick={onClick}>
      {children}
      <span
        onMouseDown={e=>onResizeStart(col,e)}
        onClick={e=>e.stopPropagation()}
        style={{position:'absolute',right:-4,top:0,bottom:0,width:8,cursor:'col-resize',zIndex:10,display:'flex',alignItems:'center',justifyContent:'center'}}
      >
        <span style={{width:3,height:'60%',background:'#bbb',borderRadius:2,pointerEvents:'none'}}/>
      </span>
    </th>
  );
  const ProdTable=({list})=>{
    if(!list.length)return<div style={{textAlign:'center',padding:'2rem',color:'#aaa',fontSize:13}}>{tab==='alerts'?t.noAlerts:t.noProducts}</div>;
    return(
      <div style={{overflowX:'auto',border:'1px solid #eee',borderRadius:12}}>
        <table style={{tableLayout:'fixed',width:Object.values(colWidths).reduce((a,b)=>a+b,0),borderCollapse:'collapse',fontSize:12}}>
          <colgroup>
            {Object.entries(colWidths).map(([col,w])=><col key={col} style={{width:w}}/>)}
          </colgroup>
          <thead><tr>
            <ResizeTh col="status" onClick={()=>toggleSort('status')}>{t.status} <span style={{color:'#4a90e2'}}>{sortIcon('status')}</span></ResizeTh>
            <ResizeTh col="name" onClick={()=>toggleSort('name')}>{t.product} <span style={{color:'#4a90e2'}}>{sortIcon('name')}</span></ResizeTh>
            <ResizeTh col="sku">{t.rootSku}</ResizeTh>
            <ResizeTh col="stock" onClick={()=>toggleSort('stock')}>{t.stock} <span style={{color:'#4a90e2'}}>{sortIcon('stock')}</span></ResizeTh>
            <ResizeTh col="days" onClick={()=>toggleSort('days')}>{t.days} <span style={{color:'#4a90e2'}}>{sortIcon('days')}</span></ResizeTh>
            <ResizeTh col="cost" onClick={()=>toggleSort('cost')}>{t.cost} <span style={{color:'#4a90e2'}}>{sortIcon('cost')}</span></ResizeTh>
            <ResizeTh col="price" onClick={()=>toggleSort('price')}>{t.price} <span style={{color:'#4a90e2'}}>{sortIcon('price')}</span></ResizeTh>
            <ResizeTh col="actions">{t.actions}</ResizeTh>
          </tr></thead>
          <tbody>
            {list.map(p=>{const st=gs(p),dl=gd(p);return(
              <tr key={p.id}>
                <td style={S.td}><Badge status={st} t={t}/></td>
                <td style={{...S.td,fontWeight:500,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={p.name}>{p.name}</td>
                <td style={S.td}><code style={{fontSize:10,background:'#f5f5f5',padding:'1px 5px',borderRadius:4}}>{p.sku||'—'}</code></td>
                <td style={{...S.td,color:p.stock<0?'#dc3545':'#111',cursor:'pointer',textDecoration:'underline',textDecorationStyle:'dotted',textDecorationColor:'#aaa'}} onClick={()=>setLocationModal(p)} title='Click to see stock by location'>{p.stock}</td>
                <td style={{...S.td,color:st==='crit'?'#dc3545':st==='low'?'#856404':undefined}}>{dl!=null?`${dl}d`:'—'}</td>
                <td style={{...S.td,cursor:p.product_type==='packaged'?'pointer':'default',color:p.product_type==='packaged'?'#4a90e2':undefined,textDecoration:p.product_type==='packaged'?'underline':undefined}} onClick={()=>{if(p.product_type==='packaged')setCostModal(p)}}>{fm(p.product_type==='packaged'?calcCost(p,globalSettings).total:p.cost)}{p.product_type==='packaged'&&<span style={{fontSize:9,marginLeft:3}}>📊</span>}</td>
                <td style={S.td}>{fm(p.price)}</td>
                <td style={{...S.td,whiteSpace:'nowrap'}}>
                  <button style={{...S.btn,padding:'3px 7px'}} onClick={()=>{setMdata({form:{...p}});setModal('product')}}>✏️</button>
                  <button style={{...S.btn,padding:'3px 7px',marginLeft:3,color:'#dc3545',borderColor:'#f5c6cb'}} onClick={()=>deleteProduct(p.id)}>🗑</button>
                </td>
              </tr>
            );})}
          </tbody>
        </table>
      </div>
    );
  };

  // ── RENDER ───────────────────────────────────────────────────
  return(
    <div style={S.app}>
      <nav style={S.nav}>
        <div style={{display:'flex',alignItems:'center',gap:'1rem'}}>
          <span style={{fontWeight:700,fontSize:16}}>{t.title}</span>
          {[['dashboard',t.dashboard],['chat',t.chat],['calculator',lang==='es'?'💰 Calculadora':'💰 Calculator'],['customers',t.customers],['reports',t.reports],['notes',t.notes]].map(([k,l])=>(
            <button key={k} onClick={()=>setPage(k)} style={{...S.btn,background:page===k?'rgba(255,255,255,.15)':'transparent',color:'#fff',border:'none',fontSize:12}}>{l}</button>
          ))}
        </div>
        <div style={{display:'flex',gap:8,alignItems:'center'}}>
          <button style={{...S.btn,color:'#fff',border:'1px solid rgba(255,255,255,.3)',fontSize:12}} onClick={()=>setLang(l=>l==='en'?'es':'en')}>{lang==='en'?'🇲🇽 Español':'🇺🇸 English'}</button>
          <button style={{...S.btn,color:'#aaa',border:'1px solid rgba(255,255,255,.15)',fontSize:11}} onClick={()=>supabase.auth.signOut()}>Sign out</button>
        </div>
      </nav>

      <main style={S.main}>
        {loading&&<div style={{textAlign:'center',padding:'3rem',color:'#aaa'}}>Loading...</div>}

        {/* DASHBOARD */}
        {!loading&&page==='dashboard'&&<>
          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:'1.25rem',flexWrap:'wrap',gap:8}}>
            <div style={{fontSize:20,fontWeight:600}}>{t.dashboard}</div>
            <div style={{display:'flex',gap:7,flexWrap:'wrap'}}>
              <button style={S.btn} onClick={exportCSV}>{t.exportCSV}</button>
              <button style={{...S.btn,color:'#4a90e2',borderColor:'#4a90e2'}} onClick={()=>{setMdata({step:1,mode:'add',parsed:[],errors:[]});setModal('import')}}>{t.bulkImport}</button>
              <button style={{...S.btn,color:'#28a745',borderColor:'#28a745'}} onClick={()=>{setMdata({step:1,platform:'amazon'});setModal('packing')}}>{t.uploadPacking}</button>
              <button style={{...S.btn,color:'#e67e22',borderColor:'#e67e22',fontWeight:600}} onClick={()=>{setMdata({step:1,platform:'amazon',headers:[],rows:[],skuCol:-1,qtyCol:-1,oidCol:-1,hash:'',preview:[],warnings:[]});setModal('orders')}}>{t.uploadOrders}</button>
              <button style={S.btnP} onClick={()=>{setMdata({form:ep()});setModal('product')}}>{t.addProduct}</button>
            </div>
          </div>

          <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:9,marginBottom:'1.5rem'}}>
            {[{l:t.totalProducts,v:prods.length},{l:t.inventoryValue,v:'$'+Math.round(totalVal).toLocaleString()},{l:t.lowStock,v:lowN,a:lowN>0},{l:t.critical,v:critN,a:critN>0}].map(m=>(
              <div key={m.l} style={{background:m.a?'#FFF3CD':'#fff',borderRadius:12,padding:'14px 16px',boxShadow:'0 1px 3px rgba(0,0,0,.08)'}}>
                <div style={{fontSize:11,color:m.a?'#856404':'#888',marginBottom:4}}>{m.l}</div>
                <div style={{fontSize:22,fontWeight:700,color:m.a?'#856404':'#111'}}>{m.v}</div>
              </div>
            ))}
          </div>

          <div style={{display:'flex',gap:6,marginBottom:'1rem',flexWrap:'wrap'}}>
            {[['all',t.allProducts],['alerts',`${t.alerts} (${alerts.length})`],['channels',t.channels],['log',`📋 ${t.log} (${logEntries.length})`]].map(([k,l])=>(
              <button key={k} onClick={()=>setTab(k)} style={{...S.btn,background:tab===k?'#111':'transparent',color:tab===k?'#fff':'#555',border:tab===k?'none':'1px solid #ddd'}}>{l}</button>
            ))}
          </div>

          <div style={S.card}>
            {tab==='all'&&<><FilterBar/><ProdTable list={filteredProds}/></>}
            {tab==='alerts'&&<ProdTable list={alerts}/>}
            {tab==='channels'&&(
              <div style={{overflowX:'auto'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:12}}>
                  <thead><tr>{[t.product,'Root SKU','Amazon (pack)','Walmart (pack)','Target (pack)','Temu (pack)',t.other].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                  <tbody>
                    {!prods.length&&<tr><td colSpan={7} style={{textAlign:'center',padding:'2rem',color:'#aaa'}}>{t.noProducts}</td></tr>}
                    {prods.map(p=><tr key={p.id}>
                      <td style={{...S.td,fontWeight:500}}>{p.name}</td>
                      <td style={S.td}><code style={{fontSize:10,background:'#f5f5f5',padding:'1px 5px',borderRadius:4}}>{p.sku||'—'}</code></td>
                      {[['amz','amz_pack_size','amazon'],['wmt','wmt_pack_size','walmart'],['tgt','tgt_pack_size','target'],['temu','temu_pack_size','temu'],['other_sku','other_pack_size','other']].map(([k,pk,platK])=>(
                        <td key={k} style={S.td}>{p[k]?<span style={{background:PLAT[platK]?.c+'22',color:PLAT[platK]?.c,padding:'1px 7px',borderRadius:99,fontSize:11,fontWeight:500}}>{p[k]} ({p[pk]||1}pk)</span>:'—'}</td>
                      ))}
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
                    {l.platform&&<span style={{fontSize:10,background:(PLAT[l.platform]?.c||'#888')+'22',color:PLAT[l.platform]?.c||'#888',padding:'1px 6px',borderRadius:99,fontWeight:600,flexShrink:0}}>{PLAT[l.platform]?.l||l.platform}</span>}
                    <span style={{flex:1}}>{l.description}</span>
                    {l.user_email&&<span style={{fontSize:10,color:'#4a90e2',background:'#e8f0fe',padding:'1px 6px',borderRadius:99,flexShrink:0}}>{l.user_email.split('@')[0]}</span>}
                    {l.qty_change!=null&&<span style={{fontWeight:600,color:l.qty_change<0?'#dc3545':'#28a745',flexShrink:0}}>{l.qty_change>0?'+':''}{l.qty_change}</span>}
                  </div>
                ))}
              </div>
            )}
          </div>
        </>}

        {/* CHAT */}
        {!loading&&page==='chat'&&(
          <div style={{display:'flex',gap:12,height:'calc(100vh - 120px)'}}>

            {/* ── Sessions sidebar */}
            <div style={{width:240,flexShrink:0,background:'#fff',borderRadius:12,boxShadow:'0 1px 3px rgba(0,0,0,.08)',display:'flex',flexDirection:'column',overflow:'hidden'}}>
              <div style={{padding:'12px 14px',borderBottom:'1px solid #eee'}}>
                <button style={{...S.btnP,width:'100%',justifyContent:'center',gap:6}} onClick={startNewChat}>✏️ New Chat</button>
              </div>
              <div style={{flex:1,overflowY:'auto',padding:'8px 6px'}}>
                {sessionsLoading&&<div style={{textAlign:'center',padding:'1rem',fontSize:12,color:'#aaa'}}>Loading...</div>}
                {!sessionsLoading&&!chatSessions.length&&<div style={{textAlign:'center',padding:'1.5rem 1rem',fontSize:12,color:'#aaa'}}>No saved conversations yet</div>}
                {chatSessions.map(s=>(
                  <div key={s.id} onClick={()=>loadSession(s)} style={{display:'flex',alignItems:'flex-start',gap:6,padding:'8px 10px',borderRadius:8,cursor:'pointer',background:currentSessionId===s.id?'#f0f0f0':'transparent',marginBottom:2,transition:'background .15s'}}>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:12,fontWeight:currentSessionId===s.id?600:400,color:'#111',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{s.title||'Untitled'}</div>
                      <div style={{fontSize:10,color:'#aaa',marginTop:2}}>{new Date(s.updated_at).toLocaleDateString()} {new Date(s.updated_at).toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'})}</div>
                    </div>
                    <button style={{...S.btn,padding:'2px 5px',fontSize:10,color:'#dc3545',borderColor:'#f5c6cb',flexShrink:0,opacity:0.6}} onClick={e=>deleteSession(s.id,e)} title="Delete">🗑</button>
                  </div>
                ))}
              </div>
            </div>

            {/* ── Chat window */}
            <div style={{...S.card,flex:1,display:'flex',flexDirection:'column',minWidth:0}}>
              <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem',paddingBottom:'0.75rem',borderBottom:'1px solid #eee',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
                <span>{currentSessionId?chatSessions.find(s=>s.id===currentSessionId)?.title||t.chat:t.chat}</span>
                {currentSessionId&&<span style={{fontSize:10,color:'#aaa',fontWeight:400}}>Auto-saved</span>}
              </div>
              <div style={{flex:1,overflowY:'auto',display:'flex',flexDirection:'column',gap:12,paddingBottom:'1rem'}}>
                {chatMsgs.map((m,i)=>(
                  <div key={i} style={{display:'flex',justifyContent:m.role==='user'?'flex-end':'flex-start'}}>
                    <div style={{maxWidth:'78%',padding:'10px 14px',borderRadius:12,background:m.role==='user'?'#111':'#f0f0f0',color:m.role==='user'?'#fff':'#111',fontSize:13,lineHeight:1.5,whiteSpace:'pre-wrap'}}>{m.content}</div>
                  </div>
                ))}
                {chatLoading&&<div style={{display:'flex',justifyContent:'flex-start'}}><div style={{padding:'10px 14px',borderRadius:12,background:'#f0f0f0',fontSize:13,color:'#888'}}>{t.thinking}</div></div>}
                <div ref={chatEndRef}/>
              </div>
              {pendingDangerousTool&&(
                <div style={{background:'#FFF3CD',border:'1px solid #ffc107',borderRadius:10,padding:'12px 14px',marginBottom:'0.75rem'}}>
                  <div style={{fontSize:13,fontWeight:600,color:'#856404',marginBottom:8}}>🔐 Password required to clear all stock</div>
                  <div style={{display:'flex',gap:8}}>
                    <input style={{...S.inp,flex:1,borderColor:passwordError?'#dc3545':'#ffc107'}} type="password" placeholder="Enter system password..." value={passwordInput} onChange={e=>{setPasswordInput(e.target.value);setPasswordError('');}} onKeyDown={async e=>{if(e.key==='Enter')await confirmDangerousAction();}} autoFocus/>
                    <button style={{...S.btn,background:'#dc3545',color:'#fff',border:'none'}} onClick={async()=>await confirmDangerousAction()}>Confirm</button>
                    <button style={S.btn} onClick={()=>{setPendingDangerousTool(null);setPasswordInput('');setPasswordError('');setChatMsgs(prev=>[...prev,{role:'assistant',content:'Action cancelled.'}]);}}>Cancel</button>
                  </div>
                  {passwordError&&<div style={{fontSize:11,color:'#dc3545',marginTop:5}}>{passwordError}</div>}
                </div>
              )}
              <div style={{borderTop:'1px solid #eee',paddingTop:'0.75rem'}}>
                {chatFile&&(
                  <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:8,background:'#f0f7ff',borderRadius:8,padding:'6px 10px'}}>
                    <span style={{fontSize:16}}>{chatFile.type.startsWith('image/')?'🖼️':chatFile.type==='application/pdf'?'📄':'📎'}</span>
                    <span style={{fontSize:12,color:'#1565c0',flex:1,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{chatFile.name}</span>
                    <button style={{...S.btn,padding:'2px 7px',fontSize:11,color:'#dc3545',borderColor:'#f5c6cb'}} onClick={()=>{setChatFile(null);if(chatFileRef.current)chatFileRef.current.value='';}}>✕</button>
                  </div>
                )}
                <form onSubmit={sendChat} style={{display:'flex',gap:8}}>
                  <input ref={chatFileRef} type="file" accept="image/*,.pdf,.csv,.xlsx,.xls" style={{display:'none'}} onChange={e=>{const f=e.target.files[0];if(f)setChatFile(f);}}/>
                  <button type="button" style={{...S.btn,padding:'8px 10px',fontSize:16,flexShrink:0,color:chatFile?'#1565c0':'#888',borderColor:chatFile?'#1565c0':'#ddd'}} onClick={()=>chatFileRef.current.click()} disabled={chatLoading} title="Attach file">📎</button>
                  <textarea style={{...S.inp,flex:1,resize:'none',minHeight:40,maxHeight:200,lineHeight:'1.5',padding:'8px 10px',fontFamily:'inherit',fontSize:13,overflowY:'auto'}} rows={1} value={chatInput} onChange={e=>{setChatInput(e.target.value);e.target.style.height='auto';e.target.style.height=Math.min(e.target.scrollHeight,200)+'px';}} onKeyDown={e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();sendChat(e);}}} placeholder={chatFile?'Add a message (optional)...':t.chatPlaceholder} disabled={chatLoading}/>
                  <button type="submit" style={{...S.btnP,padding:'8px 16px',alignSelf:'flex-end'}} disabled={chatLoading||(!chatInput.trim()&&!chatFile)}>→</button>
                </form>
              </div>
            </div>
          </div>
        )}

        {/* CUSTOMERS */}
        {!loading&&page==='customers'&&(
          <div style={S.card}>
            <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:'1rem'}}>
              <div style={{fontSize:16,fontWeight:600}}>{t.customers}</div>
              <button style={S.btnP} onClick={()=>{setMdata({form:{customer_name:'',customer_email:'',customer_phone:'',sale_date:new Date().toISOString().slice(0,10),platform:'direct',items:[]}});setModal('sale')}}>{t.addSale}</button>
            </div>
            {!customers.length&&<div style={{textAlign:'center',padding:'2rem',color:'#aaa',fontSize:13}}>No customer sales yet</div>}
            <div style={{overflowX:'auto',border:'1px solid #eee',borderRadius:12}}>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:12}}>
                <thead><tr>{['Customer','Date','Platform','Singles','Items'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                <tbody>
                  {customers.map(c=>(
                    <tr key={c.id}>
                      <td style={{...S.td,fontWeight:500}}>{c.customer_name}</td>
                      <td style={S.td}>{new Date(c.sale_date||c.created_at).toLocaleDateString()}</td>
                      <td style={S.td}><span style={{background:'#e8f5e9',color:'#2e7d32',padding:'1px 7px',borderRadius:99,fontSize:11,fontWeight:500}}>{c.platform||'direct'}</span></td>
                      <td style={S.td}>{c.total_singles}</td>
                      <td style={S.td}>{(c.customer_sale_items||[]).length} items</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* REPORTS */}
        {!loading&&page==='reports'&&(
          <div style={{display:'flex',flexDirection:'column',gap:16}}>
            {/* Reorder suggestions */}
            <div style={S.card}>
              <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem'}}>📦 {t.reorderSuggestions}</div>
              {!reorderItems.length?<div style={{color:'#28a745',fontSize:13}}>✅ All products are above reorder point</div>:(
                <div style={{overflowX:'auto',border:'1px solid #eee',borderRadius:12}}>
                  <table style={{width:'100%',borderCollapse:'collapse',fontSize:12}}>
                    <thead><tr>{['Product','Root SKU','Current Stock','Reorder At','Days Left','Supplier','Suggested Order'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                    <tbody>
                      {reorderItems.map(p=>{
                        const dl=gd(p);
                        const suggested=Math.max(p.reorder*3-p.stock,0);
                        return(
                          <tr key={p.id} style={{background:'#FFF5F5'}}>
                            <td style={{...S.td,fontWeight:500}}>{p.name}</td>
                            <td style={S.td}><code style={{fontSize:10,background:'#f5f5f5',padding:'1px 5px',borderRadius:4}}>{p.sku||'—'}</code></td>
                            <td style={{...S.td,color:'#dc3545',fontWeight:600}}>{p.stock}</td>
                            <td style={S.td}>{p.reorder}</td>
                            <td style={{...S.td,color:'#dc3545'}}>{dl!=null?`${dl}d`:'—'}</td>
                            <td style={S.td}>{p.supplier||'—'}</td>
                            <td style={{...S.td,color:'#e67e22',fontWeight:600}}>{suggested} singles</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
            {/* Velocity report */}
            <div style={S.card}>
              <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:'1rem'}}>
                <div style={{fontSize:15,fontWeight:600}}>📈 {t.velocityReport}</div>
                <button style={{...S.btn,fontSize:11,color:'#4a90e2',borderColor:'#4a90e2'}} onClick={async()=>{await recalculateVelocity();await loadAll();alert('Velocity recalculated from last 30 days of orders!');}}>🔄 Recalculate</button>
              </div>
              <div style={{overflowX:'auto',border:'1px solid #eee',borderRadius:12}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:12}}>
                  <thead><tr>{['Product','Monthly Sales','Daily Rate','Days of Supply','Value at Risk','Status'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                  <tbody>
                    {[...prods].sort((a,b)=>(b.velocity||0)-(a.velocity||0)).map(p=>{
                      const daily=(p.velocity||0)/30;
                      const dos=gd(p);
                      const valueAtRisk=p.stock*(p.price||0);
                      const st=gs(p);
                      return(
                        <tr key={p.id}>
                          <td style={{...S.td,fontWeight:500}}>{p.name}</td>
                          <td style={S.td}>{p.velocity||0} singles/mo</td>
                          <td style={S.td}>{daily.toFixed(1)}/day</td>
                          <td style={{...S.td,color:st==='crit'?'#dc3545':st==='low'?'#856404':'#28a745',fontWeight:500}}>{dos!=null?`${dos} days`:'∞'}</td>
                          <td style={S.td}>{fm(valueAtRisk)}</td>
                          <td style={S.td}><Badge status={st} t={t}/></td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* NOTES */}
        {!loading&&page==='notes'&&(
          <div style={S.card}>
            <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:'1rem'}}>
              <div>
                <div style={{fontSize:16,fontWeight:600}}>🧠 {t.notes}</div>
                <div style={{fontSize:12,color:'#888',marginTop:2}}>These notes are loaded into every Claude conversation — your permanent memory</div>
              </div>
              <button style={S.btnP} onClick={()=>{setMdata({form:{title:'',content:'',category:'general'}});setModal('note')}}>{t.addNote}</button>
            </div>
            {!agentNotes.length&&<div style={{textAlign:'center',padding:'2rem',color:'#aaa',fontSize:13}}>No notes yet — add instructions, customer info, supplier details, or anything Claude should always remember</div>}
            <div style={{display:'flex',flexDirection:'column',gap:10}}>
              {agentNotes.map(n=>(
                <div key={n.id} style={{border:'1px solid #eee',borderRadius:10,padding:'12px 14px',display:'flex',gap:12,alignItems:'flex-start'}}>
                  <div style={{flex:1}}>
                    <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:4}}>
                      <span style={{fontSize:11,background:'#f0f0f0',color:'#555',padding:'1px 7px',borderRadius:99,fontWeight:500}}>{n.category}</span>
                      <span style={{fontSize:13,fontWeight:600}}>{n.title}</span>
                    </div>
                    <div style={{fontSize:12,color:'#555',lineHeight:1.5}}>{n.content}</div>
                  </div>
                  <button style={{...S.btn,padding:'3px 7px',color:'#dc3545',borderColor:'#f5c6cb',flexShrink:0}} onClick={async()=>{await supabase.from('agent_notes').delete().eq('id',n.id);await loadAll();}}>🗑</button>
                </div>
              ))}
            </div>
          </div>
        )}
        {/* CALCULATOR */}
        {!loading&&page==='calculator'&&(
          <CostCalculatorPage prods={prods} globalSettings={globalSettings} calcCost={calcCost} saveGlobalSettings={saveGlobalSettings} costCalcOverrides={costCalcOverrides} setCostCalcOverrides={setCostCalcOverrides} S={S} lang={lang}/>
        )}
      </main>

      {/* MODALS */}
      {modal==='product'&&<ProductModal t={t} S={S} mdata={mdata} setMdata={setMdata} onSave={saveProduct} onClose={()=>setModal(null)} lang={lang}/>}
      {modal==='import'&&<ImportModal t={t} S={S} mdata={mdata} setMdata={setMdata} onFile={handleImpFile} onConfirm={confirmImport} onDownload={downloadTemplate} onClose={()=>setModal(null)}/>}
      {modal==='orders'&&<OrdersModal t={t} S={S} mdata={mdata} setMdata={setMdata} onFile={handleOrdFile} onPreview={buildPreview} onConfirm={confirmOrders} onApply={applyOrders} hashes={hashes} onClose={()=>setModal(null)}/>}
      {modal==='packing'&&<PackingModal t={t} S={S} mdata={mdata} setMdata={setMdata} onFile={handlePackingFile} onApply={applyPackingList} onClose={()=>setModal(null)} prods={prods}/>}
      {locationModal&&<LocationModal prod={locationModal} locations={locations} S={S} lang={lang} onClose={()=>setLocationModal(null)} onSave={async(rows)=>{
  for(const row of rows){
    if(row.id){
      await supabase.from('inventory_locations').update({qty:row.qty,notes:row.notes,updated_at:new Date().toISOString()}).eq('id',row.id);
    }else if((parseFloat(row.qty)||0)>0){
      await supabase.from('inventory_locations').insert({product_id:locationModal.id,location:row.location,qty:row.qty,notes:row.notes||''});
    }
  }
  // Sum all locations and update product stock
  const totalStock=rows.reduce((a,r)=>a+(parseFloat(r.qty)||0),0);
  const oldStock=locationModal.stock;
  await supabase.from('products').update({stock:totalStock}).eq('id',locationModal.id);
  await supabase.from('change_log').insert({description:`Location update: ${locationModal.name} stock set to ${totalStock} (Warehouse+EVI+Tripolac)`,qty_change:totalStock-oldStock,user_email:userEmail});
  await loadAll();
  setLocationModal(null);
}} onDelete={async(id)=>{await supabase.from('inventory_locations').delete().eq('id',id);await loadAll();}}/> }
      {costModal&&<CostBreakdownModal prod={costModal} globalSettings={globalSettings} S={S} lang={lang} onClose={()=>setCostModal(null)} onSaveSettings={saveGlobalSettings}/>}
      {modal==='note'&&<NoteModal t={t} S={S} mdata={mdata} setMdata={setMdata} onSave={async(form)=>{if(form.id)await supabase.from('agent_notes').update(form).eq('id',form.id);else await supabase.from('agent_notes').insert(form);await loadAll();setModal(null);}} onClose={()=>setModal(null)}/>}
    </div>
  );
}

// ── COST CALCULATOR PAGE ──────────────────────────────────────
const PLATFORMS_CALC=[
  {key:'amz',label:'Amazon',feeKey:'fee_amz',color:'#FF9900',priceKey:'amz_sell',packKey:'amz_pack_size'},
  {key:'wmt',label:'Walmart',feeKey:'fee_wmt',color:'#0071CE',priceKey:'wmt_sell',packKey:'wmt_pack_size'},
  {key:'tgt',label:'Target',feeKey:'fee_tgt',color:'#CC0000',priceKey:'tgt_sell',packKey:'tgt_pack_size'},
  {key:'temu',label:'Temu',feeKey:'fee_temu',color:'#FF6533',priceKey:'temu_sell',packKey:'temu_pack_size'},
];

function CostCalculatorPage({prods,globalSettings,calcCost,saveGlobalSettings,costCalcOverrides,setCostCalcOverrides,S,lang}){
  const isES=lang==='es';
  const[editFees,setEditFees]=useState(false);
  const[fees,setFees]=useState({
    fee_amz:parseFloat(globalSettings.fee_amz)||0.15,
    fee_wmt:parseFloat(globalSettings.fee_wmt)||0.15,
    fee_tgt:parseFloat(globalSettings.fee_tgt)||0.15,
    fee_temu:parseFloat(globalSettings.fee_temu)||0.12,
  });
  const[filterProd,setFilterProd]=useState('');
  const[expandedProd,setExpandedProd]=useState(null);

  // Per-product per-platform overrides: {prodId: {amz:{ad,shipping}, wmt:{ad,shipping},...}}
  function getOverride(prodId,platKey){
    return costCalcOverrides[prodId]?.[platKey]||{ad:0,shipping:0};
  }
  function setOverride(prodId,platKey,field,val){
    setCostCalcOverrides(prev=>({
      ...prev,
      [prodId]:{
        ...(prev[prodId]||{}),
        [platKey]:{...((prev[prodId]||{})[platKey]||{}), [field]:parseFloat(val)||0}
      }
    }));
  }

  function calcPlatform(prod,plat){
    const baseCost=calcCost(prod,globalSettings).total;
    const packSize=parseFloat(prod[plat.packKey])||1;
    const costPerPack=baseCost*packSize;
    // selling price: use product's platform price field
    const sellPrice=parseFloat(prod[plat.key+'_sell'])||parseFloat(prod.price)||0;
    const feeRate=parseFloat(fees[plat.feeKey])||0;
    const{ad,shipping}=getOverride(prod.id,plat.key);
    const platformFee=sellPrice*feeRate;
    const totalCost=costPerPack+platformFee+ad+shipping;
    const profit=sellPrice-totalCost;
    const margin=sellPrice>0?(profit/sellPrice)*100:0;
    const breakEven=totalCost>0?totalCost/(1-feeRate):0;
    return{baseCost,costPerPack,sellPrice,platformFee,ad,shipping,totalCost,profit,margin,breakEven,packSize};
  }

  const marginColor=(m)=>m>=30?'#28a745':m>=15?'#856404':m>=0?'#e67e22':'#dc3545';
  const marginBg=(m)=>m>=30?'#d4edda':m>=15?'#FFF3CD':m>=0?'#fdebd0':'#f8d7da';

  const filteredProds=prods.filter(p=>!filterProd||p.name.toLowerCase().includes(filterProd.toLowerCase())||p.sku.toLowerCase().includes(filterProd.toLowerCase()));

  return(
    <div>
      {/* Header + fee settings */}
      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:'1.25rem',flexWrap:'wrap',gap:8}}>
        <div>
          <div style={{fontSize:20,fontWeight:600}}>💰 {isES?'Calculadora de Costos':'Cost Calculator'}</div>
          <div style={{fontSize:12,color:'#888',marginTop:2}}>{isES?'Analiza rentabilidad por canal y producto':'Analyze profitability by channel and product'}</div>
        </div>
        <button style={{...S.btn,gap:6}} onClick={()=>setEditFees(f=>!f)}>⚙️ {isES?'Tarifas de plataforma':'Platform Fees'}</button>
      </div>

      {/* Platform fee editor */}
      {editFees&&(
        <div style={{...S.card,marginBottom:'1rem',padding:'1rem 1.25rem'}}>
          <div style={{fontSize:13,fontWeight:600,marginBottom:'0.75rem'}}>⚙️ {isES?'Tarifas globales por defecto (%)':'Global Default Fees (%)'}</div>
          <div style={{display:'flex',gap:12,flexWrap:'wrap',alignItems:'flex-end'}}>
            {PLATFORMS_CALC.map(plat=>(
              <div key={plat.key} style={{display:'flex',flexDirection:'column',gap:4,minWidth:100}}>
                <label style={{fontSize:11,fontWeight:600,color:plat.color}}>{plat.label}</label>
                <div style={{display:'flex',alignItems:'center',gap:4}}>
                  <input style={{...S.inp,width:70,textAlign:'center'}} type="number" step="0.1" min="0" max="100"
                    value={((fees[plat.feeKey]||0)*100).toFixed(1)}
                    onChange={e=>setFees(prev=>({...prev,[plat.feeKey]:parseFloat(e.target.value)/100||0}))}/>
                  <span style={{fontSize:12,color:'#888'}}>%</span>
                </div>
              </div>
            ))}
            <button style={{...S.btnP,padding:'7px 16px',alignSelf:'flex-end'}} onClick={async()=>{
              await saveGlobalSettings({...globalSettings,...fees});
              setEditFees(false);
            }}>Save</button>
          </div>
        </div>
      )}

      {/* Search */}
      <div style={{marginBottom:'1rem'}}>
        <input style={{...S.inp,maxWidth:320}} placeholder={isES?'Buscar producto...':'Search product...'} value={filterProd} onChange={e=>setFilterProd(e.target.value)}/>
      </div>

      {/* Summary cards — all platforms all products */}
      <div style={{display:'flex',gap:10,marginBottom:'1.25rem',flexWrap:'wrap'}}>
        {PLATFORMS_CALC.map(plat=>{
          const margins=filteredProds.map(p=>calcPlatform(p,plat).margin).filter(m=>!isNaN(m));
          const avg=margins.length?margins.reduce((a,b)=>a+b,0)/margins.length:0;
          const losing=margins.filter(m=>m<0).length;
          return(
            <div key={plat.key} style={{background:'#fff',borderRadius:10,padding:'12px 16px',boxShadow:'0 1px 3px rgba(0,0,0,.07)',minWidth:140,flex:1}}>
              <div style={{fontSize:11,fontWeight:700,color:plat.color,marginBottom:4}}>{plat.label}</div>
              <div style={{fontSize:20,fontWeight:700,color:marginColor(avg)}}>{avg.toFixed(1)}%</div>
              <div style={{fontSize:11,color:'#888'}}>avg margin</div>
              {losing>0&&<div style={{fontSize:11,color:'#dc3545',marginTop:2}}>⚠️ {losing} losing money</div>}
            </div>
          );
        })}
      </div>

      {/* Product rows */}
      <div style={{display:'flex',flexDirection:'column',gap:8}}>
        {filteredProds.map(prod=>{
          const isExpanded=expandedProd===prod.id;
          const baseCost=calcCost(prod,globalSettings).total;
          return(
            <div key={prod.id} style={{background:'#fff',borderRadius:12,boxShadow:'0 1px 3px rgba(0,0,0,.07)',overflow:'hidden'}}>
              {/* Product header row */}
              <div style={{display:'flex',alignItems:'center',gap:12,padding:'12px 16px',cursor:'pointer',borderBottom:isExpanded?'1px solid #eee':'none'}} onClick={()=>setExpandedProd(isExpanded?null:prod.id)}>
                <div style={{flex:'0 0 220px',minWidth:0}}>
                  <div style={{fontSize:13,fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{prod.name}</div>
                  <div style={{fontSize:10,color:'#888'}}>{prod.sku} · base cost: {fm2(baseCost)}</div>
                </div>
                <div style={{display:'flex',gap:8,flex:1,flexWrap:'wrap'}}>
                  {PLATFORMS_CALC.map(plat=>{
                    const r=calcPlatform(prod,plat);
                    if(!r.sellPrice)return(
                      <div key={plat.key} style={{flex:1,minWidth:80,background:'#f8f8f8',borderRadius:8,padding:'6px 10px',textAlign:'center'}}>
                        <div style={{fontSize:10,color:plat.color,fontWeight:600}}>{plat.label}</div>
                        <div style={{fontSize:11,color:'#ccc',marginTop:2}}>no price</div>
                      </div>
                    );
                    return(
                      <div key={plat.key} style={{flex:1,minWidth:80,background:marginBg(r.margin),borderRadius:8,padding:'6px 10px',textAlign:'center'}}>
                        <div style={{fontSize:10,color:plat.color,fontWeight:600}}>{plat.label}</div>
                        <div style={{fontSize:15,fontWeight:700,color:marginColor(r.margin)}}>{r.margin.toFixed(1)}%</div>
                        <div style={{fontSize:10,color:'#666'}}>{fm2(r.profit)} profit</div>
                      </div>
                    );
                  })}
                </div>
                <div style={{fontSize:14,color:'#aaa',flexShrink:0}}>{isExpanded?'▲':'▼'}</div>
              </div>

              {/* Expanded breakdown */}
              {isExpanded&&(
                <div style={{padding:'16px'}}>
                  <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(260px,1fr))',gap:12}}>
                    {PLATFORMS_CALC.map(plat=>{
                      const r=calcPlatform(prod,plat);
                      const ov=getOverride(prod.id,plat.key);
                      return(
                        <div key={plat.key} style={{border:`1.5px solid ${plat.color}33`,borderRadius:10,padding:'12px 14px'}}>
                          <div style={{fontSize:12,fontWeight:700,color:plat.color,marginBottom:10}}>{plat.label} {r.packSize>1?`(${r.packSize}-pack)`:''}</div>
                          {/* Editable fields */}
                          <div style={{display:'flex',flexDirection:'column',gap:6,marginBottom:10}}>
                            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:8}}>
                              <span style={{fontSize:11,color:'#666',whiteSpace:'nowrap'}}>📦 Selling price</span>
                              <span style={{fontSize:12,fontWeight:600}}>{fm(r.sellPrice)}</span>
                            </div>
                            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:8}}>
                              <span style={{fontSize:11,color:'#666',whiteSpace:'nowrap'}}>🏭 Product cost</span>
                              <span style={{fontSize:12}}>{fm2(r.costPerPack)}</span>
                            </div>
                            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:8}}>
                              <span style={{fontSize:11,color:'#666',whiteSpace:'nowrap'}}>💳 Platform fee ({((fees[plat.feeKey]||0)*100).toFixed(0)}%)</span>
                              <span style={{fontSize:12,color:'#dc3545'}}>−{fm2(r.platformFee)}</span>
                            </div>
                            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:8}}>
                              <span style={{fontSize:11,color:'#666',whiteSpace:'nowrap'}}>📣 Ad cost $</span>
                              <input style={{...S.inp,width:72,textAlign:'right',fontSize:11,padding:'2px 6px'}} type="number" min="0" step="0.01" value={ov.ad||''} placeholder="0.00"
                                onChange={e=>setOverride(prod.id,plat.key,'ad',e.target.value)}/>
                            </div>
                            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:8}}>
                              <span style={{fontSize:11,color:'#666',whiteSpace:'nowrap'}}>🚚 Shipping $</span>
                              <input style={{...S.inp,width:72,textAlign:'right',fontSize:11,padding:'2px 6px'}} type="number" min="0" step="0.01" value={ov.shipping||''} placeholder="0.00"
                                onChange={e=>setOverride(prod.id,plat.key,'shipping',e.target.value)}/>
                            </div>
                          </div>
                          <div style={{borderTop:'1px solid #eee',paddingTop:8}}>
                            <div style={{display:'flex',justifyContent:'space-between',marginBottom:4}}>
                              <span style={{fontSize:11,color:'#666'}}>Break-even price</span>
                              <span style={{fontSize:11,fontWeight:600,color:'#555'}}>{fm(r.breakEven)}</span>
                            </div>
                            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                              <span style={{fontSize:12,fontWeight:600}}>Net profit</span>
                              <span style={{fontSize:15,fontWeight:700,color:marginColor(r.margin)}}>{fm2(r.profit)}</span>
                            </div>
                            <div style={{marginTop:6,background:marginBg(r.margin),borderRadius:6,padding:'4px 8px',textAlign:'center'}}>
                              <span style={{fontSize:14,fontWeight:700,color:marginColor(r.margin)}}>{r.margin.toFixed(1)}% margin</span>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── PRODUCT MODAL ─────────────────────────────────────────────
function ProductModal({t,S,mdata,setMdata,onSave,onClose,lang}){
  const[form,setForm]=useState({...{id:null,name:'',sku:'',category:'',stock:'',velocity:'',cost:'',price:'',reorder:'',supplier:'',amz:'',wmt:'',tgt:'',temu:'',other_sku:'',amz_pack_size:1,wmt_pack_size:1,tgt_pack_size:1,temu_pack_size:1,other_pack_size:1,product_type:'finished',weight_oz:'',raw_material_cost_per_kg:'',packaging_cost:'',box_cost:'',jumbo_box_cost:''},...(mdata.form||{})});
  const isEdit=!!form.id;
  return(
    <div style={S.overlay} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={S.sheet}>
        <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem'}}>{isEdit?t.edit:t.addProduct}</div>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
          <div style={{gridColumn:'1/-1',display:'flex',gap:8,marginBottom:4,alignItems:'center'}}>
          <label style={{fontSize:11,color:'#888',fontWeight:500}}>Product Type:</label>
          {['finished','packaged'].map(pt=>(
            <label key={pt} style={{display:'flex',alignItems:'center',gap:4,fontSize:12,cursor:'pointer'}}>
              <input type="radio" name="product_type" value={pt} checked={(form.product_type||'finished')===pt} onChange={()=>setForm(prev=>({...prev,product_type:pt}))}/>
              {pt==='finished'?(lang==='es'?'Terminado (comprado)':'Finished (bought)'):(lang==='es'?'Empacado (nosotros)':'Packaged (we package)')}
            </label>
          ))}
        </div>
      {[{l:t.productName,k:'name',full:true,ph:'Lactose Free 1.5lb'},{l:'Root SKU',k:'sku',ph:'BSL-LACT-150'},{l:t.category,k:'category',ph:'Lactose Free'},{l:`${t.stock}`,k:'stock',t:'number',ph:'0'},{l:t.monthlyS,k:'velocity',t:'number',ph:'0'},{l:`${t.price} ($)`,k:'price',t:'number',ph:'0.00'},{l:t.reorder,k:'reorder',t:'number',ph:'0'},{l:t.supplier,k:'supplier',full:true,ph:'EVI Labs'}].map(f=>(
            <div key={f.k} style={{gridColumn:f.full?'1/-1':undefined,display:'flex',flexDirection:'column',gap:3}}>
              <label style={{fontSize:11,color:'#888',fontWeight:500}}>{f.l}</label>
              <input style={S.inp} type={f.t||'text'} placeholder={f.ph} value={form[f.k]||''} onChange={e=>setForm(prev=>({...prev,[f.k]:e.target.value}))}/>
            </div>
          ))}
          {/* Cost fields - shown based on product type */}
          <div style={{gridColumn:'1/-1',fontSize:10,fontWeight:600,color:'#888',textTransform:'uppercase',letterSpacing:'.05em',marginTop:6,paddingTop:8,borderTop:'1px solid #eee'}}>
            {(form.product_type||'finished')==='packaged'?(lang==='es'?'Desglose de Costos':'Cost Breakdown'):(lang==='es'?'Costo':'Cost')}
          </div>
          {(form.product_type||'finished')==='finished'&&(
            <div style={{display:'flex',flexDirection:'column',gap:3}}>
              <label style={{fontSize:11,color:'#888',fontWeight:500}}>{t.cost} ($)</label>
              <input style={S.inp} type="number" placeholder="0.00" value={form.cost||''} onChange={e=>setForm(prev=>({...prev,cost:e.target.value}))}/>
            </div>
          )}
          {(form.product_type||'finished')==='packaged'&&<>
            {[
              {l:lang==='es'?'Peso (oz)':'Weight (oz)',k:'weight_oz',ph:'12.6'},
              {l:lang==='es'?'Precio materia prima/kg ($)':'Raw material price/kg ($)',k:'raw_material_cost_per_kg',ph:'4.24'},
              {l:lang==='es'?'Costo empaque ($)':'Packaging cost ($)',k:'packaging_cost',ph:'0.50'},
              {l:lang==='es'?'Costo caja ($)':'Box cost ($)',k:'box_cost',ph:'0.05'},
              {l:lang==='es'?'Costo caja jumbo ($) (opcional)':'Jumbo box cost ($) (optional)',k:'jumbo_box_cost',ph:'0.00'},
            ].map(f=>(
              <div key={f.k} style={{display:'flex',flexDirection:'column',gap:3}}>
                <label style={{fontSize:11,color:'#888',fontWeight:500}}>{f.l}</label>
                <input style={S.inp} type="number" placeholder={f.ph} value={form[f.k]||''} onChange={e=>setForm(prev=>({...prev,[f.k]:e.target.value}))}/>
              </div>
            ))}
          </>}
          <div style={{gridColumn:'1/-1',fontSize:10,fontWeight:600,color:'#888',textTransform:'uppercase',letterSpacing:'.05em',marginTop:6,paddingTop:8,borderTop:'1px solid #eee'}}>{t.marketplaceSKUs}</div>
          {[{l:t.amazonSku,k:'amz',pk:'amz_pack_size'},{l:t.walmartSku,k:'wmt',pk:'wmt_pack_size'},{l:t.targetSku,k:'tgt',pk:'tgt_pack_size'},{l:t.temuSku,k:'temu',pk:'temu_pack_size'},{l:t.otherSku,k:'other_sku',pk:'other_pack_size'}].map(f=>(
            <div key={f.k} style={{display:'flex',flexDirection:'column',gap:3}}>
              <label style={{fontSize:11,color:'#888',fontWeight:500}}>{f.l}</label>
              <div style={{display:'flex',gap:4}}>
                <input style={{...S.inp,flex:2}} placeholder={f.k.toUpperCase()+'-001'} value={form[f.k]||''} onChange={e=>setForm(prev=>({...prev,[f.k]:e.target.value}))}/>
                <input style={{...S.inp,flex:1,width:50}} type="number" placeholder="Pack" title="Pack size" value={form[f.pk]||1} onChange={e=>setForm(prev=>({...prev,[f.pk]:e.target.value}))}/>
              </div>
            </div>
          ))}
        </div>
        <div style={{fontSize:11,color:'#aaa',marginTop:8}}>Pack size = how many singles per unit sold on that platform</div>
        <div style={{display:'flex',gap:8,justifyContent:'flex-end',marginTop:'1rem'}}>
          <button style={S.btn} onClick={onClose}>{t.cancel}</button>
          <button style={S.btnP} onClick={()=>onSave(form)}>{t.save}</button>
        </div>
      </div>
    </div>
  );
}

// ── IMPORT MODAL ──────────────────────────────────────────────
function ImportModal({t,S,mdata,setMdata,onFile,onConfirm,onDownload,onClose}){
  const fRef=useRef();
  const step=mdata.step||1;
  return(
    <div style={S.overlay} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={S.sheet}>
        <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem'}}>Bulk import — step {step} of 2</div>
        {step===1&&<>
          <div style={{border:'2px dashed #ddd',borderRadius:12,padding:'2rem',textAlign:'center',cursor:'pointer',background:'#fafafa',marginBottom:'1rem'}} onClick={()=>fRef.current.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f){const dt=new DataTransfer();dt.items.add(f);fRef.current.files=dt.files;onFile({target:fRef.current})}}}>
            <div style={{fontSize:26,marginBottom:6}}>📂</div>
            <div style={{fontSize:13,fontWeight:600,marginBottom:4}}>Drop file or click to browse</div>
            <div style={{fontSize:11,color:'#aaa'}}>.csv or .xlsx</div>
            <input ref={fRef} type="file" accept=".csv,.xlsx,.xls" style={{display:'none'}} onChange={onFile}/>
          </div>
          <div style={{display:'flex',justifyContent:'space-between'}}>
            <button style={{...S.btn,color:'#4a90e2',borderColor:'#4a90e2',fontSize:11}} onClick={onDownload}>⬇ Download BSL template</button>
            <button style={S.btn} onClick={onClose}>{t.cancel}</button>
          </div>
        </>}
        {step===2&&<>
          {(mdata.errors||[]).length>0&&<div style={{background:'#FFF3CD',border:'1px solid #ffc107',borderRadius:8,padding:'9px 12px',marginBottom:'1rem',fontSize:12,color:'#856404'}}>{mdata.errors.length} row(s) skipped</div>}
          <div style={{display:'flex',gap:8,marginBottom:'1rem'}}>
            {['add','replace'].map(m=><div key={m} onClick={()=>setMdata(prev=>({...prev,mode:m}))} style={{flex:1,border:`2px solid ${(mdata.mode||'add')===m?'#111':'#eee'}`,borderRadius:10,padding:'9px 12px',cursor:'pointer',background:(mdata.mode||'add')===m?'#f8f8f8':'#fff'}}>
              <div style={{fontSize:13,fontWeight:600,marginBottom:2}}>{m==='add'?t.addUpdate||'Add/update':t.replaceAll||'Replace all'}</div>
            </div>)}
          </div>
          <div style={{fontSize:12,color:'#555',marginBottom:8}}>{(mdata.parsed||[]).length} products ready</div>
          <div style={{overflowX:'auto',border:'1px solid #eee',borderRadius:10,marginBottom:'1rem',maxHeight:200,overflowY:'auto'}}>
            <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
              <thead><tr>{['Name','Root SKU','Stock','AMZ (pack)','WMT (pack)'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
              <tbody>{(mdata.parsed||[]).map((p,i)=><tr key={i}><td style={S.td}>{p.name}</td><td style={S.td}>{p.sku||'—'}</td><td style={S.td}>{p.stock}</td><td style={S.td}>{p.amz||'—'} ({p.amz_pack_size||1}pk)</td><td style={S.td}>{p.wmt||'—'} ({p.wmt_pack_size||1}pk)</td></tr>)}</tbody>
            </table>
          </div>
          <div style={{display:'flex',gap:8,justifyContent:'space-between'}}>
            <button style={S.btn} onClick={()=>setMdata(prev=>({...prev,step:1}))}>← Back</button>
            <div style={{display:'flex',gap:8}}>
              <button style={S.btn} onClick={onClose}>{t.cancel}</button>
              <button style={S.btnP} onClick={onConfirm}>Import {(mdata.parsed||[]).length} products</button>
            </div>
          </div>
        </>}
      </div>
    </div>
  );
}

// ── ORDERS MODAL ──────────────────────────────────────────────
function OrdersModal({t,S,mdata,setMdata,onFile,onPreview,onConfirm,onApply,hashes,onClose}){
  const oRef=useRef();
  const step=mdata.step||1;const plat=mdata.platform||'amazon';
  const PLAT_LOCAL={amazon:{l:'Amazon',c:'#FF9900',sku:['sku','seller-sku','asin'],qty:['quantity','quantity-purchased','qty'],oid:['amazon-order-id','order-id']},walmart:{l:'Walmart',c:'#0071DC',sku:['sku','item-sku'],qty:['qty','quantity'],oid:['order_id','order-id']},target:{l:'Target Plus',c:'#CC0000',sku:['sku','tcin'],qty:['quantity','qty'],oid:['order-id']},temu:{l:'Temu',c:'#7B2D8B',sku:['sku','product-sku'],qty:['quantity','qty'],oid:['order-id']},other:{l:'Other',c:'#888',sku:['sku'],qty:['quantity','qty'],oid:['order-id']}};
  return(
    <div style={S.overlay} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={S.sheet}>
        <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem'}}>Upload orders — step {Math.min(step,3)} / 3</div>
        {step===1&&<>
          <div style={{fontSize:12,color:'#555',marginBottom:12}}>Select platform then upload order file</div>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:12}}>
            {Object.entries(PLAT_LOCAL).map(([k,v])=><div key={k} onClick={()=>setMdata(prev=>({...prev,platform:k}))} style={{border:`2px solid ${plat===k?v.c:'#eee'}`,borderRadius:10,padding:'10px 12px',cursor:'pointer',background:plat===k?v.c+'11':'#fff',display:'flex',alignItems:'center',gap:8}}>
              <div style={{width:8,height:8,borderRadius:'50%',background:v.c}}/><span style={{fontSize:13,fontWeight:plat===k?600:400}}>{v.l}</span>
            </div>)}
          </div>
          <div style={{background:'#f0f7ff',borderRadius:8,padding:'10px 12px',marginBottom:12,fontSize:12,color:'#1565c0'}}>
            💡 Pack sizes are configured per product. Orders will automatically deduct singles (packs × pack size).
          </div>
          <div style={{border:'2px dashed #ddd',borderRadius:12,padding:'1.5rem',textAlign:'center',cursor:'pointer',background:'#fafafa'}} onClick={()=>oRef.current.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f){const dt=new DataTransfer();dt.items.add(f);oRef.current.files=dt.files;onFile({target:oRef.current})}}}>
            <div style={{fontSize:24,marginBottom:6}}>📂</div>
            <div style={{fontSize:13,fontWeight:600,marginBottom:3}}>Drop file or click to browse</div>
            <div style={{fontSize:11,color:'#aaa'}}>{PLAT_LOCAL[plat].l} orders — .csv or .xlsx</div>
            <input ref={oRef} type="file" accept=".csv,.xlsx,.xls" style={{display:'none'}} onChange={onFile}/>
          </div>
          <div style={{display:'flex',justifyContent:'flex-end',marginTop:12}}><button style={S.btn} onClick={onClose}>{t.cancel}</button></div>
        </>}
        {step===2&&<>
          <div style={{fontSize:12,color:'#555',marginBottom:12}}>Confirm columns from {PLAT_LOCAL[plat].l} file</div>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:10,marginBottom:12}}>
            {[{l:'SKU column *',k:'skuCol'},{l:'Quantity column *',k:'qtyCol'},{l:'Order ID (dupe check)',k:'oidCol'}].map(f=>(
              <div key={f.k} style={{display:'flex',flexDirection:'column',gap:3}}>
                <label style={{fontSize:11,color:'#888',fontWeight:500}}>{f.l}</label>
                <select style={{...S.inp,width:'auto'}} value={mdata[f.k]??-1} onChange={e=>setMdata(prev=>({...prev,[f.k]:parseInt(e.target.value)}))}>
                  <option value={-1}>— not mapped —</option>
                  {(mdata.headers||[]).map((h,i)=><option key={i} value={i}>{h}</option>)}
                </select>
              </div>
            ))}
          </div>
          <div style={{display:'flex',gap:8,justifyContent:'space-between'}}>
            <button style={S.btn} onClick={()=>setMdata(prev=>({...prev,step:1}))}>← Back</button>
            <div style={{display:'flex',gap:8}}>
              <button style={S.btn} onClick={onClose}>{t.cancel}</button>
              <button style={{...S.btn,background:'#e67e22',color:'#fff',border:'none'}} onClick={onPreview}>Preview →</button>
            </div>
          </div>
        </>}
        {step===3&&<>
          <div style={{fontSize:12,color:'#555',marginBottom:8}}>{PLAT_LOCAL[plat].l} — {(mdata.preview||[]).filter(p=>p.prod).length} matched, {(mdata.preview||[]).filter(p=>!p.prod).length} not found</div>
          {(mdata.warnings||[]).length>0&&<div style={{background:'#FFF3CD',border:'1px solid #ffc107',borderRadius:8,padding:'9px 12px',marginBottom:'1rem',fontSize:12,color:'#856404'}}>{mdata.warnings.slice(0,4).join(' · ')}</div>}
          <div style={{overflowX:'auto',border:'1px solid #eee',borderRadius:10,marginBottom:'1rem',maxHeight:240,overflowY:'auto'}}>
            <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
              <thead><tr>{['SKU','Product','Packs','× Pack Size','= Singles','Current','New Stock'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
              <tbody>{(mdata.preview||[]).map((item,i)=>(
                <tr key={i} style={{background:!item.prod?'#FFF5F5':item.newStock<0?'#FFFBF0':''}}>
                  <td style={S.td}><code style={{fontSize:10,background:'#f5f5f5',padding:'1px 4px',borderRadius:3}}>{item.sku}</code></td>
                  <td style={{...S.td,fontWeight:500,color:item.prod?'#111':'#dc3545'}}>{item.prod?item.prod.name:'Not found'}</td>
                  <td style={S.td}>{item.packsOrdered}</td>
                  <td style={S.td}>{item.packSize}</td>
                  <td style={{...S.td,color:'#dc3545',fontWeight:600}}>-{item.singlesDeducted}</td>
                  <td style={S.td}>{item.prod?item.prod.stock:'—'}</td>
                  <td style={{...S.td,fontWeight:600,color:item.newStock<0?'#dc3545':item.newStock<=10?'#856404':'#28a745'}}>{item.prod!=null?item.newStock:'—'}</td>
                </tr>
              ))}</tbody>
            </table>
          </div>
          <div style={{display:'flex',gap:8,justifyContent:'space-between'}}>
            <button style={S.btn} onClick={()=>setMdata(prev=>({...prev,step:2}))}>← Back</button>
            <div style={{display:'flex',gap:8}}>
              <button style={S.btn} onClick={onClose}>{t.cancel}</button>
              <button style={{...S.btn,background:'#28a745',color:'#fff',border:'none'}} onClick={onConfirm}>✓ Apply {(mdata.preview||[]).filter(p=>p.prod).length} deductions</button>
            </div>
          </div>
        </>}
        {step===4&&<>
          <div style={{textAlign:'center',padding:'1.5rem 0'}}>
            <div style={{fontSize:36,marginBottom:12}}>⚠️</div>
            <div style={{fontSize:15,fontWeight:600,marginBottom:8}}>Possible duplicate upload</div>
            <div style={{fontSize:12,color:'#666',marginBottom:'1.5rem'}}>This file looks like it was already uploaded. Processing again will double-deduct quantities.</div>
            <div style={{display:'flex',gap:10,justifyContent:'center'}}>
              <button style={S.btn} onClick={()=>setMdata(prev=>({...prev,step:3}))}>← Go back</button>
              <button style={{...S.btn,background:'#dc3545',color:'#fff',border:'none'}} onClick={onApply}>Process anyway</button>
            </div>
          </div>
        </>}
      </div>
    </div>
  );
}

// ── PACKING LIST MODAL ────────────────────────────────────────
function PackingModal({t,S,mdata,setMdata,onFile,onApply,onClose,prods}){
  const pRef=useRef();
  const step=mdata.step||1;
  const parsed=mdata.parsed||{};

  function updateItem(i,field,value){
    const items=[...(parsed.items||[])];
    items[i]={...items[i],[field]:value};
    // If SKU changed, update matchedSku and matchedName
    if(field==='matchedSku'){
      const prod=prods.find(p=>p.sku===value);
      items[i].matchedName=prod?.name||value;
      if(prod) items[i].description=prod.name;
    }
    setMdata(prev=>({...prev,parsed:{...prev.parsed,items,totalSingles:items.filter(x=>x.included!==false).reduce((a,x)=>a+(parseFloat(x.singles)||0),0)}}));
  }

  function toggleItem(i){
    const items=[...(parsed.items||[])];
    items[i]={...items[i],included:items[i].included===false?true:false};
    setMdata(prev=>({...prev,parsed:{...prev.parsed,items,totalSingles:items.filter(x=>x.included!==false).reduce((a,x)=>a+(parseFloat(x.singles)||0),0)}}));
  }

  const includedItems=(parsed.items||[]).filter(x=>x.included!==false);
  const totalSingles=includedItems.reduce((a,x)=>a+(parseFloat(x.singles)||0),0);

  return(
    <div style={S.overlay} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={{...S.sheet,width:680,maxWidth:'96vw'}}>
        <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem'}}>📸 {t.packingListTitle} — step {Math.min(step,3)} / 3</div>
        {step===1&&<>
          <div style={{fontSize:12,color:'#555',marginBottom:12}}>Upload a photo or PDF of the packing list. Claude will read it automatically.</div>
          <div style={{border:'2px dashed #ddd',borderRadius:12,padding:'2rem',textAlign:'center',cursor:'pointer',background:'#fafafa',marginBottom:'1rem'}} onClick={()=>pRef.current.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f){const dt=new DataTransfer();dt.items.add(f);pRef.current.files=dt.files;onFile({target:pRef.current})}}}>
            <div style={{fontSize:36,marginBottom:8}}>📸</div>
            <div style={{fontSize:13,fontWeight:600,marginBottom:4}}>Drop photo or PDF here, or click to browse</div>
            <div style={{fontSize:11,color:'#aaa'}}>JPG, PNG, or PDF — Claude will read and extract all items</div>
            <input ref={pRef} type="file" accept="image/*,.pdf" style={{display:'none'}} onChange={onFile}/>
          </div>
          <div style={{display:'flex',justifyContent:'flex-end'}}><button style={S.btn} onClick={onClose}>{t.cancel}</button></div>
        </>}
        {step===2&&<>
          <div style={{textAlign:'center',padding:'2rem'}}>
            <div style={{fontSize:32,marginBottom:12}}>🔍</div>
            <div style={{fontSize:14,fontWeight:600,marginBottom:4}}>{t.processing}</div>
            <div style={{fontSize:12,color:'#888'}}>Claude is reading your packing list...</div>
          </div>
        </>}
        {step===3&&<>
          {/* Header info — editable */}
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:'1rem'}}>
            <div style={{display:'flex',flexDirection:'column',gap:3}}>
              <label style={{fontSize:11,color:'#888',fontWeight:500}}>Supplier</label>
              <input style={S.inp} value={parsed.supplier||''} onChange={e=>setMdata(prev=>({...prev,parsed:{...prev.parsed,supplier:e.target.value}}))}/>
            </div>
            <div style={{display:'flex',flexDirection:'column',gap:3}}>
              <label style={{fontSize:11,color:'#888',fontWeight:500}}>Customer / Ship To</label>
              <input style={S.inp} value={parsed.customer||''} onChange={e=>setMdata(prev=>({...prev,parsed:{...prev.parsed,customer:e.target.value}}))}/>
            </div>
            <div style={{display:'flex',flexDirection:'column',gap:3}}>
              <label style={{fontSize:11,color:'#888',fontWeight:500}}>Date</label>
              <input style={S.inp} value={parsed.date||''} onChange={e=>setMdata(prev=>({...prev,parsed:{...prev.parsed,date:e.target.value}}))}/>
            </div>
            <div style={{display:'flex',flexDirection:'column',gap:3}}>
              <label style={{fontSize:11,color:'#888',fontWeight:500}}>Direction</label>
              <select style={S.inp} value={parsed.direction||'outbound'} onChange={e=>setMdata(prev=>({...prev,parsed:{...prev.parsed,direction:e.target.value}}))}>
                <option value="outbound">📤 Outbound (deduct stock)</option>
                <option value="inbound">📥 Inbound (add stock)</option>
              </select>
            </div>
          </div>

          {/* Customer info for outbound */}
          {parsed.direction==='outbound'&&<>
            <div style={{fontSize:12,fontWeight:500,color:'#555',marginBottom:6}}>Customer details (optional)</div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:'1rem'}}>
              <input style={S.inp} placeholder="Customer email" value={mdata.customerEmail||''} onChange={e=>setMdata(prev=>({...prev,customerEmail:e.target.value}))}/>
              <input style={S.inp} placeholder="Customer phone" value={mdata.customerPhone||''} onChange={e=>setMdata(prev=>({...prev,customerPhone:e.target.value}))}/>
            </div>
          </>}

          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:6}}>
            <div style={{fontSize:12,fontWeight:500,color:'#555'}}>Items found ({(parsed.items||[]).length}) — uncheck to exclude</div>
            <div style={{fontSize:11,color:'#888'}}>Click any field to edit</div>
          </div>

          <div style={{border:'1px solid #eee',borderRadius:10,marginBottom:'1rem',maxHeight:280,overflowY:'auto'}}>
            {(parsed.items||[]).map((item,i)=>{
              const excluded=item.included===false;
              return(
                <div key={i} style={{padding:'10px 12px',borderBottom:'1px solid #f5f5f5',background:excluded?'#f9f9f9':item.confidence==='low'||!item.matchedSku?'#FFF5F5':item.confidence==='medium'?'#FFFBF0':'#fff',opacity:excluded?0.5:1}}>
                  <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:6}}>
                    <input type="checkbox" checked={!excluded} onChange={()=>toggleItem(i)} style={{width:15,height:15,cursor:'pointer',flexShrink:0}}/>
                    <span style={{fontSize:10,background:item.confidence==='high'?'#D4EDDA':item.confidence==='medium'?'#FFF3CD':'#F8D7DA',color:item.confidence==='high'?'#155724':item.confidence==='medium'?'#856404':'#721C24',padding:'1px 7px',borderRadius:99,flexShrink:0}}>{item.confidence||'—'}</span>
                    <span style={{fontSize:11,color:'#555',flex:1}} title={item.description}>{item.description}</span>
                  </div>
                  <div style={{display:'grid',gridTemplateColumns:'1fr 2fr 80px',gap:8,paddingLeft:23}}>
                    <div style={{display:'flex',flexDirection:'column',gap:2}}>
                      <label style={{fontSize:10,color:'#aaa'}}>Matched SKU</label>
                      <select style={{...S.inp,fontSize:11,padding:'4px 6px'}} value={item.matchedSku||''} onChange={e=>updateItem(i,'matchedSku',e.target.value)} disabled={excluded}>
                        <option value="">— No match —</option>
                        {prods.map(p=><option key={p.id} value={p.sku}>{p.sku} — {p.name}</option>)}
                      </select>
                    </div>
                    <div style={{display:'flex',flexDirection:'column',gap:2}}>
                      <label style={{fontSize:10,color:'#aaa'}}>Description</label>
                      <input style={{...S.inp,fontSize:11,padding:'4px 6px'}} value={item.description||''} onChange={e=>updateItem(i,'description',e.target.value)} disabled={excluded}/>
                    </div>
                    <div style={{display:'flex',flexDirection:'column',gap:2}}>
                      <label style={{fontSize:10,color:'#aaa'}}>Singles</label>
                      <input style={{...S.inp,fontSize:11,padding:'4px 6px',fontWeight:600,color:parsed.direction==='outbound'?'#dc3545':'#28a745'}} type="number" value={item.singles||0} onChange={e=>updateItem(i,'singles',parseFloat(e.target.value)||0)} disabled={excluded}/>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>

          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:'1rem'}}>
            <div style={{fontSize:12,color:'#555'}}>
              <span style={{fontWeight:600}}>{includedItems.length}</span> of {(parsed.items||[]).length} items included
            </div>
            <div style={{fontSize:13,fontWeight:700,color:parsed.direction==='outbound'?'#dc3545':'#28a745'}}>
              {parsed.direction==='outbound'?'−':'+' }{totalSingles} singles
            </div>
          </div>

          <div style={{display:'flex',gap:8,justifyContent:'space-between'}}>
            <button style={S.btn} onClick={()=>setMdata(prev=>({...prev,step:1,parsed:null}))}>← Back</button>
            <div style={{display:'flex',gap:8}}>
              <button style={S.btn} onClick={onClose}>{t.cancel}</button>
              <button style={{...S.btn,background:'#28a745',color:'#fff',border:'none'}} onClick={onApply} disabled={includedItems.length===0}>{t.applyDeductions} ({includedItems.length} items)</button>
            </div>
          </div>
        </>}
      </div>
    </div>
  );
}

// ── COST BREAKDOWN MODAL ─────────────────────────────────────
function CostBreakdownModal({prod,globalSettings,S,lang,onClose,onSaveSettings}){
  const[gs,setGs]=useState({...globalSettings});
  const[editingGs,setEditingGs]=useState(false);
  const isES=lang==='es';

  const OZ_PER_KG_MODAL=35.274;
  function calc(p,g){
    const rmWaste=g.raw_material_waste_pct||0.005;
    const pkgWaste=g.packaging_waste_pct||0.005;
    const filling=g.filling_cost||1.15;
    const pricePerKg=parseFloat(p.raw_material_cost_per_kg)||0;
    const pricePerOz=pricePerKg/OZ_PER_KG_MODAL;
    const weightOz=parseFloat(p.weight_oz)||0;
    const rawMaterial=pricePerOz*weightOz;
    const rawWithWaste=rawMaterial*(1+rmWaste);
    const pkgCost=parseFloat(p.packaging_cost)||0;
    const pkgWithWaste=pkgCost*(1+pkgWaste);
    const box=parseFloat(p.box_cost)||0;
    const jumboBox=parseFloat(p.jumbo_box_cost)||0;
    const total=rawWithWaste+pkgWithWaste+filling+box+jumboBox;
    return{pricePerKg,pricePerOz,weightOz,rawMaterial,rawWithWaste,pkgCost,pkgWithWaste,filling,box,jumboBox,total};
  }

  const c=calc(prod,gs);
  const fm4=n=>'$'+parseFloat(n||0).toFixed(4);
  const pct=n=>((n||0)*100).toFixed(2)+'%';

  const row=(label,value,sub,highlight)=>(
    <div style={{display:'flex',justifyContent:'space-between',alignItems:'baseline',padding:'6px 0',borderBottom:'0.5px solid #f0f0f0',background:highlight?'#f8f9ff':'transparent',paddingLeft:highlight?0:0}}>
      <span style={{fontSize:12,color:highlight?'#111':'#555',fontWeight:highlight?600:400}}>{label}</span>
      <div style={{textAlign:'right'}}>
        <span style={{fontSize:12,fontWeight:highlight?700:500,color:highlight?'#111':'#333'}}>{value}</span>
        {sub&&<span style={{fontSize:10,color:'#aaa',marginLeft:4}}>{sub}</span>}
      </div>
    </div>
  );

  return(
    <div style={{position:'fixed',top:0,left:0,right:0,bottom:0,background:'rgba(0,0,0,.45)',zIndex:300,display:'flex',alignItems:'center',justifyContent:'center'}} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={{background:'#fff',borderRadius:16,padding:'1.5rem',width:480,maxWidth:'95vw',maxHeight:'90vh',overflowY:'auto'}}>
        
        {/* Header */}
        <div style={{marginBottom:'1.25rem'}}>
          <div style={{fontSize:15,fontWeight:700,marginBottom:2}}>{prod.name}</div>
          <div style={{fontSize:11,color:'#888'}}>{isES?'Desglose de Costos':'Cost Breakdown'} — {prod.sku}</div>
        </div>

        {/* Product specs */}
        <div style={{background:'#f8f8f8',borderRadius:8,padding:'10px 12px',marginBottom:'1rem',display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
          <div><div style={{fontSize:10,color:'#aaa'}}>{isES?'Peso':'Weight'}</div><div style={{fontSize:13,fontWeight:600}}>{prod.weight_oz||0} oz</div></div>
          <div><div style={{fontSize:10,color:'#aaa'}}>{isES?'Precio venta':'Selling price'}</div><div style={{fontSize:13,fontWeight:600}}>${prod.price||0}</div></div>
          <div><div style={{fontSize:10,color:'#aaa'}}>{isES?'Precio materia prima/kg':'Raw material/kg'}</div><div style={{fontSize:13,fontWeight:600}}>${prod.raw_material_cost_per_kg||0}/kg</div></div>
          <div><div style={{fontSize:10,color:'#aaa'}}>{isES?'Costo por oz':'Cost per oz'}</div><div style={{fontSize:13,fontWeight:600}}>${((parseFloat(prod.raw_material_cost_per_kg)||0)/35.274).toFixed(4)}/oz</div></div>
          <div><div style={{fontSize:10,color:'#aaa'}}>{isES?'Costo empaque':'Packaging'}</div><div style={{fontSize:13,fontWeight:600}}>${prod.packaging_cost||0}</div></div>
        </div>

        {/* Cost breakdown table */}
        <div style={{marginBottom:'1rem'}}>
          <div style={{fontSize:11,fontWeight:600,color:'#888',textTransform:'uppercase',letterSpacing:'.04em',marginBottom:6}}>{isES?'Materia Prima':'Raw Material'}</div>
          {row(isES?`Materia prima (${prod.category||prod.name.split(' ')[0]})`:`Raw material (${prod.category||prod.name.split(' ')[0]})`, fm2(c.rawMaterial), `$${c.pricePerOz.toFixed(4)}/oz × ${c.weightOz}oz`)}
          {row(isES?`Merma (${pct(gs.raw_material_waste_pct)})`:`Waste (${pct(gs.raw_material_waste_pct)})`, fm2(c.rawWithWaste-c.rawMaterial))}
          {row(isES?'Costo MP + Merma':'Raw Material + Waste', fm2(c.rawWithWaste), null, true)}
        </div>

        <div style={{marginBottom:'1rem'}}>
          <div style={{fontSize:11,fontWeight:600,color:'#888',textTransform:'uppercase',letterSpacing:'.04em',marginBottom:6}}>{isES?'Material de Empaque':'Packaging Material'}</div>
          {row(isES?'Costo empaque':'Packaging cost', fm2(prod.packaging_cost||0))}
          {row(isES?`Merma empaque (${pct(gs.packaging_waste_pct)})`:`Packaging waste (${pct(gs.packaging_waste_pct)})`, fm2(c.pkgWithWaste-c.pkgCost))}
          {row(isES?'Empaque + Merma':'Packaging + Waste', fm2(c.pkgWithWaste))}
          {row(isES?'Envasado/Sellado':'Filling/Sealing', fm2(c.filling))}
          {row(isES?'Caja':'Box cost', fm2(c.box))}
          {c.jumboBox>0&&row(isES?'Caja Jumbo':'Jumbo Box cost', fm2(c.jumboBox))}
          {row(isES?'Total Empaque':'Total Packaging', fm2(c.pkgWithWaste+c.filling+c.box+c.jumboBox), null, true)}
        </div>

        <div style={{background:'#111',borderRadius:10,padding:'12px 14px',marginBottom:'1rem',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
          <span style={{color:'#fff',fontSize:14,fontWeight:600}}>{isES?'Costo Producto Terminado':'Total Product Cost'}</span>
          <span style={{color:'#fff',fontSize:20,fontWeight:700}}>${c.total.toFixed(4)}</span>
        </div>

        {prod.price>0&&(
          <div style={{background:'#D4EDDA',borderRadius:8,padding:'10px 12px',marginBottom:'1rem',display:'flex',justifyContent:'space-between'}}>
            <span style={{fontSize:12,color:'#155724',fontWeight:500}}>{isES?'Margen bruto':'Gross margin'}</span>
            <span style={{fontSize:13,color:'#155724',fontWeight:700}}>${(prod.price-c.total).toFixed(4)} ({(((prod.price-c.total)/prod.price)*100).toFixed(1)}%)</span>
          </div>
        )}

        {/* Global settings toggle */}
        <div style={{borderTop:'1px solid #eee',paddingTop:'1rem',marginBottom:'1rem'}}>
          <button style={{...S.btn,fontSize:11,color:'#4a90e2',borderColor:'#4a90e2'}} onClick={()=>setEditingGs(!editingGs)}>
            ⚙️ {isES?'Editar configuración global':'Edit global settings'}
          </button>
          {editingGs&&(
            <div style={{marginTop:'1rem',display:'flex',flexDirection:'column',gap:8}}>
              <div style={{fontSize:11,color:'#888',marginBottom:4}}>{isES?'Estos valores aplican a todos los productos':'These values apply to all packaged products'}</div>
              {[
                {k:'raw_material_waste_pct',l:isES?'Merma materia prima %':'Raw material waste %',mult:100},
                {k:'packaging_waste_pct',l:isES?'Merma empaque %':'Packaging waste %',mult:100},
                {k:'filling_cost',l:isES?'Costo envasado ($)':'Filling/sealing cost ($)',mult:1},
                {k:'packaging_cost_default',l:isES?'Empaque default ($)':'Packaging default ($)',mult:1},
              ].map(f=>(
                <div key={f.k} style={{display:'flex',alignItems:'center',gap:8,justifyContent:'space-between'}}>
                  <label style={{fontSize:12,color:'#555'}}>{f.l}</label>
                  <input style={{...S.inp,width:80,textAlign:'right'}} type="number" step="0.01"
                    value={f.mult>1?(gs[f.k]*f.mult).toFixed(2):gs[f.k]}
                    onChange={e=>setGs(prev=>({...prev,[f.k]:parseFloat(e.target.value)/(f.mult)||0}))}/>
                </div>
              ))}
              <button style={{...S.btn,background:'#111',color:'#fff',border:'none',alignSelf:'flex-end'}} onClick={()=>{onSaveSettings(gs);setEditingGs(false);}}>
                {isES?'Guardar configuración':'Save settings'}
              </button>
            </div>
          )}
        </div>

        <div style={{display:'flex',justifyContent:'flex-end'}}>
          <button style={{...S.btn,background:'#111',color:'#fff',border:'none'}} onClick={onClose}>{isES?'Cerrar':'Close'}</button>
        </div>
      </div>
    </div>
  );
}

// ── LOCATION MODAL ───────────────────────────────────────────
const LOCATIONS=['Warehouse','EVI','Tripolac'];
function LocationModal({prod,locations,S,lang,onClose,onSave,onDelete}){
  const isES=lang==='es';
  const prodLocs=locations.filter(l=>l.product_id===prod.id);
  // Build editable rows: one per known location, pre-filled if exists
  const[dirty,setDirty]=useState(false);
  const[rows,setRows]=useState(()=>LOCATIONS.map(loc=>{
    const existing=prodLocs.find(l=>l.location===loc);
    return existing?{...existing}:{location:loc,qty:0,notes:'',_new:true};
  }));
  const[saving,setSaving]=useState(false);
  const total=rows.reduce((a,r)=>a+(parseFloat(r.qty)||0),0);

  function setRow(i,field,val){setDirty(true);setRows(prev=>prev.map((r,ri)=>ri===i?{...r,[field]:val}:r));}

  async function handleSave(){
    setSaving(true);
    // Only save rows that have been touched (qty>0 or have id)
    const toSave=rows.filter(r=>!r._new||(parseFloat(r.qty)||0)>0);
    await onSave(toSave);
    setSaving(false);
  }

  const loc_colors={Warehouse:'#4a90e2',EVI:'#28a745',Tripolac:'#e67e22'};

  return(
    <div style={{position:'fixed',top:0,left:0,right:0,bottom:0,background:'rgba(0,0,0,.45)',zIndex:300,display:'flex',alignItems:'center',justifyContent:'center'}} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={{background:'#fff',borderRadius:16,padding:'1.5rem',width:460,maxWidth:'95vw',maxHeight:'90vh',overflowY:'auto'}}>

        {/* Header */}
        <div style={{marginBottom:'1.25rem'}}>
          <div style={{fontSize:15,fontWeight:700,marginBottom:2}}>{prod.name}</div>
          <div style={{fontSize:11,color:'#888'}}>{isES?'Stock por Ubicación':'Stock by Location'} — {prod.sku}</div>
        </div>

        {/* Location bars */}
        <div style={{display:'flex',gap:8,marginBottom:'1.25rem'}}>
          {rows.map((r,i)=>{
            const pct=total>0?((parseFloat(r.qty)||0)/total*100):0;
            return(
              <div key={r.location} style={{flex:1,textAlign:'center'}}>
                <div style={{fontSize:10,color:'#888',marginBottom:4,fontWeight:600,textTransform:'uppercase'}}>{r.location}</div>
                <div style={{background:'#f0f0f0',borderRadius:6,height:6,overflow:'hidden',marginBottom:4}}>
                  <div style={{width:pct+'%',height:'100%',background:loc_colors[r.location]||'#888',borderRadius:6,transition:'width .3s'}}/>
                </div>
                <div style={{fontSize:16,fontWeight:700,color:loc_colors[r.location]||'#333'}}>{parseFloat(r.qty)||0}</div>
                <div style={{fontSize:10,color:'#aaa'}}>{pct.toFixed(0)}%</div>
              </div>
            );
          })}
        </div>

        {/* Total */}
        <div style={{background:'#f8f8f8',borderRadius:8,padding:'8px 12px',marginBottom:'1.25rem',display:'flex',justifyContent:'space-between',alignItems:'center'}}>
          <span style={{fontSize:12,color:'#555',fontWeight:500}}>{isES?'Total en sistema':'Total in system'}</span>
          <span style={{fontSize:14,fontWeight:700}}>{prod.stock} singles</span>
        </div>
        {!dirty&&total!==prod.stock&&<div style={{background:'#FFF3CD',borderRadius:8,padding:'7px 12px',marginBottom:'1rem',fontSize:12,color:'#856404'}}>⚠️ {isES?'La suma de ubicaciones':'Location sum'} ({total}) {isES?'no coincide con el stock total':'does not match total stock'} ({prod.stock})</div>}

        {/* Edit rows */}
        <div style={{display:'flex',flexDirection:'column',gap:10,marginBottom:'1.25rem'}}>
          {rows.map((r,i)=>(
            <div key={r.location} style={{border:`1.5px solid ${loc_colors[r.location]||'#eee'}22`,borderRadius:10,padding:'10px 12px',background:'#fafafa'}}>
              <div style={{display:'flex',alignItems:'center',gap:10}}>
                <div style={{width:10,height:10,borderRadius:'50%',background:loc_colors[r.location]||'#888',flexShrink:0}}/>
                <span style={{fontSize:13,fontWeight:600,flex:1}}>{r.location}</span>
                <div style={{display:'flex',alignItems:'center',gap:6}}>
                  <button style={{...S.btn,padding:'2px 8px',fontSize:16,lineHeight:1}} onClick={()=>setRow(i,'qty',Math.max(0,(parseFloat(r.qty)||0)-1))}>−</button>
                  <input style={{...S.inp,width:70,textAlign:'center',fontWeight:700,fontSize:14}} type="number" min="0" value={r.qty||0} onChange={e=>setRow(i,'qty',parseFloat(e.target.value)||0)}/>
                  <button style={{...S.btn,padding:'2px 8px',fontSize:16,lineHeight:1}} onClick={()=>setRow(i,'qty',(parseFloat(r.qty)||0)+1)}>+</button>
                </div>
              </div>
              <input style={{...S.inp,marginTop:7,fontSize:11,color:'#888'}} placeholder={isES?'Notas (opcional)':'Notes (optional)'} value={r.notes||''} onChange={e=>setRow(i,'notes',e.target.value)}/>
            </div>
          ))}
        </div>

        <div style={{display:'flex',gap:8,justifyContent:'flex-end'}}>
          <button style={S.btn} onClick={onClose}>{isES?'Cancelar':'Cancel'}</button>
          <button style={{...S.btn,background:'#111',color:'#fff',border:'none'}} onClick={handleSave} disabled={saving}>{saving?'Saving...':isES?'Guardar':'Save'}</button>
        </div>
      </div>
    </div>
  );
}

// ── NOTE MODAL ────────────────────────────────────────────────
function NoteModal({t,S,mdata,onSave,onClose}){
  const[form,setForm]=useState(mdata.form||{title:'',content:'',category:'general'});
  const categories=['general','customer','supplier','product','pricing','operations'];
  return(
    <div style={S.overlay} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={S.sheet}>
        <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem'}}>🧠 {t.addNote}</div>
        <div style={{display:'flex',flexDirection:'column',gap:10}}>
          <div style={{display:'flex',flexDirection:'column',gap:3}}>
            <label style={{fontSize:11,color:'#888',fontWeight:500}}>{t.noteCategory}</label>
            <select style={S.inp} value={form.category} onChange={e=>setForm(prev=>({...prev,category:e.target.value}))}>
              {categories.map(c=><option key={c} value={c}>{c.charAt(0).toUpperCase()+c.slice(1)}</option>)}
            </select>
          </div>
          <div style={{display:'flex',flexDirection:'column',gap:3}}>
            <label style={{fontSize:11,color:'#888',fontWeight:500}}>{t.noteTitle}</label>
            <input style={S.inp} placeholder="e.g. EVI Labs minimum order" value={form.title} onChange={e=>setForm(prev=>({...prev,title:e.target.value}))}/>
          </div>
          <div style={{display:'flex',flexDirection:'column',gap:3}}>
            <label style={{fontSize:11,color:'#888',fontWeight:500}}>{t.noteContent}</label>
            <textarea style={{...S.inp,minHeight:100,resize:'vertical'}} placeholder="e.g. EVI Labs minimum order is 120 units per SKU. Contact: info@evicorporation.com" value={form.content} onChange={e=>setForm(prev=>({...prev,content:e.target.value}))}/>
          </div>
          <div style={{background:'#f0f7ff',borderRadius:8,padding:'10px 12px',fontSize:12,color:'#1565c0'}}>
            💡 This note will be included in every conversation with Claude — use it to store important instructions, customer preferences, supplier info, or anything Claude should always know.
          </div>
        </div>
        <div style={{display:'flex',gap:8,justifyContent:'flex-end',marginTop:'1rem'}}>
          <button style={S.btn} onClick={onClose}>{t.cancel}</button>
          <button style={S.btnP} onClick={()=>onSave(form)}>{t.save}</button>
        </div>
      </div>
    </div>
  );
}
