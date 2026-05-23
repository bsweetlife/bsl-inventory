// BSL Inventory v3.3 - supabase-auth-login
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
const COL_MAP={"name":"name","product name":"name","product":"name","root sku":"sku","sku":"sku","category":"category","cat":"category","stock":"stock","stock (singles)":"stock","current stock":"stock","qty":"stock","quantity":"stock","monthly sales":"velocity","monthly sales (singles)":"velocity","velocity":"velocity","cost":"cost","cost per single ($)":"cost","cost price":"cost","price":"price","selling price single ($)":"price","selling price":"price","reorder":"reorder","reorder point":"reorder","reorder point (singles)":"reorder","supplier":"supplier","supplier name":"supplier","amazon sku":"amz","amz":"amz","amz pack size":"amz_pack_size","walmart sku":"wmt","wmt":"wmt","wmt pack size":"wmt_pack_size","target sku":"tgt","tgt":"tgt","tgt pack size":"tgt_pack_size","temu sku":"temu","temu":"temu","temu pack size":"temu_pack_size","other sku":"other_sku","other":"other_sku","other pack size":"other_pack_size","product notes":"notes_field"};

const gs=p=>{const d=(p.velocity||0)/30;if(!d)return'ok';const v=p.stock/d;return v<=15?'crit':v<=30?'low':'ok'};
const gd=p=>{const d=(p.velocity||0)/30;return d?Math.round(p.stock/d):null};
const fm=n=>(n!=null&&n!=='')?'$'+parseFloat(n).toFixed(2):'—';
const hs=s=>{let h=0;for(let i=0;i<Math.min(s.length,500);i++)h=(Math.imul(31,h)+s.charCodeAt(i))|0;return h.toString()};
const fc=(hdrs,cs)=>{for(const c of cs){const i=hdrs.findIndex(h=>h.toLowerCase().replace(/[\s_-]+/g,'-')===c);if(i>=0)return i;}for(const c of cs){const i=hdrs.findIndex(h=>h.toLowerCase().includes(c.replace(/-/g,'')));if(i>=0)return i;}return -1};
const ep=()=>({id:null,name:'',sku:'',category:'',stock:'',velocity:'',cost:'',price:'',reorder:'',supplier:'',amz:'',wmt:'',tgt:'',temu:'',other_sku:'',amz_pack_size:1,wmt_pack_size:1,tgt_pack_size:1,temu_pack_size:1,other_pack_size:1});

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
  const[lang,setLang]=useState('en');
  const[page,setPage]=useState('dashboard');
  const[tab,setTab]=useState('all');
  const[prods,setProds]=useState([]);
  const[logEntries,setLog]=useState([]);
  const[hashes,setHashes]=useState([]);
  const[customers,setCustomers]=useState([]);
  const[agentNotes,setAgentNotes]=useState([]);
  const[loading,setLoading]=useState(true);
  const[modal,setModal]=useState(null);
  const[mdata,setMdata]=useState({});
  const[chatMsgs,setChatMsgs]=useState([]);
  const[chatInput,setChatInput]=useState('');
  const[chatLoading,setChatLoading]=useState(false);
  const chatEndRef=useRef(null);
  const t=T[lang];

  useEffect(()=>{loadAll();},[]);
  useEffect(()=>{if(chatMsgs.length===0)setChatMsgs([{role:'assistant',content:T[lang].chatPlaceholder}]);},[lang]);
  useEffect(()=>{chatEndRef.current?.scrollIntoView({behavior:'smooth'});},[chatMsgs]);

  async function loadAll(){
    setLoading(true);
    try{
      const[{data:p},{data:l},{data:h},{data:c},{data:n}]=await Promise.all([
        supabase.from('products').select('*').order('name'),
        supabase.from('change_log').select('*').order('created_at',{ascending:false}).limit(200),
        supabase.from('uploaded_files').select('file_hash'),
        supabase.from('customer_sales').select('*,customer_sale_items(*)').order('created_at',{ascending:false}).limit(100),
        supabase.from('agent_notes').select('*').order('created_at',{ascending:false}),
      ]);
      if(p)setProds(p);
      if(l)setLog(l);
      if(h)setHashes(h.map(x=>x.file_hash));
      if(c)setCustomers(c);
      if(n)setAgentNotes(n);
    }catch(e){console.error(e);}
    setLoading(false);
    setChatMsgs([{role:'assistant',content:`Hi! I'm Claude, your BSL inventory manager. I have full access to your inventory, customer sales, and notes. Ask me anything — stock levels, reorder suggestions, sales trends, or upload a packing list and I'll process it automatically.`}]);
  }

  async function saveProduct(form){
    const entry={...form,stock:parseFloat(form.stock)||0,velocity:parseFloat(form.velocity)||0,cost:parseFloat(form.cost)||0,price:parseFloat(form.price)||0,reorder:parseFloat(form.reorder)||0,amz_pack_size:parseFloat(form.amz_pack_size)||1,wmt_pack_size:parseFloat(form.wmt_pack_size)||1,tgt_pack_size:parseFloat(form.tgt_pack_size)||1,temu_pack_size:parseFloat(form.temu_pack_size)||1,other_pack_size:parseFloat(form.other_pack_size)||1};
    delete entry.id;
    if(form.id){
      const old=prods.find(p=>p.id===form.id);
      await supabase.from('products').update(entry).eq('id',form.id);
      const diff=(parseFloat(form.stock)||0)-old.stock;
      await supabase.from('change_log').insert({description:diff!==0?`Stock: ${form.name} ${old.stock}→${form.stock}`:`Updated: ${form.name}`,qty_change:diff||null});
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
    await supabase.from('change_log').insert({description:`Deleted: ${p.name}`});
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
        parsed.push({name:o.name,sku:o.sku||'',category:o.category||'',stock:parseFloat(o.stock)||0,velocity:parseFloat(o.velocity)||0,cost:parseFloat(o.cost)||0,price:parseFloat(o.price)||0,reorder:parseFloat(o.reorder)||0,supplier:o.supplier||'',amz:o.amz||'',wmt:o.wmt||'',tgt:o.tgt||'',temu:o.temu||'',other_sku:o.other_sku||'',amz_pack_size:parseFloat(o.amz_pack_size)||1,wmt_pack_size:parseFloat(o.wmt_pack_size)||1,tgt_pack_size:parseFloat(o.tgt_pack_size)||1,temu_pack_size:parseFloat(o.temu_pack_size)||1,other_pack_size:parseFloat(o.other_pack_size)||1});
      });
      setMdata(prev=>({...prev,step:2,parsed,errors}));
    });
    e.target.value='';
  }

  async function confirmImport(){
    const{parsed,mode}=mdata;
    if(mode==='replace'){await supabase.from('products').delete().neq('id',0);await supabase.from('products').insert(parsed);}
    else{for(const p of parsed){const existing=prods.find(x=>x.sku&&x.sku===p.sku);if(existing)await supabase.from('products').update(p).eq('id',existing.id);else await supabase.from('products').insert(p);}}
    await supabase.from('change_log').insert({description:`Bulk import: ${parsed.length} products (${mode||'add'})`,qty_change:parsed.length});
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

  async function applyOrders(){
    const platLabel=PLAT[mdata.platform]?.l;
    for(const item of mdata.preview){
      if(!item.prod)continue;
      const newStock=item.prod.stock-item.singlesDeducted;
      await supabase.from('products').update({stock:newStock}).eq('id',item.prod.id);
      await supabase.from('orders_log').insert({platform:mdata.platform,order_sku:item.sku,product_id:item.prod.id,product_name:item.prod.name,qty_sold:item.singlesDeducted,stock_before:item.prod.stock,stock_after:newStock});
      await supabase.from('change_log').insert({description:`${platLabel}: ${item.prod.name} [${item.sku}] ${item.packsOrdered} packs × ${item.packSize} = ${item.singlesDeducted} singles. ${item.prod.stock}→${newStock}`,qty_change:-item.singlesDeducted,platform:mdata.platform});
    }
    await supabase.from('change_log').insert({description:`${platLabel} upload complete`});
    if(mdata.hash)await supabase.from('uploaded_files').insert({file_hash:mdata.hash});
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
        if(!item.matchedSku)continue;
        const prod=prods.find(p=>p.sku===item.matchedSku);
        if(!prod)continue;
        const newStock=prod.stock-item.singles;
        await supabase.from('products').update({stock:newStock}).eq('id',prod.id);
        await supabase.from('customer_sale_items').insert({sale_id:sale.id,product_id:prod.id,product_name:prod.name,root_sku:prod.sku,pack_size:item.packSize||1,packs_ordered:item.totalUnits/(item.packSize||1),singles_deducted:item.singles});
        await supabase.from('change_log').insert({description:`Direct sale to ${parsed.customer||customerName}: ${prod.name} -${item.singles} singles`,qty_change:-item.singles,platform:'direct'});
      }
    }else{
      // Inbound — add stock
      for(const item of(parsed.items||[])){
        if(!item.matchedSku)continue;
        const prod=prods.find(p=>p.sku===item.matchedSku);
        if(!prod)continue;
        const newStock=prod.stock+item.singles;
        await supabase.from('products').update({stock:newStock}).eq('id',prod.id);
        await supabase.from('change_log').insert({description:`Received from ${parsed.supplier}: ${prod.name} +${item.singles} singles`,qty_change:item.singles,platform:'inbound'});
      }
    }
    await loadAll();setModal(null);
  }

  // ── EXPORT CSV ──────────────────────────────────────────────
  function exportCSV(){
    const headers=['Name','Root SKU','Category','Stock (Singles)','Monthly Sales','Cost','Price','Reorder','Supplier','Amazon SKU','AMZ Pack','Walmart SKU','WMT Pack','Target SKU','TGT Pack','Temu SKU','Temu Pack','Other SKU','Other Pack','Status','Days Left'];
    const rows=prods.map(p=>[p.name,p.sku,p.category,p.stock,p.velocity,p.cost,p.price,p.reorder,p.supplier,p.amz,p.amz_pack_size||1,p.wmt,p.wmt_pack_size||1,p.tgt,p.tgt_pack_size||1,p.temu,p.temu_pack_size||1,p.other_sku,p.other_pack_size||1,gs(p),gd(p)??'—'].map(v=>`"${String(v??'').replace(/"/g,'""')}"`).join(','));
    const csv=[headers.join(','),...rows].join('\n');
    const a=document.createElement('a');a.href='data:text/csv;charset=utf-8,'+encodeURIComponent(csv);a.download='bsl_inventory_'+new Date().toISOString().slice(0,10)+'.csv';a.click();
  }

  function downloadTemplate(){
    const h=['Root SKU','Product Name','Category','Supplier','Stock (Singles)','Monthly Sales (Singles)','Cost Per Single ($)','Selling Price Single ($)','Reorder Point (Singles)','Amazon SKU','AMZ Pack Size','AMZ Selling Price ($)','Walmart SKU','WMT Pack Size','WMT Selling Price ($)','Target SKU','TGT Pack Size','TGT Selling Price ($)','Temu SKU','Temu Pack Size','Temu Selling Price ($)','Other SKU','Other Pack Size','Other Selling Price ($)','Product Notes'];
    const ex=['BSL-LACT-150','Lactose Free 1.5lb','Lactose Free','EVI Labs','0','0','0.00','0.00','0','AMZ-LF150-001','6','0.00','WMT-LF150-001','6','0.00','','1','0.00','','1','0.00','','1','0.00','EVI Labs product'];
    const csv=[h,ex].map(r=>r.map(v=>`"${v}"`).join(',')).join('\n');
    const a=document.createElement('a');a.href='data:text/csv;charset=utf-8,'+encodeURIComponent(csv);a.download='BSL_SKU_Template.csv';a.click();
  }

  // ── AI CHAT ─────────────────────────────────────────────────
  async function sendChat(e){
    e.preventDefault();
    if(!chatInput.trim()||chatLoading)return;
    const userMsg=chatInput.trim();
    setChatInput('');
    const newMsgs=[...chatMsgs,{role:'user',content:userMsg}];
    setChatMsgs(newMsgs);
    setChatLoading(true);
    try{
      const notesContext=agentNotes.length?`\nPermanent notes & instructions:\n${agentNotes.map(n=>`[${n.category}] ${n.title}: ${n.content}`).join('\n')}`:'';
      const inventoryContext=`Current inventory (${prods.length} products, all quantities in SINGLES):\n${prods.map(p=>`- ${p.name} | Root SKU: ${p.sku} | Stock: ${p.stock} singles | Velocity: ${p.velocity}/mo | Cost: $${p.cost} | Price: $${p.price} | Reorder at: ${p.reorder} | Status: ${gs(p)} | Days left: ${gd(p)??'N/A'} | Supplier: ${p.supplier||'—'} | AMZ: ${p.amz||'—'} (${p.amz_pack_size||1}/pack) | WMT: ${p.wmt||'—'} (${p.wmt_pack_size||1}/pack) | TGT: ${p.tgt||'—'} | Temu: ${p.temu||'—'}`).join('\n')}`;
      const recentLog=`\nRecent changes:\n${logEntries.slice(0,15).map(l=>`- ${new Date(l.created_at).toLocaleDateString()}: ${l.description}`).join('\n')}`;
      const customerContext=customers.length?`\nRecent customer sales (${customers.length} total):\n${customers.slice(0,5).map(c=>`- ${c.customer_name} | ${new Date(c.created_at).toLocaleDateString()} | ${c.total_singles} singles | ${c.platform}`).join('\n')}`:'';

      const res=await fetch('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({
        model:'claude-sonnet-4-20250514',
        max_tokens:1000,
        system:`You are Claude, the inventory manager for BSL (Blooming Sweet Life Corp), a multi-channel e-commerce business. CRITICAL RULES: 1) All inventory is tracked in SINGLES. 2) Pack sizes vary by platform — always multiply packs × pack size to get singles. 3) When suggesting reorders, consider velocity and days of supply. 4) Products come from different suppliers — EVI Labs handles lactose free milk products. 5) Direct sales go through packing lists (like EVI Labs). 6) Marketplace sales go through Amazon FBA and Walmart WFS. Respond in ${lang==='es'?'Spanish':'English'}.${notesContext}\n\n${inventoryContext}${recentLog}${customerContext}`,
        messages:newMsgs.filter(m=>m.role!=='system').map(m=>({role:m.role,content:m.content})),
      })});
      const data=await res.json();
      const reply=data.content?.[0]?.text||'Sorry, I could not process that.';
      setChatMsgs(prev=>[...prev,{role:'assistant',content:reply}]);
    }catch(err){setChatMsgs(prev=>[...prev,{role:'assistant',content:'Error connecting to Claude API.'}]);}
    setChatLoading(false);
  }

  // ── COMPUTED ─────────────────────────────────────────────────
  const totalVal=prods.reduce((a,p)=>a+(p.stock*(p.cost||0)),0);
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
  const ProdTable=({list})=>{
    if(!list.length)return<div style={{textAlign:'center',padding:'2rem',color:'#aaa',fontSize:13}}>{tab==='alerts'?t.noAlerts:t.noProducts}</div>;
    const SortTh=({col,children})=><th style={{...S.th,cursor:'pointer',userSelect:'none'}} onClick={()=>toggleSort(col)}>{children} <span style={{color:'#4a90e2'}}>{sortIcon(col)}</span></th>;
    return(
      <div style={{overflowX:'auto',border:'1px solid #eee',borderRadius:12}}>
        <table style={{width:'100%',borderCollapse:'collapse',fontSize:12}}>
          <thead><tr>
            <SortTh col="status">{t.status}</SortTh>
            <SortTh col="name">{t.product}</SortTh>
            <th style={S.th}>{t.rootSku}</th>
            <SortTh col="stock">{t.stock}</SortTh>
            <SortTh col="days">{t.days}</SortTh>
            <SortTh col="cost">{t.cost}</SortTh>
            <SortTh col="price">{t.price}</SortTh>
            <th style={S.th}>{t.actions}</th>
          </tr></thead>
          <tbody>
            {list.map(p=>{const st=gs(p),dl=gd(p);return(
              <tr key={p.id}>
                <td style={S.td}><Badge status={st} t={t}/></td>
                <td style={{...S.td,fontWeight:500,maxWidth:160}} title={p.name}>{p.name}</td>
                <td style={S.td}><code style={{fontSize:10,background:'#f5f5f5',padding:'1px 5px',borderRadius:4}}>{p.sku||'—'}</code></td>
                <td style={{...S.td,color:p.stock<0?'#dc3545':undefined}}>{p.stock}</td>
                <td style={{...S.td,color:st==='crit'?'#dc3545':st==='low'?'#856404':undefined}}>{dl!=null?`${dl}d`:'—'}</td>
                <td style={S.td}>{fm(p.cost)}</td>
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
          {[['dashboard',t.dashboard],['chat',t.chat],['customers',t.customers],['reports',t.reports],['notes',t.notes]].map(([k,l])=>(
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
            {[['all',t.allProducts],['alerts',`${t.alerts} (${alerts.length})`],['channels',t.channels],['log',t.log]].map(([k,l])=>(
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
                    {l.qty_change!=null&&<span style={{fontWeight:600,color:l.qty_change<0?'#dc3545':'#28a745',flexShrink:0}}>{l.qty_change>0?'+':''}{l.qty_change}</span>}
                  </div>
                ))}
              </div>
            )}
          </div>
        </>}

        {/* CHAT */}
        {!loading&&page==='chat'&&(
          <div style={{...S.card,display:'flex',flexDirection:'column',height:'calc(100vh - 120px)'}}>
            <div style={{fontSize:16,fontWeight:600,marginBottom:'1rem',paddingBottom:'0.75rem',borderBottom:'1px solid #eee'}}>{t.chat}</div>
            <div style={{flex:1,overflowY:'auto',display:'flex',flexDirection:'column',gap:12,paddingBottom:'1rem'}}>
              {chatMsgs.map((m,i)=>(
                <div key={i} style={{display:'flex',justifyContent:m.role==='user'?'flex-end':'flex-start'}}>
                  <div style={{maxWidth:'78%',padding:'10px 14px',borderRadius:12,background:m.role==='user'?'#111':'#f0f0f0',color:m.role==='user'?'#fff':'#111',fontSize:13,lineHeight:1.5,whiteSpace:'pre-wrap'}}>{m.content}</div>
                </div>
              ))}
              {chatLoading&&<div style={{display:'flex',justifyContent:'flex-start'}}><div style={{padding:'10px 14px',borderRadius:12,background:'#f0f0f0',fontSize:13,color:'#888'}}>{t.thinking}</div></div>}
              <div ref={chatEndRef}/>
            </div>
            <form onSubmit={sendChat} style={{display:'flex',gap:8,paddingTop:'0.75rem',borderTop:'1px solid #eee'}}>
              <input style={{...S.inp,flex:1}} value={chatInput} onChange={e=>setChatInput(e.target.value)} placeholder={t.chatPlaceholder} disabled={chatLoading}/>
              <button type="submit" style={{...S.btnP,padding:'8px 16px'}} disabled={chatLoading||!chatInput.trim()}>→</button>
            </form>
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
              <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem'}}>📈 {t.velocityReport}</div>
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
      </main>

      {/* MODALS */}
      {modal==='product'&&<ProductModal t={t} S={S} mdata={mdata} setMdata={setMdata} onSave={saveProduct} onClose={()=>setModal(null)}/>}
      {modal==='import'&&<ImportModal t={t} S={S} mdata={mdata} setMdata={setMdata} onFile={handleImpFile} onConfirm={confirmImport} onDownload={downloadTemplate} onClose={()=>setModal(null)}/>}
      {modal==='orders'&&<OrdersModal t={t} S={S} mdata={mdata} setMdata={setMdata} onFile={handleOrdFile} onPreview={buildPreview} onConfirm={confirmOrders} onApply={applyOrders} hashes={hashes} onClose={()=>setModal(null)}/>}
      {modal==='packing'&&<PackingModal t={t} S={S} mdata={mdata} setMdata={setMdata} onFile={handlePackingFile} onApply={applyPackingList} onClose={()=>setModal(null)}/>}
      {modal==='note'&&<NoteModal t={t} S={S} mdata={mdata} setMdata={setMdata} onSave={async(form)=>{if(form.id)await supabase.from('agent_notes').update(form).eq('id',form.id);else await supabase.from('agent_notes').insert(form);await loadAll();setModal(null);}} onClose={()=>setModal(null)}/>}
    </div>
  );
}

// ── PRODUCT MODAL ─────────────────────────────────────────────
function ProductModal({t,S,mdata,setMdata,onSave,onClose}){
  const[form,setForm]=useState(mdata.form||{id:null,name:'',sku:'',category:'',stock:'',velocity:'',cost:'',price:'',reorder:'',supplier:'',amz:'',wmt:'',tgt:'',temu:'',other_sku:'',amz_pack_size:1,wmt_pack_size:1,tgt_pack_size:1,temu_pack_size:1,other_pack_size:1});
  const isEdit=!!form.id;
  return(
    <div style={S.overlay} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={S.sheet}>
        <div style={{fontSize:15,fontWeight:600,marginBottom:'1rem'}}>{isEdit?t.edit:t.addProduct}</div>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
          {[{l:t.productName,k:'name',full:true,ph:'Lactose Free 1.5lb'},{l:'Root SKU',k:'sku',ph:'BSL-LACT-150'},{l:t.category,k:'category',ph:'Lactose Free'},{l:`${t.stock}`,k:'stock',t:'number',ph:'0'},{l:t.monthlyS,k:'velocity',t:'number',ph:'0'},{l:`${t.cost} ($)`,k:'cost',t:'number',ph:'0.00'},{l:`${t.price} ($)`,k:'price',t:'number',ph:'0.00'},{l:t.reorder,k:'reorder',t:'number',ph:'0'},{l:t.supplier,k:'supplier',full:true,ph:'EVI Labs'}].map(f=>(
            <div key={f.k} style={{gridColumn:f.full?'1/-1':undefined,display:'flex',flexDirection:'column',gap:3}}>
              <label style={{fontSize:11,color:'#888',fontWeight:500}}>{f.l}</label>
              <input style={S.inp} type={f.t||'text'} placeholder={f.ph} value={form[f.k]||''} onChange={e=>setForm(prev=>({...prev,[f.k]:e.target.value}))}/>
            </div>
          ))}
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
function PackingModal({t,S,mdata,setMdata,onFile,onApply,onClose}){
  const pRef=useRef();
  const step=mdata.step||1;
  const parsed=mdata.parsed||{};
  return(
    <div style={S.overlay} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={S.sheet}>
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
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10,marginBottom:'1rem'}}>
            <div style={{background:'#f8f8f8',borderRadius:8,padding:'10px 12px',fontSize:12}}>
              <div style={{color:'#888',marginBottom:2}}>Supplier</div>
              <div style={{fontWeight:600}}>{parsed.supplier||'—'}</div>
            </div>
            <div style={{background:'#f8f8f8',borderRadius:8,padding:'10px 12px',fontSize:12}}>
              <div style={{color:'#888',marginBottom:2}}>Customer / Ship To</div>
              <div style={{fontWeight:600}}>{parsed.customer||'—'}</div>
            </div>
            <div style={{background:'#f8f8f8',borderRadius:8,padding:'10px 12px',fontSize:12}}>
              <div style={{color:'#888',marginBottom:2}}>Date</div>
              <div style={{fontWeight:600}}>{parsed.date||'—'}</div>
            </div>
            <div style={{background:parsed.direction==='outbound'?'#FFF3CD':'#D4EDDA',borderRadius:8,padding:'10px 12px',fontSize:12}}>
              <div style={{color:'#888',marginBottom:2}}>Direction</div>
              <div style={{fontWeight:600,color:parsed.direction==='outbound'?'#856404':'#155724'}}>{parsed.direction==='outbound'?'📤 Outbound (deduct)':'📥 Inbound (add)'}</div>
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

          <div style={{fontSize:12,fontWeight:500,color:'#555',marginBottom:6}}>Items found ({(parsed.items||[]).length})</div>
          <div style={{overflowX:'auto',border:'1px solid #eee',borderRadius:10,marginBottom:'1rem',maxHeight:220,overflowY:'auto'}}>
            <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
              <thead><tr>{['Description','Singles','Matched SKU','Confidence'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
              <tbody>{(parsed.items||[]).map((item,i)=>(
                <tr key={i} style={{background:!item.matchedSku?'#FFF5F5':''}}>
                  <td style={{...S.td,maxWidth:200}} title={item.description}>{item.description}</td>
                  <td style={{...S.td,fontWeight:600,color:parsed.direction==='outbound'?'#dc3545':'#28a745'}}>{parsed.direction==='outbound'?'-':'+' }{item.singles}</td>
                  <td style={S.td}>{item.matchedSku?<code style={{fontSize:10,background:'#f5f5f5',padding:'1px 4px',borderRadius:3}}>{item.matchedSku}</code>:<span style={{color:'#dc3545'}}>No match</span>}</td>
                  <td style={S.td}><span style={{background:item.confidence==='high'?'#D4EDDA':item.confidence==='medium'?'#FFF3CD':'#F8D7DA',color:item.confidence==='high'?'#155724':item.confidence==='medium'?'#856404':'#721C24',padding:'1px 6px',borderRadius:99,fontSize:10}}>{item.confidence||'—'}</span></td>
                </tr>
              ))}</tbody>
            </table>
          </div>
          <div style={{fontSize:12,fontWeight:600,color:'#555',marginBottom:'1rem'}}>Total singles: {parsed.totalSingles||0}</div>
          <div style={{display:'flex',gap:8,justifyContent:'space-between'}}>
            <button style={S.btn} onClick={()=>setMdata(prev=>({...prev,step:1,parsed:null}))}>← Back</button>
            <div style={{display:'flex',gap:8}}>
              <button style={S.btn} onClick={onClose}>{t.cancel}</button>
              <button style={{...S.btn,background:'#28a745',color:'#fff',border:'none'}} onClick={onApply}>{t.applyDeductions}</button>
            </div>
          </div>
        </>}
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
