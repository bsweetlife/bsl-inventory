// BSL Daily Inventory Summary — daily-summary.js
// Runs via GitHub Actions at 6:00 AM EST every day
// Requires env vars: SUPABASE_URL, SUPABASE_SERVICE_KEY, RESEND_API_KEY, EMAIL_TO, EMAIL_FROM

import { createClient } from '@supabase/supabase-js';
import { Resend } from 'resend';
import ws from 'ws';

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY,
  { realtime: { transport: ws } }
);
const resend = new Resend(process.env.RESEND_API_KEY);

// ─── Helpers ────────────────────────────────────────────────
const fm = n => n != null ? '$' + parseFloat(n).toFixed(2) : '—';
const OZ_PER_KG = 35.274;

function calcCost(p, gs = {}) {
  const rmWaste  = gs.raw_material_waste_pct  ?? 0.005;
  const pkgWaste = gs.packaging_waste_pct     ?? 0.005;
  const filling  = gs.filling_cost            ?? 1.15;
  const defPkg   = gs.packaging_cost_default  ?? 0.45;

  if (p.product_type !== 'packaged') return parseFloat(p.cost) || 0;
  const pricePerKg = parseFloat(p.raw_material_cost_per_kg) || 0;
  const weightOz   = parseFloat(p.weight_oz) || 0;
  if (pricePerKg <= 0 || weightOz <= 0) return parseFloat(p.cost) || 0;

  const pricePerOz = pricePerKg / OZ_PER_KG;
  const rawMat     = pricePerOz * weightOz;
  const pkgCost    = (parseFloat(p.packaging_cost) || defPkg) * (1 + pkgWaste);
  const boxCost    = parseFloat(p.box_cost) || 0;
  return rawMat * (1 + rmWaste) + pkgCost + filling + boxCost;
}

function statusLabel(p) {
  const daily = (p.velocity || 0) / 30;
  if (!daily) return { label: 'OK', color: '#28a745', emoji: '✅' };
  const days = p.stock / daily;
  if (days <= 15) return { label: 'CRITICAL', color: '#dc3545', emoji: '🔴' };
  if (days <= 30) return { label: 'LOW',      color: '#e67e22', emoji: '🟡' };
  return                { label: 'OK',        color: '#28a745', emoji: '✅' };
}

function daysOfStock(p) {
  const daily = (p.velocity || 0) / 30;
  return daily ? Math.round(p.stock / daily) : null;
}

// ─── Fetch all data ──────────────────────────────────────────
async function fetchAll() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const ystStr = yesterday.toISOString().split('T')[0];

  const [
    { data: products },
    { data: gsRows },
    { data: locations },
    { data: yesterdayOrders },
    { data: openPOs },
    { data: recentLog },
    { data: customerSales },
  ] = await Promise.all([
    supabase.from('products').select('*').order('name'),
    supabase.from('global_settings').select('*'),
    supabase.from('inventory_locations').select('*'),
    supabase.from('orders_log').select('*').gte('order_date', ystStr).lte('order_date', ystStr),
    supabase.from('purchase_orders').select('*,purchase_order_items(*)').in('status', ['ordered', 'in_transit']),
    supabase.from('change_log').select('*').gte('created_at', yesterday.toISOString()).order('created_at', { ascending: false }),
    supabase.from('customer_sales').select('*,customer_sale_items(*)').gte('created_at', yesterday.toISOString()),
  ]);

  const gs = {};
  (gsRows || []).forEach(r => gs[r.key] = parseFloat(r.value));

  return { products: products || [], gs, locations: locations || [], yesterdayOrders: yesterdayOrders || [], openPOs: openPOs || [], recentLog: recentLog || [], customerSales: customerSales || [] };
}

// ─── Build HTML email ────────────────────────────────────────
function buildEmail({ products, gs, locations, yesterdayOrders, openPOs, recentLog, customerSales }) {
  const now = new Date();
  const dateStr = now.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', timeZone: 'America/New_York' });

  // Totals
  const totalCostVal  = products.reduce((a, p) => a + (p.stock || 0) * calcCost(p, gs), 0);
  const totalSellVal  = products.reduce((a, p) => a + (p.stock || 0) * (parseFloat(p.price) || 0), 0);
  const totalUnits    = products.reduce((a, p) => a + (p.stock || 0), 0);
  const criticalProds = products.filter(p => statusLabel(p).label === 'CRITICAL');
  const lowProds      = products.filter(p => statusLabel(p).label === 'LOW');

  // Location drift detection
  const locationMap = {};
  locations.forEach(l => {
    if (!locationMap[l.product_id]) locationMap[l.product_id] = 0;
    locationMap[l.product_id] += parseFloat(l.qty) || 0;
  });
  const driftProds = products.filter(p => locationMap[p.id] !== undefined && Math.abs((locationMap[p.id] || 0) - (p.stock || 0)) > 0);

  // Yesterday sales by channel
  const salesByChannel = {};
  yesterdayOrders.forEach(o => {
    if (!salesByChannel[o.platform]) salesByChannel[o.platform] = { units: 0, orders: new Set() };
    salesByChannel[o.platform].units += parseFloat(o.qty) || 0;
    if (o.order_id) salesByChannel[o.platform].orders.add(o.order_id);
  });

  // PO summary
  const incomingUnits = openPOs.reduce((a, po) => (po.purchase_order_items || []).reduce((b, i) => b + (parseFloat(i.qty) || 0), a), 0);
  const incomingValue = openPOs.reduce((a, po) => (po.purchase_order_items || []).reduce((b, i) => b + (parseFloat(i.qty) || 0) * (parseFloat(i.unit_cost) || 0), a), 0);

  // Overdue POs
  const today = new Date().toISOString().split('T')[0];
  const overduePOs = openPOs.filter(po => po.expected_date && po.expected_date < today);

  // Direct customer sales
  const directSalesTotal = customerSales.reduce((a, s) => a + (parseFloat(s.total) || 0), 0);

  // Color palette
  const C = { bg: '#f5f7fa', card: '#ffffff', primary: '#1a1a2e', accent: '#4a90e2', green: '#28a745', red: '#dc3545', orange: '#e67e22', purple: '#7b2d8b' };

  const card = (content, borderColor = '#eee') =>
    `<div style="background:${C.card};border-radius:12px;padding:20px 24px;margin-bottom:16px;border:1px solid ${borderColor};box-shadow:0 1px 4px rgba(0,0,0,.06)">${content}</div>`;

  const badge = (text, color, bg) =>
    `<span style="background:${bg};color:${color};padding:2px 10px;border-radius:99px;font-size:11px;font-weight:700;display:inline-block">${text}</span>`;

  const sectionTitle = (emoji, title) =>
    `<div style="font-size:16px;font-weight:700;color:${C.primary};margin-bottom:14px;border-bottom:2px solid #f0f0f0;padding-bottom:8px">${emoji} ${title}</div>`;

  // ── Section 1: Headline stats ──
  const statsRow = `
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      ${[
        ['Total Products', products.length, C.primary],
        ['Total Units', totalUnits.toLocaleString(), C.accent],
        ['Inventory Cost', fm(totalCostVal), C.orange],
        ['Sellable Value', fm(totalSellVal), C.green],
      ].map(([l, v, c]) => `
        <td style="text-align:center;padding:12px;background:#f9f9f9;border-radius:10px;width:25%">
          <div style="font-size:11px;color:#888;margin-bottom:4px">${l}</div>
          <div style="font-size:22px;font-weight:800;color:${c}">${v}</div>
        </td>`).join('<td width="8"></td>')}
    </tr></table>`;

  // ── Section 2: Alerts ──
  let alertsHtml = '';
  if (criticalProds.length || lowProds.length || overduePOs.length || driftProds.length) {
    const rows = [
      ...criticalProds.map(p => `<tr><td style="padding:6px 8px">🔴 <strong>${p.name}</strong></td><td style="padding:6px 8px;color:${C.red};font-weight:700">${p.stock} singles left</td><td style="padding:6px 8px;color:#888">${daysOfStock(p) ?? '∞'} days</td><td style="padding:6px 8px">${badge('CRITICAL', C.red, '#fff0f0')}</td></tr>`),
      ...lowProds.map(p => `<tr><td style="padding:6px 8px">🟡 <strong>${p.name}</strong></td><td style="padding:6px 8px;color:${C.orange};font-weight:700">${p.stock} singles left</td><td style="padding:6px 8px;color:#888">${daysOfStock(p) ?? '∞'} days</td><td style="padding:6px 8px">${badge('LOW', C.orange, '#fff8f0')}</td></tr>`),
      ...overduePOs.map(po => `<tr><td style="padding:6px 8px">⏰ <strong>PO #${po.po_number}</strong></td><td style="padding:6px 8px;color:${C.red}">${po.supplier}</td><td style="padding:6px 8px;color:${C.red}">Expected ${po.expected_date}</td><td style="padding:6px 8px">${badge('OVERDUE', '#fff', C.red)}</td></tr>`),
      ...driftProds.map(p => `<tr><td style="padding:6px 8px">⚠️ <strong>${p.name}</strong></td><td style="padding:6px 8px;color:#856404">Location sum: ${locationMap[p.id] ?? 0} vs total: ${p.stock}</td><td style="padding:6px 8px;color:#888">Fix in app</td><td style="padding:6px 8px">${badge('DRIFT', '#856404', '#FFF3CD')}</td></tr>`),
    ];
    alertsHtml = card(`
      ${sectionTitle('🚨', 'Alerts Requiring Attention')}
      <table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px">
        <thead><tr style="background:#f5f5f5">
          <th style="padding:6px 8px;text-align:left;font-size:11px;color:#888">Product / PO</th>
          <th style="padding:6px 8px;text-align:left;font-size:11px;color:#888">Detail</th>
          <th style="padding:6px 8px;text-align:left;font-size:11px;color:#888">Runway</th>
          <th style="padding:6px 8px;text-align:left;font-size:11px;color:#888">Status</th>
        </tr></thead>
        <tbody>${rows.join('')}</tbody>
      </table>`, '#ffd3d3');
  }

  // ── Section 3: Yesterday sales ──
  const channelColors = { amazon: '#FF9900', walmart: '#0071DC', target: '#CC0000', temu: '#7B2D8B', other: '#888' };
  const channelKeys = Object.keys(salesByChannel);
  const totalYestUnits = Object.values(salesByChannel).reduce((a, v) => a + v.units, 0);
  const salesHtml = card(`
    ${sectionTitle('📊', `Yesterday's Sales (${new Date(yesterday()).toLocaleDateString('en-US', { month: 'short', day: 'numeric' })})`)}
    ${channelKeys.length === 0
      ? '<div style="color:#aaa;font-size:13px;text-align:center;padding:16px">No marketplace orders recorded yesterday</div>'
      : `<table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px;margin-bottom:12px">
          <thead><tr style="background:#f5f5f5">
            ${['Channel', 'Orders', 'Units Sold'].map(h => `<th style="padding:7px 10px;text-align:left;font-size:11px;color:#888">${h}</th>`).join('')}
          </tr></thead>
          <tbody>
            ${channelKeys.map(ch => `<tr>
              <td style="padding:7px 10px"><span style="color:${channelColors[ch]||'#888'};font-weight:700">${ch.toUpperCase()}</span></td>
              <td style="padding:7px 10px">${salesByChannel[ch].orders.size}</td>
              <td style="padding:7px 10px;font-weight:700">${salesByChannel[ch].units} singles</td>
            </tr>`).join('')}
            <tr style="background:#f9f9f9;font-weight:700">
              <td style="padding:7px 10px">TOTAL</td>
              <td style="padding:7px 10px">${Object.values(salesByChannel).reduce((a, v) => a + v.orders.size, 0)}</td>
              <td style="padding:7px 10px;color:${C.green}">${totalYestUnits} singles</td>
            </tr>
          </tbody>
        </table>`}
    ${directSalesTotal > 0 ? `<div style="font-size:12px;color:#888;padding-top:8px;border-top:1px solid #eee">Direct customer sales yesterday: <strong>${fm(directSalesTotal)}</strong></div>` : ''}`);

  // ── Section 4: Full inventory table ──
  const inventoryRows = products.map((p, i) => {
    const st = statusLabel(p);
    const cost = calcCost(p, gs);
    const days = daysOfStock(p);
    const totalCostP = (p.stock || 0) * cost;
    const totalSellP = (p.stock || 0) * (parseFloat(p.price) || 0);
    const locs = locations.filter(l => l.product_id === p.id);
    const locStr = locs.length
      ? locs.filter(l => l.qty > 0).map(l => `${l.location}: ${l.qty}`).join(' · ')
      : '—';
    return `<tr style="background:${i % 2 ? '#fafafa' : '#fff'}">
      <td style="padding:7px 8px;font-size:12px;font-weight:500">${st.emoji} ${p.name}</td>
      <td style="padding:7px 8px;font-size:11px;color:#888;font-family:monospace">${p.sku || '—'}</td>
      <td style="padding:7px 8px;font-size:12px;font-weight:700;color:${st.color}">${p.stock || 0}</td>
      <td style="padding:7px 8px;font-size:11px;color:#888">${days != null ? days + 'd' : '—'}</td>
      <td style="padding:7px 8px;font-size:12px">${fm(cost)}</td>
      <td style="padding:7px 8px;font-size:12px">${fm(p.price)}</td>
      <td style="padding:7px 8px;font-size:12px;color:#666">${fm(totalCostP)}</td>
      <td style="padding:7px 8px;font-size:12px;font-weight:600">${fm(totalSellP)}</td>
      <td style="padding:7px 8px;font-size:10px;color:#999">${locStr}</td>
    </tr>`;
  }).join('');

  const inventoryHtml = card(`
    ${sectionTitle('📦', 'Full Inventory Snapshot')}
    <div style="overflow-x:auto">
      <table width="100%" cellpadding="0" cellspacing="0" style="font-size:12px;min-width:700px">
        <thead><tr style="background:#f0f0f0">
          ${['Product', 'SKU', 'Stock', 'Days', 'Unit Cost', 'Price', 'Total Cost', 'Total Value', 'Locations'].map(h => `<th style="padding:7px 8px;text-align:left;font-size:10px;color:#888;white-space:nowrap">${h}</th>`).join('')}
        </tr></thead>
        <tbody>${inventoryRows}</tbody>
        <tfoot><tr style="background:#f5f5f5;font-weight:700">
          <td style="padding:7px 8px;font-size:12px" colspan="2">TOTAL</td>
          <td style="padding:7px 8px;font-size:12px">${totalUnits.toLocaleString()}</td>
          <td></td><td></td><td></td>
          <td style="padding:7px 8px;font-size:12px;color:${C.orange}">${fm(totalCostVal)}</td>
          <td style="padding:7px 8px;font-size:12px;color:${C.green}">${fm(totalSellVal)}</td>
          <td></td>
        </tr></tfoot>
      </table>
    </div>`);

  // ── Section 5: Open Purchase Orders ──
  let poHtml = '';
  if (openPOs.length) {
    const poRows = openPOs.map((po, i) => {
      const items = po.purchase_order_items || [];
      const qty = items.reduce((a, i) => a + (parseFloat(i.qty) || 0), 0);
      const val = items.reduce((a, i) => a + (parseFloat(i.qty) || 0) * (parseFloat(i.unit_cost) || 0), 0);
      const isOverdue = po.expected_date && po.expected_date < today;
      return `<tr style="background:${i % 2 ? '#fafafa' : '#fff'}">
        <td style="padding:7px 8px;font-size:12px;font-weight:600">#${po.po_number}</td>
        <td style="padding:7px 8px;font-size:12px">${po.supplier || '—'}</td>
        <td style="padding:7px 8px;font-size:12px">${badge(po.status === 'in_transit' ? '🚚 In Transit' : '📋 Ordered', po.status === 'in_transit' ? '#1565c0' : '#e67e22', po.status === 'in_transit' ? '#e3f2fd' : '#fff8f0')}</td>
        <td style="padding:7px 8px;font-size:12px;color:${isOverdue ? C.red : '#555'}">${po.expected_date || '—'}${isOverdue ? ' ⚠️' : ''}</td>
        <td style="padding:7px 8px;font-size:12px">${qty} singles</td>
        <td style="padding:7px 8px;font-size:12px;font-weight:600">${fm(val)}</td>
        <td style="padding:7px 8px;font-size:11px;color:#888">${items.map(i => `${i.product_name} (${i.qty})`).join(', ')}</td>
      </tr>`;
    }).join('');
    poHtml = card(`
      ${sectionTitle('🚢', `Open Purchase Orders — ${incomingUnits.toLocaleString()} singles / ${fm(incomingValue)} incoming`)}
      <table width="100%" cellpadding="0" cellspacing="0" style="font-size:12px">
        <thead><tr style="background:#f0f0f0">
          ${['PO #', 'Supplier', 'Status', 'Expected', 'Qty', 'Value', 'Products'].map(h => `<th style="padding:7px 8px;text-align:left;font-size:10px;color:#888">${h}</th>`).join('')}
        </tr></thead>
        <tbody>${poRows}</tbody>
      </table>`, '#d0e8ff');
  }

  // ── Section 6: Recent activity log ──
  const logHtml = recentLog.length ? card(`
    ${sectionTitle('📋', `Activity Yesterday (${recentLog.length} changes)`)}
    <div style="max-height:300px;overflow-y:auto">
      ${recentLog.slice(0, 30).map(l => `
        <div style="padding:6px 0;border-bottom:1px solid #f5f5f5;font-size:12px">
          <span style="color:#888;font-size:10px">${new Date(l.created_at).toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', timeZone: 'America/New_York' })}</span>
          <span style="margin-left:8px">${l.description || '—'}</span>
          ${l.qty_change != null && l.qty_change !== 0 ? `<span style="margin-left:8px;font-weight:700;color:${l.qty_change > 0 ? C.green : C.red}">${l.qty_change > 0 ? '+' : ''}${l.qty_change}</span>` : ''}
          ${l.user_email ? `<span style="margin-left:8px;color:#bbb;font-size:10px">${l.user_email}</span>` : ''}
        </div>`).join('')}
    </div>`) : '';

  function yesterday() {
    const d = new Date(); d.setDate(d.getDate() - 1); return d;
  }

  // ── Assemble ──
  return `<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:${C.bg};font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif">
  <div style="max-width:800px;margin:0 auto;padding:24px 16px">

    <!-- Header -->
    <div style="background:linear-gradient(135deg,#1a1a2e 0%,#16213e 100%);border-radius:14px;padding:24px 28px;margin-bottom:20px;color:#fff">
      <div style="font-size:22px;font-weight:800;margin-bottom:4px">🌸 BSL Inventory Daily Summary</div>
      <div style="font-size:13px;color:#aaa">${dateStr}</div>
      <div style="margin-top:16px">${statsRow}</div>
      ${criticalProds.length || lowProds.length
        ? `<div style="margin-top:14px;padding:10px 14px;background:rgba(220,53,69,.15);border:1px solid rgba(220,53,69,.3);border-radius:8px;font-size:12px;color:#ff8b96">
            ⚠️ ${criticalProds.length} critical · ${lowProds.length} low stock products need attention
           </div>`
        : `<div style="margin-top:14px;padding:10px 14px;background:rgba(40,167,69,.15);border:1px solid rgba(40,167,69,.3);border-radius:8px;font-size:12px;color:#7ed4a0">
            ✅ All products at healthy stock levels
           </div>`}
    </div>

    ${alertsHtml}
    ${salesHtml}
    ${inventoryHtml}
    ${poHtml}
    ${logHtml}

    <!-- Footer -->
    <div style="text-align:center;padding:20px;color:#bbb;font-size:11px">
      BSL Inventory Management · Auto-generated at 6:00 AM EST<br>
      <a href="https://bsl-inventory.vercel.app" style="color:#4a90e2">Open App</a>
    </div>
  </div>
</body>
</html>`;
}

// ─── Main ────────────────────────────────────────────────────
async function main() {
  console.log('Fetching BSL inventory data...');
  const data = await fetchAll();
  console.log(`Loaded ${data.products.length} products, ${data.yesterdayOrders.length} yesterday orders, ${data.openPOs.length} open POs`);

  const html = buildEmail(data);

  const critCount = data.products.filter(p => statusLabel(p).label === 'CRITICAL').length;
  const lowCount  = data.products.filter(p => statusLabel(p).label === 'LOW').length;
  const today = new Date().toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', timeZone: 'America/New_York' });

  let subjectPrefix = '✅';
  if (critCount > 0) subjectPrefix = '🔴';
  else if (lowCount > 0) subjectPrefix = '🟡';

  const subject = `${subjectPrefix} BSL Inventory Summary — ${today}${critCount ? ` · ${critCount} CRITICAL` : ''}${lowCount ? ` · ${lowCount} Low` : ''}`;

  console.log('Sending email:', subject);
  const { data: emailData, error } = await resend.emails.send({
    from: process.env.EMAIL_FROM,
    to: process.env.EMAIL_TO.split(',').map(e => e.trim()),
    subject,
    html,
  });

  if (error) {
    console.error('Email send failed:', error);
    process.exit(1);
  }
  console.log('Email sent successfully:', emailData?.id);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
