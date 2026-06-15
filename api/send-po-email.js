// api/send-po-email.js
// Sends a purchase order email to a location's contact via Resend.
// Requires env var RESEND_API_KEY and EMAIL_FROM in Vercel project settings.

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { po, location, contact } = req.body || {};

  if (!po || !location || !contact?.email) {
    return res.status(400).json({ error: 'Missing po, location, or contact email' });
  }

  const RESEND_API_KEY = process.env.RESEND_API_KEY;
  const EMAIL_FROM = process.env.EMAIL_FROM || 'BSL Inventory <onboarding@resend.dev>';

  if (!RESEND_API_KEY) {
    return res.status(500).json({ error: 'RESEND_API_KEY not configured' });
  }

  const items = po.items || [];
  const totalQty = items.reduce((a, i) => a + (parseFloat(i.qty) || 0), 0);
  const totalCost = items.reduce((a, i) => a + (parseFloat(i.qty) || 0) * (parseFloat(i.unit_cost) || 0), 0);

  const rows = items.map((i, idx) => `
    <tr style="background:${idx % 2 ? '#fafafa' : '#fff'}">
      <td style="padding:8px 12px;border-bottom:1px solid #eee">${i.product_name || '—'}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right">${i.qty} singles</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right">${i.unit_cost ? '$' + parseFloat(i.unit_cost).toFixed(2) : '—'}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right">${i.unit_cost ? '$' + ((parseFloat(i.qty) || 0) * parseFloat(i.unit_cost)).toFixed(2) : '—'}</td>
    </tr>`).join('');

  const html = `<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#f5f7fa;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif">
  <div style="max-width:640px;margin:0 auto;padding:24px 16px">
    <div style="background:linear-gradient(135deg,#1a1a2e 0%,#16213e 100%);border-radius:12px;padding:20px 24px;margin-bottom:20px;color:#fff">
      <div style="font-size:20px;font-weight:800">🌸 Purchase Order #${po.po_number}</div>
      <div style="font-size:13px;color:#aaa;margin-top:4px">Sent to: ${location}${contact.contact_name ? ' — ' + contact.contact_name : ''}</div>
    </div>

    <div style="background:#fff;border-radius:12px;padding:20px 24px;margin-bottom:16px;box-shadow:0 1px 4px rgba(0,0,0,.06)">
      <table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px;margin-bottom:12px">
        <tr><td style="color:#888;padding:4px 0">Supplier</td><td style="font-weight:600;padding:4px 0">${po.supplier || '—'}</td></tr>
        <tr><td style="color:#888;padding:4px 0">Expected Date</td><td style="font-weight:600;padding:4px 0">${po.expected_date || '—'}</td></tr>
        ${po.notes ? `<tr><td style="color:#888;padding:4px 0">Notes</td><td style="padding:4px 0">${po.notes}</td></tr>` : ''}
      </table>

      <table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px">
        <thead><tr style="background:#f0f0f0">
          <th style="padding:8px 12px;text-align:left;font-size:11px;color:#888">Product</th>
          <th style="padding:8px 12px;text-align:right;font-size:11px;color:#888">Qty</th>
          <th style="padding:8px 12px;text-align:right;font-size:11px;color:#888">Unit Cost</th>
          <th style="padding:8px 12px;text-align:right;font-size:11px;color:#888">Subtotal</th>
        </tr></thead>
        <tbody>${rows}</tbody>
        <tfoot><tr style="background:#f5f5f5;font-weight:700">
          <td style="padding:8px 12px">TOTAL</td>
          <td style="padding:8px 12px;text-align:right">${totalQty} singles</td>
          <td></td>
          <td style="padding:8px 12px;text-align:right;color:#1565c0">$${totalCost.toFixed(2)}</td>
        </tr></tfoot>
      </table>
    </div>

    <div style="text-align:center;padding:16px;color:#bbb;font-size:11px">
      BSL Inventory Management · Purchase Order Notification
    </div>
  </div>
</body>
</html>`;

  try {
    const resendRes = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${RESEND_API_KEY}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        from: EMAIL_FROM,
        to: [contact.email],
        subject: `Purchase Order #${po.po_number} — ${location}`,
        html,
      }),
    });

    const data = await resendRes.json();

    if (!resendRes.ok) {
      return res.status(resendRes.status).json({ error: data });
    }

    return res.status(200).json({ success: true, id: data.id });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
