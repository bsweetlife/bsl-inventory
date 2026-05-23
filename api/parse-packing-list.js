export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }
  try {
    const { imageBase64, mediaType, products } = req.body;
    const productList = products.map(p => 
      `- ${p.name} | Root SKU: ${p.sku} | AMZ Pack: ${p.amz_pack_size||1} | WMT Pack: ${p.wmt_pack_size||1}`
    ).join('\n');

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.REACT_APP_ANTHROPIC_KEY,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 2000,
        messages: [{
          role: 'user',
          content: [
            {
              type: 'image',
              source: { type: 'base64', media_type: mediaType, data: imageBase64 }
            },
            {
              type: 'text',
              text: `You are reading a packing list or packing slip document. Extract all line items and match them to our product inventory.

Our products:
${productList}

Extract and return ONLY a JSON object with this exact structure, no other text:
{
  "supplier": "supplier name from document",
  "customer": "ship to name if present",
  "packingListNumber": "packing list number if present",
  "poNumber": "PO number if present",
  "date": "date from document",
  "direction": "outbound if shipping TO a customer, inbound if receiving FROM supplier",
  "items": [
    {
      "description": "product description from document",
      "totalUnits": 120,
      "packSize": 12,
      "singles": 120,
      "matchedSku": "BSL-LACT-150 or null if no match",
      "matchedName": "matched product name or null",
      "confidence": "high/medium/low",
      "lotNumber": "lot number if present"
    }
  ],
  "totalSingles": 240,
  "notes": "any other relevant notes from the document"
}`
            }
          ]
        }]
      }),
    });
    const data = await response.json();
    const text = data.content?.[0]?.text || '{}';
    try {
      const clean = text.replace(/```json|```/g, '').trim();
      const parsed = JSON.parse(clean);
      res.status(200).json(parsed);
    } catch {
      res.status(200).json({ error: 'Could not parse document', raw: text });
    }
  } catch (error) {
    res.status(500).json({ error: 'Failed to process packing list' });
  }
}
