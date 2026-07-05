require('./load-env')(); // loads .env into process.env

const express = require('express');
const twilio = require('twilio');
const Groq = require('groq-sdk');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

const TWILIO_ACCOUNT_SID = process.env.TWILIO_ACCOUNT_SID;
const TWILIO_AUTH_TOKEN = process.env.TWILIO_AUTH_TOKEN;
const TWILIO_FROM_NUMBER = process.env.TWILIO_FROM_NUMBER;
const GROQ_API_KEY = process.env.GROQ_API_KEY;
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL;

const REQUIRED_ENV = { TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, TWILIO_FROM_NUMBER, GROQ_API_KEY, PUBLIC_BASE_URL };
const missing = Object.entries(REQUIRED_ENV).filter(([, v]) => !v).map(([k]) => k);
if (missing.length) {
  console.error('Missing required environment variables:', missing.join(', '));
  console.error('Copy .env.example to .env and fill in your real values.');
  process.exit(1);
}

const groq = new Groq({ apiKey: GROQ_API_KEY });
const twilioClient = twilio(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN);
const calls = {};
const pendingCalls = {};
const EXCEL_FILE = path.join(__dirname, 'orders.xlsx');

let orderCounter = 1;
if (fs.existsSync(EXCEL_FILE)) {
  try {
    const wb = XLSX.readFile(EXCEL_FILE);
    const ws = wb.Sheets['Orders'];
    if (ws) {
      const data = XLSX.utils.sheet_to_json(ws);
      if (data.length > 0) orderCounter = data.length + 1;
    }
  } catch(e) {}
}

let productData = [];
let productList = '';
try {
  const wb2 = XLSX.readFile('products.xlsx');
  const ws2 = wb2.Sheets['Products'];
  productData = XLSX.utils.sheet_to_json(ws2);
  // Purani products.xlsx mein Stock column nahi hoga to default 100 laga do,
  // taaki purani file ke saath bhi crash na ho.
  for (const p of productData) {
    if (p.Stock == null || isNaN(p.Stock)) p.Stock = 100;
  }
  // Har product ke saare variant prices group karo (size mention nahi karna,
  // sirf price options AI ko dena hai taaki customer ko "100, 300, 500 wala"
  // bata sake, "100ml, 200ml" nahi).
  const variantPrices = {};
  const variantStock = {};
  for (const p of productData) {
    if (!p.Product) continue;
    if (!variantPrices[p.Product]) variantPrices[p.Product] = [];
    if (p['Price (Rs)'] != null) variantPrices[p.Product].push(p['Price (Rs)']);
    variantStock[p.Product] = (variantStock[p.Product] || 0) + (p.Stock || 0);
  }
  // Har product ke naam ke saath uska total bacha hua stock bhi bhej rahe hain
  // (jaise "Dettol Soap:Rs45(Stock:80)"), taaki AI zyada stock waale products
  // ko upsell mein priority de sake aur stock jaldi khaali ho.
  productList = Object.entries(variantPrices)
    .map(([name, prices]) => `${name}:Rs${prices.join('/Rs')}(Stock:${variantStock[name] || 0})`)
    .join(', ');
  console.log('Products loaded:', productData.length, '| Unique:', Object.keys(variantPrices).length);
} catch(e) {
  console.log('Product file nahi mili:', e.message);
}

// productData (jisme har row ka Stock number hai) ko wapas products.xlsx mein
// likh deta hai. Order ke baad stock ghatane ke liye ye call hota hai, taaki
// website refresh karne par naya stock turant dikhe aur server restart hone
// par bhi stock wahi rahe jahan chhoda tha.
function saveProductsFile() {
  try {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(productData);
    ws['!cols'] = [{wch:5},{wch:25},{wch:15},{wch:10},{wch:12},{wch:10}];
    XLSX.utils.book_append_sheet(wb, ws, 'Products');
    XLSX.writeFile(wb, 'products.xlsx');
  } catch(e) {
    console.log('Stock save karne mein error:', e.message);
  }
}

// Parses an order string that may contain multiple items, e.g.
// "Parle G x2, Colgate x1" and returns a per-item breakdown plus a total.
// Also returns the matched product ROW itself (not just its price) so that
// stock can be decremented from the exact same variant that was priced —
// the pricing and the stock-deduction always agree on which row it was.
function getProductPrice(orderDetails) {
  if (!orderDetails) return { breakdown: '', total: 0, items: [] };

  // Split on commas/semicolons/"and" so each segment is (ideally) one item.
  const segments = orderDetails.split(/,|;| and /i).map(s => s.trim()).filter(Boolean);
  const segmentsToProcess = segments.length ? segments : [orderDetails];

  const lines = [];
  const items = [];
  let total = 0;

  for (const segment of segmentsToProcess) {
    const lowerSegment = segment.toLowerCase();
    let matched = null;
    for (const p of productData) {
      if (p.Product && lowerSegment.includes(p.Product.toLowerCase())) {
        matched = p;
        break;
      }
    }
    if (!matched) continue; // skip segments we can't match to a known product

    const qtyMatch = segment.match(/x\s*(\d+)/i) || segment.match(/(\d+)\s*x/i);
    const qty = qtyMatch ? parseInt(qtyMatch[1]) : 1;
    const rate = matched['Price (Rs)'] || 0;
    const lineTotal = rate * qty;

    lines.push(`${matched.Product} x${qty} @${rate} = ${lineTotal}`);
    total += lineTotal;
    items.push({ row: matched, qty });
  }

  return { breakdown: lines.join(' | '), total, items };
}

// Ghatata hai stock exact variant row se jo order mein match hua tha
// (Product AND Size dono match, sirf naam se nahi, kyunki row khud
// getProductPrice se pass hoti hai). 0 se neeche kabhi nahi jaayega.
// Aakhir mein products.xlsx file mein save kar deta hai taaki website
// refresh karte hi naya stock dikhe.
function decrementStock(items) {
  if (!items || !items.length) return;
  for (const item of items) {
    item.row.Stock = Math.max(0, (item.row.Stock || 0) - item.qty);
  }
  saveProductsFile();
}

function saveOrder(orderData) {
  let wb, ws, existingData = [];
  if (fs.existsSync(EXCEL_FILE)) {
    wb = XLSX.readFile(EXCEL_FILE);
    ws = wb.Sheets['Orders'];
    if (ws) existingData = XLSX.utils.sheet_to_json(ws);
  } else {
    wb = XLSX.utils.book_new();
  }
  const { breakdown, total, items } = getProductPrice(orderData.order);
  existingData.push({
    'Order No': `RT-${String(orderCounter++).padStart(4, '0')}`,
    'Date': new Date().toLocaleDateString('en-IN'),
    'Time': new Date().toLocaleTimeString('en-IN'),
    'Customer Name': orderData.naam,
    'Phone': orderData.phone,
    'Address': orderData.address,
    'Order Details': orderData.order,
    'Item Breakdown': breakdown,
    'Total Amount (Rs)': total,
    'Status': 'Confirmed',
    'Delivery': 'Pending'
  });
  ws = XLSX.utils.json_to_sheet(existingData);
  ws['!cols'] = [
    {wch:10},{wch:12},{wch:10},{wch:20},
    {wch:15},{wch:25},{wch:30},{wch:40},
    {wch:18},{wch:12},{wch:12}
  ];
  if (wb.SheetNames.includes('Orders')) { wb.Sheets['Orders'] = ws; }
  else { XLSX.utils.book_append_sheet(wb, ws, 'Orders'); }
  XLSX.writeFile(wb, EXCEL_FILE);
  console.log('Order saved! Breakdown:', breakdown, '| Total:', total);

  // Order confirm hote hi us exact variant ka stock kam kar do.
  decrementStock(items);
}

async function sendSMS(phone, naam, order) {
  try {
    await twilioClient.messages.create({
      body: `Namaste ${naam} ji! Aapka order confirm ho gaya.\nOrder: ${order}\nRajat Traders, Sagarpur Delhi.\nDhanyawad!`,
      from: TWILIO_FROM_NUMBER,
      to: phone
    });
    console.log('SMS sent!');
  } catch(e) {
    console.log('SMS error:', e.message);
  }
}

app.post('/call', async (req, res) => {
  const callSid = req.body.CallSid;
  const customerPhone = pendingCalls[callSid] || req.body.To;
  calls[callSid] = { history: [], callerNumber: customerPhone, startTime: new Date().toLocaleString('en-IN') };
  delete pendingCalls[callSid];
  console.log('Call started, customer phone:', customerPhone);
  const twiml = new twilio.twiml.VoiceResponse();
  const gather = twiml.gather({
    input: 'speech',
    action: '/respond',
    method: 'POST',
    language: 'hi-IN',
    speechTimeout: '2',
    timeout: '15'
  });
  gather.say({ language: 'hi-IN' }, 'Namaste! Rajat Traders mein aapka swagat hai. Kya chahiye aapko?');
  res.type('text/xml');
  res.send(twiml.toString());
});

app.post('/respond', async (req, res) => {
  const callSid = req.body.CallSid;
  const userSpeech = req.body.SpeechResult || '';
  const twiml = new twilio.twiml.VoiceResponse();
  if (!calls[callSid]) calls[callSid] = { history: [], callerNumber: req.body.To, startTime: new Date().toLocaleString('en-IN') };

  if (!userSpeech) {
    const gather = twiml.gather({ input: 'speech', action: '/respond', method: 'POST', language: 'hi-IN', speechTimeout: '2', timeout: '15' });
    gather.say({ language: 'hi-IN' }, 'Ji boliye?');
    res.type('text/xml');
    res.send(twiml.toString());
    return;
  }

  calls[callSid].history.push({ role: 'user', parts: [{ text: userSpeech }] });

  try {
    const systemPrompt = `You are Rajat Traders assistant. Sagarpur Delhi general store.
Available products (name:Rs price1/Rs price2/Rs price3 for different variants): ${productList}

STRICT RULES:
- ONLY offer or confirm products that appear in the list above. NEVER invent a product or a price that is not in the list.
- If customer asks for something NOT in the list, clearly say it's not available, then suggest 1-2 similar products from the list.
- Always talk in Rupees (Rs) only. NEVER mention ml, g, kg, L, or any size/weight unit, even if customer asks for "variants" or "options" — just list the different Rs prices.
  Example: if customer asks "Himalaya shampoo ke kaunse variant hain?", reply like "Rs75, Rs140, aur Rs230 wala hai, kaunsa chahiye?" — NOT "100ml, 200ml, 400ml".
- Always assume the cheapest/first price by default if customer doesn't specify. Do not ask which variant unless the customer asks or wants to choose.
- When confirming any item, state the price in the same sentence, e.g. "Patanjali Kesh Kanti liya, Rs60."
- SELLING IS YOUR TOP PRIORITY. After EVERY item the customer confirms, immediately suggest ONE more relevant complementary product from the list before moving on (e.g. after soap, suggest shampoo; after biscuits, suggest a cold drink).
- The list shows each product's remaining stock as (Stock:N). When choosing what to upsell, PREFER products with higher stock numbers — the goal is to sell as much of the high-stock inventory as possible during this call.
- If the customer declines one upsell suggestion, immediately try ONE different product instead of giving up completely. Only stop upselling after the customer has declined twice in a row, or clearly says they don't want anything else / are in a hurry.
- Before ending the call (CALL_END), always do one final check: ask "Aur kuch chahiye?" or similar, and if a call is likely to end soon, prioritize confirming any item the customer already showed interest in over starting a new pitch.
- NEVER let the call end while an item the customer has verbally agreed to ("haan", "theek hai", "de do") is still unconfirmed — immediately fold it into the order and update the SAVE| line before saying goodbye.
- ALWAYS reply ONLY in Roman script Hinglish (Hindi + English words written in English letters). NEVER output Devanagari (हिंदी) characters, not even one word.
- Example of CORRECT style: "Namaste! Aapko kya chahiye?"
- Example of WRONG style (do NOT do this): "नमस्ते! आपको क्या चाहिए?"
- Ask ONE question at a time only
- Order first, then name, then address
- Max 2 short sentences per reply
- Stay focused: do not repeat or re-ask things already confirmed
- When order+name+address collected, output on new line: SAVE|Name|Address|Product xQty
- If customer orders multiple items, separate them with commas in the same field, e.g. SAVE|Name|Address|Parle G x2, Colgate x1
- Name and address in English/Roman letters only (transliterate)
- Product name exactly as in list
- CALL_END only when customer says bye/shukriya/dhanyawad/thank you`;

    const messages = [
      { role: 'system', content: systemPrompt },
      ...calls[callSid].history.slice(0, -1).map(h => ({
        role: h.role === 'model' ? 'assistant' : 'user',
        content: h.parts[0].text
      })),
      { role: 'user', content: userSpeech }
    ];

    const result = await groq.chat.completions.create({
      model: 'llama-3.3-70b-versatile',
      messages: messages,
      max_tokens: 150
    });

    let aiResponse = result.choices[0].message.content;

    // Safety net: if the model slipped into Devanagari script, retry once with a firmer reminder.
    const devanagariRegex = /[\u0900-\u097F]/;
    if (devanagariRegex.test(aiResponse)) {
      const retryResult = await groq.chat.completions.create({
        model: 'llama-3.3-70b-versatile',
        messages: [
          ...messages,
          { role: 'assistant', content: aiResponse },
          { role: 'user', content: 'Reminder: reply ONLY in Roman/English letters (Hinglish), no Devanagari script at all. Rewrite your last reply in Roman letters only.' }
        ],
        max_tokens: 150
      });
      aiResponse = retryResult.choices[0].message.content;
    }

    calls[callSid].history.push({ role: 'model', parts: [{ text: aiResponse }] });
    console.log('AI:', aiResponse);

    if (aiResponse.includes('CALL_END')) {
      const clean = aiResponse.replace('CALL_END', '').trim();
      let naam = '', address = '', order = '';
      const saveMatch = [...calls[callSid].history].reverse().find(h => h.role === 'model' && h.parts[0].text.includes('SAVE|'));
      if (saveMatch) {
        const parts = saveMatch.parts[0].text.split('|');
        naam = parts[1] ? parts[1].trim() : '';
        address = parts[2] ? parts[2].trim() : '';
        order = parts[3] ? parts[3].trim() : '';
      }
      saveOrder({ naam, phone: calls[callSid].callerNumber, address, order });
      if (naam && order) await sendSMS(calls[callSid].callerNumber, naam, order);
      twiml.say({ language: 'hi-IN' }, clean || 'Bahut shukriya! Aapka order note ho gaya. SMS abhi aa jayega. Phir milenge!');
      twiml.hangup();
      delete calls[callSid];
    } else {
      const gather = twiml.gather({ input: 'speech', action: '/respond', method: 'POST', language: 'hi-IN', speechTimeout: '2', timeout: '15' });
      gather.say({ language: 'hi-IN' }, aiResponse);
    }
  } catch (err) {
    console.error('ERROR:', err.message);
    const gather = twiml.gather({ input: 'speech', action: '/respond', method: 'POST', language: 'hi-IN', speechTimeout: '2', timeout: '15' });
    gather.say({ language: 'hi-IN' }, 'Ek second, dobara boliye please.');
  }
  res.type('text/xml');
  res.send(twiml.toString());
});

app.use(express.static(path.join(__dirname, 'public')));

app.post('/make-call', async (req, res) => {
  try {
    const { phone } = req.body;
    const call = await twilioClient.calls.create({
      url: `${PUBLIC_BASE_URL}/call`,
      to: phone,
      from: TWILIO_FROM_NUMBER
    });
    pendingCalls[call.sid] = phone;
    console.log('Call created, SID:', call.sid, 'Phone:', phone);
    res.json({ success: true });
  } catch(e) {
    res.json({ success: false, error: e.message });
  }
});

app.get('/orders', (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) return res.json([]);
    const wb = XLSX.readFile(EXCEL_FILE);
    const data = XLSX.utils.sheet_to_json(wb.Sheets['Orders']);
    res.json(data);
  } catch(e) { res.json([]); }
});

// Website ke "Product availability" page ke liye — har variant (product +
// size) ka naam, price, aur bacha hua stock deta hai.
app.get('/availability', (req, res) => {
  try {
    const data = productData.map(p => ({
      Product: p.Product,
      Category: p.Category,
      Size: p.Size,
      'Price (Rs)': p['Price (Rs)'],
      Stock: p.Stock
    }));
    res.json(data);
  } catch(e) { res.json([]); }
});

app.get('/download', (req, res) => {
  if (fs.existsSync(EXCEL_FILE)) res.download(EXCEL_FILE);
  else res.status(404).send('No orders yet');
});

app.listen(3000, () => console.log('Server chal raha hai port 3000 pe!'));