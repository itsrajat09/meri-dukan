require('./load-env')(); // loads .env into process.env

const TWILIO_ACCOUNT_SID = process.env.TWILIO_ACCOUNT_SID;
const TWILIO_AUTH_TOKEN = process.env.TWILIO_AUTH_TOKEN;
const TWILIO_FROM_NUMBER = process.env.TWILIO_FROM_NUMBER;
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL;

// Pass the customer's number as a command-line arg:
//   node call.js +917011641451
const toNumber = process.argv[2];

if (!TWILIO_ACCOUNT_SID || !TWILIO_AUTH_TOKEN || !TWILIO_FROM_NUMBER || !PUBLIC_BASE_URL) {
  console.error('Missing env vars. Copy .env.example to .env and fill it in.');
  process.exit(1);
}
if (!toNumber) {
  console.error('Usage: node call.js +91XXXXXXXXXX');
  process.exit(1);
}

const t = require('twilio')(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN);
t.calls.create({
  to: toNumber,
  from: TWILIO_FROM_NUMBER,
  url: `${PUBLIC_BASE_URL}/call`
}).then(c => console.log('Call ja rahi hai!', c.sid))
.catch(e => console.log('Error:', e.message));