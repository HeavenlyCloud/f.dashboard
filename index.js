require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

const lineConfig = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

const client = new line.messagingApi.MessagingApiClient({
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
});

let googleKey;
if (process.env.GOOGLE_CREDENTIALS_JSON) {
  googleKey = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
} else {
  googleKey = require('./google-credentials.json');
}

const serviceAccountAuth = new JWT({
  email: googleKey.client_email,
  key: googleKey.private_key,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID, serviceAccountAuth);
let isDocLoaded = false;

async function getTransactionsSheet() {
  if (!isDocLoaded) {
    await doc.loadInfo();
    isDocLoaded = true;
  }
  return doc.sheetsByTitle['Transactions'];
}

const app = express();

app.get('/health', (req, res) => {
  res.status(200).send('OK');
});

app.post('/webhook', line.middleware(lineConfig), (req, res) => {
  res.status(200).end();

  Promise.all(req.body.events.map(handleEvent)).catch((err) => {
    console.error('Error handling webhook event in background:', err);
  });
});

async function handleEvent(event) {
  if (event.type !== 'message' || event.message.type !== 'text') {
    return Promise.resolve(null);
  }

  const text = event.message.text.trim();
  const match = text.match(/^(\d+)\s+(.+)$/);

  if (!match) {
    return client.replyMessage({
      replyToken: event.replyToken,
      messages: [
        {
          type: 'text',
          text: 'Format unrecognized. Send transactions like:\n120 night market',
        },
      ],
    });
  }

  const amount = parseInt(match[1], 10);
  const note = match[2];
  const today = new Date().toISOString().split('T')[0];

  try {
    const sheet = await getTransactionsSheet();
    await sheet.addRow({
      Date: today,
      Amount: amount,
      Category: 'Uncategorized',
      Description: note,
      Source: 'Manual',
      Status: 'manual_unverified',
    });

    return client.replyMessage({
      replyToken: event.replyToken,
      messages: [
        {
          type: 'text',
          text: `Logged: $${amount} TWD for "${note}"`,
        },
      ],
    });
  } catch (err) {
    console.error('Failed to write to Google Sheets:', err);
  }
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Finance bot server running on port ${PORT}`);
});