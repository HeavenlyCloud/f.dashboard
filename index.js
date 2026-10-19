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

const serviceAccountAuth = new JWT({
  email: require('./google-credentials.json').client_email,
  key: require('./google-credentials.json').private_key,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const app = express();

app.post('/webhook', line.middleware(lineConfig), async (req, res) => {
  try {
    const results = await Promise.all(req.body.events.map(handleEvent));
    res.json(results);
  } catch (err) {
    console.error('Error handling webhook event:', err);
    res.status(500).end();
  }
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

  const doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID, serviceAccountAuth);
  await doc.loadInfo();
  const sheet = doc.sheetsByTitle['Transactions'];

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
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Finance bot server running on port ${PORT}`);
});