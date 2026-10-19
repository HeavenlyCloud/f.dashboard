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
let sheetInstance = null;

async function getSheet() {
  if (!sheetInstance) {
    await doc.loadInfo();
    sheetInstance = doc.sheetsByTitle['Transactions'];
  }
  return sheetInstance;
}

const grocerySessions = {};

const app = express();

app.get('/health', (req, res) => {
  res.status(200).send('OK');
});

app.post('/webhook', line.middleware(lineConfig), (req, res) => {
  res.status(200).end();

  Promise.all(req.body.events.map(handleEvent)).catch((err) => {
    console.error('Error processing webhook event in background:', err);
  });
});

async function handleEvent(event) {
  if (event.type !== 'message' || event.message.type !== 'text') {
    return null;
  }

  const userId = event.source.userId || 'Unknown';
  const text = event.message.text.trim();
  const lowerText = text.toLowerCase();

  if (lowerText === 'start glist') {
    grocerySessions[userId] = {
      items: [],
      active: true,
      startTime: new Date(),
    };

    return client.replyMessage({
      replyToken: event.replyToken,
      messages: [
        {
          type: 'text',
          text: `🛒 Grocery List session started!\n\nAdd items like:\n"120 apples"\n"85 milk"\n\nCommands:\n• "partial total" - view current cart\n• "remove apples" / "del apples"\n• "undo" / "pop" - remove last item\n• "total" / "end glist" - save and complete`,
        },
      ],
    });
  }


  const session = grocerySessions[userId];

  if (session && session.active) {

    if (lowerText === 'partial total') {
      const currentTotal = session.items.reduce((sum, item) => sum + item.amount, 0);
      const listText = session.items.length > 0
        ? session.items.map((i, idx) => `${idx + 1}. ${i.note}: $${i.amount} TWD`).join('\n')
        : 'Cart is currently empty.';

      return client.replyMessage({
        replyToken: event.replyToken,
        messages: [
          {
            type: 'text',
            text: `🛒 Current Cart:\n${listText}\n\n💰 Partial Total: $${currentTotal} TWD`,
          },
        ],
      });
    }

    const removeMatch = text.match(/^(?:remove|delete|del)\s+(.+)$/i);
    if (removeMatch) {
      const itemToRemove = removeMatch[1].toLowerCase().trim();
      const initialLength = session.items.length;

      session.items = session.items.filter((i) => i.note.toLowerCase() !== itemToRemove);

      if (session.items.length < initialLength) {
        const runningTotal = session.items.reduce((sum, item) => sum + item.amount, 0);
        return client.replyMessage({
          replyToken: event.replyToken,
          messages: [
            {
              type: 'text',
              text: `🗑️ Removed "${itemToRemove}" from cart.\n🛒 Running Total: $${runningTotal} TWD`,
            },
          ],
        });
      } else {
        return client.replyMessage({
          replyToken: event.replyToken,
          messages: [
            {
              type: 'text',
              text: `⚠️ Item "${itemToRemove}" not found in current cart.`,
            },
          ],
        });
      }
    }

    if (lowerText === 'undo' || lowerText === 'pop') {
      if (session.items.length === 0) {
        return client.replyMessage({
          replyToken: event.replyToken,
          messages: [{ type: 'text', text: '🛒 Cart is already empty.' }],
        });
      }

      const poppedItem = session.items.pop();
      const runningTotal = session.items.reduce((sum, item) => sum + item.amount, 0);

      return client.replyMessage({
        replyToken: event.replyToken,
        messages: [
          {
            type: 'text',
            text: `↩️ Removed last item: "${poppedItem.note}" ($${poppedItem.amount} TWD).\n🛒 Running Total: $${runningTotal} TWD`,
          },
        ],
      });
    }

    if (lowerText === 'total' || lowerText === 'end glist') {
      const finalTotal = session.items.reduce((sum, item) => sum + item.amount, 0);

      if (session.items.length === 0) {
        delete grocerySessions[userId];
        return client.replyMessage({
          replyToken: event.replyToken,
          messages: [{ type: 'text', text: '🛒 Grocery list ended (0 items logged).' }],
        });
      }

      const summaryNote = session.items.map((i) => `${i.note} ($${i.amount})`).join(', ');
      const today = new Date().toISOString().split('T')[0];

      const replyPromise = client.replyMessage({
        replyToken: event.replyToken,
        messages: [
          {
            type: 'text',
            text: `✅ Grocery trip completed!\n\nItemization:\n${session.items.map((i) => `• ${i.note}:$${i.amount}`).join('\n')}\n\n💳 Total Logged: $${finalTotal} TWD`,
          },
        ],
      });

      const sheetPromise = (async () => {
        try {
          const sheet = await getSheet();
          await sheet.addRow({
            Date: today,
            User: userId,
            Amount: finalTotal,
            Category: 'Groceries',
            Description: `Grocery Trip: ${summaryNote}`,
            Source: 'Manual (glist)',
            Status: 'manual_unverified',
          });
        } catch (err) {
          console.error('Failed to log grocery trip to Google Sheets:', err);
        }
      })();

      delete grocerySessions[userId];

      await replyPromise;
      await sheetPromise;
      return;
    }

    const match = text.match(/^(\d+)\s+(.+)$/);
    if (match) {
      const amount = parseInt(match[1], 10);
      const note = match[2];

      session.items.push({ amount, note });
      const runningTotal = session.items.reduce((sum, item) => sum + item.amount, 0);

      return client.replyMessage({
        replyToken: event.replyToken,
        messages: [
          {
            type: 'text',
            text: `Added: $${amount} TWD for "${note}"\n🛒 Running Total: $${runningTotal} TWD`,
          },
        ],
      });
    }
  }

  const match = text.match(/^(\d+)\s+(.+)$/);
  if (!match) {
    return client.replyMessage({
      replyToken: event.replyToken,
      messages: [
        {
          type: 'text',
          text: 'Format unrecognized. Send transactions like:\n120 night market\n\nOr start grocery mode with:\nstart glist',
        },
      ],
    });
  }

  const amount = parseInt(match[1], 10);
  const note = match[2];
  const today = new Date().toISOString().split('T')[0];

  const replyPromise = client.replyMessage({
    replyToken: event.replyToken,
    messages: [
      {
        type: 'text',
        text: `Logged: $${amount} TWD for "${note}"`,
      },
    ],
  });

  const sheetPromise = (async () => {
    try {
      const sheet = await getSheet();
      await sheet.addRow({
        Date: today,
        User: userId,
        Amount: amount,
        Category: 'Uncategorized',
        Description: note,
        Source: 'Manual',
        Status: 'manual_unverified',
      });
    } catch (err) {
      console.error('Failed to write transaction to Google Sheets:', err);
    }
  })();

  await replyPromise;
  await sheetPromise;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Finance bot server running on port ${PORT}`);
});