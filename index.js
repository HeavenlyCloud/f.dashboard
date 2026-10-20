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
let transactionsSheet = null;
let subscriptionSheet = null;

async function getTransactionsSheet() {
  if (!transactionsSheet) {
    await doc.loadInfo();
    transactionsSheet = doc.sheetsByTitle['Transactions'];
  }
  return transactionsSheet;
}

async function getSubscriptionSheet() {
  if (!subscriptionSheet) {
    await doc.loadInfo();
    subscriptionSheet = doc.sheetsByTitle['Subscriptions'];
  }
  return subscriptionSheet;
}

const KNOWN_USERS = {};

function getSanitizedUserName(userId) {
  if (!userId || userId === 'Unknown') return 'Guest';
  if (KNOWN_USERS[userId]) return KNOWN_USERS[userId];
  return `User_${userId.slice(-4)}`;
}

const grocerySessions = {};

const app = express();

app.get('/health', (req, res) => {
  res.status(200).send('OK');
});

app.get('/check-subscriptions', async (req, res) => {
  try {
    const sheet = await getSubscriptionSheet();
    const rows = await sheet.getRows();
    const today = new Date();

    const upcomingAlerts = [];

    rows.forEach((row) => {
      const name = row.get('Name');
      const renewalDateStr = row.get('Next Renewal Date');
      const status = row.get('Status');
      const currency = row.get('Currency') || 'TWD';
      const origAmount = row.get('Original Amount');
      const twdAmount = row.get('Amount (TWD)');

      if (status === 'Active' && renewalDateStr) {
        const renewalDate = new Date(renewalDateStr);
        const diffTime = renewalDate - today;
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        if (diffDays === 3) {
          upcomingAlerts.push(
            `🔔 Subscription Renewal in 3 Days!\n` +
            `• Service: ${name}\n` +
            `• Cost: ${origAmount} ${currency} (~$${Math.round(twdAmount || 0)} TWD)\n` +
            `• Date: ${renewalDateStr}`
          );
        }
      }
    });

    if (upcomingAlerts.length > 0 && process.env.USER_LINE_ID) {
      await client.pushMessage({
        to: process.env.USER_LINE_ID,
        messages: [{ type: 'text', text: upcomingAlerts.join('\n\n') }],
      });
    }

    res.status(200).send({ status: 'Checked', alertsSent: upcomingAlerts.length });
  } catch (err) {
    console.error('Failed checking subscription renewals:', err);
    res.status(500).send({ error: err.message });
  }
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

  const rawUserId = event.source.userId || 'Unknown';
  const userName = getSanitizedUserName(rawUserId);
  const text = event.message.text.trim();
  const lowerText = text.toLowerCase();

  if (lowerText === 'start glist') {
    grocerySessions[rawUserId] = {
      items: [],
      active: true,
      startTime: new Date(),
    };

    return client.replyMessage({
      replyToken: event.replyToken,
      messages: [
        {
          type: 'text',
          text: `🛒 Grocery List session started!\n\nAdd items like:\n"120 apples"\n"85 milk"\n\nCommands:\n• "partial total" - view current cart\n• "remove apples"\n• "undo" / "pop" - remove last item\n• "total" / "end glist" - save and complete`,
        },
      ],
    });
  }

  const session = grocerySessions[rawUserId];

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
        delete grocerySessions[rawUserId];
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
          const sheet = await getTransactionsSheet();
          await sheet.addRow({
            Date: today,
            User: userName,
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

      delete grocerySessions[rawUserId];

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
      const sheet = await getTransactionsSheet();
      await sheet.addRow({
        Date: today,
        User: userName,
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