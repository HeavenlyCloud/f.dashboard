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
const DEFAULT_BASE_CURRENCY = process.env.BASE_CURRENCY || 'TWD';

async function ensureDocLoaded() {
  if (!isDocLoaded) {
    await doc.loadInfo();
    isDocLoaded = true;
  }
}

async function getUserOrMainSheet(sheetName = 'Transactions') {
  await ensureDocLoaded();
  let sheet = doc.sheetsByTitle[sheetName];
  if (!sheet) {
    sheet = await doc.addSheet({
      title: sheetName,
      headerValues: sheetName === 'Cart' 
        ? ['Date', 'User', 'Item', 'Amount', 'Currency']
        : ['Date', 'User', 'Amount', 'Currency', 'Amount (Base)', 'Category', 'Description', 'Source', 'Status']
    });
  }
  return sheet;
}

async function getSubscriptionSheet() {
  await ensureDocLoaded();
  return doc.sheetsByTitle['Subscriptions'];
}

function categorizeExpense(note) {
  const lower = note.toLowerCase();
  const rules = [
    { category: 'Food & Dining', keywords: ['lunch', 'dinner', 'breakfast', 'ramen', 'coffee', 'tea', 'food', 'night market', 'boba', 'restaurant', 'meal', 'cafe'] },
    { category: 'Groceries', keywords: ['supermarket', 'market', 'apples', 'milk', 'fruit', 'vegetables', 'meat', 'eggs', 'carrefour', 'pxmart', 'grocery'] },
    { category: 'Transportation', keywords: ['uber', 'taxi', 'mrt', 'train', 'bus', 'flight', 'gas', 'petrol', 'parking', 'metro'] },
    { category: 'Entertainment', keywords: ['movie', 'game', 'steam', 'concert', 'spotify', 'netflix', 'youtube', 'cinema'] },
    { category: 'Shopping', keywords: ['clothes', 'shoes', 'amazon', 'electronics', 'books', 'store'] },
    { category: 'Utilities & Bills', keywords: ['water', 'electricity', 'rent', 'internet', 'phone', 'bill', 'icloud', 'aws'] }
  ];

  for (const rule of rules) {
    if (rule.keywords.some((kw) => lower.includes(kw))) {
      return rule.category;
    }
  }

  return 'Uncategorized';
}

const userNameCache = {};

async function fetchRealUserName(userId) {
  if (!userId || userId === 'Unknown') return 'Guest';
  if (userNameCache[userId]) return userNameCache[userId];

  try {
    const profile = await client.getProfile(userId);
    const sanitizedName = profile.displayName.replace(/[^a-zA-Z0-9_\-\s]/g, '').trim() || `User_${userId.slice(-4)}`;
    userNameCache[userId] = sanitizedName;
    return sanitizedName;
  } catch (err) {
    return `User_${userId.slice(-4)}`;
  }
}

function createReceiptFlexCard(title, amount, currency, note, user, category = 'Uncategorized') {
  return {
    type: 'flex',
    altText: `${title}: ${amount} ${currency} for "${note}"`,
    contents: {
      type: 'bubble',
      styles: { header: { backgroundColor: '#1DB954' } },
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: 'HeavenlyCloud 雲雲', weight: 'bold', color: '#FFFFFF', size: 'xs' },
          { type: 'text', text: title, weight: 'bold', color: '#FFFFFF', size: 'xl' }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'md',
        contents: [
          {
            type: 'box',
            layout: 'horizontal',
            contents: [
              { type: 'text', text: 'Amount', color: '#8C8C8C', size: 'sm' },
              { type: 'text', text: `${amount} ${currency}`, weight: 'bold', align: 'end', size: 'lg', color: '#1DB954' }
            ]
          },
          {
            type: 'box',
            layout: 'horizontal',
            contents: [
              { type: 'text', text: 'Description', color: '#8C8C8C', size: 'sm' },
              { type: 'text', text: note, align: 'end', size: 'sm', wrap: true }
            ]
          },
          {
            type: 'box',
            layout: 'horizontal',
            contents: [
              { type: 'text', text: 'Category', color: '#8C8C8C', size: 'sm' },
              { type: 'text', text: category, align: 'end', size: 'sm' }
            ]
          },
          {
            type: 'box',
            layout: 'horizontal',
            contents: [
              { type: 'text', text: 'User', color: '#8C8C8C', size: 'sm' },
              { type: 'text', text: user, align: 'end', size: 'sm' }
            ]
          }
        ]
      },
      footer: {
        type: 'box',
        layout: 'horizontal',
        spacing: 'sm',
        contents: [
          {
            type: 'button',
            style: 'secondary',
            height: 'sm',
            action: { type: 'message', label: '🛒 Start Grocery', text: 'start glist' }
          },
          {
            type: 'button',
            style: 'primary',
            color: '#1DB954',
            height: 'sm',
            action: { type: 'message', label: '📊 Summary', text: 'summary' }
          }
        ]
      }
    }
  };
}

function createCartFlexCard(cartItems, runningTotal, userName, isFinal = false) {
  const itemRows = cartItems.map((item, idx) => ({
    type: 'box',
    layout: 'horizontal',
    contents: [
      { type: 'text', text: `${idx + 1}. ${item.note}`, size: 'sm', color: '#333333', flex: 3, wrap: true },
      { type: 'text', text: `$${item.amount} ${DEFAULT_BASE_CURRENCY}`, size: 'sm', align: 'end', weight: 'bold', flex: 2 }
    ]
  }));

  const footerButtons = isFinal
    ? [
        {
          type: 'button',
          style: 'primary',
          color: '#1DB954',
          height: 'sm',
          action: { type: 'message', label: '🛒 New Grocery Trip', text: 'start glist' }
        }
      ]
    : [
        {
          type: 'button',
          style: 'secondary',
          height: 'sm',
          action: { type: 'message', label: '📊 View Cart', text: 'partial total' }
        },
        {
          type: 'button',
          style: 'secondary',
          height: 'sm',
          action: { type: 'message', label: '↩️ Undo', text: 'undo' }
        },
        {
          type: 'button',
          style: 'primary',
          color: '#1DB954',
          height: 'sm',
          action: { type: 'message', label: '✅ Finish Trip', text: 'end glist' }
        }
      ];

  return {
    type: 'flex',
    altText: `Cart Total: $${runningTotal} ${DEFAULT_BASE_CURRENCY}`,
    contents: {
      type: 'bubble',
      styles: { header: { backgroundColor: isFinal ? '#1DB954' : '#2D3748' } },
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: `HeavenlyCloud 雲雲 (${userName})`, weight: 'bold', color: '#FFFFFF', size: 'xs' },
          { type: 'text', text: isFinal ? '🛒 Grocery Summary' : '🛒 Active Grocery Cart', weight: 'bold', color: '#FFFFFF', size: 'lg' }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'md',
        contents: [
          ...(itemRows.length > 0 ? itemRows : [{ type: 'text', text: 'Cart is empty.', size: 'sm', color: '#8C8C8C' }]),
          { type: 'separator', margin: 'md' },
          {
            type: 'box',
            layout: 'horizontal',
            margin: 'md',
            contents: [
              { type: 'text', text: 'Running Total', weight: 'bold', size: 'md' },
              { type: 'text', text: `$${runningTotal} ${DEFAULT_BASE_CURRENCY}`, weight: 'bold', size: 'lg', align: 'end', color: '#1DB954' }
            ]
          }
        ]
      },
      footer: {
        type: 'box',
        layout: 'horizontal',
        spacing: 'xs',
        contents: footerButtons
      }
    }
  };
}

function createCartItemAddedFlexCard(lastItem, runningTotal, cartCount, userName) {
  return {
    type: 'flex',
    altText: `Added ${lastItem.note} ($${lastItem.amount}) - Cart Total: $${runningTotal}`,
    contents: {
      type: 'bubble',
      styles: { header: { backgroundColor: '#2D3748' } },
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: `HeavenlyCloud 雲雲 (${userName})`, weight: 'bold', color: '#FFFFFF', size: 'xs' },
          { type: 'text', text: '🛒 Item Added to Cart', weight: 'bold', color: '#FFFFFF', size: 'lg' }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'md',
        contents: [
          {
            type: 'box',
            layout: 'horizontal',
            contents: [
              { type: 'text', text: 'Item Added', color: '#8C8C8C', size: 'sm' },
              { type: 'text', text: lastItem.note, weight: 'bold', align: 'end', size: 'sm', wrap: true }
            ]
          },
          {
            type: 'box',
            layout: 'horizontal',
            contents: [
              { type: 'text', text: 'Item Cost', color: '#8C8C8C', size: 'sm' },
              { type: 'text', text: `$${lastItem.amount} ${DEFAULT_BASE_CURRENCY}`, weight: 'bold', align: 'end', size: 'sm' }
            ]
          },
          { type: 'separator' },
          {
            type: 'box',
            layout: 'horizontal',
            contents: [
              { type: 'text', text: `Cart (${cartCount} items)`, color: '#8C8C8C', size: 'sm' },
              { type: 'text', text: `$${runningTotal} ${DEFAULT_BASE_CURRENCY}`, weight: 'bold', align: 'end', size: 'md', color: '#1DB954' }
            ]
          }
        ]
      },
      footer: {
        type: 'box',
        layout: 'horizontal',
        spacing: 'xs',
        contents: [
          {
            type: 'button',
            style: 'secondary',
            height: 'sm',
            action: { type: 'message', label: '📊 View Cart', text: 'partial total' }
          },
          {
            type: 'button',
            style: 'secondary',
            height: 'sm',
            action: { type: 'message', label: '↩️ Undo', text: 'undo' }
          },
          {
            type: 'button',
            style: 'primary',
            color: '#1DB954',
            height: 'sm',
            action: { type: 'message', label: '✅ Finish', text: 'end glist' }
          }
        ]
      }
    }
  };
}

function createSummaryFlexCard(monthStr, totalSpent, categoryTotals) {
  const rows = Object.entries(categoryTotals).map(([cat, amt]) => ({
    type: 'box',
    layout: 'horizontal',
    contents: [
      { type: 'text', text: cat, size: 'sm', color: '#555555', flex: 3 },
      { type: 'text', text: `$${amt} ${DEFAULT_BASE_CURRENCY}`, size: 'sm', align: 'end', weight: 'bold', flex: 2 }
    ]
  }));

  return {
    type: 'flex',
    altText: `Monthly Summary (${monthStr}): Total $${totalSpent} ${DEFAULT_BASE_CURRENCY}`,
    contents: {
      type: 'bubble',
      styles: { header: { backgroundColor: '#1DB954' } },
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: 'HeavenlyCloud 雲雲', weight: 'bold', color: '#FFFFFF', size: 'xs' },
          { type: 'text', text: `📊 Spending Report (${monthStr})`, weight: 'bold', color: '#FFFFFF', size: 'lg' }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'md',
        contents: rows.length > 0 ? rows : [{ type: 'text', text: 'No expenses logged this month.', size: 'sm', color: '#8C8C8C' }]
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'separator' },
          {
            type: 'box',
            layout: 'horizontal',
            margin: 'md',
            contents: [
              { type: 'text', text: 'Total Spent', weight: 'bold', size: 'md' },
              { type: 'text', text: `$${totalSpent} ${DEFAULT_BASE_CURRENCY}`, weight: 'bold', size: 'lg', align: 'end', color: '#1DB954' }
            ]
          }
        ]
      }
    }
  };
}

function createErrorFlexCard(title, message) {
  return {
    type: 'flex',
    altText: title,
    contents: {
      type: 'bubble',
      styles: { header: { backgroundColor: '#E53E3E' } },
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: 'HeavenlyCloud 雲雲', weight: 'bold', color: '#FFFFFF', size: 'xs' },
          { type: 'text', text: title, weight: 'bold', color: '#FFFFFF', size: 'lg' }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: message, size: 'sm', color: '#4A5568', wrap: true }
        ]
      },
      footer: {
        type: 'box',
        layout: 'horizontal',
        spacing: 'sm',
        contents: [
          {
            type: 'button',
            style: 'secondary',
            height: 'sm',
            action: { type: 'message', label: '🛒 Start Grocery', text: 'start glist' }
          },
          {
            type: 'button',
            style: 'primary',
            color: '#1DB954',
            height: 'sm',
            action: { type: 'message', label: '📊 Summary', text: 'summary' }
          }
        ]
      }
    }
  };
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
      const currency = row.get('Currency') || DEFAULT_BASE_CURRENCY;
      const origAmount = row.get('Original Amount');
      const baseAmount = row.get(`Amount (${DEFAULT_BASE_CURRENCY})`);

      if (status === 'Active' && renewalDateStr) {
        const renewalDate = new Date(renewalDateStr);
        const diffTime = renewalDate - today;
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        if (diffDays === 3) {
          upcomingAlerts.push(
            `🔔 Global Subscription Renewal Warning (3 Days)\n` +
            `• Service: ${name}\n` +
            `• Cost: ${origAmount} ${currency} (~$${Math.round(baseAmount || origAmount)} ${DEFAULT_BASE_CURRENCY})\n` +
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
  if (event.type !== 'message' || event.message.type !== 'text') return null;

  const rawUserId = event.source.userId || 'Unknown';
  const userName = await fetchRealUserName(rawUserId);
  const text = event.message.text.trim();
  const lowerText = text.toLowerCase();

  if (lowerText === 'summary' || lowerText === 'this month') {
    try {
      const sheet = await getUserOrMainSheet('Transactions');
      const rows = await sheet.getRows();
      const currentMonth = new Date().toISOString().slice(0, 7);

      let totalSpent = 0;
      const categoryTotals = {};

      rows.forEach((row) => {
        const dateStr = row.get('Date');
        if (dateStr && dateStr.startsWith(currentMonth)) {
          const amt = parseFloat(row.get('Amount (Base)') || row.get('Amount')) || 0;
          const category = row.get('Category') || 'Uncategorized';
          totalSpent += amt;
          categoryTotals[category] = (categoryTotals[category] || 0) + amt;
        }
      });

      const flexCard = createSummaryFlexCard(currentMonth, Math.round(totalSpent), categoryTotals);
      return client.replyMessage({ replyToken: event.replyToken, messages: [flexCard] });
    } catch (err) {
      console.error('Failed generating spending summary:', err);
    }
  }

  const subMatch = text.match(/^sub\s+add\s+(.+)\s+([A-Za-z]{3})\s+(\d+(?:\.\d+)?)\s+(\d{4}-\d{2}-\d{2})$/i);
  if (subMatch) {
    const name = subMatch[1].trim();
    const currency = subMatch[2].toUpperCase();
    const origAmount = parseFloat(subMatch[3]);
    const renewalDate = subMatch[4];

    const fxFormula = currency === DEFAULT_BASE_CURRENCY 
      ? origAmount 
      : `=D${2}*GOOGLEFINANCE("CURRENCY:${currency}${DEFAULT_BASE_CURRENCY}")`;

    const flexCard = createReceiptFlexCard('Subscription Added', origAmount, currency, `${name} (Next: ${renewalDate})`, userName, 'Subscriptions');

    const replyPromise = client.replyMessage({
      replyToken: event.replyToken,
      messages: [flexCard]
    });

    const sheetPromise = (async () => {
      try {
        const sheet = await getSubscriptionSheet();
        await sheet.addRow({
          Name: name,
          User: userName,
          Currency: currency,
          'Original Amount': origAmount,
          [`Amount (${DEFAULT_BASE_CURRENCY})`]: fxFormula,
          'Billing Cycle': 'Monthly',
          'Next Renewal Date': renewalDate,
          Status: 'Active'
        });
      } catch (err) {
        console.error('Failed adding subscription to Google Sheet:', err);
      }
    })();

    await replyPromise;
    await sheetPromise;
    return;
  }

  if (lowerText === 'start glist') {
    grocerySessions[rawUserId] = { items: [], active: true, startTime: new Date() };
    const flexCard = createCartFlexCard([], 0, userName, false);

    return client.replyMessage({
      replyToken: event.replyToken,
      messages: [flexCard]
    });
  }

  const session = grocerySessions[rawUserId];
  if (session && session.active) {
    if (lowerText === 'partial total') {
      const currentTotal = session.items.reduce((sum, item) => sum + item.amount, 0);
      const flexCard = createCartFlexCard(session.items, currentTotal, userName, false);

      return client.replyMessage({ replyToken: event.replyToken, messages: [flexCard] });
    }

    const removeMatch = text.match(/^(?:remove|delete|del)\s+(.+)$/i);
    if (removeMatch) {
      const itemToRemove = removeMatch[1].toLowerCase().trim();
      const initialLength = session.items.length;

      session.items = session.items.filter((i) => i.note.toLowerCase() !== itemToRemove);

      if (session.items.length < initialLength) {
        const runningTotal = session.items.reduce((sum, item) => sum + item.amount, 0);
        const flexCard = createCartFlexCard(session.items, runningTotal, userName, false);

        return client.replyMessage({ replyToken: event.replyToken, messages: [flexCard] });
      } else {
        const flexCard = createErrorFlexCard('Item Not Found', `Item "${itemToRemove}" was not found in your active cart.`);
        return client.replyMessage({ replyToken: event.replyToken, messages: [flexCard] });
      }
    }

    if (lowerText === 'undo' || lowerText === 'pop') {
      if (session.items.length === 0) {
        const flexCard = createErrorFlexCard('Cart Empty', 'Your grocery cart is already empty.');
        return client.replyMessage({ replyToken: event.replyToken, messages: [flexCard] });
      }

      session.items.pop();
      const runningTotal = session.items.reduce((sum, item) => sum + item.amount, 0);
      const flexCard = createCartFlexCard(session.items, runningTotal, userName, false);

      return client.replyMessage({ replyToken: event.replyToken, messages: [flexCard] });
    }

    if (lowerText === 'total' || lowerText === 'end glist') {
      const finalTotal = session.items.reduce((sum, item) => sum + item.amount, 0);

      if (session.items.length === 0) {
        delete grocerySessions[rawUserId];
        const flexCard = createErrorFlexCard('Grocery Session Ended', 'Shopping trip ended with 0 items logged.');
        return client.replyMessage({ replyToken: event.replyToken, messages: [flexCard] });
      }

      const summaryNote = session.items.map((i) => `${i.note} ($${i.amount})`).join(', ');
      const today = new Date().toISOString().split('T')[0];
      const flexCard = createCartFlexCard(session.items, finalTotal, userName, true);

      const replyPromise = client.replyMessage({ replyToken: event.replyToken, messages: [flexCard] });

      const sheetPromise = (async () => {
        try {
          const txSheet = await getUserOrMainSheet('Transactions');
          await txSheet.addRow({
            Date: today,
            User: userName,
            Amount: finalTotal,
            Currency: DEFAULT_BASE_CURRENCY,
            'Amount (Base)': finalTotal,
            Category: 'Groceries',
            Description: `Grocery Trip: ${summaryNote}`,
            Source: 'Manual (glist)',
            Status: 'manual_unverified',
          });

          const cartSheet = await getUserOrMainSheet('Cart');
          for (const item of session.items) {
            await cartSheet.addRow({
              Date: today,
              User: userName,
              Item: item.note,
              Amount: item.amount,
              Currency: DEFAULT_BASE_CURRENCY
            });
          }
        } catch (err) {
          console.error('Failed to log grocery trip to Google Sheets:', err);
        }
      })();

      delete grocerySessions[rawUserId];
      await replyPromise;
      await sheetPromise;
      return;
    }

    const match = text.match(/^(\d+(?:\.\d+)?)\s+(.+)$/);
    if (match) {
      const amount = parseFloat(match[1]);
      const note = match[2];

      session.items.push({ amount, note });
      const runningTotal = session.items.reduce((sum, item) => sum + item.amount, 0);

      const flexCard = createCartItemAddedFlexCard({ amount, note }, runningTotal, session.items.length, userName);

      return client.replyMessage({ replyToken: event.replyToken, messages: [flexCard] });
    }
  }

  const multiCurrMatch = text.match(/^(\d+(?:\.\d+)?)\s+([A-Za-z]{3})\s+(.+)$/i);
  const stdMatch = text.match(/^(\d+(?:\.\d+)?)\s+(.+)$/);

  let amount = 0;
  let currency = DEFAULT_BASE_CURRENCY;
  let note = '';

  if (multiCurrMatch) {
    amount = parseFloat(multiCurrMatch[1]);
    currency = multiCurrMatch[2].toUpperCase();
    note = multiCurrMatch[3];
  } else if (stdMatch) {
    amount = parseFloat(stdMatch[1]);
    note = stdMatch[2];
  } else {
    const flexCard = createErrorFlexCard(
      'Format Unrecognized',
      'Send transactions like:\n• 120 night market\n• 15.49 USD netflix\n\nOr query budget with:\n• summary'
    );
    return client.replyMessage({ replyToken: event.replyToken, messages: [flexCard] });
  }

  const category = categorizeExpense(note);
  const today = new Date().toISOString().split('T')[0];
  const fxFormula = currency === DEFAULT_BASE_CURRENCY 
    ? amount 
    : `=C${2}*GOOGLEFINANCE("CURRENCY:${currency}${DEFAULT_BASE_CURRENCY}")`;

  const flexCard = createReceiptFlexCard('Expense Logged', amount, currency, note, userName, category);

  const replyPromise = client.replyMessage({
    replyToken: event.replyToken,
    messages: [flexCard]
  });

  const sheetPromise = (async () => {
    try {
      const sheet = await getUserOrMainSheet('Transactions');
      await sheet.addRow({
        Date: today,
        User: userName,
        Amount: amount,
        Currency: currency,
        'Amount (Base)': fxFormula,
        Category: category,
        Description: note,
        Source: 'Manual',
        Status: 'manual_unverified',
      });
    } catch (err) {
      console.error('Failed writing global transaction to Google Sheet:', err);
    }
  })();

  await replyPromise;
  await sheetPromise;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Finance bot server running on port ${PORT}`);
});