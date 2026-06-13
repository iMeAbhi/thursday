# FinanceOS 💰

Your personal finance dashboard. Your data lives in your own Google Sheet, the app lives on your phone, and a free AI advisor reads your numbers and tells you what to do next. Nothing is ever sent to anyone else's server.

Setting up takes about 15 minutes, one time. After that it's all automatic. Follow the steps in order — each one is just a few clicks.

---

## 1. Set up your Google Sheet (run the script once)

This creates your entire financial spreadsheet automatically — every tab, every formula, sample data and the daily morning email. You never have to build anything by hand.

1. On your computer, open **script.google.com** in your browser and sign in with your Google account.
2. Click **New project** (big button, top left).
3. You'll see a code window with a few lines already in it. Select all of that and delete it.
4. Open the file **Code.gs** (it came with this app), select everything inside, copy it, and paste it into the empty code window.
5. Click the **floppy disk icon** (Save) at the top.
6. In the toolbar at the top there's a dropdown that says which function will run. Make sure it says **setupFinanceOS**.
7. Click **Run** (the ▶ play button).
8. Google will ask you to authorise. Click **Review permissions**, choose your account, click **Advanced**, then **Go to Untitled project (unsafe)**, then **Allow**. This warning appears for every personal script — it's *your own* script running in *your own* account, so it's safe.
9. Wait about 30 seconds. When it finishes, open **Google Drive** — you'll find a new spreadsheet called **FinanceOS Data**. That's yours now.
10. Open it and fill in your real numbers:
    - **CONFIG** tab → your monthly take-home salary in the yellow cell
    - **ACCOUNTS** tab → your real bank balances
    - **LIABILITIES** tab → your real card limits, outstanding amounts and due dates
    - **EQUITY_MF**, **GOLD**, **FIXED_INCOME** tabs → your real investments
    - **TAX** tab → your salary details in the yellow cells

Don't worry about the formulas — they update themselves. Just type over the sample numbers.

---

## 2. Get your Sheet ID

The app needs to know which Sheet is yours. The ID is hiding in the web address.

1. Open your **FinanceOS Data** sheet in the browser.
2. Look at the **address bar at the very top of the browser**. It looks like:
   docs.google.com/spreadsheets/d/**1aBcDeFgHiJkLmNoP...**/edit
3. Copy the long jumble of letters and numbers between **/d/** and **/edit**. That's your Sheet ID.
4. Paste it somewhere handy (a note on your phone) — you'll need it in step 5.

⚠️ **Common mistake:** the link shown inside the "Publish to web" popup starts with **2PACX** — that is *not* your Sheet ID and won't work. Always take the ID from the browser's address bar, never from the publish popup. (If you accidentally paste the wrong one, the app will tell you.)

Tip: if you copy the whole address-bar URL, that's fine too — the app pulls the ID out of it.

---

## 3. Let the app read your Sheet (two quick settings)

This gives the app read-only access. Nobody can edit your data — they could only read it, and only if they somehow had the link.

**Part A — share with link:**

1. In your Sheet, click the green **Share** button (top right).
2. Under **General access**, change "Restricted" to **Anyone with the link**.
3. Make sure the role next to it says **Viewer** (it's the default).
4. Click **Done**.

This is the part the app actually uses to read your numbers — don't skip it.

**Part B — publish to web:**

1. Click **File** → **Share** → **Publish to web**.
2. First dropdown: **Entire document**. Second dropdown: **Comma-separated values (.csv)**.
3. Click **Publish**, then **OK**. Close the popup — you don't need the link it shows.

---

## 4. Get your free Gemini AI key

This powers the AI Advisor tab. It's free and takes two minutes.

1. Go to **aistudio.google.com** in your browser.
2. Sign in with your Google account.
3. Click **Get API key**.
4. Click **Create API key in new project**.
5. A long key appears. Click **Copy**.
6. You'll paste it into the app in the next step. Optionally also keep a backup copy in your Sheet's **CONFIG** tab, cell **B6** — but only do that after you've completed step 8 (private mode), otherwise anyone with your Sheet's link could read the key too.

---

## 5. Put FinanceOS on the internet (so your phone can install it)

The app files contain no personal data or secrets, so it's safe to host them publicly. GitHub Pages is free and takes five minutes:

1. Go to **github.com** and sign up (or sign in).
2. Click the **+** at the top right → **New repository**.
3. Name it **financeos**, leave it on **Public**, click **Create repository**.
4. On the new page, click the small link **"uploading an existing file"**.
5. Drag these five files in: **index.html, manifest.json, sw.js, icon-192.svg, icon-512.svg**. Click **Commit changes**.
6. Go to the repository's **Settings** tab → **Pages** (left sidebar).
7. Under "Branch" choose **main**, folder **/ (root)**, click **Save**.
8. Wait a minute or two, then refresh — a green box shows your app's address:
   **https://YOUR-USERNAME.github.io/financeos/**
   That link is your app. Bookmark it.

### Then open it on your phone

1. Open **Chrome** on your Android phone and go to that link.
2. The first-time setup wizard appears:
   - **Step 1** — paste your Sheet ID (or, in private mode, your web app link + token) and tap Next
   - **Step 2** — paste your Gemini key and tap Next
   - **Step 3** — tap Open Dashboard 🎉
3. These values are saved **on your phone only** — they are never sent anywhere.
4. Made a typo? Tap the **⚙️ gear icon** in the top bar any time to re-enter them.

---

## 6. Install it as a home screen app

1. While the app is open in Chrome, look for the **"Add to Home Screen"** message at the bottom of the screen. (If it doesn't appear, tap the **⋮ three-dot menu** at the top right and choose **Add to Home screen**.)
2. Tap **Add**, then **Add** again to confirm.
3. Done. FinanceOS now sits on your home screen with its own icon and opens full screen, exactly like a native app. It even works offline — it remembers your last numbers.

---

## 7. Daily use

- ☀️ **Morning** — an email briefing arrives around 8am: your net worth, anything due this week, and one money tip.
- 📱 **Any time** — open the app for the live dashboard. Pull down to refresh.
- ✏️ **When things change** — update the numbers in your Google Sheet (takes 5–10 minutes a month). Paid a card? Update the outstanding amount. New SIP? Add a row.
- ✨ **Once a month** — open the **AI tab** and tap **Refresh with Live Data**. The advisor reads your actual Sheet and gives you four specific actions. Ask follow-up questions in the chat below it.
- 🌑 **Light or dark** — tap the ☀️🌑 pill in the top bar. The app remembers your choice.
- 📸 **Track progress** — in your Sheet, use the **💰 FinanceOS menu** → **Export Monthly Snapshot** once a month. Your morning email will then show how much your net worth grew.

---

## 8. Lock down your Sheet (recommended once everything works)

With the basic setup, anyone who somehow obtains your Sheet's link can read it. Private mode closes that door completely: your Sheet goes back to fully **Restricted**, and the app reads your data through your own tiny "web app" instead — which only answers when it's shown a secret token saved on your phone.

**Part A — update the script (only if you set up before this feature existed):**

1. Go to **script.google.com** and open your FinanceOS project.
2. Replace all the code with the latest **Code.gs** and press Save. (Fresh setups already have it.)

**Part B — create your secret token:**

1. Open your **FinanceOS Data** sheet. In the menu bar click **💰 FinanceOS** → **🔐 Generate App Token**.
2. A long random code appears in the **CONFIG** tab, cell **B8**. Keep that tab open — you'll copy it in a minute.

**Part C — turn the script into your personal web app:**

1. Back at script.google.com, click the blue **Deploy** button (top right) → **New deployment**.
2. Click the ⚙️ gear next to "Select type" and choose **Web app**.
3. Set **Execute as: Me**, and **Who has access: Anyone**. (Sounds scary — it isn't: "anyone" can only reach a door that stays shut without your token.)
4. Click **Deploy**, authorise if asked, then **copy the Web app URL** (it ends in **/exec**).

**Part D — point the app at it:**

1. On your phone, open FinanceOS → tap **⚙️**.
2. In the first box, paste the **Web app URL** (instead of the Sheet ID). A new token box appears.
3. Paste the token from **CONFIG cell B8** into it. Next → Next → Open Dashboard.
4. Tap **⟳** — you should see "Live data loaded from your Sheet ✓".

**Part E — now actually lock the Sheet:**

1. In your Sheet: **Share** → General access → back to **Restricted** → Done.
2. **File** → **Share** → **Publish to web** → **Stop publishing**.
3. Tap **⟳** in the app once more to confirm it still works. 🎉

Your Sheet is now invisible to everyone but you. If you ever suspect the token leaked, just run **🔐 Generate App Token** again and paste the new one into ⚙️ Settings — the old one stops working instantly. One more habit: if you later update the script's code, the web app keeps running the old version until you go to **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**.

---

## 9. Turn on the Predict engine (subscriptions, card EMIs, forecasts)

The Predict tab and the new Debt sections read three extra tabs. To add them to your **existing** sheet without losing any data:

1. Go to **script.google.com**, open your FinanceOS project, replace the code with the latest **Code.gs**, and Save.
2. Open your **FinanceOS Data** sheet → menu **💰 FinanceOS** → **➕ Add new data tabs**. This adds three tabs and skips anything already there:
   - **SUBSCRIPTIONS** — your recurring services (Netflix, ChatGPT Plus…). Columns: Service, Amount, Cycle (Monthly/Annual), Billing Day, Category.
   - **CARD_EMIS** — EMIs running on a credit card. Columns: Card (must match the card name in LIABILITIES), Description, Monthly Installment, Total Months, Months Paid, Interest Rate %.
   - **CARD_SWIPES** — one-time spends this billing cycle. Columns: Card, Description, Amount, Date. (Clear this tab after each statement.)
3. Fill in your real rows (sample rows are pre-filled to show the format).
4. **Redeploy the web app so the app can see the new tabs:** Deploy → **Manage deployments** → ✏️ pencil → Version: **New version** → Deploy. (The URL stays the same.)
5. In the app, tap **⟳**. The Debt tab now shows Loans / Cards / Subscriptions, card rows expand to show EMIs & swipes, and the **Predict** tab fills in.

**Core Fixed Expenses** (no sheet needed): at the top of the Predict tab, type your monthly Rent, Utilities, Transport, and Groceries into the glass input fields. They save on your phone instantly (no Apps Script, no redeploy) and feed the runway, safe-to-spend, and expense forecasts below. Note: a panic-wipe clears these along with everything else.

The Predict tab has two views (tap the switch up top): **🔮 Forecasts** and **💸 Cash Flow**.
- **Card bill this cycle** (Forecasts) — each card's bill split into regular spend (pay in full) + this month's EMI installment.
- **Debt Freedom Countdown** — the month you hit zero debt; tap ＋/− to test extra monthly prepayment, and switch Avalanche/Snowball.
- **Jobless Runway** — months your available cash covers core expenses + loan EMIs + subscriptions if income stopped.
- **Safe-to-Spend** (Cash Flow) — checking balance minus rent, utilities, active EMIs, and subscriptions still due this cycle.
- **Next Month's Expense Forecast** — core expenses + EMIs + subscriptions + recent variable card spend.

**Logging a bill:** in the Debt tab, tap a card → "Log a bill payment" → enter what you paid → Mark paid. It subtracts that from the card's outstanding, advances any EMI by one month, and rolls the due date forward — all written to your Sheet (needs private/web-app mode).

**Emergency fund** counts only instantly-available money (bank savings + liquid funds). Your FD shows as a "backup" tier and isn't part of the SAFE score, since it takes a day or two to break. Investments (equity, gold, PPF, RD) are never counted — only the ACCOUNTS tab.

---

## If something doesn't look right

- **Numbers don't update** — tap the **⟳ refresh icon** in the top bar; the app will tell you exactly what's wrong. The two usual causes: the Sheet ID was copied from the publish popup (starts with 2PACX — use the address-bar ID instead, see step 2), or sharing wasn't set to "Anyone with the link · Viewer" (see step 3, Part A).
- **"Could not load live data"** — check both parts of step 3 were completed, and that the Sheet ID in ⚙️ Settings is correct.
- **"AI unavailable"** — check the Gemini key in ⚙️ Settings. Keys are long; make sure the whole thing was pasted.
- **Gold price shows ₹7,240 always** — the live gold-rate website occasionally blocks lookups. Just type the current rate into the GOLD tab's "LIVE GOLD RATE/g" cell.
- **No morning email** — open your Sheet once; the menu and triggers activate the first time the sheet is opened after setup. Also check spam for the first email.

That's it. Your money, your sheet, your phone. 🚀
