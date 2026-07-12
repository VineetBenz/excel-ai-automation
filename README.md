# Excel AI Expense Analyser 

> Wire AI directly into Excel using VBA — no external tools, no dedicated app, no backend.
> Just a spreadsheet that thinks.

---

## What This Does

Add an expense row in columns A to D.
Run the macro.
The AI fills in columns E to K automatically.

**Before running the macro:**

| Date | Description | Amount | Payment Method | Category | Necessity | Budget Status | Saving Suggestion | Monthly Trend | Financial Advice | Savings Potential |
|---|---|---|---|---|---|---|---|---|---|---|
| 07-Jul-26 | Swiggy | 450 | UPI | | | | | | | |

**After running the macro:**

| Date | Description | Amount | Payment Method | Category | Necessity | Budget Status | Saving Suggestion | Monthly Trend | Financial Advice | Savings Potential |
|---|---|---|---|---|---|---|---|---|---|---|
| 07-Jul-26 | Swiggy | 450 | UPI | Food | Discretionary | Within Budget | Cook at home | Increasing | Reduce food delivery | 25% |

---

## Files in This Repo

| File | What it is |
|---|---|
| `Module1_Main.bas` | Main VBA module — paste into Insert > Module |
| `Sheet_AutoDetect.bas` | Sheet module — paste into Sheet > View Code |
| `screenshots/before.png` | Sheet before running macro |
| `screenshots/after.png` | Sheet after AI fills everything in |

---

## How to Set Up — Step by Step

### Step 1 — Create a new Excel file

Open Excel and create a new file.

Set up **Row 1** with these exact headers:

| A | B | C | D | E | F | G | H | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| Date | Description | Amount | Payment Method | Category | Necessity | Budget Status | Saving Suggestion | Monthly Trend | Financial Advice | Savings Potential |

### Step 2 — Save as macro-enabled

```
File → Save As
→ File type: Excel Macro-Enabled Workbook (.xlsm)
→ Name it anything you like
→ Click Save
```

### Step 3 — Open the VBA Editor

```
Press Alt + F11
```

### Step 4 — Paste the main module

```
In VBA Editor:
1. Click Insert (top menu)
2. Click Module
3. A new Module1 window opens
4. Open Module1_Main.bas from this repo
5. Copy everything and paste it into Module1
```

### Step 5 — Add your API key

Find this line at the top of Module1:

```vba
Const API_KEY As String = "YOUR_API_KEY_HERE"
```

Replace `YOUR_API_KEY_HERE` with your actual API key.

**Where to get a free API key:**
- Gemini (default): [aistudio.google.com](https://aistudio.google.com) → Get API Key → Free tier available
- OpenAI: [platform.openai.com](https://platform.openai.com) → See OpenAI switch instructions below

### Step 6 — Paste the sheet module

```
In VBA Editor:
1. Look at the left panel
2. Find your sheet — usually called Sheet1
3. Double-click it
4. Open Sheet_AutoDetect.bas from this repo
5. Copy everything and paste it into that window
```

### Step 7 — Add some sample data

Add a few rows of expense data in columns A to D.
Leave columns E to K completely empty — the macro fills those.

Example data:

| Date | Description | Amount | Payment Method |
|---|---|---|---|
| 07-Jul-26 | Swiggy | 450 | UPI |
| 07-Jul-26 | Netflix | 649 | Credit Card |
| 07-Jul-26 | Apollo Pharmacy | 320 | Debit Card |
| 07-Jul-26 | Uber | 180 | UPI |
| 07-Jul-26 | Big Bazaar Groceries | 1200 | Cash |

### Step 8 — Run the macro

```
Press Alt + F8
→ Select "first"
→ Click Run
```

Watch columns E to K fill in automatically. ✅

---

## How to Switch from Gemini to OpenAI

Open `Module1_Main.bas` and find the `AIRequest` function.

Comment out the Gemini lines and uncomment the OpenAI lines:

```vba
' GEMINI — comment this out
' URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" & API_KEY
' JSON = "{""contents"":[{""parts"":[{""text"":""" & EscapeJSON(Prompt) & """}]}]}"

' OPENAI — uncomment these
URL = "https://api.openai.com/v1/chat/completions"
JSON = "{""model"":""gpt-4o-mini"",""messages"":[{""role"":""user"",""content"":""" & EscapeJSON(Prompt) & """}]}"
```

Also uncomment this line just below `Http.setRequestHeader "Content-Type"`:

```vba
Http.setRequestHeader "Authorization", "Bearer " & API_KEY
```

---

## How to Adapt for Your Own Use Case

The architecture works for any AI-powered Excel workflow.
Only the prompts need to change.

1. Open `Module1_Main.bas`
2. Find `BuildClassificationPrompt` and `BuildAdvicePrompt`
3. Replace the prompt text with your own instructions
4. Update the column numbers to match your sheet layout

Everything else stays exactly the same.

---

## Column Reference

| Column | Letter | Filled By |
|---|---|---|
| 1 | A | You — Date |
| 2 | B | You — Description |
| 3 | C | You — Amount |
| 4 | D | You — Payment Method |
| 5 | E | AI — Category |
| 6 | F | AI — Necessity |
| 7 | G | AI — Budget Status |
| 8 | H | AI — Saving Suggestion |
| 9 | I | AI — Monthly Trend |
| 10 | J | AI — Financial Advice |
| 11 | K | AI — Savings Potential |

---

## Smart Features Built In

| Feature | What it does |
|---|---|
| Skip already processed rows | Runs only on empty rows — safe to re-run anytime |
| Auto retry on rate limit | Retries up to 3 times if API returns 429 |
| Auto clear on edit | If you edit columns A–D, results clear automatically |
| Reprocess single row | Run `ReprocessSelectedRow` to redo just one row |
| Debug log | Press `Ctrl + G` in VBA Editor to see full logs |

---

## Requirements

- Microsoft Excel with macros enabled
- Windows OS (uses native `MSXML2.XMLHTTP`)
- Internet connection
- A free API key from Gemini or OpenAI

---

## Important

> ⚠️ Never paste your real API key into this repo or share it publicly.
> Always keep `YOUR_API_KEY_HERE` as the placeholder in any file you upload.

---

## Read the Full Article

https://medium.com/@vineetbhandari1919/turned-excel-into-an-ai-brain-no-external-tools-no-apps-just-vba-any-ai-api-97d5cf7cd5f8

---

## License

MIT — use it, adapt it, build on it freely.
