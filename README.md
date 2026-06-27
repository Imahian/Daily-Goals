# 📅 Daily Goals Tracker
> A Google Sheets + Google Apps Script system that turns a weekly emoji-based routine into a dynamic daily task list — with Google Calendar sync and efficiency tracking.

---

## ✨ What it does

- Reads your weekly **Routine** table and auto-populates a **Daily** task list every morning at midnight
- Groups consecutive emoji blocks into military-time activity ranges (e.g. `10:00 - 12:00`)
- Syncs checkboxes **bidirectionally** with Google Calendar / Google Tasks
- Saves a nightly snapshot into an **Average** table with per-category tracking and a daily efficiency percentage
- Auto-updates when you add new activity categories to the **Reference** table

---

## 🗂 Sheet Structure

You need **3 sheets** inside one Google Spreadsheet:

| Sheet name | Contains |
|---|---|
| `Schedule` | The `Routine` table — your full weekly emoji schedule |
| `Daily` | The `Daily` table — today's auto-generated task list |
| `Metrics` | The `Reference` table + the `Average` table |

---

## 📋 Tables & Named Ranges

Every table is identified by a **named range** pointing to its header row only. This makes the system position-independent — you can place each table anywhere on its sheet.

### 1. Routine — `Routine_enc`

Located on the **Schedule** sheet.

| Column | Description |
|---|---|
| `Time` | Clock emojis (🕛🕧🕐🕜…) — one per 30-min slot, 00:00 → 23:30 (48 rows) |
| `Sunday` | Activity emoji for that slot on Sunday |
| `Monday` | Activity emoji for that slot on Monday |
| `Tuesday` | … |
| `Wednesday` | … |
| `Thursday` | … |
| `Friday` | … |
| `Saturday` | Activity emoji for that slot on Saturday |

**How to fill it:** place the same emoji in consecutive rows to define a block. The system reads each day's column and groups consecutive identical emojis into a single time block.

```
Time   Monday
🕙     💻        ← Work block starts at 10:00
🕥     💻
🕚     💻
🕦     💻        ← last slot at 11:30 → block ends at 12:00
🌞     🌞        ← decorative separator, ignored automatically
🕧     🥪        ← Lunch starts at 12:00
```

> **Important:** Decorative emojis like 🌞 are automatically ignored as long as they are **not registered** in the Reference table.

---

### 2. Daily — `Daily_enc`

Located on the **Daily** sheet.

| Column | Description |
|---|---|
| `Activity` | Auto-filled: `emoji Category` (e.g. `💻 Work`) |
| `Range` | Auto-filled: military time range (e.g. `10:00 - 12:00`) |
| `Check` | Checkbox — mark as done; syncs with Google Calendar |

**The header row must be pre-formatted** with checkbox validation in the `Check` column before running the script for the first time. The script resets values to `false` but does not create the checkbox format itself.

> **Note:** If you have a formula (e.g. a progress bar) below the Daily table, set `CFG.DAILY_PROTECTED_FROM_ROW` in the script to that row number. The script will never write to or below that row.

---

### 3. Reference — `Reference_enc`

Located on the **Metrics** sheet.

| Column | Description |
|---|---|
| `Icon` | The emoji that appears in the Routine table |
| `Category` | Human-readable name shown in the Daily table and Average headers |

**This is the master registry.** Only emojis listed here will be picked up from the Routine table. Any emoji not listed is treated as a decorator and ignored.

When you add a new row here, the `onEdit` trigger automatically:
- Adds a new column to Average
- Reloads the Daily table with the updated category name

```
Icon  Category
💻    Work
📚    Study
🏋🏾   Exercise
🧘🏾‍♂️  Meditate
🥪    Lunch
🧹    Chores
```

---

### 4. Average — `Average_enc`

Located on the **Metrics** sheet.

| Column | Description |
|---|---|
| `Date` | Date of the snapshot (`yyyy-MM-dd`) |
| `[icon columns]` | One column per emoji in Reference — checkbox: ✅ if all blocks completed |
| `Efficiency` | `tasks completed / total tasks × 100%` |

Rows are inserted **newest on top** (just below the header). New emoji columns are inserted automatically as physical columns before `Efficiency`, keeping all data aligned.

> **Important:** Pre-format the emoji columns with **checkbox validation** before running the script. The script writes `true`/`false` values to trigger those checkboxes — it does not create the format itself.

---

## ⚙️ Setup

### Step 1 — Copy the script

1. Open your Google Spreadsheet
2. Go to **Extensions → Apps Script**
3. Delete the default empty function
4. Paste the full contents of `DailyGoalsTracker.gs`
5. Save (`Ctrl+S` / `Cmd+S`)

---

### Step 2 — Enable Google Tasks API

1. In the Apps Script editor, click **Services** (`+` icon on the left sidebar)
2. Find **Google Tasks API**
3. Click **Add**

> If you skip this step, the script falls back to Google Calendar events instead of Tasks. Everything still works, but tasks won't have a native checkbox on your phone.

---

### Step 3 — Create the named ranges

For each table, select **only the header row** and create a named range:

1. Select the header row of the table
2. Go to **Data → Named ranges**
3. Create the range with the exact name below:

| Named range | Points to |
|---|---|
| `Routine_enc` | Header row of the Routine table (Schedule sheet) |
| `Daily_enc` | Header row of the Daily table (Daily sheet) |
| `Reference_enc` | Header row of the Reference table (Metrics sheet) |
| `Average_enc` | Header row of the Average table (Metrics sheet) |

> The named range must cover **only the header row**, not the data rows. The script reads the header to locate each column dynamically.

---

### Step 4 — Configure the script

Open `DailyGoalsTracker.gs` and adjust the `CFG` object at the top:

```javascript
const CFG = {
  // Sheet names — must match exactly (case-sensitive)
  SHEET_SCHEDULE : "Schedule",
  SHEET_DAILY    : "Daily",
  SHEET_METRICS  : "Metrics",

  // Named ranges — must match exactly
  RANGE_ROUTINE   : "Routine_enc",
  RANGE_DAILY     : "Daily_enc",
  RANGE_REFERENCE : "Reference_enc",
  RANGE_AVERAGE   : "Average_enc",

  // Google Calendar ID
  // Use "primary" for your main calendar, or paste a specific calendar ID
  CALENDAR_ID : "primary",

  // Row number where your progress bar or any formula below Daily lives.
  // The script NEVER writes to this row or beyond.
  // Change this if you move the formula to a different row.
  DAILY_PROTECTED_FROM_ROW : 17,
};
```

---

### Step 5 — Install triggers

In the Apps Script editor:

1. Select the function `installTrigger` from the dropdown
2. Click **Run**
3. Authorize all permissions when prompted (Spreadsheets, Calendar, Tasks)

This installs **4 triggers**:

| Trigger | Time | What it does |
|---|---|---|
| `iniciarSistema` | 00:01 daily | Clears previous tasks, populates Daily, creates Calendar tasks |
| `iniciarSistema` | 23:50 daily | Saves snapshot to Average, clears Daily |
| `onEditTrigger` | On edit | Syncs checkboxes → Calendar; auto-updates when Reference changes |
| `syncCalendarToDaily` | Every 5 min | Polls Calendar/Tasks and reflects completion status back into Daily |

---

### Step 6 — Run it manually for the first time

From the custom menu **Daily Tracker** (appears in your sheet after authorizing):

1. Click **▶ Start system (full flow)**

You should see a sequence of toast notifications confirming each step.

---

## 🧪 Manual functions (via menu)

| Menu item | What it runs |
|---|---|
| Start system (full flow) | Full pipeline: clear → populate Daily → sync Calendar → check Average columns |
| Only: load today's Daily | Re-reads Routine and repopulates Daily |
| Only: sync Google Calendar | Recreates today's tasks/events in Calendar |
| Only: Daily → Calendar (checkboxes) | Pushes current checkbox state to Calendar |
| Only: Calendar → Daily (polling) | Pulls task completion state from Calendar into Daily |
| Simulate 23:55 snapshot → Average | Saves today's data to Average and clears Daily |
| Debug snapshot (check Logger) | Shows a breakdown in the Logger without saving anything |
| Diagnostic (toasts + Logger) | Validates named ranges, Reference icons, and today's detected blocks |
| Install triggers (run once) | Installs all 4 time-based and edit triggers |

---

## 📊 How Efficiency is calculated

At 23:55 each night, the script counts all activities in the Daily table:

```
Efficiency = (completed activities / total activities) × 100%
```

Each emoji category is tracked individually in Average. A category is marked ✅ only if **all blocks** of that type were completed during the day. The `Efficiency` column reflects the overall daily completion rate across all categories.

---

## 🔒 Privacy & Security

This script contains **no hardcoded credentials, tokens, or personal data**. The only user-specific value is `CALENDAR_ID : "primary"`, which refers to the signed-in Google account's default calendar — no actual ID is stored.

Each user who copies this project authorizes the script under their own Google account. No data leaves the user's own Google ecosystem.

---

## 🛠 Troubleshooting

**"Named range not found" error**
Make sure the named ranges exist and are spelled exactly as shown (case-sensitive). Go to Data → Named ranges to verify.

**No blocks detected for today**
Run **Diagnostic** from the menu and check the Logger (`Ctrl+Enter`). It shows which emojis were found in today's Routine column and whether they match any entry in Reference.

**Calendar tasks not appearing**
Make sure Google Tasks API is enabled under Services. If not, the script falls back to Calendar events — check Google Calendar on the web to confirm events were created.

**Progress bar formula gets cleared**
Verify that `CFG.DAILY_PROTECTED_FROM_ROW` is set to the correct row number in the script.

**Efficiency column out of alignment in Average**
This happens if columns were added manually instead of through the script. Run **Only: load today's Daily** to trigger `ensureAverageColumns`, which inserts physical columns and keeps Efficiency last.

---

## 📁 File structure

```
DailyGoalsTracker.gs    ← the full Apps Script (paste into Apps Script editor)
README.md               ← this file
```

---

## 📄 License

MIT — free to use, modify, and share.
