# 📬 Automated Email Attachment Organizer with Daily Dashboard

> Python 3 RPA project for macOS — automatically fetches Gmail attachments,
> organises them into category folders, logs every action to Excel, and
> visualises everything on a live Neumorphism-styled Streamlit dashboard.

---

## ✅ Feature Overview

| Phase | Feature | Status |
|---|---|---|
| 1 | Local file organiser + Excel logging | ✅ Done |
| 2 | Gmail IMAP integration (Inbox + Spam) | ✅ Done |
| 3 | Streamlit dashboard (Neumorphism UI) | ✅ Done |
| 4 | Scheduled automation (`--schedule`) | ✅ Done |

---

## 📁 Project Structure

```
rpa_email_dashboard/
├── bot.py                  # Core RPA logic — organiser + Gmail + scheduler
├── dashboard.py            # Streamlit dashboard (Neumorphism UI)
├── config.py               # Shared paths, extension map, column definitions
├── test_imap.py            # IMAP diagnostic / debug script
├── .env                    # Gmail credentials (⚠️ never commit)
├── requirements.txt        # Python dependencies
│
├── downloads/              # Organised attachments (auto-created)
│   ├── pdf/
│   ├── word/
│   ├── excel/
│   ├── image/
│   └── others/
│
├── reports/
│   └── daily_report.xlsx   # Auto-created Excel log
│
├── temp_downloads/         # Staging area for Gmail downloads (auto-cleaned)
└── test_files/             # Drop files here to simulate attachments
```

---

## ⚙️ Setup

### 1. Clone / open the project folder
```bash
cd rpa_email_dashboard
```

### 2. Create and activate a virtual environment
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure Gmail credentials

Create a `.env` file (copy from template):
```bash
cp .env.example .env   # or create it manually
```

Edit `.env`:
```env
EMAIL_ADDRESS=your_gmail@gmail.com
EMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx   # 16-char Google App Password
EMAIL_SEARCH_KEYWORD=RPA_TEST            # subject keyword to search for
```

> **How to get an App Password:**
> 1. Enable **2-Step Verification** on your Google account
> 2. Go to [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
> 3. Generate a password for **Mail / Mac**
> 4. Paste the 16-character code into `.env`
>
> Also make sure **IMAP is enabled** in Gmail:
> Gmail → ⚙️ Settings → See all settings → Forwarding and POP/IMAP → Enable IMAP

---

## 🚀 Running the Project

### Run bot once (organise local files + fetch Gmail)
```bash
venv/bin/python3 bot.py
```

### Run bot on a schedule (automatic, no cron needed)
```bash
# Every 30 minutes (default)
venv/bin/python3 bot.py --schedule

# Every 60 minutes
venv/bin/python3 bot.py --schedule --interval 60

# Every 15 minutes
venv/bin/python3 bot.py --schedule --interval 15
```
Press `Ctrl+C` to stop the scheduler.

### Run the dashboard
```bash
streamlit run dashboard.py
# then open http://localhost:8501
```

### Debug IMAP connection
```bash
venv/bin/python3 test_imap.py
```
This prints your mailbox list, tests folder access, and counts matching emails.

### Deactivate virtual environment
```bash
deactivate
```

---

## 🤖 How `bot.py` Works

```
Start
  │
  ├─ [1] Create folders (downloads/, reports/, temp_downloads/, test_files/)
  ├─ [2] Create Excel report if missing
  ├─ [3] Generate sample test files in test_files/
  ├─ [4] Organise local test files → downloads/<type>/
  └─ [5] Gmail IMAP
           │
           ├─ Connect to imap.gmail.com with App Password
           ├─ Auto-detect Spam folder via \Junk IMAP flag
           │   (works for any Gmail language — Thai, Chinese, etc.)
           ├─ Search INBOX for emails matching EMAIL_SEARCH_KEYWORD
           ├─ Search Spam folder for same keyword
           ├─ Download all attachments → temp_downloads/
           ├─ Move each file → downloads/<type>/
           └─ Log every result → reports/daily_report.xlsx
```

### Functions in `bot.py`

| Function | Purpose |
|---|---|
| `create_folders()` | Creates all required directories |
| `create_report_if_missing()` | Creates styled Excel report if not found |
| `get_file_type(file_name)` | Maps file extension → category |
| `write_log(...)` | Appends one styled row to the Excel report |
| `move_file_to_category(src, ...)` | Moves file + writes log |
| `create_sample_test_files()` | Generates 12 sample files for demo |
| `process_test_files()` | Processes every file in `test_files/` |
| `_find_spam_folder(mail)` | Auto-detects Gmail spam folder name via `\Junk` flag |
| `_process_mailbox(mail, mailbox)` | Searches one mailbox and downloads attachments |
| `download_gmail_attachments()` | Orchestrates IMAP connection + all mailbox searches |
| `gmail_job()` | Lightweight job used by the scheduler (skips one-time setup) |
| `main()` | Full single-run entry point |

---

## 📊 Excel Report Columns

| Column | Description |
|---|---|
| `DATE` | Date file was processed |
| `TIME` | Time file was processed |
| `EMAIL SUBJECT` | Subject of the source email (or "Test Run" for local files) |
| `SENDER` | Sender email address |
| `FILE NAME` | Original file name |
| `FILE TYPE` | Category: pdf / word / excel / image / others |
| `SAVE PATH` | Full path where file was saved |
| `STATUS` | `SUCCESS` or `FAILED` |
| `ERROR MESSAGE` | Error detail if status is FAILED |

---

## 🗂 File Type Mapping

| Category | Extensions |
|---|---|
| `pdf` | `.pdf` |
| `word` | `.doc`, `.docx` |
| `excel` | `.xls`, `.xlsx`, `.csv` |
| `image` | `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.webp` |
| `others` | everything else |

---

## 📈 Dashboard (`dashboard.py`)

**Design:** Neumorphism (Soft UI) — light gray + coral pink accents

**Features:**
- Date range filter (sidebar)
- Metric cards: Total Files · Success · Failed · PDFs · Images
- Bar chart: files by type
- Donut chart: success vs failed split
- Trend line: daily activity over time
- Log table with search / filter
- Manual refresh button (clears cache)

**Customise theme** — edit these variables near the top of `dashboard.py`:
```python
ACCENT      = "#f4896b"   # coral / salmon accent colour
BG          = "#e8ecf1"   # main background
SHADOW_DARK = "#b2bac5"   # dark side of neumorphic shadow
SHADOW_LITE = "#ffffff"   # light side of neumorphic shadow
```

---

## 🔒 Security Notes

- **Never commit `.env`** — it contains your Gmail App Password.
  `.env` is listed in `.gitignore`.
- The bot uses **read-only** IMAP access — it does not delete or move emails.
- Passwords are never printed or logged.

---

## 🛠 Dependencies

```
openpyxl       # Excel read/write
pandas         # Data manipulation
streamlit      # Dashboard UI
altair         # Charts
python-dotenv  # .env loading
schedule       # Job scheduler
watchdog       # Streamlit file watcher
```

Install all with:
```bash
pip install -r requirements.txt
```

---

## 📋 Quick Reference

```bash
# Activate venv
source venv/bin/activate

# Single run
venv/bin/python3 bot.py

# Scheduled (every 30 min)
venv/bin/python3 bot.py --schedule

# Scheduled (custom interval)
venv/bin/python3 bot.py --schedule --interval 60

# Start dashboard
streamlit run dashboard.py

# Debug IMAP
venv/bin/python3 test_imap.py

# Deactivate venv
deactivate
```
