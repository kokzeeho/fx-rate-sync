# FX Rate Auto-Update — Setup Guide (GitHub version)

No Azure, no API keys, no admin consent. Three steps.

---

## Step 1 — Create a GitHub repo

Create a new repo (private is fine), e.g. fx-rate-sync.

## Step 2 — Add the files

Copy in from this package:
- fx_update.py
- requirements.txt
- .github/workflows/fx-update.yml

Then rename your Excel file to SGDUSD_Tracker.xlsx and add it to the repo
root (same folder as fx_update.py). If you want a different filename, update
the FILE_PATH variable at the top of fx_update.py to match.

## Step 3 — Push

    git add .
    git commit -m "Set up FX rate automation"
    git push

That's it. No secrets to configure — GitHub Actions already has permission to
commit back to its own repo.

---

## What it does daily

1. At midnight GMT+8, GitHub spins up a runner, pulls SGD/USD, SGD/EUR, USD/RMB
   from Frankfurter (free, ECB rates, no key needed).
2. Opens SGDUSD_Tracker.xlsx, finds today's row in sheet "new", writes the
   3 rates into columns B, C, D.
3. Commits the updated file back to the repo.

You'll see a new commit appear in the repo every day. Pull the repo (or just
download the file from GitHub) whenever you want the latest copy on your Mac.

## To test it immediately

Go to the repo's Actions tab -> Daily FX Rate Update -> Run workflow.
This triggers it manually instead of waiting for midnight.

## Known limitation

Frankfurter (ECB) doesn't update on weekends/EU holidays — it'll write the
same rate as the prior day on those dates. Flag if you'd rather skip weekend
writes entirely.

---

## Keeping a synced copy in OneDrive

This makes the file also appear in your OneDrive automatically — using a
local scheduled job, no Azure setup needed.

Important: the git repo itself must live on plain local disk, NOT inside
the OneDrive folder. OneDrive (and iCloud Drive's "Desktop & Documents" sync)
virtualizes files on demand, and background processes can't reliably access
a .git folder living inside that kind of cloud-synced directory — this
caused daily sync failures ("Operation not permitted") in the original setup.
The fix: clone the repo somewhere plain (e.g. ~/fx-rate-sync-automation),
and have the script copy just the finished .xlsx file into OneDrive
afterward. Git operations touch stable local disk only; OneDrive only ever
sees a simple file copy, which it handles fine.

### Step 4 — Clone the repo to a plain local folder

    git clone https://github.com/YOUR_USERNAME/fx-rate-sync.git ~/fx-rate-sync-automation

Do NOT clone this into ~/OneDrive/... or ~/Desktop (if Desktop sync is on)
or any other iCloud/OneDrive-virtualized path.

### Step 5 — Add a login-triggered local pull job

1. Copy pull_fx_repo.sh to ~/Scripts/pull_fx_repo.sh. It should:
   - cd into the local mirror (~/fx-rate-sync-automation)
   - run git pull origin main
   - cp the resulting .xlsx into your OneDrive folder path
   - log all output to ~/Scripts/fx-pull.log
2. Make it executable:

       chmod +x ~/Scripts/pull_fx_repo.sh

3. Copy com.heritage.fxpull.plist to ~/Library/LaunchAgents/.
4. In the plist, use RunAtLoad: true (NOT StartCalendarInterval). This
   triggers the pull every time you log in or start up your Mac, rather than
   at a fixed midnight time — avoiding the case where the Mac is off or
   asleep at the scheduled hour and silently skips a day.
5. Load the job:

       launchctl load ~/Library/LaunchAgents/com.heritage.fxpull.plist

### Limitation

The file updates in OneDrive only when you next log in / start up your Mac
after GitHub's midnight run has committed. If you don't turn your Mac on for
several days, OneDrive will simply be a few days behind until you do — the
GitHub repo itself is always current regardless.

### Monthly manual step (unchanged)

Column E (退税, export tax rebate %) is manual-only — the script never
touches it. Update it about once a month via: pull in GitHub Desktop -> edit
column E in Excel -> save -> commit -> push.
