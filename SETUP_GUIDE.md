# FX Rate Auto-Update — Setup Guide (GitHub version)

No Azure, no API keys, no admin consent. Three steps.

---

## Step 1 — Create a GitHub repo

Create a new repo (private is fine), e.g. `fx-rate-sync`.

## Step 2 — Add the files

Copy in from this package:
- `fx_update.py`
- `requirements.txt`
- `.github/workflows/fx-update.yml`

Then rename your Excel file to `SGDUSD_Tracker.xlsx` and add it to the repo
root (same folder as `fx_update.py`). If you want a different filename, update
the `FILE_PATH` variable at the top of `fx_update.py` to match.

## Step 3 — Push

```
git add .
git commit -m "Set up FX rate automation"
git push
```

That's it. No secrets to configure — GitHub Actions already has permission to
commit back to its own repo.

---

## What it does daily

1. At midnight GMT+8, GitHub spins up a runner, pulls SGD/USD, SGD/EUR, USD/RMB
   from Frankfurter (free, ECB rates, no key needed).
2. Opens `SGDUSD_Tracker.xlsx`, finds today's row in sheet `new`, writes the
   3 rates into columns B, C, D.
3. Commits the updated file back to the repo.

You'll see a new commit appear in the repo every day. Pull the repo (or just
download the file from GitHub) whenever you want the latest copy on your Mac.

## To test it immediately

Go to the repo's **Actions** tab → **Daily FX Rate Update** → **Run workflow**.
This triggers it manually instead of waiting for midnight.

## Known limitation

Frankfurter (ECB) doesn't update on weekends/EU holidays — it'll write the
same rate as the prior day on those dates. Flag if you'd rather skip weekend
writes entirely.

---

## Optional: keep a synced copy in OneDrive

This makes the file also appear in your OneDrive automatically, without any
Azure setup — using a local scheduled pull instead.

### Step 4 — Clone the repo into your OneDrive folder

On your Mac:

```
cd ~/OneDrive       # or wherever your OneDrive folder is
git clone https://github.com/YOUR_USERNAME/fx-rate-sync.git
```

This creates `~/OneDrive/fx-rate-sync/SGDUSD_Tracker.xlsx`. OneDrive will sync
this file to the cloud like any other file in that folder — but only when it
changes *locally*, which is what the next step handles.

### Step 5 — Add a daily local pull job

1. Copy `pull_fx_repo.sh` to your home folder (`~/pull_fx_repo.sh`).
2. Open it and confirm `REPO_DIR` matches where you cloned the repo in Step 4.
3. Make it executable:
   ```
   chmod +x ~/pull_fx_repo.sh
   ```
4. Copy `com.heritage.fxpull.plist` to `~/Library/LaunchAgents/`.
5. Open it and replace both instances of `YOUR_USERNAME` with your actual Mac
   username.
6. Load the job:
   ```
   launchctl load ~/Library/LaunchAgents/com.heritage.fxpull.plist
   ```

This runs `git pull` every day at 00:10 (local time) — 10 minutes after the
GitHub Actions job finishes — pulling the fresh commit into your OneDrive
folder. OneDrive then syncs it like any normal file change.

### Limitation

Only works while your Mac is on. If it's asleep at 00:10, it catches up the
next time you're online — but there's a gap until then. If that's not
acceptable, that's the point where the Azure/Graph API push becomes worth
the extra setup.
