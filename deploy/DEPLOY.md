# Deploying `upload_listings.py` to a DigitalOcean droplet

This script is an **infinite polling loop** (polls every 5–10 min, runs forever).
On a server you want it to: start on boot, restart on crash, and persist its
logs. We do that with **systemd + journald**. The script also writes its own
rotated `logs/sync.log` file (see "Logs" below).

> **The one gotcha:** `get_worksheet()` calls `gspread.oauth()`, which opens a
> **web browser** on first run to authorize Google. A droplet has no browser, so
> you must authorize ONCE on your Windows machine and copy the resulting
> `authorized_user.json` token up to the server. Do **not** try to do first-run
> auth on the droplet.

---

## 0. Prepare on Windows (once)

1. Run the script locally and complete the Google login in the browser:
   ```powershell
   python upload_listings.py
   ```
   Let it finish one iteration, then `Ctrl-C`. You now have a valid
   `authorized_user.json` next to the script — this is the file that avoids the
   browser on the server.

2. Confirm `.env` exists with all secrets:
   `SP_API_REFRESH_TOKEN`, `LWA_APP_ID`, `LWA_CLIENT_SECRET`, `SELLER_ID`,
   `SPREADSHEET_ID`.

3. Confirm `DRY_RUN` in `upload_listings.py` is what you intend.
   `DRY_RUN = False` means it writes to Amazon **for real**.

4. Check the pricing looks right before the server ever pushes:
   ```powershell
   python upload_listings.py --rates 1.50
   ```
   Column G is the **wholesale cost in USD** (written by `fetch_esim_prices.py`),
   not the shelf price: the tier formula marks it up in GBP (with a
   £4-minimum-profit safeguard), the result is converted into each marketplace's
   currency, and only then rounded up to that currency's next `.99`. This command
   prints the whole breakdown and writes nothing (see "Pricing" in `CLAUDE.md`).

Files that must travel to the server:
`upload_listings.py`, `.env`, `oauth_client.json`, `authorized_user.json`.

---

## 1. Create + prepare the droplet

Smallest basic Ubuntu 22.04/24.04 droplet is plenty (this is a light script).
SSH in as root, then:

```bash
sudo apt update && sudo apt -y upgrade
sudo apt -y install python3 python3-venv python3-pip

# dedicated, non-root app user
sudo adduser --disabled-password --gecos "" youguide
sudo mkdir -p /opt/youguide
sudo chown youguide:youguide /opt/youguide
```

---

## 2. Copy files up (run from PowerShell on Windows, in the script folder)

```powershell
scp upload_listings.py .env oauth_client.json authorized_user.json `
    deploy/youguide-sync.service `
    root@YOUR_DROPLET_IP:/opt/youguide/
```

Then on the droplet, fix ownership and lock down secrets:

```bash
sudo chown youguide:youguide /opt/youguide/*
sudo chmod 600 /opt/youguide/.env /opt/youguide/oauth_client.json /opt/youguide/authorized_user.json
```

---

## 3. Install dependencies in a venv

```bash
sudo -u youguide -i
cd /opt/youguide
python3 -m venv venv
./venv/bin/pip install --upgrade pip
./venv/bin/pip install python-amazon-sp-api gspread google-auth
exit
```

Sanity check — should print logs and NOT try to open a browser. `Ctrl-C` after
one iteration:

```bash
sudo -u youguide /opt/youguide/venv/bin/python /opt/youguide/upload_listings.py
```

If it prints an auth URL / tries to open a browser, the token didn't transfer —
redo steps 0.1 and 2.

---

## 4. Install + start the systemd service

```bash
sudo cp /opt/youguide/youguide-sync.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now youguide-sync
sudo systemctl status youguide-sync
```

Common commands:

```bash
sudo systemctl restart youguide-sync   # after editing .env or the script
sudo systemctl stop youguide-sync      # pause syncing
sudo systemctl disable youguide-sync   # don't start on boot
```

After changing code/secrets, re-`scp` the file and `systemctl restart`.

---

## Exchange rates on the server

The loop converts the marked-up GBP price into each marketplace's currency, so the
droplet needs **outbound HTTPS** to the rate feeds (`open.er-api.com`,
`api.frankfurter.app`, `cdn.jsdelivr.net`) — no API key, no inbound ports. Rates
are fetched at most every `FX_TTL_HOURS` (12) and cached in
`/opt/youguide/fx_rates.json`, which the `youguide` user must be able to write
(it is, if you followed step 2's `chown`).

If every feed is unreachable the loop keeps running on the cached rates and logs
a warning; only if there is no cache at all does it skip the non-GBP markets that
pass (`NO_FX` in the status column) and try again on the next poll. Verify with:

```bash
sudo -u youguide /opt/youguide/venv/bin/python /opt/youguide/upload_listings.py --rates
```

Tune with `.env`: `AMAZON_FEE_PCT` (18) and `MIN_PROFIT` (4) for the margin
policy; `FX_MARKUP_PCT` (uplift %), `FX_ROUNDING` (`charm`/`nearest`),
`FX_TTL_HOURS`, or `FX_RATE_EUR=1.17`-style pins for a fixed rate. **Any of these
reprices listings on the next pass** — run `--rates` first, then restart.

---

## Logs

You have two log streams, both already set up:

### journald (from systemd — recommended for day-to-day)
```bash
journalctl -u youguide-sync -f             # live tail
journalctl -u youguide-sync --since today  # today only
journalctl -u youguide-sync -p warning     # warnings + errors only
```
Cap disk use by setting `SystemMaxUse=200M` in `/etc/systemd/journald.conf`,
then `sudo systemctl restart systemd-journald`.

### Rotated file (written by the script itself)
The script writes `logs/sync.log` (5 MB × up to 6 files ≈ 30 MB max, auto-rotated):
```bash
tail -f /opt/youguide/logs/sync.log
```
Download it to Windows for sharing (from PowerShell):
```powershell
scp root@YOUR_DROPLET_IP:/opt/youguide/logs/sync.log .
```
Override the location with the `LOG_DIR` env var if desired.

---

## Updating the app later

```powershell
# from Windows
scp upload_listings.py root@YOUR_DROPLET_IP:/opt/youguide/
```
```bash
# on the droplet
sudo chown youguide:youguide /opt/youguide/upload_listings.py
sudo systemctl restart youguide-sync
```
