# Deploy to Render (free, gives you a public URL anyone can open)

This guide assumes you've **never used GitHub or Render before**. Total time: ~15 minutes.

## ⚠️ Before you start

This app handles patient data (names, DOBs, insurance info). Hosting it in the cloud means that data leaves your computer. Please confirm with your hospital that this is OK. The app is protected by a password (set below), but a password is not the same as a legal/compliance review.

---

## Step 1 — Create a GitHub account (skip if you have one)

1. Go to <https://github.com/signup>
2. Use a work email, pick a username, verify the email.

## Step 2 — Create a GitHub repository

1. Go to <https://github.com/new>
2. **Repository name:** `sigorta-hesabati`
3. **Privacy:** choose **Private** (important — this is a healthcare app)
4. Leave everything else as-is. Click **Create repository**.
5. Keep this browser tab open — you'll need the URL it shows.

## Step 3 — Push your code to GitHub

Open Terminal and run these commands one by one (paste each line, press Enter, wait for it to finish):

```bash
cd ~/saf_hospital/flask_app

# Only the first time you ever use git on this Mac:
git config --global user.name "Your Name"
git config --global user.email "you@example.com"

git init -b main
git add .
git commit -m "Initial commit"

# Replace YOUR-USERNAME with your actual GitHub username:
git remote add origin https://github.com/YOUR-USERNAME/sigorta-hesabati.git
git push -u origin main
```

When it asks for a password, **don't type your GitHub password** — GitHub doesn't accept those anymore. Instead:
1. Go to <https://github.com/settings/tokens/new>
2. **Note:** `sigorta-hesabati deploy`
3. **Expiration:** 90 days
4. **Scopes:** check `repo`
5. Click **Generate token** at the bottom, copy the long string starting with `ghp_…`
6. Paste it as the password.

Refresh your GitHub repo tab — you should see all the files (`app.py`, `templates/`, etc.).

## Step 4 — Sign up for Render

1. Go to <https://render.com/>
2. Click **Get Started** → **Sign up with GitHub** (easiest).

## Step 5 — Create the web service

1. From the Render dashboard, click **New +** → **Web Service**.
2. Click **Connect** next to your `sigorta-hesabati` repo. (If you don't see it, click **Configure GitHub App** and grant Render access to that repo.)
3. Render will read `render.yaml` from your repo and auto-fill most fields. Confirm:
   - **Name:** `sigorta-hesabati` (or anything; this becomes part of your URL)
   - **Region:** Frankfurt (closest to Azerbaijan)
   - **Branch:** `main`
   - **Runtime:** Python 3
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn -w 2 -b 0.0.0.0:$PORT --timeout 180 --access-logfile - app:app`
   - **Instance type:** Free
4. Scroll down to **Environment Variables** and add:
   - **Key:** `APP_PASSWORD`
   - **Value:** pick a strong password (write it down — you'll share it with your colleagues)
5. Click **Create Web Service**.

Render will now build and deploy. Watch the log — after 3–5 minutes you'll see `Your service is live 🎉`. Your URL is shown at the top of the page, something like:

```
https://sigorta-hesabati.onrender.com
```

## Step 6 — Test it

1. Open the URL in your browser.
2. A login box pops up. Enter **any username** (Render's basic auth ignores it) and the password you set above.
3. Upload a test `.xlsx` file. Make sure preview + downloads work.

## Step 7 — Share with your team

Send them two things:
- The URL: `https://sigorta-hesabati.onrender.com`
- The password

They open the URL in any browser on any computer (Mac, Windows, phone). No installation needed.

---

## Things to know about the free tier

- **Cold start:** if no one uses the app for 15 minutes, Render puts it to sleep. The next visit takes ~30 seconds to wake up. Subsequent visits are instant. This is fine for occasional use.
- **Memory:** 512 MB. Your ~63 000-row sample file processes in well under that, but very large files (200 000+ rows) might hit the limit. If that happens, upgrade to the $7/month Starter plan.
- **Cost:** $0 for free tier. No credit card needed.

## Updating the app later

After changing any file:

```bash
cd ~/saf_hospital/flask_app
git add .
git commit -m "what changed"
git push
```

Render detects the push and redeploys automatically (1–2 minutes).

## Changing the password

Render dashboard → your service → **Environment** → edit `APP_PASSWORD` → **Save Changes**. The service restarts (~30 sec) with the new password.

## Stopping the service

Render dashboard → your service → **Settings** → scroll down → **Suspend Web Service**.
