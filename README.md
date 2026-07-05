# Attainment Register — CO‑PO Mapping Tool

A shared web tool for tagging exam questions to Course Outcomes (CO) and
Bloom's levels, uploading student marks, and auto-computing CO and PO/PSO
attainment for accreditation reporting.

This is the standalone version — it saves data to a real shared database
(Firebase Firestore, free tier) instead of Claude's built-in storage, so it
can be hosted on GitHub Pages and used by your whole faculty from any device.

---

## 1. Create a free Firebase project (5 minutes)

1. Go to https://console.firebase.google.com and sign in with any Google account.
2. Click **Add project**, give it a name (e.g. `copo-register`), finish the wizard.
3. In the left sidebar, click **Build → Firestore Database → Create database**.
   Choose **Start in test mode** for now (see the security note below), pick
   any region close to you.
4. Click the gear icon → **Project settings**. Under "Your apps", click the
   **</>** (web) icon to register a new web app. Give it any nickname.
5. Firebase will show you a `firebaseConfig` object. Copy it.
6. Open `src/firebase.js` in this project and paste your values in, replacing
   the placeholders (`YOUR_API_KEY`, `YOUR_PROJECT_ID`, etc).

### Security note (important)

Firestore's "test mode" allows **anyone** to read/write your database for 30
days, then locks it. Because this app has no real login system, before
sharing the link widely, go to **Firestore → Rules** and set something like:

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /copo_store/{document=**} {
      allow read, write: if true; // open — fine for a small trusted faculty group
    }
  }
}
```

This keeps the data open only to people who have the link, same trust model
as the original prototype. If you want real per-user login later (so only
verified faculty accounts can write), tell your developer/Claude to add
Firebase Authentication — it's a small follow-on step, not a rebuild.

---

## 2. Run it locally

```bash
npm install
npm run dev
```

Open the printed `localhost` URL. Sign in, create a course, try uploading a
sample marks Excel sheet, and confirm attainment numbers look right before
deploying.

---

## 3. Put it on GitHub

```bash
git init
git add .
git commit -m "Initial commit: CO-PO attainment register"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
git push -u origin main
```

---

## 4. Deploy to GitHub Pages

1. Open `vite.config.js` and change `base: "/your-repo-name/"` to match your
   actual repo name exactly (e.g. `base: "/copo-register/"`).
   - **Exception:** if this repo is named `YOUR_USERNAME.github.io` (a GitHub
     user/org homepage repo), set `base: "/"` instead.
2. Install the deploy helper and publish:

   ```bash
   npm install
   npm run deploy
   ```

   This builds the app and pushes it to a `gh-pages` branch.
3. On GitHub: **Settings → Pages → Build and deployment → Source**, choose
   **Deploy from a branch**, branch = `gh-pages`, folder = `/ (root)`. Save.
4. After a minute, your site is live at:
   `https://YOUR_USERNAME.github.io/YOUR_REPO_NAME/`

Share that link with your faculty. Everyone who opens it shares the same
Firestore database, so marks, CO-PO matrices, and reports stay in sync
across the department automatically.

---

## What's different from the Claude-artifact version

- Data is stored in **Firebase Firestore** (`src/storage.js`) instead of
  Claude's `window.storage` — this is the only functional change needed to
  make it work outside Claude.
- Everything else (UI, attainment formulas, Excel upload, CO-PO matrix,
  reports) is identical.

## Updating the app later

Edit files under `src/`, then re-run `npm run deploy` to publish the new
version to the same link.
