# Kyen Sales Hub — Deployment Guide (Vercel + Supabase)

## STEP 1 — Set up Supabase (10 minutes)

1. Go to **supabase.com** → Create a free account → New Project
2. Name it `kyen-sales-hub`, choose a strong database password, pick any region
3. Once created, go to **SQL Editor** (left sidebar)
4. Paste the entire contents of `schema-v2.sql` and click **Run**
5. Go to **Settings → API** and copy:
   - **Project URL** (looks like `https://xxxx.supabase.co`)
   - **anon / public key** (long string starting with `eyJ...`)

---

## STEP 2 — Set up the project locally (5 minutes)

```bash
# Install Node.js from nodejs.org if not already installed

# Create project
npm create vite@latest kyen-sales-hub -- --template react
cd kyen-sales-hub

# Install dependencies
npm install
npm install @supabase/supabase-js recharts xlsx

# Replace src/App.jsx with the App.jsx file from this folder
# Replace src/index.css with just: @tailwind base; @tailwind components; @tailwind utilities;

# Install Tailwind
npm install -D tailwindcss postcss autoprefixer
npx tailwindcss init -p
```

In `tailwind.config.js`:
```js
export default {
  content: ["./index.html","./src/**/*.{js,jsx}"],
  theme: { extend: {} },
  plugins: [],
}
```

Create `.env` file in the project root:
```
VITE_SUPABASE_URL=https://your-project-id.supabase.co
VITE_SUPABASE_ANON_KEY=your-anon-key-here
```

Test locally:
```bash
npm run dev
# App runs at http://localhost:5173
```

---

## STEP 3 — Push to GitHub (5 minutes)

1. Go to **github.com** → New repository → Name: `kyen-sales-hub` → Private → Create
2. In your project folder:

```bash
git init
git add .
git commit -m "Initial commit — Kyen Sales Hub"
git remote add origin https://github.com/YOUR_USERNAME/kyen-sales-hub.git
git push -u origin main
```

---

## STEP 4 — Deploy to Vercel (3 minutes)

1. Go to **vercel.com** → Sign in with GitHub
2. Click **Add New Project** → Import `kyen-sales-hub`
3. Before clicking Deploy, click **Environment Variables** and add:
   - `VITE_SUPABASE_URL` → your Supabase project URL
   - `VITE_SUPABASE_ANON_KEY` → your anon key
4. Click **Deploy**

Your app will be live at: `https://kyen-sales-hub.vercel.app`

---

## STEP 5 — First-time setup in the app

1. Open your Vercel URL
2. Sign in as **Admin** (password: `kyen2024`)
3. Go to **Prices** → Set retail and wholesale prices for all SKUs
4. Go to **Team** → Add each sales rep with username and password
5. Go to **Targets** → Select month → Set value, visit, and per-SKU targets per rep
6. Share the URL with your reps — they sign in as Sales Rep

---

## Updating the app in future (data stays safe)

Whenever you make changes to `App.jsx`:

```bash
git add .
git commit -m "Update: describe what changed"
git push
```

Vercel redeploys automatically. All data in Supabase is **completely separate** from the code — updates never touch the database. Data is retained permanently as long as your Supabase project exists.

---

## Adding new database columns in future

If a future update needs a new column, run a migration in Supabase SQL Editor:

```sql
-- Example: adding a new field to sales
ALTER TABLE sales ADD COLUMN IF NOT EXISTS invoice_number text;
```

This never deletes existing data. Always use `IF NOT EXISTS` and `ADD COLUMN` — never `DROP` or `RECREATE` a table that has data in it.

---

## Default credentials

| Role  | Username | Password   |
|-------|----------|------------|
| Admin | —        | kyen2024   |
| Reps  | set by admin in Team section | set by admin |

**Change the admin password immediately** in Prices → Settings after first login.

---

## Free tier limits (Supabase)

| Resource     | Free limit         | Kyen estimate     |
|--------------|--------------------|-------------------|
| Database     | 500 MB             | Years away        |
| API requests | 2M / month         | Well within range |
| Realtime     | 200 concurrent     | More than enough  |

