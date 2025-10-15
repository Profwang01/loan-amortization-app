# Loan Amortization App (React + Vite)

A zero-backend loan amortization calculator with extra payments and one-off prepayments. Exports CSV and shows a balance chart.

## Run locally
```bash
npm install
npm run dev
```
Open the URL printed by Vite (typically http://localhost:5173).

## Build for production
```bash
npm run build
npm run preview
```

## Deploy to Vercel (recommended)
1. Create a free account at https://vercel.com and install the Vercel CLI (optional).
2. Upload this folder to a new GitHub repo (or use **Import Project** in Vercel and drag-drop this folder).
3. Framework preset: **Vite** (or **Other**).  
   - Build Command: `npm run build`  
   - Output Directory: `dist`
4. Click **Deploy** — you'll get a public URL usable in Safari/Chrome/Edge/mobile.

No server is required (pure front-end).

## Files
- `src/amortization.js` — core math & schedule generation
- `src/App.jsx` — UI & chart
- `src/main.jsx` — React bootstrap
- `src/styles.css` — minimal styling
