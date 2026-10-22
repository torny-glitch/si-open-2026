# SI Open 2026 — starter website

A responsive Vietnamese tournament microsite inspired by the structure of the public IT MSB Open Championship site. This is a clean-room rebuild; it does not include the original site's source code, proprietary assets, or private data.

## Run locally

1. Install Node.js (LTS) from https://nodejs.org/
2. Open a terminal in this folder.
3. Run `npm install`
4. Run `npm run dev` and open the local URL printed in the terminal.
5. Run `npm run build` to verify a production build.

## Deploy to Vercel

1. Create a GitHub repository and upload these project files.
2. In https://vercel.com/, choose **Add New → Project** and import the repository.
3. Framework preset: **Vite**. Build command: `npm run build`. Output directory: `dist`.
4. Click **Deploy**. Vercel assigns an independent URL such as `si-open-2026-<suffix>.vercel.app`; you can edit the project name in Vercel settings.

## Data notes

- The new SI Open 2026 title is set throughout the site.
- Match schedule, draw, standings, date, venue, and tournament rules are intentionally placeholders because new official data has not been provided.
- The Excel template uses sheets `Lịch thi đấu` and `Vận động viên`.
- The included browser upload feature is a local preview convenience. It stores uploaded data only in the current browser via localStorage; it does **not** publish data to other visitors. To operate a shared public admin panel, connect a database and authentication (e.g. Supabase) or maintain match data in source and redeploy.
