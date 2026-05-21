# BSL Inventory — Setup Guide

## Step 1 — Run the database schema in Supabase
1. Go to supabase.com → your project
2. Click SQL Editor in the left sidebar
3. Click "New query"
4. Copy and paste the contents of `supabase_schema.sql`
5. Click "Run"

## Step 2 — Push this code to GitHub
1. Go to github.com → New repository → name it `bsl-inventory` → Create
2. Open Terminal (Mac) or Command Prompt (Windows) and run:
```
cd path/to/bsl-inventory
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/bsl-inventory.git
git push -u origin main
```

## Step 3 — Deploy on Vercel
1. Go to vercel.com → New Project
2. Import your `bsl-inventory` GitHub repo
3. Add these Environment Variables:
   - REACT_APP_SUPABASE_URL = https://vismrinosmyevgirrhsb.supabase.co
   - REACT_APP_SUPABASE_ANON_KEY = your anon key
   - REACT_APP_ANTHROPIC_KEY = your new anthropic key
4. Click Deploy

## Step 4 — Rotate your Anthropic API key
Go to console.anthropic.com → API Keys → delete the old one → create a new one → update in Vercel

## Your app will be live at:
https://bsl-inventory.vercel.app (or similar URL Vercel assigns)
