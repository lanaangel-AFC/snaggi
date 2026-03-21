# Deploying Facade Defect Tracker to Railway

This guide walks you through deploying the app to Railway — a free cloud hosting service. 
Once deployed, you'll have a permanent URL that works independently.

---

## Step 1: Create a GitHub Account (if you don't have one)

Go to [github.com](https://github.com) and sign up. It's free.

## Step 2: Create a Railway Account

1. Go to [railway.app](https://railway.app)
2. Click **Login** → **Login with GitHub**
3. Authorise Railway to access your GitHub

## Step 3: Push the Code to GitHub

On your computer (or ask me to help), create a new GitHub repository:

1. Go to [github.com/new](https://github.com/new)
2. Name it `facade-tracker` (or whatever you like)
3. Keep it **Private**
4. Click **Create repository**
5. Follow the instructions to push existing code (I can help with this)

## Step 4: Deploy on Railway

1. Go to [railway.app/dashboard](https://railway.app/dashboard)
2. Click **New Project** → **Deploy from GitHub Repo**
3. Select your `facade-tracker` repository
4. Railway will detect the Dockerfile and start building

## Step 5: Add Persistent Storage

This is critical — without it, your data resets on every deploy.

1. In your Railway project, click **Add New** → **Volume**
2. Set the mount path to: `/data`
3. Click **Add**

## Step 6: Set Environment Variables

In your Railway service settings → **Variables**, add:

| Variable | Value |
|----------|-------|
| `DATA_DIR` | `/data` |
| `NODE_ENV` | `production` |
| `PORT` | `3000` |

## Step 7: Get Your URL

1. Go to **Settings** → **Networking** → **Generate Domain**
2. Railway will give you a URL like `facade-tracker-production.up.railway.app`
3. This is your permanent app URL

## Step 8: Add to iPhone Home Screen

1. Open the Railway URL in Safari on your iPhone
2. Tap the Share button → **Add to Home Screen**
3. Name it "Facade Tracker" → **Add**

---

## Notes

- **Free tier**: Railway gives you $5/month in free credits, which is more than enough for this app
- **Data persists**: The `/data` volume stores your database and photos permanently
- **Auto-deploys**: Push changes to GitHub and Railway rebuilds automatically
- **Photos**: When you take photos in the app using the camera button, iOS saves the full-resolution original to your camera roll automatically
