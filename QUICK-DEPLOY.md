# Quick Netlify Deployment (30 seconds)

## Step 1: Install Netlify CLI
```bash
npm install -g netlify-cli
```

## Step 2: Deploy
```bash
cd /c/Users/ayeddula/word-autoopen-addin
netlify deploy --prod
```

Follow the prompts:
1. Authenticate with GitHub/Email
2. Create a new site (or link to existing)
3. Publish directory: `.` (current directory)

You'll get a URL like: `https://word-autoopen-addin.netlify.app`

## Step 3: Update manifest.xml

Replace all instances of `https://anirudhy.github.io/word-autoopen-addin` with your Netlify URL.

---

**Alternative: Manual Netlify Deploy (no CLI needed)**

1. Go to https://app.netlify.com/drop
2. Drag the `word-autoopen-addin` folder
3. Get your URL instantly!
