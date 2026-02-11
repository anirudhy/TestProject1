# Deployment Options

Since your Enterprise GitHub account doesn't support GitHub Pages for private repositories, here are alternative hosting options:

## Option 1: Netlify (Recommended - Free & Easy)

1. **Sign up at [netlify.com](https://netlify.com)** (it's free)

2. **Deploy via Netlify CLI:**
   ```bash
   npm install -g netlify-cli
   cd /c/Users/ayeddula/word-autoopen-addin
   netlify deploy --prod
   ```

3. **Or deploy via Netlify UI:**
   - Go to https://app.netlify.com
   - Drag and drop your project folder
   - Netlify will give you a URL like: `https://word-autoopen-addin.netlify.app`

4. **Update manifest.xml** with your Netlify URL (replace all instances of the placeholder)

## Option 2: Vercel (Also Free)

1. **Sign up at [vercel.com](https://vercel.com)**

2. **Deploy via Vercel CLI:**
   ```bash
   npm install -g vercel
   cd /c/Users/ayeddula/word-autoopen-addin
   vercel --prod
   ```

3. **Update manifest.xml** with your Vercel URL

## Option 3: Azure Static Web Apps (Microsoft)

1. **Create a Static Web App in Azure Portal**

2. **Deploy via Azure CLI:**
   ```bash
   az staticwebapp create --name word-autoopen-addin --resource-group <your-rg> --location "Central US"
   ```

3. **Update manifest.xml** with your Azure URL

## Option 4: Personal GitHub Account with Pages

If you have a personal GitHub account (not Enterprise):

1. Create a new repository at https://github.com/anirudhy
2. Push the code
3. Enable GitHub Pages in Settings → Pages
4. Your site will be at: `https://anirudhy.github.io/word-autoopen-addin`

---

## Current Status

✅ Repository created: https://github.com/ayeddula_microsoft/word-autoopen-addin
❌ GitHub Pages not available on Enterprise plan
🔄 Manifest URLs are currently set to: `https://anirudhy.github.io/word-autoopen-addin`

**Next step:** Choose a hosting option above and update the manifest.xml URLs accordingly.
