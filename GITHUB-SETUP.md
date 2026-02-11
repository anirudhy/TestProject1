# Manual GitHub Setup Instructions

Your Enterprise GitHub account (ayeddula_microsoft) doesn't allow public repositories.

## Create Repository on Personal Account (anirudhy)

### Step 1: Create the Repository

1. Go to https://github.com/anirudhy
2. Click the **"+"** button (top-right) → **"New repository"**
3. Repository name: `word-autoopen-addin`
4. Description: `Minimal Word Office Add-in with auto-open task pane`
5. Visibility: **Public** ✅
6. **Do NOT** initialize with README, .gitignore, or license
7. Click **"Create repository"**

### Step 2: Push Your Code

Run these commands:

```bash
cd /c/Users/ayeddula/word-autoopen-addin
git remote set-url origin https://github.com/anirudhy/word-autoopen-addin.git
git push -u origin master
```

When prompted for credentials:
- Username: `anirudhy`
- Password: Use a Personal Access Token (not your password)

**To create a Personal Access Token:**
1. Go to https://github.com/settings/tokens/new
2. Note: "Word Add-in deployment"
3. Expiration: 90 days (or your preference)
4. Scopes: Check **`repo`** (all repo permissions)
5. Click **"Generate token"**
6. Copy the token and use it as your password

### Step 3: Enable GitHub Pages

1. Go to https://github.com/anirudhy/word-autoopen-addin/settings/pages
2. Under **"Source"**, select:
   - Branch: **master**
   - Folder: **/ (root)**
3. Click **"Save"**
4. Wait 1-2 minutes for deployment
5. Your site will be live at: **https://anirudhy.github.io/word-autoopen-addin**

### Step 4: Verify Deployment

Visit these URLs to confirm everything works:
- https://anirudhy.github.io/word-autoopen-addin
- https://anirudhy.github.io/word-autoopen-addin/manifest.xml
- https://anirudhy.github.io/word-autoopen-addin/taskpane.html

### Step 5: Test the Add-in

1. Download the manifest: https://anirudhy.github.io/word-autoopen-addin/manifest.xml
2. Upload it to Word on the web (Insert → Add-ins → Upload My Add-in)
3. Click "Show Taskpane" in the ribbon
4. Test the auto-open functionality!

---

## Alternative: I Can Help You Push

If you'd like, you can:
1. Create the repository manually on https://github.com/anirudhy (Steps 1-3 above)
2. Tell me when it's ready, and I'll help push the code
