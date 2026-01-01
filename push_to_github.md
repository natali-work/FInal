# Instructions to Push Files to GitHub

## Prerequisites
1. Install Git: Download from https://git-scm.com/download/win
2. Create a GitHub account (if you don't have one): https://github.com
3. Create a new repository on GitHub (don't initialize with README)

## Steps to Push Files

### 1. Open PowerShell or Command Prompt in the C:\Final directory

### 2. Initialize Git Repository (if not already done)
```bash
git init
```

### 3. Configure Git (if not already configured)
```bash
git config --global user.name "Your Name"
git config --global user.email "your.email@example.com"
```

### 4. Add All Python and Batch Files
```bash
git add *.py
git add *.bat
git add .gitignore
```

### 5. Commit the Files
```bash
git commit -m "Initial commit: Add experiment analyzer scripts"
```

### 6. Add Remote Repository
Replace `YOUR_USERNAME` and `YOUR_REPO_NAME` with your GitHub username and repository name:
```bash
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
```

### 7. Push to GitHub
```bash
git branch -M main
git push -u origin main
```

## Alternative: Use the Automated Script

Run the `push_to_github.bat` script (after installing Git and creating the repository):
```bash
push_to_github.bat
```

You'll be prompted to enter:
- Your GitHub username
- Your repository name
- Your GitHub password or personal access token

