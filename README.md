# Telegram Footer Bot

A Telegram bot that adds footers to PDF and DOCX files with name, roll number, and page numbers.

## Features
- Processes PDF and DOCX files
- Adds customized footer with name and roll number
- Supports multiple users simultaneously
- Automatic error handling and retries

## Requirements
- Python 3.7+
- Dependencies listed in requirements.txt

## Environment Variables
- TELEGRAM_BOT_TOKEN: Your Telegram bot token from @BotFather

## Deployment Instructions

### Option 1: Railway.app (Recommended)
1. Fork this repository to your GitHub account
2. Go to [Railway.app](https://railway.app/)
3. Create an account/Sign in with GitHub
4. Click "New Project" → "Deploy from GitHub repo"
5. Select this repository
6. Add environment variable:
   - Add `TELEGRAM_BOT_TOKEN` with your bot token
7. Railway will automatically deploy your bot

### Option 2: Heroku
1. Create a Heroku account
2. Install Heroku CLI
3. Run the following commands:
```bash
heroku create your-bot-name
heroku config:set TELEGRAM_BOT_TOKEN=your_bot_token
git push heroku main
```

### Option 3: DigitalOcean App Platform
1. Create a DigitalOcean account
2. Go to App Platform
3. Create a new app from GitHub repo
4. Add environment variable for TELEGRAM_BOT_TOKEN
5. Deploy

## Local Development
1. Clone the repository
2. Install dependencies: `pip install -r requirements.txt`
3. Create `.env` file with your TELEGRAM_BOT_TOKEN
4. Run `python bot.py` 