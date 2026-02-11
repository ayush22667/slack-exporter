# Slack Message Fetcher

A Rust tool to fetch messages from a Slack channel and export them to an Excel spreadsheet.

## Features

- Fetches messages from Slack channel
- Filters messages by time period
- Exports to Excel with formatted headers and columns
- Handles pagination automatically

## Setup

### 1. Get a Slack Token

You need a Slack Bot Token or User Token with the following permissions:
- `channels:history` (to read channel messages)
- `channels:read` (to read channel information)

To get a token:
1. Go to https://api.slack.com/apps
2. Create a new app or select an existing one
3. Go to "OAuth & Permissions"
4. Add the required scopes under "Bot Token Scopes"
5. Install the app to your workspace
6. Copy the "Bot User OAuth Token" (starts with `xoxb-`)

### 2. Configure Environment

Copy the example environment file and add your token:

```bash
cp .env.example .env
```

Edit `.env` and replace the token with your actual Slack token:
```
SLACK_TOKEN=xoxb-your-actual-token-here
```

### 3. Build and Run

```bash
# Build the project
cargo build --release

# Run the script
cargo run --release
```

## Usage

When you run the script, it will prompt you for:

1. **Start date** (optional): Enter in format `YYYY-MM-DD HH:MM:SS` or press Enter to fetch all messages
2. **End date** (optional): Enter in format `YYYY-MM-DD HH:MM:SS` or press Enter for current time

Example:
```
Enter start date (YYYY-MM-DD HH:MM:SS) or press Enter for all messages:
2024-01-01 00:00:00

Enter end date (YYYY-MM-DD HH:MM:SS) or press Enter for now:
2024-01-31 23:59:59
```

The script will:
1. Fetch all messages from the specified time period
2. Parse Grafana alert messages
3. Export to an Excel file named `slack_messages_YYYYMMDD_HHMMSS.xlsx`

## Excel Output Format

The Excel file contains the following columns:

| Column | Description |
|--------|-------------|
| Timestamp | Message timestamp |
| User/Bot | User ID or Bot ID who sent the message |
| Connector | Connector name (from Grafana alerts) |
| Flow | Payment flow (from Grafana alerts) |
| Sub-flow | Sub-flow type (from Grafana alerts) |
| Error Code | Error code (from Grafana alerts) |
| Error Message | Error message (from Grafana alerts) |
| Full Text | Complete message text |

## Troubleshooting

### "SLACK_TOKEN environment variable not set"
Make sure you've created a `.env` file with your Slack token.

### "Slack API error: not_in_channel"
Your bot needs to be added to the channel. Invite it to channel `C0ACZK80RPB` using `/invite @your-bot-name`

### "Slack API error: missing_scope"
Make sure your bot token has the `channels:history` and `channels:read` permissions.

## Dependencies

- `reqwest` - HTTP client for Slack API
- `tokio` - Async runtime
- `serde` - JSON serialization
- `rust_xlsxwriter` - Excel file generation
- `chrono` - Date/time handling
- `dotenvy` - Environment variable management
- `anyhow` - Error handling
