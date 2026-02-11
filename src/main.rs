use anyhow::{Context, Result};
use chrono::{DateTime, TimeZone, Utc};
use reqwest::Client;
use rust_xlsxwriter::{Format, Workbook};
use serde::Deserialize;
use std::collections::HashSet;
use std::env;
use std::fs::File;
use std::io::Write;
use std::sync::Mutex;

// Global logger for writing to both console and file
static LOG_FILE: Mutex<Option<File>> = Mutex::new(None);

fn log(message: &str) {
    // Print to console
    println!("{}", message);

    // Write to file
    if let Ok(mut file_opt) = LOG_FILE.lock() {
        if let Some(ref mut file) = *file_opt {
            let _ = writeln!(file, "{}", message);
        }
    }
}

macro_rules! log {
    ($($arg:tt)*) => {
        log(&format!($($arg)*))
    };
}

#[derive(Debug, Deserialize)]
struct SlackResponse {
    ok: bool,
    messages: Option<Vec<SlackMessage>>,
    error: Option<String>,
    response_metadata: Option<ResponseMetadata>,
}

#[derive(Debug, Deserialize)]
struct ResponseMetadata {
    next_cursor: Option<String>,
}

#[derive(Debug, Deserialize, Clone)]
struct SlackMessage {
    user: Option<String>,
    text: Option<String>,
    ts: String,
    bot_id: Option<String>,
    attachments: Option<Vec<Attachment>>,
    blocks: Option<Vec<serde_json::Value>>,
}

#[derive(Debug, Deserialize, Clone)]
struct Attachment {
    text: Option<String>,
    fallback: Option<String>,
    title: Option<String>,
    pretext: Option<String>,
}

#[derive(Debug)]
struct ParsedMessage {
    _timestamp: DateTime<Utc>,
    _user: String,
    text: String,
    connector: Option<String>,
    flow: Option<String>,
    sub_flow: Option<String>,
    error_code: Option<String>,
    error_message: Option<String>,
}

impl ParsedMessage {
    fn from_slack_message_multiple(msg: &SlackMessage) -> Vec<Self> {
        let timestamp = parse_slack_timestamp(&msg.ts);
        let user = msg.user.clone()
            .or_else(|| msg.bot_id.clone())
            .unwrap_or_else(|| "Unknown".to_string());

        // Extract text from various sources
        let text = extract_message_text(msg);

        // Parse multiple alerts from the same message
        let alerts = parse_multiple_grafana_messages(&text);

        if alerts.is_empty() {
            // No alerts found, return single message with original text
            vec![Self {
                _timestamp: timestamp,
                _user: user.clone(),
                text: text.clone(),
                connector: None,
                flow: None,
                sub_flow: None,
                error_code: None,
                error_message: None,
            }]
        } else {
            // Create a ParsedMessage for each alert
            alerts.into_iter().map(|(connector, flow, sub_flow, error_code, error_message)| {
                Self {
                    _timestamp: timestamp,
                    _user: user.clone(),
                    text: text.clone(),
                    connector,
                    flow,
                    sub_flow,
                    error_code,
                    error_message,
                }
            }).collect()
        }
    }
}

fn extract_message_text(msg: &SlackMessage) -> String {
    // Try different sources of text in order of preference

    // 1. Direct text field
    if let Some(ref text) = msg.text {
        if !text.is_empty() {
            return text.clone();
        }
    }

    // 2. Try attachments
    if let Some(ref attachments) = msg.attachments {
        let mut attachment_texts = Vec::new();
        for attachment in attachments {
            if let Some(ref text) = attachment.text {
                attachment_texts.push(text.clone());
            } else if let Some(ref fallback) = attachment.fallback {
                attachment_texts.push(fallback.clone());
            } else if let Some(ref title) = attachment.title {
                attachment_texts.push(title.clone());
            } else if let Some(ref pretext) = attachment.pretext {
                attachment_texts.push(pretext.clone());
            }
        }
        if !attachment_texts.is_empty() {
            return attachment_texts.join("\n");
        }
    }

    // 3. Try blocks
    if let Some(ref blocks) = msg.blocks {
        let mut block_texts = Vec::new();
        for block in blocks {
            if let Some(text_obj) = block.get("text") {
                if let Some(text_str) = text_obj.get("text").and_then(|t| t.as_str()) {
                    block_texts.push(text_str.to_string());
                }
            }
        }
        if !block_texts.is_empty() {
            return block_texts.join("\n");
        }
    }

    // If nothing found, return empty string
    String::new()
}

fn parse_slack_timestamp(ts: &str) -> DateTime<Utc> {
    let timestamp: f64 = ts.parse().unwrap_or(0.0);
    let secs = timestamp as i64;
    let nsecs = ((timestamp - secs as f64) * 1_000_000_000.0) as u32;
    Utc.timestamp_opt(secs, nsecs).unwrap()
}

fn parse_multiple_grafana_messages(text: &str) -> Vec<(Option<String>, Option<String>, Option<String>, Option<String>, Option<String>)> {
    let mut alerts = Vec::new();
    let mut current_alert = (None, None, None, None, None);
    let mut has_data = false;

    for line in text.lines() {
        let line = line.trim();

        // Check if we're starting a new alert (when we see a Connector line after already having data)
        if line.contains("Connector:") && has_data {
            // Save the current alert
            alerts.push(current_alert.clone());
            current_alert = (None, None, None, None, None);
            has_data = false;
        }

        // Parse fields
        if line.contains("Connector:") {
            let value = line
                .split("Connector:")
                .nth(1)
                .unwrap_or("")
                .trim()
                .replace("*", "")
                .trim()
                .to_string();
            if !value.is_empty() {
                current_alert.0 = Some(value);
                has_data = true;
            }
        } else if line.contains("Flow:") && !line.contains("Sub-flow:") {
            let value = line
                .split("Flow:")
                .nth(1)
                .unwrap_or("")
                .trim()
                .replace("*", "")
                .trim()
                .to_string();
            if !value.is_empty() {
                current_alert.1 = Some(value);
            }
        } else if line.contains("Sub-flow:") {
            let value = line
                .split("Sub-flow:")
                .nth(1)
                .unwrap_or("")
                .trim()
                .replace("*", "")
                .trim()
                .to_string();
            if !value.is_empty() {
                current_alert.2 = Some(value);
            }
        } else if line.contains("Error code:") {
            let value = line
                .split("Error code:")
                .nth(1)
                .unwrap_or("")
                .trim()
                .replace("*", "")
                .trim()
                .to_string();

            let cleaned = clean_rust_option(&value);
            if !cleaned.is_empty() && cleaned != "None" {
                current_alert.3 = Some(cleaned);
            }
        } else if line.contains("Error message:") {
            let value = line
                .split("Error message:")
                .nth(1)
                .unwrap_or("")
                .trim()
                .replace("*", "")
                .trim()
                .to_string();

            let cleaned = clean_rust_option(&value);
            if !cleaned.is_empty() && cleaned != "None" {
                current_alert.4 = Some(cleaned);
            }
        }
    }

    // Don't forget to add the last alert
    if has_data {
        alerts.push(current_alert);
    }

    alerts
}

fn clean_rust_option(value: &str) -> String {
    let trimmed = value.trim();

    // Handle None
    if trimmed == "None" {
        return String::new();
    }

    // Handle Some("...")
    if trimmed.starts_with("Some(\"") && trimmed.ends_with("\")") {
        return trimmed[6..trimmed.len()-2].to_string();
    }

    // Handle Some('...')
    if trimmed.starts_with("Some('") && trimmed.ends_with("')") {
        return trimmed[6..trimmed.len()-2].to_string();
    }

    // Return as-is if no wrapper
    trimmed.to_string()
}

async fn fetch_messages(
    client: &Client,
    token: &str,
    channel_id: &str,
    oldest: Option<f64>,
    latest: Option<f64>,
) -> Result<Vec<SlackMessage>> {
    let mut all_messages = Vec::new();
    let mut cursor: Option<String> = None;
    let mut page = 1;

    loop {
        let mut url = format!(
            "https://slack.com/api/conversations.history?channel={}",
            channel_id
        );

        if let Some(oldest_ts) = oldest {
            url.push_str(&format!("&oldest={}", oldest_ts));
        }

        if let Some(latest_ts) = latest {
            url.push_str(&format!("&latest={}", latest_ts));
        }

        if let Some(ref c) = cursor {
            url.push_str(&format!("&cursor={}", c));
        }

        url.push_str("&limit=1000");

        log!("[FETCH] Page {}: Requesting messages...", page);

        let response = client
            .get(&url)
            .header("Authorization", format!("Bearer {}", token))
            .send()
            .await
            .context("Failed to send request to Slack API")?;

        let slack_response: SlackResponse = response
            .json()
            .await
            .context("Failed to parse Slack API response")?;

        if !slack_response.ok {
            anyhow::bail!(
                "Slack API error: {}",
                slack_response.error.unwrap_or_else(|| "Unknown error".to_string())
            );
        }

        if let Some(mut messages) = slack_response.messages {
            log!("[FETCH] Page {}: Received {} messages (total so far: {})",
                page, messages.len(), all_messages.len() + messages.len());
            all_messages.append(&mut messages);
        } else {
            log!("[FETCH] Page {}: No messages in response", page);
        }

        if let Some(metadata) = slack_response.response_metadata {
            if let Some(next_cursor) = metadata.next_cursor {
                if !next_cursor.is_empty() {
                    log!("[FETCH] More pages available, continuing...");
                    cursor = Some(next_cursor);
                    page += 1;
                    continue;
                }
            }
        }

        log!("[FETCH] No more pages, fetch complete");
        break;
    }

    Ok(all_messages)
}

fn export_to_excel(messages: &[ParsedMessage], filename: &str) -> Result<()> {
    log!("\n[EXCEL] Creating workbook...");
    let mut workbook = Workbook::new();

    log!("[EXCEL] Adding worksheet...");
    let worksheet = workbook.add_worksheet();

    log!("[EXCEL] Creating header format...");
    let header_format = Format::new()
        .set_bold()
        .set_background_color(rust_xlsxwriter::Color::RGB(0x4472C4))
        .set_font_color(rust_xlsxwriter::Color::White);

    log!("[EXCEL] Writing headers...");
    worksheet.write_string_with_format(0, 0, "Connector", &header_format)?;
    worksheet.write_string_with_format(0, 1, "Flow", &header_format)?;
    worksheet.write_string_with_format(0, 2, "Sub-flow", &header_format)?;
    worksheet.write_string_with_format(0, 3, "Error Code", &header_format)?;
    worksheet.write_string_with_format(0, 4, "Error Message", &header_format)?;
    log!("[EXCEL] Headers written successfully");

    log!("[EXCEL] Setting column widths...");
    worksheet.set_column_width(0, 15)?;
    worksheet.set_column_width(1, 25)?;
    worksheet.set_column_width(2, 20)?;
    worksheet.set_column_width(3, 50)?;
    worksheet.set_column_width(4, 50)?;
    log!("[EXCEL] Column widths set");

    log!("[EXCEL] Writing {} data rows...", messages.len());
    for (idx, msg) in messages.iter().enumerate() {
        let row = (idx + 1) as u32;

        if idx < 3 || idx % 100 == 0 {
            log!("[EXCEL] Writing row {} - Connector: {:?}, Flow: {:?}",
                row, msg.connector, msg.flow);
        }

        worksheet.write_string(row, 0, msg.connector.as_deref().unwrap_or(""))?;
        worksheet.write_string(row, 1, msg.flow.as_deref().unwrap_or(""))?;
        worksheet.write_string(row, 2, msg.sub_flow.as_deref().unwrap_or(""))?;
        worksheet.write_string(row, 3, msg.error_code.as_deref().unwrap_or(""))?;
        worksheet.write_string(row, 4, msg.error_message.as_deref().unwrap_or(""))?;
    }
    log!("[EXCEL] All data rows written");

    log!("[EXCEL] Saving workbook to: {}", filename);
    workbook.save(filename)?;
    log!("[EXCEL] Workbook saved successfully");

    Ok(())
}

#[tokio::main]
async fn main() -> Result<()> {
    dotenvy::dotenv().ok();

    // Initialize log file
    let log_filename = format!("slack_fetch_log_{}.txt", Utc::now().format("%Y%m%d_%H%M%S"));
    let log_file = File::create(&log_filename)?;
    *LOG_FILE.lock().unwrap() = Some(log_file);

    log!("===========================================");
    log!("Slack Message Fetcher - Run started at {}", Utc::now().format("%Y-%m-%d %H:%M:%S"));
    log!("Log file: {}", log_filename);
    log!("===========================================\n");

    let token = env::var("SLACK_TOKEN")
        .context("SLACK_TOKEN environment variable not set. Please create a .env file with your token.")?;

    let channel_id = env::var("CHANNEL_ID")
        .context("CHANNEL_ID environment variable not set. Please add it to your .env file.")?;

    log!("Fetching messages from Slack channel: {}", channel_id);

    log!("\nHow many days back do you want to fetch messages?");
    log!("  0 = only today");
    log!("  1 = today + yesterday (last 2 days)");
    log!("  2 = today + last 2 days (last 3 days)");
    log!("  Press Enter for all messages");
    log!("\nEnter number of days:");

    let mut days_input = String::new();
    std::io::stdin().read_line(&mut days_input)?;
    let days_input = days_input.trim();

    let (oldest, latest) = if !days_input.is_empty() {
        let days: i64 = days_input.parse()
            .context("Please enter a valid number")?;

        let now = Utc::now();
        let start_of_today = now.date_naive().and_hms_opt(0, 0, 0).unwrap();
        let start_of_today_utc = Utc.from_utc_datetime(&start_of_today);

        // Calculate the start date (going back 'days' number of days from start of today)
        let start_date = start_of_today_utc - chrono::Duration::days(days);

        log!("\nFetching messages from {} to now", start_date.format("%Y-%m-%d %H:%M:%S"));
        log!("(Last {} day{})", days + 1, if days == 0 { "" } else { "s" });

        (Some(start_date.timestamp() as f64), None)
    } else {
        log!("\nFetching all messages (no date filter)");
        (None, None)
    };

    let client = Client::new();

    log!("\n[FETCH] Fetching messages...");
    let messages = fetch_messages(&client, &token, &channel_id, oldest, latest).await?;

    log!("[FETCH] Fetched {} raw messages", messages.len());

    if messages.is_empty() {
        log!("[WARN] No messages found in the specified time range!");
        return Ok(());
    }

    log!("\n[PARSE] Parsing messages...");
    log!("[PARSE] Detailed logs will be saved for all messages in the log file");
    log!("[PARSE] Note: Messages with multiple alerts will be split into separate rows");

    let parsed_messages: Vec<ParsedMessage> = messages
        .iter()
        .enumerate()
        .flat_map(|(idx, msg)| {
            log!("\n[PARSE] ===== Message {} / {} =====", idx + 1, messages.len());
            log!("[PARSE]   User/Bot: {:?}", msg.user.as_ref().or(msg.bot_id.as_ref()).unwrap_or(&"Unknown".to_string()));
            log!("[PARSE]   Timestamp: {}", msg.ts);
            log!("[PARSE]   Has text field: {}", msg.text.is_some());
            log!("[PARSE]   Has attachments: {}", msg.attachments.is_some());
            log!("[PARSE]   Has blocks: {}", msg.blocks.is_some());

            let parsed_alerts = ParsedMessage::from_slack_message_multiple(msg);

            log!("[PARSE]   Extracted text ({} chars):", parsed_alerts[0].text.len());
            if parsed_alerts[0].text.len() > 0 {
                log!("[PARSE]   Text: {}", parsed_alerts[0].text);
            } else {
                log!("[PARSE]   Text: <EMPTY>");
            }

            log!("[PARSE]   Found {} alert(s) in this message:", parsed_alerts.len());
            for (alert_idx, parsed) in parsed_alerts.iter().enumerate() {
                log!("[PARSE]   Alert {} fields:", alert_idx + 1);
                log!("[PARSE]     - Connector: {:?}", parsed.connector);
                log!("[PARSE]     - Flow: {:?}", parsed.flow);
                log!("[PARSE]     - Sub-flow: {:?}", parsed.sub_flow);
                log!("[PARSE]     - Error Code: {:?}", parsed.error_code);
                log!("[PARSE]     - Error Message: {:?}", parsed.error_message);
            }

            parsed_alerts
        })
        .collect();

    log!("[PARSE] Parsed {} messages", parsed_messages.len());

    // Show statistics
    let with_connector = parsed_messages.iter().filter(|m| m.connector.is_some()).count();
    let with_flow = parsed_messages.iter().filter(|m| m.flow.is_some()).count();
    let with_sub_flow = parsed_messages.iter().filter(|m| m.sub_flow.is_some()).count();
    let with_error_code = parsed_messages.iter().filter(|m| m.error_code.is_some()).count();
    let with_error_message = parsed_messages.iter().filter(|m| m.error_message.is_some()).count();

    log!("\n[PARSE] Statistics:");
    log!("[PARSE]   Messages with Connector: {} / {}", with_connector, parsed_messages.len());
    log!("[PARSE]   Messages with Flow: {} / {}", with_flow, parsed_messages.len());
    log!("[PARSE]   Messages with Sub-flow: {} / {}", with_sub_flow, parsed_messages.len());
    log!("[PARSE]   Messages with Error Code: {} / {}", with_error_code, parsed_messages.len());
    log!("[PARSE]   Messages with Error Message: {} / {}", with_error_message, parsed_messages.len());

    if parsed_messages.is_empty() {
        log!("[WARN] No messages after parsing!");
        return Ok(());
    }

    // Deduplicate based on (Connector, Error Code, Error Message)
    log!("\n[DEDUP] Removing duplicate alerts...");
    let mut seen = HashSet::new();
    let mut unique_messages = Vec::new();
    let mut duplicate_count = 0;

    for msg in parsed_messages {
        // Create a unique key from connector, error code, and error message
        let key = (
            msg.connector.clone().unwrap_or_default(),
            msg.error_code.clone().unwrap_or_default(),
            msg.error_message.clone().unwrap_or_default(),
        );

        if seen.insert(key.clone()) {
            // This is a unique combination
            unique_messages.push(msg);
        } else {
            // This is a duplicate
            duplicate_count += 1;
            log!("[DEDUP] Skipping duplicate: Connector={}, Error Code={}, Error Message={}",
                key.0, key.1, key.2);
        }
    }

    log!("[DEDUP] Removed {} duplicates", duplicate_count);
    log!("[DEDUP] Unique alerts remaining: {}", unique_messages.len());

    let parsed_messages = unique_messages;

    let filename = format!(
        "slack_messages_{}.xlsx",
        Utc::now().format("%Y%m%d_%H%M%S")
    );

    log!("\n[EXPORT] Exporting to Excel: {}", filename);
    export_to_excel(&parsed_messages, &filename)?;

    // Verify file was created and has data
    let file_metadata = std::fs::metadata(&filename)?;
    let file_size = file_metadata.len();

    log!("\n✓ Successfully exported {} messages to {}", parsed_messages.len(), filename);
    log!("  File size: {} bytes ({:.2} KB)", file_size, file_size as f64 / 1024.0);
    log!("  File path: {}/{}", std::env::current_dir()?.display(), filename);

    // Show a sample of what was exported
    if !parsed_messages.is_empty() {
        log!("\n📋 Sample of exported data (first message with Grafana data):");
        let sample = parsed_messages.iter()
            .find(|m| m.connector.is_some())
            .unwrap_or(&parsed_messages[0]);

        log!("  Connector: {}", sample.connector.as_deref().unwrap_or("N/A"));
        log!("  Flow: {}", sample.flow.as_deref().unwrap_or("N/A"));
        log!("  Sub-flow: {}", sample.sub_flow.as_deref().unwrap_or("N/A"));
        log!("  Error Code: {}", sample.error_code.as_deref().unwrap_or("N/A"));
        log!("  Error Message: {}", sample.error_message.as_deref().unwrap_or("N/A"));
    }

    log!("\n===========================================");
    log!("Run completed at {}", Utc::now().format("%Y-%m-%d %H:%M:%S"));
    log!("Log saved to: {}", log_filename);
    log!("===========================================");

    Ok(())
}
