# Email Automation Tool

A desktop-based email automation application built with Python and Tkinter. The application allows users to send personalized bulk emails, attach PDFs, manage follow-ups, use Gmail signatures, track replies, and resume interrupted sending sessions.

## Features

### 1. Bulk Email Sending

- Send emails to multiple recipients from a Google Sheet.
- Personalize emails using recipient-specific information.
- Supports HTML and plain-text email content.
- Configure a delay between emails.
- Use multiple parallel SMTP threads for faster sending.

### 2. Email Personalization

Use placeholders in the email body to automatically insert recipient information.

Example:

```text
Hello {Name},

I am reaching out regarding a potential collaboration between our teams at {Company}.
```

Supported placeholders:

```text
{Name}
{Email}
{Company}
```

Additional columns from the Google Sheet can also be used as placeholders.

### 3. PDF Attachments

The application supports attaching PDF files to emails.

Useful for:

- Resumes
- Brochures
- Proposals
- Company profiles
- Reports
- Sponsorship documents

### 4. Gmail SMTP Sending

Emails are sent using Gmail's SMTP server.

```text
SMTP Server: smtp.gmail.com
Port: 587
Security: STARTTLS
```

For authentication, use a Gmail App Password instead of your normal Gmail password.

### 5. Gmail Signature Integration

The application can retrieve your Gmail signature using the Gmail API.

You need to configure:

```text
credentials.json
```

and authenticate with your Google account.

### 6. Text to HTML Converter

The application includes a Text to HTML conversion utility for creating formatted email content.

Useful for:

- Bold text
- Links
- Formatted paragraphs
- Professional email templates
- Basic email formatting

### 7. Manual Follow-Up System

The application provides a follow-up system for previously sent emails.

It can:

1. Load previously saved email records.
2. Identify the original email thread.
3. Check whether the recipient has replied.
4. Allow you to select recipients for follow-up.
5. Send follow-up emails within the existing Gmail conversation.

The application uses email headers such as:

```text
In-Reply-To
References
```

to maintain the original email thread.

### 8. Multiple Follow-Up Drafts

You can create multiple follow-up messages.

Example:

```text
Follow-up 1:
Just following up on my previous email.

Follow-up 2:
I wanted to check if you had a chance to review my previous message.

Follow-up 3:
Please let me know if this is something your team would be interested in discussing.
```

### 9. Automatic Reply Detection

The application can check Gmail for replies using IMAP.

It checks whether a recipient has responded before sending a follow-up.

### 10. Auto-Save

The application automatically saves email-sending progress.

This helps prevent losing progress if:

- The application is closed.
- The computer shuts down.
- The internet connection is interrupted.
- Email sending is stopped manually.

### 11. Resume Interrupted Sessions

If an email campaign is interrupted, the application can resume from the saved progress instead of starting again from the beginning.

### 12. Pause, Resume and Stop

The application provides controls for managing an active email campaign.

- **Pause:** Temporarily pauses the sending process.
- **Resume:** Continues sending from where the process was paused.
- **Stop:** Stops the current campaign.

The application keeps track of sending status so completed emails are not unnecessarily resent.

---

# Technology Stack

| Category | Technology |
|----------|------------|
| Programming Language | Python |
| GUI | Tkinter |
| Data Processing | Pandas |
| Email | SMTP, IMAP, Python Email Library |
| Google Services | Gmail API, Google Sheets API |
| Authentication | OAuth 2.0 |
| Concurrency | Python Threading |

---

# Installation

## 1. Clone the Repository

```bash
git clone <repository-url>
cd <project-folder>
```

## 2. Install Dependencies

Install the required packages:

```bash
pip install pandas
pip install gspread
pip install google-auth
pip install google-auth-oauthlib
pip install google-api-python-client
```

Or, if a `requirements.txt` file is provided:

```bash
pip install -r requirements.txt
```

---

# Gmail Configuration

## 1. Enable Two-Step Verification

Enable 2-Step Verification on your Google Account before creating an App Password.

## 2. Create a Gmail App Password

After enabling 2-Step Verification:

1. Open your Google Account.
2. Go to Security.
3. Open App Passwords.
4. Create a new App Password.
5. Give it a suitable name such as `Email Automation`.
6. Copy the generated 16-character password.

Use this App Password instead of your normal Gmail password.

---

# Gmail API Configuration

The Gmail API is required for features such as retrieving your Gmail signature.

## Step 1: Create a Google Cloud Project

Create a new project in Google Cloud Console.

## Step 2: Enable Gmail API

Enable the Gmail API for the project.

## Step 3: Configure OAuth Consent Screen

Configure the OAuth consent screen for the application.

## Step 4: Create OAuth Credentials

Create an OAuth Client ID and select:

```text
Desktop Application
```

Download the credentials file and rename it:

```text
credentials.json
```

Place it in the project directory.

---

# Authentication Files

After successful authentication, the application may create:

```text
token.pickle
```

Do not upload authentication files to GitHub.

Add the following to `.gitignore`:

```text
credentials.json
token.pickle
*.pickle
.env
```

---

# Google Sheets Setup

The application can use Google Sheets as the recipient database.

A typical sheet should contain:

| Name | Email | Company |
|------|-------|---------|
| Rahul | rahul@example.com | ABC Ltd |
| Priya | priya@example.com | XYZ Ltd |

The column names should match the placeholders used in your email template.

For example:

```text
{Name}
{Email}
{Company}
```

Additional columns can also be used as placeholders.

---

# Email Sending Workflow

A typical workflow is:

```text
1. Open the application
2. Enter sender credentials
3. Load Google Sheet
4. Select email template
5. Add placeholders
6. Select PDF attachment
7. Configure delay
8. Configure parallel threads
9. Start campaign
10. Track sending status
```

---

# Sending Delay

You can configure a delay between emails.

For example:

```text
Delay = 5 seconds
```

The application waits approximately five seconds before sending the next email from a given sending thread.

Using a delay can help avoid sending a very large number of emails in a short period.

---

# Parallel Threads

The application supports multiple SMTP sending threads.

For example:

```text
Threads = 3
```

allows multiple emails to be processed concurrently.

Increasing the number of threads can improve sending speed, but excessive concurrency may cause:

- Gmail rate limits
- SMTP errors
- Temporary account restrictions
- Connection failures

Use a reasonable number of threads.

---

# Follow-Up Workflow

The follow-up process works approximately as follows:

```text
Original Email
      |
      v
Check Gmail Thread
      |
      v
Has Recipient Replied?
      |
   +--+--+
   |     |
  Yes    No
   |     |
   v     v
 Stop   Select
        Follow-up
           |
           v
       Send Reply
           |
           v
      Same Thread
```

If a recipient has already replied, the application avoids sending an unnecessary follow-up.

---

# Generated and Saved Files

Depending on the workflow, the application may create or use files for:

- Email campaign data
- Sending progress
- Follow-up information
- Authentication tokens
- Logs
- CSV exports

Do not commit authentication or sensitive files to a public repository.

---

# Error Handling

The application includes handling for common issues such as:

- SMTP connection failures
- Authentication errors
- Network interruptions
- Invalid recipient addresses
- Gmail API authentication problems
- Google Sheets access errors
- Missing attachments
- Interrupted sending sessions

Check the application logs and error messages when troubleshooting.

---

# Troubleshooting

## Gmail Authentication Error

Make sure you are using a Gmail App Password rather than your normal Gmail password.

Check that:

```text
2-Step Verification = Enabled
```

and that the App Password is correct.

## Gmail API Authentication Problems

Check that:

```text
credentials.json
```

is present in the correct project directory.

If authentication information becomes invalid, remove the locally generated token and authenticate again.

## Google Sheets Access Error

Make sure:

- The Google Sheet URL is correct.
- The required Google APIs are enabled.
- The authenticated Google account has access to the sheet.
- The sheet contains the expected column names.

## PDF Attachment Error

Check that:

- The PDF exists.
- The selected path is correct.
- The file is not being used by another application.
- The application has permission to access the file.

## Emails Not Sending

Check:

```text
Internet connection
Gmail credentials
App Password
SMTP server
SMTP port
Recipient email addresses
Gmail sending limits
```

---

# Recommended Project Structure

```text
email-automation/
|
├── main.py
├── credentials.json
├── token.pickle
├── requirements.txt
├── README.md
|
├── attachments/
|
├── logs/
|
└── data/
```

Do not upload sensitive authentication files to GitHub.

---

# Security Considerations

This application handles sensitive information including:

- Email addresses
- Gmail credentials
- Google OAuth tokens
- Email content
- Recipient information
- PDF attachments

Follow these practices:

1. Never share your Gmail App Password.
2. Never commit `credentials.json` to GitHub.
3. Never commit `token.pickle` to GitHub.
4. Use a `.gitignore` file.
5. Keep recipient databases private.
6. Avoid storing passwords directly in source code.
7. Use environment variables for sensitive configuration where possible.

---

# Example Use Cases

## Recruitment Outreach

```text
Hello {Name},

I am a student at IIT Kharagpur interested in opportunities at {Company}.
```

## Sponsorship Outreach

```text
Hello {Name},

I am reaching out regarding a potential partnership between {Company} and our upcoming event.
```

A PDF sponsorship brochure can be attached automatically.

## Business Outreach

Use a Google Sheet containing:

```text
Name
Email
Company
Designation
```

and personalize each message automatically.

---

# Future Improvements

Potential improvements include:

- Email scheduling
- Rich HTML email editor
- Multiple Gmail account support
- Campaign analytics
- Open and click tracking
- Advanced reply classification
- Automatic follow-up scheduling
- Email templates library
- Better error recovery
- Database-backed campaign management
- Contact management
- Attachment personalization
- Campaign statistics dashboard

---

# Disclaimer

This tool is intended for legitimate email communication and outreach.

Users are responsible for:

- Obtaining appropriate consent where required.
- Following applicable email and privacy regulations.
- Respecting Gmail and Google API usage policies.
- Avoiding spam and abusive email practices.
- Managing recipient data responsibly.

Use reasonable sending rates and appropriate delays when conducting bulk outreach.

---

# Author

Developed as a Python desktop automation project.

```text
Python | Tkinter | Gmail API | SMTP | IMAP | Google Sheets | Pandas
```
