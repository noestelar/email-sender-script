# Email Campaign Sender

A Python script for sending HTML email campaigns to a list of recipients from an Excel file. The script includes features to avoid spam detection, track sent emails, and embed images in the email body.

## Features

- Read email addresses from an Excel file
- Send HTML emails with embedded images
- Send emails in batches to avoid spam detection
- Track sent emails to prevent duplicates
- Customizable email templates with placeholder support
- Command-line interface for easy use

## Requirements

- Python 3.6 or higher
- Dependencies listed in `requirements.txt`

## Installation

1. Clone or download this repository
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python script.py your_excel_file.xlsx --send --template email_template.html --subject "Your Subject" --sender your_email@gmail.com --password your_password
```

### Command-line Arguments

| Argument | Description | Default |
|----------|-------------|---------|
| `file_path` | Path to the Excel file containing email addresses | (Required) |
| `-s`, `--sheet` | Name or index of the sheet to read | 0 |
| `--send` | Send emails to the extracted addresses | False |
| `--subject` | Email subject | "Test Email" |
| `--body` | Plain text email body (when not using a template) | "This is a test email sent from Python." |
| `--template` | Path to HTML template file for email content | None |
| `--sender` | Sender email address | (Required for sending) |
| `--password` | Sender email password or app password | (Required for sending) |
| `--smtp` | SMTP server address | smtp.gmail.com |
| `--port` | SMTP server port | 587 |
| `--batch-size` | Number of emails to send in each batch | 20 |
| `--delay` | Delay in seconds between batches | 60 |
| `--tracking-file` | Path to JSON file to track sent emails | sent_emails.json |

### Excel File Format

The script expects an Excel file with a column named `E-mail` containing the email addresses.

### HTML Template

Create an HTML template for your email content. You can use placeholders in the format `{{placeholder_name}}` that will be replaced with values when sending.

Example template:

```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Email Template</title>
    <style>
        body {
            margin: 0;
            padding: 0;
            font-family: Arial, sans-serif;
        }
    </style>
</head>
<body>
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; margin: 0 auto;">
        <tr>
            <td style="padding: 10px;">
                <p>Hello!</p>
                <p>Your content here...</p>
                <p>Date: {{current_date}}</p>
            </td>
        </tr>
        <tr>
            <td style="padding: 10px;">
                <img src="cid:header_image" alt="Header Image" width="600" style="display: block; max-width: 100%;">
            </td>
        </tr>
    </table>
</body>
</html>
```

### Embedding Images

To embed images in your email, use the `cid:` protocol in your HTML template's image sources:

```html
<img src="cid:header_image" alt="Header Image">
```

The script will look for image files named:
- `info.jpeg` (referenced as `cid:header_image`)
- `footer.jpeg` (referenced as `cid:footer_image`)

Place these files in the same directory as your HTML template.

### Using with Gmail

If you're using Gmail, you'll need to:

1. Enable 2-Step Verification for your Google account
2. Generate an App Password (Google Account → Security → App Passwords)
3. Use this App Password instead of your regular password

## Examples

### Send emails using an HTML template in batches of 10 with a 120-second delay

```bash
python script.py contacts.xlsx --send --template email_template.html --subject "Important Announcement" --sender your_email@gmail.com --password your_app_password --batch-size 10 --delay 120
```

### Just extract and display email addresses from an Excel file

```bash
python script.py contacts.xlsx
```

## Troubleshooting

- **Emails not sending**: Check your SMTP settings and ensure your password is correct
- **Images not displaying**: Ensure image paths are correct and the images exist in the same directory as the template
- **Spam detection**: Try reducing batch size and increasing delay between batches

## License

This project is licensed under the MIT License - see the LICENSE file for details.
