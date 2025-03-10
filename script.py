# %% importar librerias
import pandas as pd
import argparse
import smtplib
import os
import re
import time
import json
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from pathlib import Path
import ssl

# %% leer columnas de excel
def read_excel_column(file_path, column_name=None, column_index=None, sheet_name=0):
    """
    Read an Excel file and extract values from a specified column.
    
    Parameters:
    file_path (str): Path to the Excel file
    column_name (str, optional): Name of the column to extract
    column_index (int, optional): Index of the column to extract (0-based)
    sheet_name (str or int, optional): Name or index of the sheet to read
    
    Returns:
    list: Values from the specified column
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Extract column values based on name or index
        if column_name is not None:
            if column_name in df.columns:
                values = df[column_name].tolist()
            else:
                raise ValueError(f"Column '{column_name}' not found in the Excel file")
        elif column_index is not None:
            if 0 <= column_index < len(df.columns):
                values = df.iloc[:, column_index].tolist()
            else:
                raise ValueError(f"Column index {column_index} is out of range")
        else:
            # If no column specified, show all columns and return
            print("Available columns:")
            for i, col in enumerate(df.columns):
                print(f"{i}: {col}")
            return None
        
        return values
    
    except Exception as e:
        print(f"Error: {e}")
        return None

# %% leer emails
def read_emails_from_excel(file_path, sheet_name=0, sent_emails_file=None):
    """
    Read an Excel file and extract all email addresses from the 'E-mail' column.
    Optionally filter out already sent emails.
    
    Parameters:
    file_path (str): Path to the Excel file
    sheet_name (str or int, optional): Name or index of the sheet to read
    sent_emails_file (str, optional): Path to the JSON file tracking sent emails
    
    Returns:
    tuple: (List of email addresses, DataFrame with all data)
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Check if 'E-mail' column exists
        if 'E-mail' not in df.columns:
            print("Warning: 'E-mail' column not found in the Excel file")
            return [], df
        
        # Filter out any None or empty values
        df = df[df['E-mail'].notna()]
        df = df[df['E-mail'].astype(str).str.strip() != '']
        
        # Load already sent emails if tracking file exists
        sent_emails = set()
        if sent_emails_file and os.path.exists(sent_emails_file):
            try:
                with open(sent_emails_file, 'r') as f:
                    sent_data = json.load(f)
                    sent_emails = set(sent_data.get('sent_emails', []))
                print(f"Loaded {len(sent_emails)} previously sent emails from tracking file")
            except Exception as e:
                print(f"Warning: Could not load sent emails tracking file: {e}")
        
        # Filter out already sent emails
        if sent_emails:
            original_count = len(df)
            df = df[~df['E-mail'].isin(sent_emails)]
            filtered_count = len(df)
            print(f"Filtered out {original_count - filtered_count} already sent emails")
        
        # Extract email addresses
        emails = df['E-mail'].tolist()
        
        return emails, df
    
    except Exception as e:
        print(f"Error reading emails: {e}")
        return [], None

# %% leer plantilla HTML
def read_html_template(template_path, placeholder_values=None):
    """
    Read an HTML template file and replace placeholders with values.
    
    Parameters:
    template_path (str): Path to the HTML template file
    placeholder_values (dict, optional): Dictionary of placeholder values to replace in the template
    
    Returns:
    tuple: (HTML content, list of image paths found in the template)
    """
    try:
        with open(template_path, 'r', encoding='utf-8') as file:
            html_content = file.read()
        
        # Replace placeholders if provided
        if placeholder_values and isinstance(placeholder_values, dict):
            for key, value in placeholder_values.items():
                placeholder = f"{{{{{key}}}}}"
                html_content = html_content.replace(placeholder, str(value))
        
        # Find all image paths in the HTML content
        # This regex looks for <img src="..." or <img src='...' patterns
        img_paths = []
        img_pattern = re.compile(r'<img[^>]+src=[\'"]([^\'"]+)[\'"]', re.IGNORECASE)
        
        # Find all matches and extract the image paths
        for match in img_pattern.finditer(html_content):
            img_path = match.group(1)
            # Skip Content-ID references as they'll be handled separately
            if img_path.startswith('cid:'):
                continue
            # Only include local file paths, not URLs
            if not img_path.startswith(('http://', 'https://', 'data:')):
                # Convert relative paths to absolute paths
                if not os.path.isabs(img_path):
                    template_dir = os.path.dirname(os.path.abspath(template_path))
                    img_path = os.path.join(template_dir, img_path)
                img_paths.append(img_path)
        
        return html_content, img_paths
    
    except Exception as e:
        print(f"Error reading HTML template: {e}")
        return None, []

# %% enviar emails
def send_emails(email_list, subject, message_body=None, template_path=None, placeholder_values=None, 
                sender_email=None, sender_password=None, smtp_server="mail.isosdrive.com", smtp_port=465,
                use_ssl=True, use_direct_ip=False, direct_ip=None):
    """
    Send emails to a list of recipients using either plain text or an HTML template.
    
    Parameters:
    email_list (list): List of recipient email addresses
    subject (str): Email subject
    message_body (str, optional): Plain text email body content (used if template_path is None)
    template_path (str, optional): Path to HTML template file
    placeholder_values (dict, optional): Dictionary of values to replace placeholders in the template
    sender_email (str): Sender's email address
    sender_password (str): Sender's email password or app password
    smtp_server (str, optional): SMTP server address
    smtp_port (int, optional): SMTP server port
    use_ssl (bool, optional): Whether to use SSL/TLS connection (True) or STARTTLS (False)
    use_direct_ip (bool, optional): Whether to use a direct IP address instead of hostname
    direct_ip (str, optional): Direct IP address of the mail server if use_direct_ip is True
    
    Returns:
    dict: Results of the email sending operation
    """
    results = {
        "success": [],
        "failed": []
    }
    
    try:
        # Set up the SMTP server with SSL
        context = ssl.create_default_context()
        
        if use_ssl:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context, timeout=30)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
            server.starttls(context=context)
        
        # server.set_debuglevel(1)  # Show all communication with the server
        
        # Try login with full email address
        server.login(sender_email, sender_password)
        
        # Determine if we're using HTML template or plain text
        html_content = None
        img_paths = []
        
        if template_path:
            html_content, img_paths = read_html_template(template_path, placeholder_values)
            if not html_content:
                print(f"Error: Could not read HTML template from {template_path}")
                return results
                
            # Check for specific image files that should be embedded with Content-ID
            template_dir = os.path.dirname(os.path.abspath(template_path))
            header_img_path = os.path.join(template_dir, "info.jpeg")
            footer_img_path = os.path.join(template_dir, "footer.jpeg")
            
            if os.path.exists(header_img_path) and header_img_path not in img_paths:
                img_paths.append(header_img_path)
            
            if os.path.exists(footer_img_path) and footer_img_path not in img_paths:
                img_paths.append(footer_img_path)
        
        # Send emails to each recipient
        for recipient in email_list:
            try:
                # Create message
                msg = MIMEMultipart('related')  # Use 'related' for HTML with inline images
                msg['From'] = sender_email
                msg['To'] = recipient
                msg['Subject'] = subject
                
                # Create alternative part for the email (text/html)
                alt_part = MIMEMultipart('alternative')
                msg.attach(alt_part)
                
                # Attach the message body
                if html_content:
                    # Create a copy of the HTML content for this recipient
                    recipient_html = html_content
                    
                    # Replace recipient-specific placeholders if needed
                    if placeholder_values and isinstance(placeholder_values, dict):
                        recipient_placeholder_values = placeholder_values.copy()
                        # You can add recipient-specific values here if needed
                        recipient_placeholder_values['recipient_email'] = recipient
                        
                        for key, value in recipient_placeholder_values.items():
                            placeholder = f"{{{{{key}}}}}"
                            recipient_html = recipient_html.replace(placeholder, str(value))
                    
                    # Add plain text alternative first (for clients that don't support HTML)
                    plain_text = "Este correo contiene contenido HTML que tu cliente de correo no puede mostrar."
                    alt_part.attach(MIMEText(plain_text, 'plain'))
                    
                    # Then add the HTML version
                    alt_part.attach(MIMEText(recipient_html, 'html'))
                    
                    # Embed images with Content-ID references
                    for i, img_path in enumerate(img_paths):
                        try:
                            with open(img_path, 'rb') as img_file:
                                img_data = img_file.read()
                                
                            # Create image attachment with Content-ID
                            img_filename = os.path.basename(img_path)
                            img = MIMEImage(img_data)
                            
                            # Set Content-ID based on filename
                            if img_filename == "info.jpeg":
                                img_id = "header_image"
                            elif img_filename == "footer.jpeg":
                                img_id = "footer_image"
                            else:
                                img_id = f"image_{i}"
                                
                            img.add_header('Content-ID', f'<{img_id}>')
                            img.add_header('Content-Disposition', 'inline', filename=img_filename)
                            
                            # Add the image to the message
                            msg.attach(img)
                        except Exception as img_error:
                            print(f"Warning: Could not embed image {img_path}: {img_error}")
                else:
                    # Use plain text if no HTML template
                    alt_part.attach(MIMEText(message_body, 'plain'))
                
                # Send the email
                server.send_message(msg)
                results["success"].append(recipient)
                print(f"Email sent successfully to {recipient}")
                
            except Exception as e:
                results["failed"].append({"email": recipient, "error": str(e)})
                print(f"Failed to send email to {recipient}: {e}")
        
        # Close the connection
        server.quit()
        
    except Exception as e:
        print(f"Error setting up email server: {e}")
        # Add more detailed error information
        import traceback
        traceback.print_exc()
    
    return results

# %% enviar emails en lotes
def send_emails_in_batches(email_list, subject, message_body=None, template_path=None, placeholder_values=None, 
                          sender_email=None, sender_password=None, smtp_server="smtpout.secureserver.net", smtp_port=465,
                          batch_size=20, delay_between_batches=60, sent_emails_file=None, max_emails=None):
    """
    Send emails to a list of recipients in batches to avoid spam detection.
    
    Parameters:
    email_list (list): List of recipient email addresses
    subject (str): Email subject
    message_body (str, optional): Plain text email body content (used if template_path is None)
    template_path (str, optional): Path to HTML template file
    placeholder_values (dict, optional): Dictionary of values to replace placeholders in the template
    sender_email (str): Sender's email address
    sender_password (str): Sender's email password or app password
    smtp_server (str, optional): SMTP server address
    smtp_port (int, optional): SMTP server port
    batch_size (int, optional): Number of emails to send in each batch
    delay_between_batches (int, optional): Delay in seconds between batches
    sent_emails_file (str, optional): Path to JSON file to track sent emails
    max_emails (int, optional): Maximum number of emails to send before stopping
    
    Returns:
    dict: Results of the email sending operation
    """
    all_results = {
        "success": [],
        "failed": []
    }
    
    # Load already sent emails if tracking file exists
    sent_emails = set()
    if sent_emails_file and os.path.exists(sent_emails_file):
        try:
            with open(sent_emails_file, 'r') as f:
                sent_data = json.load(f)
                sent_emails = set(sent_data.get('sent_emails', []))
        except Exception as e:
            print(f"Warning: Could not load sent emails tracking file: {e}")
    
    # Filter out already sent emails
    email_list = [email for email in email_list if email not in sent_emails]
    
    # Apply max_emails limit if specified
    if max_emails is not None and max_emails > 0:
        if len(email_list) > max_emails:
            print(f"Limiting to {max_emails} emails as specified by max_emails parameter")
            email_list = email_list[:max_emails]
    
    # Process emails in batches
    total_emails = len(email_list)
    print(f"Preparing to send {total_emails} emails in batches of {batch_size}")
    
    for i in range(0, total_emails, batch_size):
        batch = email_list[i:i+batch_size]
        batch_num = (i // batch_size) + 1
        total_batches = (total_emails + batch_size - 1) // batch_size
        
        print(f"\nSending batch {batch_num}/{total_batches} ({len(batch)} emails)...")
        
        # Send the current batch
        batch_results = send_emails(
            batch, 
            subject, 
            message_body, 
            template_path, 
            placeholder_values, 
            sender_email, 
            sender_password, 
            smtp_server, 
            smtp_port
        )
        
        # Update overall results
        all_results["success"].extend(batch_results["success"])
        all_results["failed"].extend(batch_results["failed"])
        
        # Update sent emails tracking file
        if sent_emails_file and batch_results["success"]:
            sent_emails.update(batch_results["success"])
            try:
                with open(sent_emails_file, 'w') as f:
                    json.dump({"sent_emails": list(sent_emails), "last_updated": datetime.now().isoformat()}, f, indent=2)
                print(f"Updated sent emails tracking file with {len(batch_results['success'])} new emails")
            except Exception as e:
                print(f"Warning: Could not update sent emails tracking file: {e}")
        
        # Wait between batches if not the last batch
        if i + batch_size < total_emails:
            print(f"Waiting {delay_between_batches} seconds before sending next batch...")
            time.sleep(delay_between_batches)
    
    print(f"\nEmail sending completed:")
    print(f"Successfully sent: {len(all_results['success'])}")
    print(f"Failed: {len(all_results['failed'])}")
    
    return all_results

# %% ejecutar proceso principal
def main():
    # Set up command line arguments
    parser = argparse.ArgumentParser(description='Extract email addresses from an Excel file and optionally send emails')
    parser.add_argument('file_path', help='Path to the Excel file')
    parser.add_argument('-s', '--sheet', default=0, help='Name or index of the sheet to read')
    parser.add_argument('--send', action='store_true', help='Send emails to the extracted addresses')
    parser.add_argument('--subject', default='Test Email', help='Email subject (when sending emails)')
    parser.add_argument('--body', default='This is a test email sent from Python.', help='Email body content (when sending emails)')
    parser.add_argument('--template', help='Path to HTML template file for email content')
    parser.add_argument('--sender', help='Sender email address (required when sending emails)')
    parser.add_argument('--password', help='Sender email password (required when sending emails)')
    parser.add_argument('--smtp', default='smtpout.secureserver.net', help='SMTP server address')
    parser.add_argument('--port', type=int, default=465, help='SMTP server port')
    parser.add_argument('--batch-size', type=int, default=20, help='Number of emails to send in each batch')
    parser.add_argument('--delay', type=int, default=60, help='Delay in seconds between batches')
    parser.add_argument('--tracking-file', default='sent_emails.json', help='Path to JSON file to track sent emails')
    parser.add_argument('--max-emails', type=int, help='Maximum number of emails to send before stopping')
    
    args = parser.parse_args()
    
    # Create tracking file directory if it doesn't exist
    tracking_file_dir = os.path.dirname(args.tracking_file)
    if tracking_file_dir and not os.path.exists(tracking_file_dir):
        os.makedirs(tracking_file_dir)
    
    # Extract email addresses
    emails, _ = read_emails_from_excel(args.file_path, args.sheet, args.tracking_file)
    
    # Print the extracted email addresses
    if emails:
        print(f"Found {len(emails)} email addresses to send:")
        for email in emails[:5]:  # Show only first 5 for brevity
            print(email)
        if len(emails) > 5:
            print(f"... and {len(emails) - 5} more")
    else:
        print("No email addresses found, or all emails have already been sent.")
        return
    
    # Add SMTP server validation
    if args.send:
        print(f"Will attempt to connect to SMTP server: {args.smtp} on port {args.port}")
        try:
            # Test SMTP connection before sending emails
            if args.port == 465:
                test_server = smtplib.SMTP_SSL(args.smtp, args.port, timeout=10)
            else:
                test_server = smtplib.SMTP(args.smtp, args.port, timeout=10)
                test_server.starttls()
            
            test_server.ehlo()
            print("SMTP server connection test successful")
            
            if args.sender and args.password:
                test_server.login(args.sender, args.password)
                print("SMTP authentication test successful")
            
            test_server.quit()
        except Exception as e:
            print(f"SMTP server connection test failed: {e}")
            print("You may need to:")
            print("1. Check your SMTP server settings")
            print("2. Verify your username and password")
            print("3. Check if your email provider allows SMTP access")
            print("4. Check if you need to enable 'less secure apps' or generate an app password")
            import traceback
            traceback.print_exc()
            return
    
    # Send emails if requested
    if args.send:
        if not args.sender or not args.password:
            print("Error: Sender email and password are required to send emails.")
            return
        
        print(f"\nPreparing to send emails to {len(emails)} recipients...")
        if args.max_emails:
            print(f"Will stop after sending {args.max_emails} emails as specified")
        
        # Determine if we're using a template or plain text
        if args.template:
            if not os.path.exists(args.template):
                print(f"Error: Template file {args.template} does not exist.")
                return
                
            # Example placeholder values - you can customize this
            placeholder_values = {
                "current_date": datetime.now().strftime("%Y-%m-%d"),
                "sender_name": args.sender.split('@')[0] if '@' in args.sender else args.sender,
                # Add more placeholder values as needed
            }
            
            results = send_emails_in_batches(
                emails, 
                args.subject, 
                template_path=args.template,
                placeholder_values=placeholder_values,
                sender_email=args.sender, 
                sender_password=args.password,
                smtp_server=args.smtp,
                smtp_port=args.port,
                batch_size=args.batch_size,
                delay_between_batches=args.delay,
                sent_emails_file=args.tracking_file,
                max_emails=args.max_emails
            )
        else:
            # Use plain text body
            results = send_emails_in_batches(
                emails, 
                args.subject, 
                args.body, 
                sender_email=args.sender, 
                sender_password=args.password,
                smtp_server=args.smtp,
                smtp_port=args.port,
                batch_size=args.batch_size,
                delay_between_batches=args.delay,
                sent_emails_file=args.tracking_file,
                max_emails=args.max_emails
            )
        
        print(f"\nEmail sending summary:")
        print(f"Successfully sent: {len(results['success'])}")
        print(f"Failed: {len(results['failed'])}")
        
        if results['failed']:
            print("\nFailed emails:")
            for failed in results['failed'][:5]:  # Show only first 5 failures for brevity
                print(f"  {failed['email']}: {failed['error']}")
            if len(results['failed']) > 5:
                print(f"  ... and {len(results['failed']) - 5} more")

if __name__ == "__main__":
    main()