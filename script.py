# %% importar librerias
import pandas as pd
import argparse
import smtplib
import os
import re
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from pathlib import Path

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
def read_emails_from_excel(file_path, sheet_name=0):
    """
    Read an Excel file and extract all email addresses from the 'E-mail' column.
    
    Parameters:
    file_path (str): Path to the Excel file
    sheet_name (str or int, optional): Name or index of the sheet to read
    
    Returns:
    list: Email addresses from the 'E-mail' column
    """
    try:
        # Read the Excel file
        emails = read_excel_column(file_path, column_name="E-mail", sheet_name=sheet_name)
        
        # Filter out any None or empty values
        if emails:
            emails = [email for email in emails if email and isinstance(email, str)]
        
        return emails
    
    except Exception as e:
        print(f"Error reading emails: {e}")
        return []

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
                sender_email=None, sender_password=None, smtp_server="smtp.gmail.com", smtp_port=587):
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
    
    Returns:
    dict: Results of the email sending operation
    """
    results = {
        "success": [],
        "failed": []
    }
    
    try:
        # Set up the SMTP server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection
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
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = recipient
                msg['Subject'] = subject
                
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
                    
                    # Attach HTML content after embedding images
                    msg.attach(MIMEText(recipient_html, 'html'))
                else:
                    # Use plain text if no HTML template
                    msg.attach(MIMEText(message_body, 'plain'))
                
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
    
    return results

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
    parser.add_argument('--smtp', default='smtp.gmail.com', help='SMTP server address')
    parser.add_argument('--port', type=int, default=587, help='SMTP server port')
    
    args = parser.parse_args()
    
    # Extract email addresses
    emails = read_emails_from_excel(args.file_path, args.sheet)
    
    # Print the extracted email addresses
    if emails:
        print(f"Found {len(emails)} email addresses:")
        for email in emails:
            print(email)
    else:
        print("No email addresses found or the 'E-mail' column doesn't exist.")
        return
    
    # Send emails if requested
    if args.send:
        if not args.sender or not args.password:
            print("Error: Sender email and password are required to send emails.")
            return
        
        print(f"\nSending emails to {len(emails)} recipients...")
        
        # Determine if we're using a template or plain text
        if args.template:
            if not os.path.exists(args.template):
                print(f"Error: Template file {args.template} does not exist.")
                return
                
            # Example placeholder values - you can customize this
            placeholder_values = {
                "current_date": pd.Timestamp.now().strftime("%Y-%m-%d"),
                "sender_name": args.sender.split('@')[0] if '@' in args.sender else args.sender,
                # Add more placeholder values as needed
            }
            
            results = send_emails(
                emails, 
                args.subject, 
                template_path=args.template,
                placeholder_values=placeholder_values,
                sender_email=args.sender, 
                sender_password=args.password,
                smtp_server=args.smtp,
                smtp_port=args.port
            )
        else:
            # Use plain text body
            results = send_emails(
                emails, 
                args.subject, 
                args.body, 
                sender_email=args.sender, 
                sender_password=args.password,
                smtp_server=args.smtp,
                smtp_port=args.port
            )
        
        print(f"\nEmail sending results:")
        print(f"Successfully sent: {len(results['success'])}")
        print(f"Failed: {len(results['failed'])}")

if __name__ == "__main__":
    main()