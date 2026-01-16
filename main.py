#!/usr/bin/env python3
"""
Simplified PDF Merge and Email System
"""

import os
import sys
import time
import logging
import configparser
from datetime import datetime
from pathlib import Path

# PDF Processing
try:
    from PyPDF2 import PdfReader, PdfWriter
except ImportError:
    print("ERROR: PyPDF2 not installed. Run: pip install PyPDF2")
    sys.exit(1)

# Email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

class SimplePDFMerger:
    """Simple PDF Merger and Email Sender"""
    
    def __init__(self):
        self.setup_directories()
        self.setup_logging()
        self.load_config()
        
    def setup_directories(self):
        """Create necessary directories"""
        dirs = ['incoming_pdfs', 'archive', 'merged_pdfs', 'logs']
        for dir_name in dirs:
            Path(dir_name).mkdir(exist_ok=True)
        print("✓ Created all directories")
    
    def setup_logging(self):
        """Setup logging system"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('logs/pdf_processor.log'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger('PDFProcessor')
        self.logger.info("PDF Processor started")
    
    def load_config(self):
        """Load configuration from config.ini"""
        if not os.path.exists('config.ini'):
            print("ERROR: config.ini not found!")
            print("Please copy config.ini.template to config.ini and edit it")
            sys.exit(1)
        
        self.config = configparser.ConfigParser()
        self.config.read('config.ini')
        
        # Get settings with defaults
        self.source_folder = self.config.get('PATHS', 'source_folder', fallback='incoming_pdfs')
        self.archive_folder = self.config.get('PATHS', 'archive_folder', fallback='archive')
        self.merged_folder = self.config.get('PATHS', 'merged_folder', fallback='merged_pdfs')
        
        self.logger.info("Configuration loaded successfully")
    
    def find_pdf_files(self):
        """Find all PDF files in source folder"""
        pdf_files = []
        source_path = Path(self.source_folder)
        
        if not source_path.exists():
            self.logger.error(f"Source folder not found: {self.source_folder}")
            return pdf_files
        
        for file in source_path.glob("*.pdf"):
            pdf_files.append(str(file))
        
        self.logger.info(f"Found {len(pdf_files)} PDF files")
        return sorted(pdf_files)
    
    def merge_pdfs(self, pdf_files):
        """Merge PDF files into one"""
        if not pdf_files:
            self.logger.warning("No PDF files to merge")
            return None
        
        try:
            merger = PdfWriter()
            
            for pdf_file in pdf_files:
                try:
                    reader = PdfReader(pdf_file)
                    for page in reader.pages:
                        merger.add_page(page)
                    self.logger.info(f"Added: {os.path.basename(pdf_file)}")
                except Exception as e:
                    self.logger.error(f"Error reading {pdf_file}: {e}")
                    continue
            
            # Create output filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"merged_{timestamp}.pdf"
            output_path = os.path.join(self.merged_folder, output_file)
            
            # Save merged PDF
            with open(output_path, 'wb') as output:
                merger.write(output)
            
            self.logger.info(f"✓ Merged {len(pdf_files)} files into {output_file}")
            return output_path
            
        except Exception as e:
            self.logger.error(f"Error merging PDFs: {e}")
            return None
    
    def send_email(self, attachment_path):
        """Send email with merged PDF"""
        try:
            # Get email settings
            smtp_server = self.config.get('EMAIL', 'smtp_server', fallback='')
            smtp_port = self.config.getint('EMAIL', 'smtp_port', fallback=587)
            sender_email = self.config.get('EMAIL', 'sender_email', fallback='')
            sender_password = self.config.get('EMAIL', 'sender_password', fallback='')
            
            if not smtp_server or not sender_email:
                self.logger.warning("Email configuration incomplete. Skipping email.")
                return False
            
            # Get recipients
            recipients_str = self.config.get('EMAIL', 'recipients', fallback='')
            if not recipients_str:
                self.logger.warning("No recipients configured. Skipping email.")
                return False
            
            recipients = [r.strip() for r in recipients_str.split(',')]
            
            # Create email
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = ', '.join(recipients)
            msg['Subject'] = self.config.get('EMAIL', 'subject', fallback='Daily Merged PDF')
            
            # Add body
            body = self.config.get('EMAIL', 'body', fallback='Please find attached the merged PDF.')
            msg.attach(MIMEText(body, 'plain'))
            
            # Add attachment
            if attachment_path and os.path.exists(attachment_path):
                with open(attachment_path, 'rb') as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
                msg.attach(part)
            
            # Send email
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                if sender_password:
                    server.login(sender_email, sender_password)
                server.send_message(msg)
            
            self.logger.info(f"✓ Email sent to {len(recipients)} recipients")
            return True
            
        except Exception as e:
            self.logger.error(f"Error sending email: {e}")
            return False
    
    def archive_files(self, pdf_files):
        """Move processed files to archive"""
        for file in pdf_files:
            try:
                filename = os.path.basename(file)
                archive_path = os.path.join(self.archive_folder, filename)
                
                # If file exists, add timestamp
                if os.path.exists(archive_path):
                    name, ext = os.path.splitext(filename)
                    timestamp = datetime.now().strftime("%H%M%S")
                    archive_path = os.path.join(self.archive_folder, f"{name}_{timestamp}{ext}")
                
                os.rename(file, archive_path)
                self.logger.info(f"Archived: {filename}")
            except Exception as e:
                self.logger.error(f"Error archiving {file}: {e}")
    
    def run(self):
        """Main execution method"""
        self.logger.info("=" * 50)
        self.logger.info("Starting PDF merge process")
        
        # Step 1: Find PDFs
        pdf_files = self.find_pdf_files()
        if not pdf_files:
            self.logger.info("No PDF files found. Exiting.")
            return False
        
        # Step 2: Merge PDFs
        merged_file = self.merge_pdfs(pdf_files)
        if not merged_file:
            self.logger.error("Failed to merge PDFs")
            return False
        
        # Step 3: Send email
        email_sent = self.send_email(merged_file)
        
        # Step 4: Archive files
        self.archive_files(pdf_files)
        
        self.logger.info(f"Process completed successfully!")
        self.logger.info(f"- Merged {len(pdf_files)} files")
        self.logger.info(f"- Email sent: {'Yes' if email_sent else 'No'}")
        self.logger.info(f"- Output: {os.path.basename(merged_file)}")
        
        return True

def main():
    """Main function"""
    print("\n" + "=" * 50)
    print("AUTOMATED PDF MERGE SYSTEM")
    print("=" * 50)
    
    # Check dependencies
    try:
        import configparser
        import PyPDF2
    except ImportError as e:
        print(f"\n Missing dependency: {e}")
        print("Please install required packages:")
        print("pip install PyPDF2 configparser")
        return
    
    # Run the processor
    processor = SimplePDFMerger()
    success = processor.run()
    
    if success:
        print("\n Process completed successfully!")
        print("Check 'logs/pdf_processor.log' for details")
    else:
        print("\n        pyinstaller --onefile --name pdf_merge_email --add-data "config.ini;." --add-data "incoming_pdfs;incoming_pdfs" --add-data "archive;archive" --add-data "merged_pdfs;merged_pdfs" --add-data "logs;logs" --hidden-import PyPDF2 --console [main.py](http://_vscodecontentref_/0) Process completed with errors")
        print("Check 'logs/pdf_processor.log' for details")
    
    print("\nPress Enter to exit...")
    input()

if __name__ == "__main__":
    main()