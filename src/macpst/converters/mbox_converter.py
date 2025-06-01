"""
MBOX format converter
"""

import mailbox
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
import logging
from typing import List
from ..core.pst_parser import EmailMessage
from .base_converter import BaseConverter


logger = logging.getLogger(__name__)


class MBOXConverter(BaseConverter):
    """Converter for MBOX format"""
    
    def get_file_extension(self) -> str:
        return "mbox"
    
    def convert_message(self, message: EmailMessage, output_filename: str) -> bool:
        """Convert single email message to MBOX format"""
        return self.convert_messages_to_single_mbox([message], output_filename)
    
    def convert_messages_to_single_mbox(self, messages: List[EmailMessage], 
                                      output_filename: str) -> bool:
        """Convert multiple messages to a single MBOX file"""
        try:
            output_path = self.get_output_path(output_filename)
            mbox = mailbox.mbox(str(output_path))
            
            for message in messages:
                try:
                    email_msg = self._create_email_message(message)
                    mbox.add(email_msg)
                except Exception as e:
                    logger.warning(f"Error adding message to MBOX: {e}")
                    continue
            
            mbox.close()
            logger.debug(f"Converted {len(messages)} messages to MBOX: {output_filename}")
            return True
            
        except Exception as e:
            logger.error(f"Error converting messages to MBOX: {e}")
            return False
    
    def _create_email_message(self, message: EmailMessage):
        """Create email.message.EmailMessage from EmailMessage"""
        if message.body_html and message.body_text:
            msg = MIMEMultipart('alternative')
            text_part = MIMEText(self._sanitize_text(message.body_text), 'plain', 'utf-8')
            html_part = MIMEText(self._sanitize_text(message.body_html), 'html', 'utf-8')
            msg.attach(text_part)
            msg.attach(html_part)
        elif message.body_html:
            msg = MIMEText(self._sanitize_text(message.body_html), 'html', 'utf-8')
        else:
            msg = MIMEText(self._sanitize_text(message.body_text), 'plain', 'utf-8')
        
        msg['Subject'] = self._sanitize_text(message.subject)
        msg['From'] = self._sanitize_text(message.sender)
        
        if message.recipients:
            msg['To'] = ', '.join(self._sanitize_text(r) for r in message.recipients)
        
        if message.cc_recipients:
            msg['Cc'] = ', '.join(self._sanitize_text(r) for r in message.cc_recipients)
        
        if message.sent_time:
            msg['Date'] = formatdate(message.sent_time.timestamp(), localtime=True)
        elif message.received_time:
            msg['Date'] = formatdate(message.received_time.timestamp(), localtime=True)
        
        if message.message_id:
            msg['Message-ID'] = self._sanitize_text(message.message_id)
        
        return msg
    
    def convert_messages(self, messages: List[EmailMessage], 
                        progress_callback=None, single_file=True) -> dict:
        """Convert messages to MBOX format"""
        if single_file:
            filename = f"pst_export_{len(messages)}_messages.mbox"
            success = self.convert_messages_to_single_mbox(messages, filename)
            
            return {
                'total_messages': len(messages),
                'converted_count': len(messages) if success else 0,
                'error_count': 0 if success else len(messages),
                'errors': [] if success else ['Failed to create MBOX file'],
                'output_directory': str(self.output_dir)
            }
        else:
            return super().convert_messages(messages, progress_callback)