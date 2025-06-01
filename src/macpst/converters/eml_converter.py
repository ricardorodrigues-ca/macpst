"""
EML format converter
"""

import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate, parseaddr
import logging
from ..core.pst_parser import EmailMessage
from .base_converter import BaseConverter


logger = logging.getLogger(__name__)


class EMLConverter(BaseConverter):
    """Converter for EML format"""
    
    def get_file_extension(self) -> str:
        return "eml"
    
    def convert_message(self, message: EmailMessage, output_filename: str) -> bool:
        """Convert email message to EML format"""
        try:
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
            
            output_path = self.get_output_path(output_filename)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(msg.as_string())
            
            logger.debug(f"Converted message to EML: {output_filename}")
            return True
            
        except Exception as e:
            logger.error(f"Error converting message to EML: {e}")
            return False