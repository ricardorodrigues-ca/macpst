"""
PDF format converter
"""

import html
import logging
from datetime import datetime
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from ..core.pst_parser import EmailMessage
from .base_converter import BaseConverter


logger = logging.getLogger(__name__)


class PDFConverter(BaseConverter):
    """Converter for PDF format"""
    
    def get_file_extension(self) -> str:
        return "pdf"
    
    def convert_message(self, message: EmailMessage, output_filename: str) -> bool:
        """Convert email message to PDF format"""
        try:
            output_path = self.get_output_path(output_filename)
            doc = SimpleDocTemplate(str(output_path), pagesize=A4,
                                  rightMargin=72, leftMargin=72,
                                  topMargin=72, bottomMargin=18)
            
            styles = getSampleStyleSheet()
            story = []
            
            header_style = ParagraphStyle(
                'CustomHeader',
                parent=styles['Heading1'],
                fontSize=16,
                spaceAfter=12,
                textColor=colors.darkblue
            )
            
            meta_style = ParagraphStyle(
                'CustomMeta',
                parent=styles['Normal'],
                fontSize=10,
                spaceAfter=6,
                textColor=colors.darkgray
            )
            
            body_style = ParagraphStyle(
                'CustomBody',
                parent=styles['Normal'],
                fontSize=11,
                spaceAfter=12,
                leading=14
            )
            
            story.append(Paragraph("Email Message", header_style))
            story.append(Spacer(1, 12))
            
            subject = self._sanitize_text(message.subject) or "No Subject"
            story.append(Paragraph(f"<b>Subject:</b> {html.escape(subject)}", meta_style))
            
            sender = self._sanitize_text(message.sender) or "Unknown Sender"
            story.append(Paragraph(f"<b>From:</b> {html.escape(sender)}", meta_style))
            
            if message.recipients:
                recipients = ', '.join(self._sanitize_text(r) for r in message.recipients)
                story.append(Paragraph(f"<b>To:</b> {html.escape(recipients)}", meta_style))
            
            if message.cc_recipients:
                cc_recipients = ', '.join(self._sanitize_text(r) for r in message.cc_recipients)
                story.append(Paragraph(f"<b>CC:</b> {html.escape(cc_recipients)}", meta_style))
            
            if message.sent_time:
                sent_time = message.sent_time.strftime("%Y-%m-%d %H:%M:%S")
                story.append(Paragraph(f"<b>Sent:</b> {sent_time}", meta_style))
            elif message.received_time:
                received_time = message.received_time.strftime("%Y-%m-%d %H:%M:%S")
                story.append(Paragraph(f"<b>Received:</b> {received_time}", meta_style))
            
            if message.folder_path:
                story.append(Paragraph(f"<b>Folder:</b> {html.escape(message.folder_path)}", meta_style))
            
            if message.attachments:
                attachments = ', '.join(att.get('name', 'Unknown') for att in message.attachments)
                story.append(Paragraph(f"<b>Attachments:</b> {html.escape(attachments)}", meta_style))
            
            story.append(Spacer(1, 24))
            
            story.append(Paragraph("<b>Message Body:</b>", meta_style))
            story.append(Spacer(1, 12))
            
            body_text = ""
            if message.body_text:
                body_text = self._sanitize_text(message.body_text)
            elif message.body_html:
                body_text = self._html_to_text(message.body_html)
            else:
                body_text = "(No message body)"
            
            body_paragraphs = body_text.split('\n')
            for para in body_paragraphs:
                if para.strip():
                    story.append(Paragraph(html.escape(para), body_style))
                else:
                    story.append(Spacer(1, 6))
            
            doc.build(story)
            logger.debug(f"Converted message to PDF: {output_filename}")
            return True
            
        except Exception as e:
            logger.error(f"Error converting message to PDF: {e}")
            return False
    
    def _html_to_text(self, html_content: str) -> str:
        """Convert HTML content to plain text"""
        try:
            import re
            
            html_content = re.sub(r'<script[^>]*>.*?</script>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
            html_content = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
            html_content = re.sub(r'<[^>]+>', '', html_content)
            html_content = html.unescape(html_content)
            
            lines = html_content.split('\n')
            cleaned_lines = [line.strip() for line in lines if line.strip()]
            
            return '\n'.join(cleaned_lines)
            
        except Exception as e:
            logger.warning(f"Error converting HTML to text: {e}")
            return self._sanitize_text(html_content)