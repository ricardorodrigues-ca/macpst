"""
Unit tests for converters
"""

import unittest
import tempfile
import os
import shutil
from pathlib import Path
from datetime import datetime
from unittest.mock import patch, Mock

from src.macpst.core.pst_parser import EmailMessage
from src.macpst.converters.base_converter import BaseConverter
from src.macpst.converters.eml_converter import EMLConverter
from src.macpst.converters.mbox_converter import MBOXConverter
from src.macpst.converters.pdf_converter import PDFConverter
from src.macpst.core.converter import Converter


class MockConverter(BaseConverter):
    """Mock converter for testing base functionality"""
    
    def get_file_extension(self):
        return "mock"
    
    def convert_message(self, message, output_filename):
        output_path = self.get_output_path(output_filename)
        with open(output_path, 'w') as f:
            f.write(f"Mock conversion of: {message.subject}")
        return True


class TestBaseConverter(unittest.TestCase):
    """Test BaseConverter functionality"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.converter = MockConverter(self.temp_dir)
        
        self.sample_message = EmailMessage(
            subject="Test Subject",
            sender="sender@example.com",
            recipients=["recipient@example.com"],
            body_text="Test message body",
            sent_time=datetime(2023, 1, 1, 12, 0, 0)
        )
    
    def tearDown(self):
        """Clean up test fixtures"""
        shutil.rmtree(self.temp_dir)
    
    def test_converter_initialization(self):
        """Test converter initialization"""
        self.assertEqual(self.converter.output_dir, Path(self.temp_dir))
        self.assertTrue(self.converter.output_dir.exists())
        self.assertEqual(self.converter.converted_count, 0)
        self.assertEqual(self.converter.error_count, 0)
    
    def test_generate_filename(self):
        """Test filename generation"""
        filename = self.converter._generate_filename(self.sample_message, 0)
        
        self.assertTrue(filename.endswith('.mock'))
        self.assertIn('Test_Subject', filename)
        self.assertIn('20230101', filename)
    
    def test_generate_filename_with_special_characters(self):
        """Test filename generation with special characters"""
        message = EmailMessage(
            subject="Test: Special/Characters\\In|Subject",
            sender="sender@example.com",
            recipients=["recipient@example.com"]
        )
        
        filename = self.converter._generate_filename(message, 0)
        
        # Should not contain special characters
        self.assertNotIn(':', filename)
        self.assertNotIn('/', filename)
        self.assertNotIn('\\', filename)
        self.assertNotIn('|', filename)
    
    def test_generate_filename_collision_handling(self):
        """Test filename collision handling"""
        # Create a file to cause collision
        first_filename = self.converter._generate_filename(self.sample_message, 0)
        collision_path = self.converter.get_output_path(first_filename)
        collision_path.touch()
        
        # Generate filename again - should get different name
        second_filename = self.converter._generate_filename(self.sample_message, 0)
        
        self.assertNotEqual(first_filename, second_filename)
        self.assertTrue(second_filename.endswith('.mock'))
    
    def test_sanitize_text(self):
        """Test text sanitization"""
        test_text = "Normal text with émojis 🚀 and special chars"
        sanitized = self.converter._sanitize_text(test_text)
        
        self.assertIsInstance(sanitized, str)
        self.assertTrue(len(sanitized) > 0)
    
    def test_convert_single_message(self):
        """Test converting a single message"""
        filename = "test_message.mock"
        result = self.converter.convert_message(self.sample_message, filename)
        
        self.assertTrue(result)
        output_path = self.converter.get_output_path(filename)
        self.assertTrue(output_path.exists())
        
        with open(output_path, 'r') as f:
            content = f.read()
            self.assertIn("Test Subject", content)
    
    def test_convert_multiple_messages(self):
        """Test converting multiple messages"""
        messages = [
            self.sample_message,
            EmailMessage(
                subject="Second Message",
                sender="sender2@example.com",
                recipients=["recipient2@example.com"],
                body_text="Second message body"
            )
        ]
        
        result = self.converter.convert_messages(messages)
        
        self.assertEqual(result['total_messages'], 2)
        self.assertEqual(result['converted_count'], 2)
        self.assertEqual(result['error_count'], 0)
        self.assertEqual(len(result['errors']), 0)


class TestEMLConverter(unittest.TestCase):
    """Test EML converter"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.converter = EMLConverter(self.temp_dir)
        
        self.sample_message = EmailMessage(
            subject="Test EML Subject",
            sender="sender@example.com",
            recipients=["recipient@example.com"],
            cc_recipients=["cc@example.com"],
            body_text="Plain text body",
            body_html="<html><body>HTML body</body></html>",
            sent_time=datetime(2023, 1, 1, 12, 0, 0),
            message_id="<test123@example.com>"
        )
    
    def tearDown(self):
        """Clean up test fixtures"""
        shutil.rmtree(self.temp_dir)
    
    def test_get_file_extension(self):
        """Test EML file extension"""
        self.assertEqual(self.converter.get_file_extension(), "eml")
    
    def test_convert_message_with_both_text_and_html(self):
        """Test converting message with both text and HTML"""
        filename = "test.eml"
        result = self.converter.convert_message(self.sample_message, filename)
        
        self.assertTrue(result)
        
        output_path = self.converter.get_output_path(filename)
        self.assertTrue(output_path.exists())
        
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
            self.assertIn("Subject: Test EML Subject", content)
            self.assertIn("From: sender@example.com", content)
            self.assertIn("To: recipient@example.com", content)
            self.assertIn("Cc: cc@example.com", content)
            self.assertIn("Message-ID: <test123@example.com>", content)
            self.assertIn("Plain text body", content)
            self.assertIn("HTML body", content)
    
    def test_convert_message_text_only(self):
        """Test converting message with text only"""
        message = EmailMessage(
            subject="Text Only",
            sender="sender@example.com",
            recipients=["recipient@example.com"],
            body_text="Only plain text content"
        )
        
        filename = "text_only.eml"
        result = self.converter.convert_message(message, filename)
        
        self.assertTrue(result)
        
        output_path = self.converter.get_output_path(filename)
        with open(output_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
            self.assertIn("Only plain text content", content)
            self.assertNotIn("multipart", content.lower())


class TestMBOXConverter(unittest.TestCase):
    """Test MBOX converter"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.converter = MBOXConverter(self.temp_dir)
        
        self.sample_messages = [
            EmailMessage(
                subject="First Message",
                sender="sender1@example.com",
                recipients=["recipient@example.com"],
                body_text="First message body",
                sent_time=datetime(2023, 1, 1, 12, 0, 0)
            ),
            EmailMessage(
                subject="Second Message",
                sender="sender2@example.com",
                recipients=["recipient@example.com"],
                body_text="Second message body",
                sent_time=datetime(2023, 1, 2, 12, 0, 0)
            )
        ]
    
    def tearDown(self):
        """Clean up test fixtures"""
        shutil.rmtree(self.temp_dir)
    
    def test_get_file_extension(self):
        """Test MBOX file extension"""
        self.assertEqual(self.converter.get_file_extension(), "mbox")
    
    def test_convert_messages_to_single_mbox(self):
        """Test converting multiple messages to single MBOX file"""
        filename = "test.mbox"
        result = self.converter.convert_messages_to_single_mbox(self.sample_messages, filename)
        
        self.assertTrue(result)
        
        output_path = self.converter.get_output_path(filename)
        self.assertTrue(output_path.exists())
        
        # Verify MBOX content
        import mailbox
        mbox = mailbox.mbox(str(output_path))
        messages = list(mbox)
        
        self.assertEqual(len(messages), 2)
        self.assertEqual(messages[0]['Subject'], "First Message")
        self.assertEqual(messages[1]['Subject'], "Second Message")
        
        mbox.close()
    
    def test_convert_messages_single_file_mode(self):
        """Test batch conversion in single file mode"""
        result = self.converter.convert_messages(self.sample_messages, single_file=True)
        
        self.assertEqual(result['total_messages'], 2)
        self.assertEqual(result['converted_count'], 2)
        self.assertEqual(result['error_count'], 0)


class TestPDFConverter(unittest.TestCase):
    """Test PDF converter"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.converter = PDFConverter(self.temp_dir)
        
        self.sample_message = EmailMessage(
            subject="Test PDF Subject",
            sender="sender@example.com",
            recipients=["recipient@example.com"],
            cc_recipients=["cc@example.com"],
            body_text="Plain text body for PDF conversion",
            body_html="<html><body><p>HTML body for PDF</p></body></html>",
            sent_time=datetime(2023, 1, 1, 12, 0, 0),
            folder_path="/Inbox",
            attachments=[{"name": "attachment.pdf", "size": 1024}]
        )
    
    def tearDown(self):
        """Clean up test fixtures"""
        shutil.rmtree(self.temp_dir)
    
    def test_get_file_extension(self):
        """Test PDF file extension"""
        self.assertEqual(self.converter.get_file_extension(), "pdf")
    
    @patch('src.macpst.converters.pdf_converter.SimpleDocTemplate')
    def test_convert_message(self, mock_doc_class):
        """Test PDF message conversion"""
        mock_doc = Mock()
        mock_doc_class.return_value = mock_doc
        
        filename = "test.pdf"
        result = self.converter.convert_message(self.sample_message, filename)
        
        self.assertTrue(result)
        mock_doc.build.assert_called_once()
    
    def test_html_to_text_conversion(self):
        """Test HTML to text conversion"""
        html_content = """
        <html>
            <head><title>Test</title></head>
            <body>
                <p>First paragraph</p>
                <div>Second section</div>
                <script>alert('test');</script>
            </body>
        </html>
        """
        
        text_result = self.converter._html_to_text(html_content)
        
        self.assertIn("First paragraph", text_result)
        self.assertIn("Second section", text_result)
        self.assertNotIn("<p>", text_result)
        self.assertNotIn("alert", text_result)  # Script should be removed


class TestConverter(unittest.TestCase):
    """Test main Converter class"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.converter = Converter(self.temp_dir)
        
        self.sample_messages = [
            EmailMessage(
                subject="Test Message 1",
                sender="sender1@example.com",
                recipients=["recipient@example.com"],
                body_text="First test message"
            ),
            EmailMessage(
                subject="Test Message 2",
                sender="sender2@example.com",
                recipients=["recipient@example.com"],
                body_text="Second test message"
            )
        ]
    
    def tearDown(self):
        """Clean up test fixtures"""
        shutil.rmtree(self.temp_dir)
    
    def test_get_supported_formats(self):
        """Test getting supported formats"""
        formats = self.converter.get_supported_formats()
        
        self.assertIn('eml', formats)
        self.assertIn('mbox', formats)
        self.assertIn('pdf', formats)
    
    def test_convert_messages_unsupported_format(self):
        """Test conversion with unsupported format"""
        with self.assertRaises(ValueError):
            self.converter.convert_messages(self.sample_messages, 'unsupported')
    
    @patch('src.macpst.converters.eml_converter.EMLConverter.convert_messages')
    def test_convert_messages_eml(self, mock_convert):
        """Test EML conversion through main converter"""
        mock_convert.return_value = {
            'total_messages': 2,
            'converted_count': 2,
            'error_count': 0,
            'errors': [],
            'output_directory': str(self.temp_dir)
        }
        
        result = self.converter.convert_messages(self.sample_messages, 'eml')
        
        mock_convert.assert_called_once()
        self.assertEqual(result['converted_count'], 2)
    
    def test_batch_convert_multiple_formats(self):
        """Test batch conversion to multiple formats"""
        with patch('src.macpst.converters.eml_converter.EMLConverter.convert_messages') as mock_eml, \
             patch('src.macpst.converters.pdf_converter.PDFConverter.convert_messages') as mock_pdf:
            
            mock_eml.return_value = {
                'total_messages': 2, 'converted_count': 2, 'error_count': 0,
                'errors': [], 'output_directory': str(self.temp_dir)
            }
            mock_pdf.return_value = {
                'total_messages': 2, 'converted_count': 2, 'error_count': 0,
                'errors': [], 'output_directory': str(self.temp_dir)
            }
            
            results = self.converter.batch_convert(self.sample_messages, ['eml', 'pdf'])
            
            self.assertIn('eml', results)
            self.assertIn('pdf', results)
            self.assertEqual(results['eml']['converted_count'], 2)
            self.assertEqual(results['pdf']['converted_count'], 2)


if __name__ == '__main__':
    unittest.main()