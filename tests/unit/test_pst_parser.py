"""
Unit tests for PST Parser
"""

import unittest
import tempfile
import os
from pathlib import Path
from datetime import datetime
from unittest.mock import Mock, patch, MagicMock

from src.macpst.core.pst_parser import PSTParser, EmailMessage, PSTFolder


class TestEmailMessage(unittest.TestCase):
    """Test EmailMessage dataclass"""
    
    def test_email_message_creation(self):
        """Test EmailMessage creation with default values"""
        message = EmailMessage(
            subject="Test Subject",
            sender="test@example.com",
            recipients=["recipient@example.com"]
        )
        
        self.assertEqual(message.subject, "Test Subject")
        self.assertEqual(message.sender, "test@example.com")
        self.assertEqual(message.recipients, ["recipient@example.com"])
        self.assertEqual(message.cc_recipients, [])
        self.assertEqual(message.bcc_recipients, [])
        self.assertEqual(message.attachments, [])
    
    def test_email_message_with_all_fields(self):
        """Test EmailMessage with all fields populated"""
        sent_time = datetime.now()
        received_time = datetime.now()
        
        message = EmailMessage(
            subject="Test Subject",
            sender="sender@example.com",
            recipients=["recipient1@example.com", "recipient2@example.com"],
            cc_recipients=["cc@example.com"],
            bcc_recipients=["bcc@example.com"],
            body_text="Plain text body",
            body_html="<html>HTML body</html>",
            sent_time=sent_time,
            received_time=received_time,
            attachments=[{"name": "test.pdf", "size": 1024}],
            message_id="msg123",
            folder_path="/Inbox"
        )
        
        self.assertEqual(message.subject, "Test Subject")
        self.assertEqual(message.sender, "sender@example.com")
        self.assertEqual(len(message.recipients), 2)
        self.assertEqual(len(message.cc_recipients), 1)
        self.assertEqual(len(message.bcc_recipients), 1)
        self.assertEqual(message.body_text, "Plain text body")
        self.assertEqual(message.body_html, "<html>HTML body</html>")
        self.assertEqual(message.sent_time, sent_time)
        self.assertEqual(message.received_time, received_time)
        self.assertEqual(len(message.attachments), 1)
        self.assertEqual(message.message_id, "msg123")
        self.assertEqual(message.folder_path, "/Inbox")


class TestPSTFolder(unittest.TestCase):
    """Test PSTFolder dataclass"""
    
    def test_pst_folder_creation(self):
        """Test PSTFolder creation"""
        folder = PSTFolder(
            name="Inbox",
            path="/Inbox",
            message_count=10
        )
        
        self.assertEqual(folder.name, "Inbox")
        self.assertEqual(folder.path, "/Inbox")
        self.assertEqual(folder.message_count, 10)
        self.assertEqual(folder.subfolders, [])
    
    def test_pst_folder_with_subfolders(self):
        """Test PSTFolder with subfolders"""
        subfolder = PSTFolder(name="Sent", path="/Sent", message_count=5)
        folder = PSTFolder(
            name="Root",
            path="/",
            message_count=0,
            subfolders=[subfolder]
        )
        
        self.assertEqual(len(folder.subfolders), 1)
        self.assertEqual(folder.subfolders[0].name, "Sent")


class TestPSTParser(unittest.TestCase):
    """Test PSTParser class"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.test_pst_file = os.path.join(self.temp_dir, "test.pst")
        
        # Create a mock PST file with proper signature
        with open(self.test_pst_file, 'wb') as f:
            # Write PST signature
            f.write(b'!BDN')
            # Write file type (Unicode PST)
            f.write(b'\x00' * 6)  # padding
            f.write(b'\x17\x00')  # Unicode PST type
            # Write rest of header (500+ bytes)
            f.write(b'\x00' * 500)
    
    def tearDown(self):
        """Clean up test fixtures"""
        import shutil
        shutil.rmtree(self.temp_dir)
    
    def test_parser_creation_with_valid_file(self):
        """Test PSTParser creation with valid PST file"""
        parser = PSTParser(self.test_pst_file)
        self.assertEqual(parser.pst_file_path, Path(self.test_pst_file))
        self.assertIsNone(parser.file_handle)
        self.assertFalse(parser.is_unicode)
    
    def test_parser_creation_with_nonexistent_file(self):
        """Test PSTParser creation with non-existent file"""
        with self.assertRaises(FileNotFoundError):
            PSTParser("/nonexistent/file.pst")
    
    def test_parser_creation_with_invalid_signature(self):
        """Test PSTParser creation with invalid signature"""
        invalid_file = os.path.join(self.temp_dir, "invalid.pst")
        with open(invalid_file, 'wb') as f:
            f.write(b'INVALID_SIG')
        
        with self.assertRaises(ValueError):
            PSTParser(invalid_file)
    
    def test_parser_open_close(self):
        """Test opening and closing PST file"""
        parser = PSTParser(self.test_pst_file)
        
        # Test opening
        parser.open()
        self.assertIsNotNone(parser.file_handle)
        self.assertTrue(parser.is_unicode)
        self.assertIsNotNone(parser.header)
        
        # Test closing
        parser.close()
        self.assertIsNone(parser.file_handle)
    
    def test_parser_context_manager(self):
        """Test PSTParser as context manager"""
        with PSTParser(self.test_pst_file) as parser:
            self.assertIsNotNone(parser.file_handle)
            self.assertTrue(parser.is_unicode)
        
        # File should be closed after context
        self.assertIsNone(parser.file_handle)
    
    def test_get_statistics(self):
        """Test getting PST file statistics"""
        parser = PSTParser(self.test_pst_file)
        
        with parser:
            stats = parser.get_statistics()
            
            self.assertIn('file_path', stats)
            self.assertIn('file_size', stats)
            self.assertIn('is_unicode', stats)
            self.assertIn('total_messages', stats)
            self.assertIn('total_folders', stats)
            
            self.assertEqual(stats['file_path'], str(parser.pst_file_path))
            self.assertTrue(stats['file_size'] > 0)
            self.assertTrue(stats['is_unicode'])
    
    @patch('src.macpst.core.pst_parser.pypff')
    def test_get_folder_tree_with_pypff(self, mock_pypff):
        """Test folder tree extraction with pypff available"""
        # Mock pypff objects
        mock_file = Mock()
        mock_root_folder = Mock()
        mock_root_folder.name = "Root"
        mock_root_folder.number_of_messages = 10
        mock_root_folder.number_of_sub_folders = 1
        
        mock_subfolder = Mock()
        mock_subfolder.name = "Inbox"
        mock_subfolder.number_of_messages = 5
        mock_subfolder.number_of_sub_folders = 0
        
        mock_root_folder.get_sub_folder.return_value = mock_subfolder
        mock_file.get_root_folder.return_value = mock_root_folder
        mock_pypff.file.return_value = mock_file
        
        parser = PSTParser(self.test_pst_file)
        
        with parser:
            folder_tree = parser.get_folder_tree()
            
            self.assertIsInstance(folder_tree, PSTFolder)
            self.assertEqual(folder_tree.name, "Root")
            self.assertEqual(folder_tree.message_count, 10)
            self.assertEqual(len(folder_tree.subfolders), 1)
            self.assertEqual(folder_tree.subfolders[0].name, "Inbox")
    
    def test_get_folder_tree_fallback(self):
        """Test folder tree extraction fallback without pypff"""
        parser = PSTParser(self.test_pst_file)
        
        with parser:
            # Mock ImportError for pypff
            with patch('src.macpst.core.pst_parser.pypff', side_effect=ImportError):
                folder_tree = parser.get_folder_tree()
                
                self.assertIsInstance(folder_tree, PSTFolder)
                self.assertEqual(folder_tree.name, "Root")
                self.assertEqual(folder_tree.path, "/")
                self.assertTrue(len(folder_tree.subfolders) > 0)
                
                # Check for standard folders
                folder_names = [f.name for f in folder_tree.subfolders]
                self.assertIn("Inbox", folder_names)
                self.assertIn("Sent Items", folder_names)
    
    @patch('src.macpst.core.pst_parser.pypff')
    def test_extract_messages_with_pypff(self, mock_pypff):
        """Test message extraction with pypff available"""
        # Mock pypff objects
        mock_file = Mock()
        mock_folder = Mock()
        mock_message = Mock()
        
        mock_message.subject = "Test Subject"
        mock_message.sender_name = "Test Sender"
        mock_message.sender_email_address = "sender@example.com"
        mock_message.plain_text_body = "Test body"
        mock_message.html_body = "<html>Test HTML body</html>"
        mock_message.number_of_recipients = 1
        mock_message.number_of_attachments = 0
        
        mock_recipient = Mock()
        mock_recipient.type = 1  # TO recipient
        mock_recipient.name = "Recipient"
        mock_recipient.email_address = "recipient@example.com"
        mock_message.get_recipient.return_value = mock_recipient
        
        mock_folder.number_of_messages = 1
        mock_folder.get_message.return_value = mock_message
        mock_folder.name = "Inbox"
        mock_folder.number_of_sub_folders = 0
        
        mock_file.get_root_folder.return_value = mock_folder
        mock_pypff.file.return_value = mock_file
        
        parser = PSTParser(self.test_pst_file)
        
        with parser:
            messages = list(parser.extract_messages())
            
            self.assertEqual(len(messages), 1)
            message = messages[0]
            
            self.assertIsInstance(message, EmailMessage)
            self.assertEqual(message.subject, "Test Subject")
            self.assertEqual(message.sender, "Test Sender <sender@example.com>")
            self.assertEqual(message.body_text, "Test body")
            self.assertEqual(message.body_html, "<html>Test HTML body</html>")
    
    def test_extract_messages_fallback(self):
        """Test message extraction fallback without pypff"""
        parser = PSTParser(self.test_pst_file)
        
        with parser:
            with patch('src.macpst.core.pst_parser.pypff', side_effect=ImportError):
                messages = list(parser.extract_messages())
                
                # Should return at least one fallback message
                self.assertTrue(len(messages) >= 1)
                message = messages[0]
                
                self.assertIsInstance(message, EmailMessage)
                self.assertIn("Sample Email", message.subject)
                self.assertIn("Limited Parser", message.subject)


if __name__ == '__main__':
    unittest.main()