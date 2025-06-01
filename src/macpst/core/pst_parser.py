"""
PST File Parser - Core module for parsing Outlook PST files
"""

import os
import struct
import logging
from typing import Dict, List, Optional, Generator, Tuple
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path

logger = logging.getLogger(__name__)


@dataclass
class EmailMessage:
    """Represents an email message extracted from PST"""
    subject: str
    sender: str
    recipients: List[str]
    cc_recipients: List[str] = None
    bcc_recipients: List[str] = None
    body_text: str = ""
    body_html: str = ""
    sent_time: Optional[datetime] = None
    received_time: Optional[datetime] = None
    attachments: List[Dict] = None
    message_id: str = ""
    folder_path: str = ""
    
    def __post_init__(self):
        if self.cc_recipients is None:
            self.cc_recipients = []
        if self.bcc_recipients is None:
            self.bcc_recipients = []
        if self.attachments is None:
            self.attachments = []


@dataclass
class PSTFolder:
    """Represents a folder in the PST file"""
    name: str
    path: str
    message_count: int
    subfolders: List['PSTFolder'] = None
    
    def __post_init__(self):
        if self.subfolders is None:
            self.subfolders = []


class PSTParser:
    """Parser for Microsoft Outlook PST files"""
    
    PST_SIGNATURE = b'!BDN'
    UNICODE_PST = 0x17
    ANSI_PST = 0x0E
    
    def __init__(self, pst_file_path: str):
        self.pst_file_path = Path(pst_file_path)
        self.file_handle = None
        self.is_unicode = False
        self.header = None
        self.root_folder = None
        self._validate_file()
    
    def _validate_file(self):
        """Validate PST file format and signature"""
        if not self.pst_file_path.exists():
            raise FileNotFoundError(f"PST file not found: {self.pst_file_path}")
        
        if not self.pst_file_path.is_file():
            raise ValueError(f"Path is not a file: {self.pst_file_path}")
        
        with open(self.pst_file_path, 'rb') as f:
            signature = f.read(4)
            if signature != self.PST_SIGNATURE:
                raise ValueError("Invalid PST file signature")
    
    def open(self):
        """Open PST file for reading"""
        try:
            self.file_handle = open(self.pst_file_path, 'rb')
            self._read_header()
            logger.info(f"Successfully opened PST file: {self.pst_file_path}")
        except Exception as e:
            if self.file_handle:
                self.file_handle.close()
                self.file_handle = None
            raise e
    
    def close(self):
        """Close PST file"""
        if self.file_handle:
            self.file_handle.close()
            self.file_handle = None
            logger.info("PST file closed")
    
    def __enter__(self):
        self.open()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
    
    def _read_header(self):
        """Read and parse PST file header"""
        if not self.file_handle:
            raise RuntimeError("PST file not opened")
        
        self.file_handle.seek(0)
        header_data = self.file_handle.read(512)
        
        signature = header_data[:4]
        if signature != self.PST_SIGNATURE:
            raise ValueError("Invalid PST signature in header")
        
        file_type = struct.unpack('<H', header_data[10:12])[0]
        
        if file_type == self.UNICODE_PST:
            self.is_unicode = True
            logger.info("Detected Unicode PST file")
        elif file_type == self.ANSI_PST:
            self.is_unicode = False
            logger.info("Detected ANSI PST file")
        else:
            raise ValueError(f"Unsupported PST file type: {file_type}")
        
        self.header = {
            'signature': signature,
            'file_type': file_type,
            'is_unicode': self.is_unicode,
            'file_size': os.path.getsize(self.pst_file_path)
        }
    
    def get_folder_tree(self) -> PSTFolder:
        """Get the folder structure of the PST file"""
        if not self.file_handle:
            raise RuntimeError("PST file not opened")
        
        try:
            import pypff
            
            pff_file = pypff.file()
            pff_file.open(str(self.pst_file_path))
            
            root = pff_file.get_root_folder()
            self.root_folder = self._build_folder_tree(root)
            
            pff_file.close()
            return self.root_folder
            
        except ImportError:
            logger.warning("libpff-python not available, using basic folder detection")
            return self._get_basic_folder_structure()
        except Exception as e:
            logger.error(f"Error reading folder structure: {e}")
            return self._get_basic_folder_structure()
    
    def _build_folder_tree(self, pff_folder) -> PSTFolder:
        """Build folder tree from pypff folder object"""
        # Get message count safely - try multiple methods
        message_count = 0
        try:
            # Try sub_messages first (what this API uses)
            if hasattr(pff_folder, 'number_of_sub_messages'):
                message_count = pff_folder.number_of_sub_messages
            elif hasattr(pff_folder, 'get_number_of_sub_messages'):
                message_count = pff_folder.get_number_of_sub_messages()
            # Fallback to regular messages
            elif hasattr(pff_folder, 'number_of_messages'):
                message_count = pff_folder.number_of_messages
            elif hasattr(pff_folder, 'get_number_of_messages'):
                message_count = pff_folder.get_number_of_messages()
            # Try sub_items as they might contain messages
            elif hasattr(pff_folder, 'number_of_sub_items'):
                sub_item_count = pff_folder.number_of_sub_items
                # Count how many sub_items are actually messages
                message_count = 0
                for i in range(sub_item_count):
                    try:
                        item = pff_folder.get_sub_item(i)
                        if item and self._is_email_item(item):
                            message_count += 1
                    except:
                        continue
        except:
            message_count = 0
        
        folder = PSTFolder(
            name=pff_folder.name or "Root",
            path=self._get_folder_path(pff_folder),
            message_count=message_count
        )
        
        # Get subfolder count safely
        num_subfolders = 0
        try:
            if hasattr(pff_folder, 'number_of_sub_folders'):
                num_subfolders = pff_folder.number_of_sub_folders
            elif hasattr(pff_folder, 'get_number_of_sub_folders'):
                num_subfolders = pff_folder.get_number_of_sub_folders()
        except:
            num_subfolders = 0
        
        for i in range(num_subfolders):
            try:
                subfolder = pff_folder.get_sub_folder(i)
                if subfolder:
                    folder.subfolders.append(self._build_folder_tree(subfolder))
            except Exception as e:
                logger.warning(f"Error accessing subfolder {i}: {e}")
                continue
        
        return folder
    
    def _is_email_item(self, item):
        """Check if an item is an email message"""
        try:
            # Check for email-like attributes
            email_attributes = ['subject', 'sender_name', 'plain_text_body', 'html_body', 'recipients']
            return any(hasattr(item, attr) for attr in email_attributes)
        except:
            return False
    
    def _get_folder_path(self, pff_folder) -> str:
        """Get full path of folder"""
        path_parts = []
        current = pff_folder
        
        while current and current.name:
            path_parts.append(current.name)
            parent = getattr(current, 'parent', None)
            if parent:
                current = parent
            else:
                break
        
        path_parts.reverse()
        return "/" + "/".join(path_parts) if path_parts else "/"
    
    def _get_basic_folder_structure(self) -> PSTFolder:
        """Fallback method for basic folder structure"""
        return PSTFolder(
            name="Root",
            path="/",
            message_count=0,
            subfolders=[
                PSTFolder(name="Inbox", path="/Inbox", message_count=0),
                PSTFolder(name="Sent Items", path="/Sent Items", message_count=0),
                PSTFolder(name="Deleted Items", path="/Deleted Items", message_count=0),
                PSTFolder(name="Drafts", path="/Drafts", message_count=0)
            ]
        )
    
    def extract_messages(self, folder_path: str = None) -> Generator[EmailMessage, None, None]:
        """Extract email messages from PST file"""
        if not self.file_handle:
            raise RuntimeError("PST file not opened")
        
        try:
            import pypff
            
            logger.info(f"Opening PST file with libpff-python: {self.pst_file_path}")
            pff_file = pypff.file()
            pff_file.open(str(self.pst_file_path))
            
            root_folder = pff_file.get_root_folder()
            logger.info(f"Got root folder: {root_folder.name if root_folder else 'None'}")
            
            if folder_path:
                folder = self._find_folder_by_path(root_folder, folder_path)
                if folder:
                    logger.info(f"Extracting from specific folder: {folder_path}")
                    yield from self._extract_messages_from_folder(folder, folder_path)
                else:
                    logger.warning(f"Folder not found: {folder_path}")
            else:
                logger.info("Extracting from all folders")
                message_count = 0
                for message in self._extract_all_messages(root_folder):
                    message_count += 1
                    logger.debug(f"Yielding message {message_count}: {message.subject}")
                    yield message
                logger.info(f"Total messages extracted: {message_count}")
            
            pff_file.close()
            
        except ImportError:
            logger.warning("libpff-python not available, using basic PST parsing")
            yield from self._extract_messages_basic()
        except Exception as e:
            logger.error(f"Error during PST message extraction: {e}")
            logger.info("Falling back to basic PST parsing")
            yield from self._extract_messages_basic()
    
    def _find_folder_by_path(self, root_folder, target_path):
        """Find folder by path string"""
        if target_path == "/" or target_path == root_folder.name:
            return root_folder
        
        parts = target_path.strip("/").split("/")
        current = root_folder
        
        for part in parts:
            found = False
            for i in range(current.number_of_sub_folders):
                subfolder = current.get_sub_folder(i)
                if subfolder.name == part:
                    current = subfolder
                    found = True
                    break
            if not found:
                return None
        
        return current
    
    def _extract_all_messages(self, folder):
        """Recursively extract all messages from all folders"""
        folder_path = self._get_folder_path(folder)
        folder_name = getattr(folder, 'name', 'Unknown') or 'Root'
        
        logger.info(f"Processing folder: {folder_name} at path: {folder_path}")
        
        # Extract messages from current folder
        message_count = 0
        for message in self._extract_messages_from_folder(folder, folder_path):
            message_count += 1
            yield message
        
        logger.info(f"Found {message_count} messages in folder: {folder_name}")
        
        # Get subfolder count safely
        num_subfolders = 0
        try:
            if hasattr(folder, 'number_of_sub_folders'):
                num_subfolders = folder.number_of_sub_folders
            elif hasattr(folder, 'get_number_of_sub_folders'):
                num_subfolders = folder.get_number_of_sub_folders()
        except:
            num_subfolders = 0
        
        logger.info(f"Folder {folder_name} has {num_subfolders} subfolders")
        
        for i in range(num_subfolders):
            try:
                subfolder = folder.get_sub_folder(i)
                if subfolder:
                    subfolder_name = getattr(subfolder, 'name', f'Subfolder_{i}')
                    logger.info(f"Processing subfolder {i}: {subfolder_name}")
                    yield from self._extract_all_messages(subfolder)
                else:
                    logger.warning(f"Subfolder {i} is None")
            except Exception as e:
                logger.warning(f"Error accessing subfolder {i} in {folder_path}: {e}")
                continue
    
    def _extract_messages_from_folder(self, folder, folder_path):
        """Extract messages from a specific folder"""
        # Try multiple methods to get messages
        messages_found = 0
        
        # Method 1: Try sub_messages (API method)
        try:
            if hasattr(folder, 'number_of_sub_messages'):
                num_sub_messages = folder.number_of_sub_messages
                logger.info(f"Found {num_sub_messages} sub_messages in {folder_path}")
                
                for i in range(num_sub_messages):
                    try:
                        message = folder.get_sub_message(i)
                        if message:
                            subject = getattr(message, 'subject', f'SubMessage {i}') or f'SubMessage {i}'
                            logger.debug(f"Processing sub_message {i}: {subject}")
                            email_msg = self._convert_to_email_message(message, folder_path)
                            if email_msg:
                                messages_found += 1
                                yield email_msg
                    except Exception as e:
                        logger.warning(f"Error extracting sub_message {i} from {folder_path}: {e}")
                        continue
        except Exception as e:
            logger.debug(f"Error accessing sub_messages: {e}")
        
        # Method 2: Try sub_items (might contain messages)
        try:
            if hasattr(folder, 'number_of_sub_items'):
                num_sub_items = folder.number_of_sub_items
                logger.info(f"Found {num_sub_items} sub_items in {folder_path}")
                
                for i in range(num_sub_items):
                    try:
                        item = folder.get_sub_item(i)
                        if item and self._is_email_item(item):
                            subject = getattr(item, 'subject', f'Item {i}') or f'Item {i}'
                            logger.debug(f"Processing email item {i}: {subject}")
                            email_msg = self._convert_to_email_message(item, folder_path)
                            if email_msg:
                                messages_found += 1
                                yield email_msg
                    except Exception as e:
                        logger.warning(f"Error extracting sub_item {i} from {folder_path}: {e}")
                        continue
        except Exception as e:
            logger.debug(f"Error accessing sub_items: {e}")
        
        # Method 3: Fallback to regular messages (old API)
        if messages_found == 0:
            try:
                num_messages = 0
                if hasattr(folder, 'number_of_messages'):
                    num_messages = folder.number_of_messages
                elif hasattr(folder, 'get_number_of_messages'):
                    num_messages = folder.get_number_of_messages()
                
                logger.info(f"Trying regular messages: {num_messages} in {folder_path}")
                
                for i in range(num_messages):
                    try:
                        message = folder.get_message(i)
                        if message:
                            subject = getattr(message, 'subject', f'Message {i}') or f'Message {i}'
                            logger.debug(f"Processing message {i}: {subject}")
                            email_msg = self._convert_to_email_message(message, folder_path)
                            if email_msg:
                                messages_found += 1
                                yield email_msg
                    except Exception as e:
                        logger.warning(f"Error extracting message {i} from {folder_path}: {e}")
                        continue
            except Exception as e:
                logger.debug(f"Error accessing regular messages: {e}")
        
        logger.info(f"Total messages extracted from {folder_path}: {messages_found}")
    
    def _convert_to_email_message(self, pff_message, folder_path) -> EmailMessage:
        """Convert pypff message to EmailMessage object"""
        try:
            subject = getattr(pff_message, 'subject', '') or ''
            sender = self._get_sender_info(pff_message)
            recipients = self._get_recipients(pff_message)
            
            # Try different body text attributes
            body_text = ''
            for attr in ['plain_text_body', 'body', 'text_body']:
                if hasattr(pff_message, attr):
                    body_text = getattr(pff_message, attr, '') or ''
                    if body_text:
                        break
            
            # Try different HTML body attributes
            body_html = ''
            for attr in ['html_body', 'rtf_body']:
                if hasattr(pff_message, attr):
                    body_html = getattr(pff_message, attr, '') or ''
                    if body_html:
                        break
            
            # Get message ID
            message_id = ''
            for attr in ['message_identifier', 'message_id', 'internet_message_id']:
                if hasattr(pff_message, attr):
                    message_id = getattr(pff_message, attr, '') or ''
                    if message_id:
                        break
            
            email_msg = EmailMessage(
                subject=subject,
                sender=sender,
                recipients=recipients,
                cc_recipients=self._get_cc_recipients(pff_message),
                bcc_recipients=self._get_bcc_recipients(pff_message),
                body_text=body_text,
                body_html=body_html,
                sent_time=self._get_sent_time(pff_message),
                received_time=self._get_received_time(pff_message),
                attachments=self._get_attachments(pff_message),
                message_id=message_id,
                folder_path=folder_path
            )
            
            logger.debug(f"Converted message: subject='{subject}', sender='{sender}', "
                        f"body_length={len(body_text)}, html_length={len(body_html)}")
            
            return email_msg
            
        except Exception as e:
            logger.error(f"Error converting message to EmailMessage: {e}")
            return None
    
    def _get_sender_info(self, message):
        """Extract sender information"""
        sender = getattr(message, 'sender_name', '')
        email = getattr(message, 'sender_email_address', '')
        
        if sender and email:
            return f"{sender} <{email}>"
        return sender or email or ''
    
    def _get_recipients(self, message):
        """Extract recipient list"""
        recipients = []
        try:
            # Get recipient count safely
            num_recipients = 0
            if hasattr(message, 'number_of_recipients'):
                num_recipients = message.number_of_recipients
            elif hasattr(message, 'get_number_of_recipients'):
                num_recipients = message.get_number_of_recipients()
            
            for i in range(num_recipients):
                try:
                    recipient = message.get_recipient(i)
                    if recipient and getattr(recipient, 'type', 0) == 1:  # TO recipient
                        name = getattr(recipient, 'name', '')
                        email = getattr(recipient, 'email_address', '')
                        if name and email:
                            recipients.append(f"{name} <{email}>")
                        else:
                            recipients.append(name or email or '')
                except Exception as e:
                    logger.debug(f"Error getting recipient {i}: {e}")
                    continue
        except Exception as e:
            logger.debug(f"Error getting recipients: {e}")
        return recipients
    
    def _get_cc_recipients(self, message):
        """Extract CC recipient list"""
        recipients = []
        try:
            # Get recipient count safely
            num_recipients = 0
            if hasattr(message, 'number_of_recipients'):
                num_recipients = message.number_of_recipients
            elif hasattr(message, 'get_number_of_recipients'):
                num_recipients = message.get_number_of_recipients()
            
            for i in range(num_recipients):
                try:
                    recipient = message.get_recipient(i)
                    if recipient and getattr(recipient, 'type', 0) == 2:  # CC recipient
                        name = getattr(recipient, 'name', '')
                        email = getattr(recipient, 'email_address', '')
                        if name and email:
                            recipients.append(f"{name} <{email}>")
                        else:
                            recipients.append(name or email or '')
                except Exception as e:
                    logger.debug(f"Error getting CC recipient {i}: {e}")
                    continue
        except Exception as e:
            logger.debug(f"Error getting CC recipients: {e}")
        return recipients
    
    def _get_bcc_recipients(self, message):
        """Extract BCC recipient list"""
        recipients = []
        try:
            # Get recipient count safely
            num_recipients = 0
            if hasattr(message, 'number_of_recipients'):
                num_recipients = message.number_of_recipients
            elif hasattr(message, 'get_number_of_recipients'):
                num_recipients = message.get_number_of_recipients()
            
            for i in range(num_recipients):
                try:
                    recipient = message.get_recipient(i)
                    if recipient and getattr(recipient, 'type', 0) == 3:  # BCC recipient
                        name = getattr(recipient, 'name', '')
                        email = getattr(recipient, 'email_address', '')
                        if name and email:
                            recipients.append(f"{name} <{email}>")
                        else:
                            recipients.append(name or email or '')
                except Exception as e:
                    logger.debug(f"Error getting BCC recipient {i}: {e}")
                    continue
        except Exception as e:
            logger.debug(f"Error getting BCC recipients: {e}")
        return recipients
    
    def _get_sent_time(self, message):
        """Extract sent time"""
        try:
            return getattr(message, 'delivery_time', None)
        except:
            return None
    
    def _get_received_time(self, message):
        """Extract received time"""
        try:
            return getattr(message, 'creation_time', None)
        except:
            return None
    
    def _get_attachments(self, message):
        """Extract attachment information"""
        attachments = []
        try:
            # Get attachment count safely
            num_attachments = 0
            if hasattr(message, 'number_of_attachments'):
                num_attachments = message.number_of_attachments
            elif hasattr(message, 'get_number_of_attachments'):
                num_attachments = message.get_number_of_attachments()
            
            for i in range(num_attachments):
                try:
                    attachment = message.get_attachment(i)
                    if attachment:
                        attachments.append({
                            'name': getattr(attachment, 'name', f'attachment_{i}'),
                            'size': getattr(attachment, 'size', 0),
                            'type': getattr(attachment, 'type', 'unknown')
                        })
                except Exception as e:
                    logger.debug(f"Error getting attachment {i}: {e}")
                    continue
        except Exception as e:
            logger.debug(f"Error getting attachments: {e}")
        return attachments
    
    def _extract_messages_basic(self):
        """Basic PST message extraction without pypff"""
        logger.info("Using basic PST parsing - extracting sample data")
        
        # Try to read some basic information from PST file
        messages_found = self._parse_pst_basic()
        
        if messages_found:
            for msg in messages_found:
                yield msg
        else:
            # Fallback to sample message
            yield EmailMessage(
                subject="Sample Email (Basic Parser)",
                sender="noreply@example.com",
                recipients=["user@example.com"],
                body_text="PST file detected but advanced parsing requires libpff-python library.\n"
                         "Install libpff-python for full message extraction:\n"
                         "pip install libpff-python",
                folder_path="/Inbox"
            )
    
    def _parse_pst_basic(self):
        """Basic PST parsing using binary format knowledge"""
        messages = []
        
        try:
            # Read PST file in chunks to look for email-like patterns
            self.file_handle.seek(512)  # Skip header
            
            # This is a simplified approach - look for common email patterns
            chunk_size = 8192
            message_count = 0
            
            while message_count < 10:  # Limit to 10 sample messages
                chunk = self.file_handle.read(chunk_size)
                if not chunk:
                    break
                
                # Look for email-like patterns in the binary data
                if self._contains_email_pattern(chunk):
                    message_count += 1
                    
                    # Extract what we can from the chunk
                    subject = self._extract_text_pattern(chunk, b'Subject:', 50)
                    sender = self._extract_text_pattern(chunk, b'From:', 50)
                    
                    if subject or sender:
                        messages.append(EmailMessage(
                            subject=subject or f"Message {message_count}",
                            sender=sender or "unknown@example.com",
                            recipients=["user@example.com"],
                            body_text=f"Message extracted using basic parser (#{message_count})",
                            folder_path="/Inbox"
                        ))
                
                # Move to next chunk with overlap
                current_pos = self.file_handle.tell()
                if current_pos >= len(chunk):
                    self.file_handle.seek(current_pos - 100)  # Small overlap
            
        except Exception as e:
            logger.debug(f"Basic parsing encountered issue: {e}")
        
        return messages
    
    def _contains_email_pattern(self, data):
        """Check if data chunk contains email-like patterns"""
        email_indicators = [
            b'Subject:',
            b'From:',
            b'To:',
            b'Message-ID:',
            b'@',
            b'.com',
            b'.org',
            b'Content-Type:'
        ]
        
        return any(indicator in data for indicator in email_indicators)
    
    def _extract_text_pattern(self, data, pattern, max_length=100):
        """Extract text following a pattern"""
        try:
            start = data.find(pattern)
            if start == -1:
                return None
            
            start += len(pattern)
            end = start + max_length
            
            # Find the end of the text (null byte, newline, or non-printable)
            text_data = data[start:end]
            
            # Convert to string and clean up
            try:
                text = text_data.decode('utf-8', errors='ignore')
            except:
                text = text_data.decode('latin-1', errors='ignore')
            
            # Clean up the text
            text = ''.join(c for c in text if c.isprintable() or c.isspace())
            text = text.split('\n')[0].split('\r')[0].strip()
            
            return text if text and len(text) > 2 else None
            
        except Exception:
            return None
    
    def get_statistics(self) -> Dict:
        """Get PST file statistics"""
        stats = {
            'file_path': str(self.pst_file_path),
            'file_size': self.header['file_size'] if self.header else 0,
            'is_unicode': self.is_unicode,
            'total_messages': 0,
            'total_folders': 0
        }
        
        if self.root_folder:
            stats['total_messages'] = self._count_messages(self.root_folder)
            stats['total_folders'] = self._count_folders(self.root_folder)
        
        return stats
    
    def _count_messages(self, folder: PSTFolder) -> int:
        """Count total messages in folder tree"""
        count = folder.message_count
        for subfolder in folder.subfolders:
            count += self._count_messages(subfolder)
        return count
    
    def _count_folders(self, folder: PSTFolder) -> int:
        """Count total folders in folder tree"""
        count = 1
        for subfolder in folder.subfolders:
            count += self._count_folders(subfolder)
        return count