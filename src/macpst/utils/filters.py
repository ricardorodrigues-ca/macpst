"""
Filtering utilities for PST conversion
"""

import logging
from datetime import datetime, date
from typing import List, Optional, Callable, Dict, Any
from ..core.pst_parser import EmailMessage


logger = logging.getLogger(__name__)


class MessageFilter:
    """Filter for email messages based on various criteria"""
    
    def __init__(self):
        self.date_from: Optional[datetime] = None
        self.date_to: Optional[datetime] = None
        self.sender_filters: List[str] = []
        self.subject_filters: List[str] = []
        self.folder_filters: List[str] = []
        self.has_attachments: Optional[bool] = None
        self.attachment_types: List[str] = []
        self.size_min: Optional[int] = None
        self.size_max: Optional[int] = None
        self.exclude_folders: List[str] = []
    
    def set_date_range(self, date_from: Optional[datetime] = None, 
                      date_to: Optional[datetime] = None):
        """Set date range filter"""
        self.date_from = date_from
        self.date_to = date_to
        logger.debug(f"Date filter set: {date_from} to {date_to}")
    
    def add_sender_filter(self, sender_pattern: str):
        """Add sender filter pattern"""
        self.sender_filters.append(sender_pattern.lower())
        logger.debug(f"Added sender filter: {sender_pattern}")
    
    def add_subject_filter(self, subject_pattern: str):
        """Add subject filter pattern"""
        self.subject_filters.append(subject_pattern.lower())
        logger.debug(f"Added subject filter: {subject_pattern}")
    
    def add_folder_filter(self, folder_pattern: str):
        """Add folder filter pattern"""
        self.folder_filters.append(folder_pattern.lower())
        logger.debug(f"Added folder filter: {folder_pattern}")
    
    def exclude_folder(self, folder_pattern: str):
        """Add folder to exclude from filtering"""
        self.exclude_folders.append(folder_pattern.lower())
        logger.debug(f"Added folder exclusion: {folder_pattern}")
    
    def set_attachment_filter(self, has_attachments: Optional[bool] = None,
                            attachment_types: List[str] = None):
        """Set attachment-related filters"""
        self.has_attachments = has_attachments
        if attachment_types:
            self.attachment_types = [ext.lower().strip('.') for ext in attachment_types]
        logger.debug(f"Attachment filter set: has_attachments={has_attachments}, types={attachment_types}")
    
    def set_size_filter(self, size_min: Optional[int] = None, 
                       size_max: Optional[int] = None):
        """Set message size filter (in bytes)"""
        self.size_min = size_min
        self.size_max = size_max
        logger.debug(f"Size filter set: {size_min} to {size_max} bytes")
    
    def matches(self, message: EmailMessage) -> bool:
        """Check if message matches all filter criteria"""
        
        if not self._check_date_range(message):
            return False
        
        if not self._check_sender_filters(message):
            return False
        
        if not self._check_subject_filters(message):
            return False
        
        if not self._check_folder_filters(message):
            return False
        
        if not self._check_folder_exclusions(message):
            return False
        
        if not self._check_attachment_filters(message):
            return False
        
        return True
    
    def _check_date_range(self, message: EmailMessage) -> bool:
        """Check if message falls within date range"""
        if not self.date_from and not self.date_to:
            return True
        
        message_date = message.sent_time or message.received_time
        if not message_date:
            return True  # Include messages without dates
        
        if self.date_from and message_date < self.date_from:
            return False
        
        if self.date_to and message_date > self.date_to:
            return False
        
        return True
    
    def _check_sender_filters(self, message: EmailMessage) -> bool:
        """Check sender filters"""
        if not self.sender_filters:
            return True
        
        sender_lower = message.sender.lower() if message.sender else ""
        return any(pattern in sender_lower for pattern in self.sender_filters)
    
    def _check_subject_filters(self, message: EmailMessage) -> bool:
        """Check subject filters"""
        if not self.subject_filters:
            return True
        
        subject_lower = message.subject.lower() if message.subject else ""
        return any(pattern in subject_lower for pattern in self.subject_filters)
    
    def _check_folder_filters(self, message: EmailMessage) -> bool:
        """Check folder filters"""
        if not self.folder_filters:
            return True
        
        folder_lower = message.folder_path.lower() if message.folder_path else ""
        return any(pattern in folder_lower for pattern in self.folder_filters)
    
    def _check_folder_exclusions(self, message: EmailMessage) -> bool:
        """Check folder exclusions"""
        if not self.exclude_folders:
            return True
        
        folder_lower = message.folder_path.lower() if message.folder_path else ""
        return not any(pattern in folder_lower for pattern in self.exclude_folders)
    
    def _check_attachment_filters(self, message: EmailMessage) -> bool:
        """Check attachment filters"""
        if self.has_attachments is None and not self.attachment_types:
            return True
        
        has_attachments = bool(message.attachments)
        
        if self.has_attachments is not None:
            if self.has_attachments != has_attachments:
                return False
        
        if self.attachment_types and has_attachments:
            message_extensions = set()
            for attachment in message.attachments:
                name = attachment.get('name', '')
                if '.' in name:
                    ext = name.split('.')[-1].lower()
                    message_extensions.add(ext)
            
            if not any(ext in message_extensions for ext in self.attachment_types):
                return False
        
        return True
    
    def filter_messages(self, messages: List[EmailMessage]) -> List[EmailMessage]:
        """Filter a list of messages"""
        filtered = [msg for msg in messages if self.matches(msg)]
        logger.info(f"Filtered {len(messages)} messages down to {len(filtered)}")
        return filtered
    
    def get_filter_summary(self) -> Dict[str, Any]:
        """Get a summary of current filter settings"""
        return {
            'date_range': {
                'from': self.date_from.isoformat() if self.date_from else None,
                'to': self.date_to.isoformat() if self.date_to else None
            },
            'sender_filters': self.sender_filters,
            'subject_filters': self.subject_filters,
            'folder_filters': self.folder_filters,
            'exclude_folders': self.exclude_folders,
            'attachment_settings': {
                'has_attachments': self.has_attachments,
                'attachment_types': self.attachment_types
            },
            'size_range': {
                'min': self.size_min,
                'max': self.size_max
            }
        }


class DuplicateDetector:
    """Detect and handle duplicate email messages"""
    
    def __init__(self):
        self.check_subject = True
        self.check_sender = True
        self.check_recipients = True
        self.check_body = False
        self.check_date = False
        self.date_tolerance_minutes = 5
    
    def configure(self, check_subject: bool = True, check_sender: bool = True,
                 check_recipients: bool = True, check_body: bool = False,
                 check_date: bool = False, date_tolerance_minutes: int = 5):
        """Configure duplicate detection criteria"""
        self.check_subject = check_subject
        self.check_sender = check_sender
        self.check_recipients = check_recipients
        self.check_body = check_body
        self.check_date = check_date
        self.date_tolerance_minutes = date_tolerance_minutes
        
        logger.debug(f"Duplicate detection configured: subject={check_subject}, "
                    f"sender={check_sender}, recipients={check_recipients}, "
                    f"body={check_body}, date={check_date}")
    
    def get_message_signature(self, message: EmailMessage) -> str:
        """Get a signature string for message comparison"""
        signature_parts = []
        
        if self.check_subject and message.subject:
            signature_parts.append(f"subj:{message.subject.lower().strip()}")
        
        if self.check_sender and message.sender:
            signature_parts.append(f"from:{message.sender.lower().strip()}")
        
        if self.check_recipients and message.recipients:
            recipients = sorted([r.lower().strip() for r in message.recipients])
            signature_parts.append(f"to:{','.join(recipients)}")
        
        if self.check_body:
            body = message.body_text or message.body_html or ""
            body_preview = body[:100].lower().strip()
            signature_parts.append(f"body:{body_preview}")
        
        if self.check_date and (message.sent_time or message.received_time):
            date_obj = message.sent_time or message.received_time
            # Round to nearest time window for tolerance
            minutes = date_obj.minute
            rounded_minutes = (minutes // self.date_tolerance_minutes) * self.date_tolerance_minutes
            date_rounded = date_obj.replace(minute=rounded_minutes, second=0, microsecond=0)
            signature_parts.append(f"date:{date_rounded.isoformat()}")
        
        return "|".join(signature_parts)
    
    def find_duplicates(self, messages: List[EmailMessage]) -> Dict[str, List[EmailMessage]]:
        """Find duplicate messages grouped by signature"""
        signature_groups = {}
        
        for message in messages:
            signature = self.get_message_signature(message)
            if signature not in signature_groups:
                signature_groups[signature] = []
            signature_groups[signature].append(message)
        
        # Filter to only groups with duplicates
        duplicates = {sig: msgs for sig, msgs in signature_groups.items() if len(msgs) > 1}
        
        total_duplicates = sum(len(msgs) - 1 for msgs in duplicates.values())
        logger.info(f"Found {len(duplicates)} duplicate groups with {total_duplicates} duplicate messages")
        
        return duplicates
    
    def remove_duplicates(self, messages: List[EmailMessage], 
                         keep_strategy: str = 'first') -> List[EmailMessage]:
        """Remove duplicate messages, keeping one from each group"""
        duplicate_groups = self.find_duplicates(messages)
        
        if not duplicate_groups:
            return messages
        
        kept_messages = []
        processed_signatures = set()
        
        for message in messages:
            signature = self.get_message_signature(message)
            
            if signature in duplicate_groups and signature not in processed_signatures:
                # This is the first occurrence of a duplicate group
                group = duplicate_groups[signature]
                kept_message = self._select_message_to_keep(group, keep_strategy)
                kept_messages.append(kept_message)
                processed_signatures.add(signature)
                
            elif signature not in duplicate_groups:
                # This is a unique message
                kept_messages.append(message)
        
        removed_count = len(messages) - len(kept_messages)
        logger.info(f"Removed {removed_count} duplicate messages, kept {len(kept_messages)}")
        
        return kept_messages
    
    def _select_message_to_keep(self, duplicate_group: List[EmailMessage], 
                              keep_strategy: str) -> EmailMessage:
        """Select which message to keep from a duplicate group"""
        if keep_strategy == 'first':
            return duplicate_group[0]
        elif keep_strategy == 'last':
            return duplicate_group[-1]
        elif keep_strategy == 'newest':
            # Keep the one with the most recent date
            with_dates = [msg for msg in duplicate_group 
                         if msg.sent_time or msg.received_time]
            if with_dates:
                return max(with_dates, 
                          key=lambda m: m.sent_time or m.received_time)
            return duplicate_group[0]
        elif keep_strategy == 'oldest':
            # Keep the one with the oldest date
            with_dates = [msg for msg in duplicate_group 
                         if msg.sent_time or msg.received_time]
            if with_dates:
                return min(with_dates, 
                          key=lambda m: m.sent_time or m.received_time)
            return duplicate_group[0]
        else:
            return duplicate_group[0]