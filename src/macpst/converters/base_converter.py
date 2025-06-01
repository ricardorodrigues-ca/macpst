"""
Base converter class for all output formats
"""

import os
import logging
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional
from pathlib import Path
from ..core.pst_parser import EmailMessage


logger = logging.getLogger(__name__)


class BaseConverter(ABC):
    """Abstract base class for all format converters"""
    
    def __init__(self, output_dir: str):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.converted_count = 0
        self.error_count = 0
        self.errors = []
    
    @abstractmethod
    def convert_message(self, message: EmailMessage, output_filename: str) -> bool:
        """Convert a single email message to the target format"""
        pass
    
    @abstractmethod
    def get_file_extension(self) -> str:
        """Get the file extension for this format"""
        pass
    
    def convert_messages(self, messages: List[EmailMessage], 
                        progress_callback: Optional[callable] = None) -> Dict[str, Any]:
        """Convert multiple messages with progress tracking"""
        total_messages = len(messages)
        self.converted_count = 0
        self.error_count = 0
        self.errors = []
        
        for i, message in enumerate(messages):
            try:
                filename = self._generate_filename(message, i)
                success = self.convert_message(message, filename)
                
                if success:
                    self.converted_count += 1
                else:
                    self.error_count += 1
                    self.errors.append(f"Failed to convert message: {message.subject}")
                
            except Exception as e:
                self.error_count += 1
                self.errors.append(f"Error converting '{message.subject}': {str(e)}")
                logger.error(f"Conversion error: {e}")
            
            if progress_callback:
                progress = (i + 1) / total_messages * 100
                progress_callback(progress, i + 1, total_messages)
        
        return {
            'total_messages': total_messages,
            'converted_count': self.converted_count,
            'error_count': self.error_count,
            'errors': self.errors,
            'output_directory': str(self.output_dir)
        }
    
    def _generate_filename(self, message: EmailMessage, index: int) -> str:
        """Generate a safe filename for the message"""
        subject = message.subject or f"message_{index}"
        
        safe_subject = "".join(c for c in subject if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_subject = safe_subject[:50]  # Limit length
        
        if not safe_subject:
            safe_subject = f"message_{index}"
        
        timestamp = ""
        if message.sent_time:
            timestamp = message.sent_time.strftime("_%Y%m%d_%H%M%S")
        elif message.received_time:
            timestamp = message.received_time.strftime("_%Y%m%d_%H%M%S")
        
        base_filename = f"{safe_subject}{timestamp}"
        extension = self.get_file_extension()
        
        counter = 1
        filename = f"{base_filename}.{extension}"
        
        while (self.output_dir / filename).exists():
            filename = f"{base_filename}_{counter}.{extension}"
            counter += 1
        
        return filename
    
    def _sanitize_text(self, text: str) -> str:
        """Sanitize text for safe output"""
        if not text:
            return ""
        
        try:
            return text.encode('utf-8', errors='replace').decode('utf-8')
        except Exception:
            return str(text)
    
    def get_output_path(self, filename: str) -> Path:
        """Get full output path for a filename"""
        return self.output_dir / filename