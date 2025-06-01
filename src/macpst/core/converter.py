"""
Main converter class that coordinates all format conversions
"""

import logging
from typing import List, Dict, Any, Optional, Type
from pathlib import Path
from ..converters.base_converter import BaseConverter
from ..converters.eml_converter import EMLConverter
from ..converters.mbox_converter import MBOXConverter
from ..converters.pdf_converter import PDFConverter
from .pst_parser import EmailMessage


logger = logging.getLogger(__name__)


class Converter:
    """Main converter class for PST file conversion"""
    
    SUPPORTED_FORMATS = {
        'eml': EMLConverter,
        'mbox': MBOXConverter,
        'pdf': PDFConverter,
    }
    
    def __init__(self, output_directory: str):
        self.output_directory = Path(output_directory)
        self.output_directory.mkdir(parents=True, exist_ok=True)
        
    def get_supported_formats(self) -> List[str]:
        """Get list of supported output formats"""
        return list(self.SUPPORTED_FORMATS.keys())
    
    def convert_messages(self, messages: List[EmailMessage], 
                        output_format: str,
                        progress_callback: Optional[callable] = None,
                        **kwargs) -> Dict[str, Any]:
        """Convert messages to specified format"""
        
        if output_format.lower() not in self.SUPPORTED_FORMATS:
            raise ValueError(f"Unsupported format: {output_format}. "
                           f"Supported formats: {', '.join(self.get_supported_formats())}")
        
        converter_class = self.SUPPORTED_FORMATS[output_format.lower()]
        
        format_output_dir = self.output_directory / output_format.lower()
        converter = converter_class(str(format_output_dir))
        
        logger.info(f"Converting {len(messages)} messages to {output_format.upper()} format")
        
        if output_format.lower() == 'mbox' and kwargs.get('single_file', True):
            result = converter.convert_messages(messages, progress_callback, single_file=True)
        else:
            result = converter.convert_messages(messages, progress_callback)
        
        logger.info(f"Conversion completed: {result['converted_count']} successful, "
                   f"{result['error_count']} errors")
        
        return result
    
    def batch_convert(self, messages: List[EmailMessage], 
                     formats: List[str],
                     progress_callback: Optional[callable] = None) -> Dict[str, Dict[str, Any]]:
        """Convert messages to multiple formats"""
        
        results = {}
        total_operations = len(formats)
        
        for i, format_name in enumerate(formats):
            try:
                logger.info(f"Starting conversion to {format_name} ({i+1}/{total_operations})")
                
                def format_progress(progress, current, total):
                    overall_progress = (i * 100 + progress) / total_operations
                    if progress_callback:
                        progress_callback(overall_progress, format_name, current, total)
                
                result = self.convert_messages(messages, format_name, format_progress)
                results[format_name] = result
                
            except Exception as e:
                logger.error(f"Error converting to {format_name}: {e}")
                results[format_name] = {
                    'total_messages': len(messages),
                    'converted_count': 0,
                    'error_count': len(messages),
                    'errors': [f"Conversion failed: {str(e)}"],
                    'output_directory': str(self.output_directory / format_name)
                }
        
        return results