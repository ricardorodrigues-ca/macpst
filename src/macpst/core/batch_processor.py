"""
Batch processing functionality for multiple PST files
"""

import os
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional, Callable
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
from .pst_parser import PSTParser, EmailMessage
from .converter import Converter
from ..utils.filters import MessageFilter, DuplicateDetector


logger = logging.getLogger(__name__)


class BatchProcessor:
    """Batch processor for multiple PST files"""
    
    def __init__(self, output_directory: str, max_workers: int = 4):
        self.output_directory = Path(output_directory)
        self.output_directory.mkdir(parents=True, exist_ok=True)
        self.max_workers = max_workers
        self.converter = Converter(str(self.output_directory))
        self.message_filter = MessageFilter()
        self.duplicate_detector = DuplicateDetector()
        self.remove_duplicates = False
        self.stats = {
            'files_processed': 0,
            'files_failed': 0,
            'total_messages': 0,
            'filtered_messages': 0,
            'duplicates_removed': 0,
            'conversion_results': {}
        }
        self._lock = threading.Lock()
    
    def configure_filter(self, message_filter: MessageFilter):
        """Configure message filtering"""
        self.message_filter = message_filter
        logger.info("Message filter configured")
    
    def configure_duplicate_detection(self, duplicate_detector: DuplicateDetector, 
                                    remove_duplicates: bool = False):
        """Configure duplicate detection and removal"""
        self.duplicate_detector = duplicate_detector
        self.remove_duplicates = remove_duplicates
        logger.info(f"Duplicate detection configured (remove={remove_duplicates})")
    
    def process_files(self, pst_files: List[str], output_formats: List[str],
                     progress_callback: Optional[Callable] = None) -> Dict[str, Any]:
        """Process multiple PST files with batch operations"""
        
        logger.info(f"Starting batch processing of {len(pst_files)} PST files")
        self._reset_stats()
        
        total_steps = len(pst_files) + 1  # +1 for final conversion step
        current_step = 0
        
        # Step 1: Extract messages from all PST files
        all_messages = []
        
        if self.max_workers == 1:
            # Sequential processing
            for i, pst_file in enumerate(pst_files):
                current_step += 1
                if progress_callback:
                    progress = (current_step / total_steps) * 100
                    progress_callback(progress, f"Processing {Path(pst_file).name}", 
                                    current_step, total_steps)
                
                messages = self._process_single_file(pst_file)
                if messages is not None:
                    all_messages.extend(messages)
        else:
            # Parallel processing
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                future_to_file = {
                    executor.submit(self._process_single_file, pst_file): pst_file 
                    for pst_file in pst_files
                }
                
                for future in as_completed(future_to_file):
                    pst_file = future_to_file[future]
                    current_step += 1
                    
                    if progress_callback:
                        progress = (current_step / total_steps) * 100
                        progress_callback(progress, f"Processed {Path(pst_file).name}", 
                                        current_step, total_steps)
                    
                    try:
                        messages = future.result()
                        if messages is not None:
                            all_messages.extend(messages)
                    except Exception as e:
                        logger.error(f"Error processing {pst_file}: {e}")
                        with self._lock:
                            self.stats['files_failed'] += 1
        
        with self._lock:
            self.stats['total_messages'] = len(all_messages)
        
        logger.info(f"Extracted {len(all_messages)} total messages from {len(pst_files)} files")
        
        # Step 2: Apply filters and duplicate removal
        processed_messages = self._apply_processing(all_messages)
        
        # Step 3: Convert to output formats
        current_step += 1
        if progress_callback:
            progress = (current_step / total_steps) * 100
            progress_callback(progress, "Converting to output formats", 
                            current_step, total_steps)
        
        def conversion_progress(prog, format_name=None, current=None, total=None):
            # Adjust progress to fit within the conversion step
            base_progress = (current_step / total_steps) * 100
            step_progress = (prog / 100) * (100 / total_steps)
            overall_progress = base_progress + step_progress
            
            if progress_callback:
                status = f"Converting to {format_name}" if format_name else "Converting"
                progress_callback(overall_progress, status, current or 0, total or len(processed_messages))
        
        conversion_results = self.converter.batch_convert(
            processed_messages, output_formats, conversion_progress
        )
        
        with self._lock:
            self.stats['conversion_results'] = conversion_results
        
        logger.info("Batch processing completed")
        return self._get_final_stats()
    
    def _process_single_file(self, pst_file: str) -> Optional[List[EmailMessage]]:
        """Process a single PST file and return messages"""
        try:
            logger.debug(f"Processing PST file: {pst_file}")
            
            with PSTParser(pst_file) as parser:
                messages = list(parser.extract_messages())
                
                with self._lock:
                    self.stats['files_processed'] += 1
                
                logger.debug(f"Extracted {len(messages)} messages from {Path(pst_file).name}")
                return messages
                
        except Exception as e:
            logger.error(f"Error processing PST file {pst_file}: {e}")
            with self._lock:
                self.stats['files_failed'] += 1
            return None
    
    def _apply_processing(self, messages: List[EmailMessage]) -> List[EmailMessage]:
        """Apply filtering and duplicate removal to messages"""
        processed_messages = messages
        
        # Apply message filters
        if self.message_filter:
            logger.info("Applying message filters")
            processed_messages = self.message_filter.filter_messages(processed_messages)
            
            with self._lock:
                self.stats['filtered_messages'] = len(processed_messages)
            
            logger.info(f"Filtered down to {len(processed_messages)} messages")
        
        # Remove duplicates if configured
        if self.remove_duplicates and self.duplicate_detector:
            logger.info("Removing duplicate messages")
            original_count = len(processed_messages)
            processed_messages = self.duplicate_detector.remove_duplicates(processed_messages)
            
            duplicates_removed = original_count - len(processed_messages)
            with self._lock:
                self.stats['duplicates_removed'] = duplicates_removed
            
            logger.info(f"Removed {duplicates_removed} duplicate messages")
        
        return processed_messages
    
    def _reset_stats(self):
        """Reset processing statistics"""
        with self._lock:
            self.stats = {
                'files_processed': 0,
                'files_failed': 0,
                'total_messages': 0,
                'filtered_messages': 0,
                'duplicates_removed': 0,
                'conversion_results': {}
            }
    
    def _get_final_stats(self) -> Dict[str, Any]:
        """Get final processing statistics"""
        with self._lock:
            final_stats = self.stats.copy()
        
        # Add summary information
        final_stats['summary'] = {
            'files_total': final_stats['files_processed'] + final_stats['files_failed'],
            'files_success_rate': (
                final_stats['files_processed'] / 
                max(1, final_stats['files_processed'] + final_stats['files_failed'])
            ) * 100,
            'messages_processed': final_stats.get('filtered_messages', final_stats['total_messages']),
            'output_directory': str(self.output_directory)
        }
        
        return final_stats
    
    def get_stats(self) -> Dict[str, Any]:
        """Get current processing statistics"""
        with self._lock:
            return self.stats.copy()


class ProgressTracker:
    """Progress tracking utility for batch operations"""
    
    def __init__(self):
        self.current_progress = 0.0
        self.current_status = "Ready"
        self.current_file = ""
        self.current_step = 0
        self.total_steps = 0
        self.callbacks = []
        self._lock = threading.Lock()
    
    def add_callback(self, callback: Callable):
        """Add a progress callback function"""
        with self._lock:
            self.callbacks.append(callback)
    
    def update_progress(self, progress: float, status: str = "", 
                       current_step: int = 0, total_steps: int = 0):
        """Update progress and notify callbacks"""
        with self._lock:
            self.current_progress = progress
            self.current_status = status
            self.current_step = current_step
            self.total_steps = total_steps
            
            # Notify all callbacks
            for callback in self.callbacks:
                try:
                    callback(progress, status, current_step, total_steps)
                except Exception as e:
                    logger.warning(f"Progress callback error: {e}")
    
    def get_progress(self) -> Dict[str, Any]:
        """Get current progress information"""
        with self._lock:
            return {
                'progress': self.current_progress,
                'status': self.current_status,
                'current_step': self.current_step,
                'total_steps': self.total_steps
            }