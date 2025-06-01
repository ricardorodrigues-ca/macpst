#!/usr/bin/env python3
"""
Debug script to test PST parsing
"""

import sys
import logging
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from macpst.core.pst_parser import PSTParser

def debug_pst_file(pst_file_path):
    """Debug PST file parsing"""
    
    # Set up detailed logging
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    logger = logging.getLogger(__name__)
    
    try:
        logger.info(f"Starting PST debug for: {pst_file_path}")
        
        # Test PST file opening
        with PSTParser(pst_file_path) as parser:
            logger.info("PST file opened successfully")
            
            # Get statistics
            stats = parser.get_statistics()
            logger.info(f"PST Statistics: {stats}")
            
            # Get folder structure
            logger.info("Getting folder structure...")
            folder_tree = parser.get_folder_tree()
            logger.info(f"Root folder: {folder_tree.name}, message_count: {folder_tree.message_count}")
            
            def print_folder_tree(folder, indent=0):
                prefix = "  " * indent
                logger.info(f"{prefix}{folder.name} ({folder.message_count} messages)")
                for subfolder in folder.subfolders:
                    print_folder_tree(subfolder, indent + 1)
            
            print_folder_tree(folder_tree)
            
            # Extract messages
            logger.info("Extracting messages...")
            message_count = 0
            for message in parser.extract_messages():
                message_count += 1
                logger.info(f"Message {message_count}: {message.subject[:50]}... from {message.sender}")
                
                if message_count >= 5:  # Limit to first 5 for debugging
                    logger.info("Stopping at 5 messages for debug")
                    break
            
            logger.info(f"Total messages found: {message_count}")
            
    except Exception as e:
        logger.error(f"Error debugging PST file: {e}", exc_info=True)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python debug_pst.py <path_to_pst_file>")
        sys.exit(1)
    
    pst_file = sys.argv[1]
    if not Path(pst_file).exists():
        print(f"PST file not found: {pst_file}")
        sys.exit(1)
    
    debug_pst_file(pst_file)