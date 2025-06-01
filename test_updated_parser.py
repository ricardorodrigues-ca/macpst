#!/usr/bin/env python3
"""
Test the updated PST parser
"""

import sys
import logging
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from macpst.core.pst_parser import PSTParser

def test_updated_parser(pst_file_path):
    """Test the updated PST parser"""
    
    # Set up detailed logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    logger = logging.getLogger(__name__)
    
    try:
        logger.info(f"Testing updated parser on: {pst_file_path}")
        
        with PSTParser(pst_file_path) as parser:
            logger.info("PST file opened successfully")
            
            # Get folder structure first
            logger.info("Getting folder structure...")
            folder_tree = parser.get_folder_tree()
            
            def print_folder_tree(folder, indent=0):
                prefix = "  " * indent
                logger.info(f"{prefix}{folder.name} ({folder.message_count} messages)")
                for subfolder in folder.subfolders:
                    print_folder_tree(subfolder, indent + 1)
            
            print_folder_tree(folder_tree)
            
            # Extract messages with the updated parser
            logger.info("Extracting messages with updated parser...")
            message_count = 0
            for message in parser.extract_messages():
                message_count += 1
                logger.info(f"Message {message_count}: '{message.subject[:50]}...' from {message.sender}")
                
                if message_count >= 10:  # Limit to first 10 for testing
                    logger.info("Stopping at 10 messages for test")
                    break
            
            logger.info(f"Total messages found: {message_count}")
            
            if message_count > 0:
                logger.info("SUCCESS: Messages found with updated parser!")
            else:
                logger.warning("No messages found - may need further investigation")
            
    except Exception as e:
        logger.error(f"Error testing updated parser: {e}", exc_info=True)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python test_updated_parser.py <path_to_pst_file>")
        sys.exit(1)
    
    pst_file = sys.argv[1]
    if not Path(pst_file).exists():
        print(f"PST file not found: {pst_file}")
        sys.exit(1)
    
    test_updated_parser(pst_file)