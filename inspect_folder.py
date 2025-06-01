#!/usr/bin/env python3
"""
Inspect folder contents in detail
"""

import sys
from pathlib import Path

def inspect_folder_details(pst_file):
    """Inspect folder contents in detail"""
    try:
        import pypff
        
        print(f"Opening PST file: {pst_file}")
        pff_file = pypff.file()
        pff_file.open(str(pst_file))
        
        root_folder = pff_file.get_root_folder()
        
        def inspect_folder(folder, indent=0):
            prefix = "  " * indent
            folder_name = getattr(folder, 'name', 'Unnamed')
            
            print(f"{prefix}=== FOLDER: {folder_name} ===")
            
            # Print all available attributes
            print(f"{prefix}Available attributes:")
            for attr in sorted(dir(folder)):
                if not attr.startswith('_'):
                    try:
                        value = getattr(folder, attr)
                        if callable(value):
                            print(f"{prefix}  {attr}() - method")
                        else:
                            print(f"{prefix}  {attr} = {value} ({type(value).__name__})")
                    except Exception as e:
                        print(f"{prefix}  {attr} - error: {e}")
            
            # Try different ways to get message count
            print(f"{prefix}Message count attempts:")
            for method in ['number_of_messages', 'get_number_of_messages', 'message_count']:
                if hasattr(folder, method):
                    try:
                        attr_value = getattr(folder, method)
                        if callable(attr_value):
                            count = attr_value()
                            print(f"{prefix}  {method}() = {count}")
                        else:
                            count = attr_value
                            print(f"{prefix}  {method} = {count}")
                    except Exception as e:
                        print(f"{prefix}  {method} - error: {e}")
            
            # Try to manually access messages
            print(f"{prefix}Manual message access:")
            for i in range(10):  # Try first 10 message slots
                try:
                    message = folder.get_message(i)
                    if message:
                        subject = getattr(message, 'subject', f'Message {i}')
                        print(f"{prefix}  Message {i}: {subject}")
                    else:
                        print(f"{prefix}  Message {i}: None")
                except Exception as e:
                    print(f"{prefix}  Message {i}: Error - {e}")
                    if i == 0:
                        print(f"{prefix}  (No messages found)")
                        break
            
            # Process subfolders (limit to important ones)
            if folder_name in ['Inbox', 'Sent Items', 'Drafts'] or indent == 0:
                try:
                    num_subfolders = 0
                    if hasattr(folder, 'number_of_sub_folders'):
                        num_subfolders = folder.number_of_sub_folders
                    elif hasattr(folder, 'get_number_of_sub_folders'):
                        num_subfolders = folder.get_number_of_sub_folders()
                    
                    print(f"{prefix}Subfolders: {num_subfolders}")
                    
                    for i in range(min(num_subfolders, 5)):  # Limit to first 5
                        try:
                            subfolder = folder.get_sub_folder(i)
                            if subfolder:
                                subfolder_name = getattr(subfolder, 'name', f'Subfolder_{i}')
                                if subfolder_name in ['Inbox', 'Sent Items', 'Drafts']:
                                    inspect_folder(subfolder, indent + 1)
                        except Exception as e:
                            print(f"{prefix}  Subfolder {i}: Error - {e}")
                            
                except Exception as e:
                    print(f"{prefix}Error accessing subfolders: {e}")
            
            print(f"{prefix}=== END FOLDER: {folder_name} ===\n")
        
        inspect_folder(root_folder)
        
        pff_file.close()
        
    except ImportError:
        print("pypff not available")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python inspect_folder.py <pst_file>")
        sys.exit(1)
    
    pst_file = sys.argv[1]
    inspect_folder_details(pst_file)