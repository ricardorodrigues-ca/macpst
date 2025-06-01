#!/usr/bin/env python3
"""
Check all types of items in PST file
"""

import sys
from pathlib import Path

def check_all_items(pst_file):
    """Check all types of items in PST file"""
    try:
        import pypff
        
        print(f"Opening PST file: {pst_file}")
        pff_file = pypff.file()
        pff_file.open(str(pst_file))
        
        root_folder = pff_file.get_root_folder()
        
        def check_folder_items(folder, folder_name="", level=0):
            indent = "  " * level
            name = getattr(folder, 'name', folder_name or 'Unknown')
            
            print(f"{indent}Checking folder: {name}")
            
            # Check for different types of items
            item_types = [
                ('messages', 'get_message'),
                ('sub_messages', 'get_sub_message'), 
                ('items', 'get_item'),
                ('records', 'get_record'),
                ('entries', 'get_entry')
            ]
            
            for item_type, get_method in item_types:
                if hasattr(folder, get_method):
                    print(f"{indent}  Has method: {get_method}")
                    
                    # Try to get count
                    count_methods = [
                        f'number_of_{item_type}',
                        f'get_number_of_{item_type}',
                        f'{item_type}_count'
                    ]
                    
                    for count_method in count_methods:
                        if hasattr(folder, count_method):
                            try:
                                attr = getattr(folder, count_method)
                                if callable(attr):
                                    count = attr()
                                else:
                                    count = attr
                                print(f"{indent}    {count_method}: {count}")
                                
                                if count > 0:
                                    # Try to access first few items
                                    print(f"{indent}    Found {count} {item_type}!")
                                    for i in range(min(count, 3)):
                                        try:
                                            item = getattr(folder, get_method)(i)
                                            if item:
                                                # Try to get some info about the item
                                                info = "Item found"
                                                if hasattr(item, 'subject'):
                                                    info = f"Subject: {getattr(item, 'subject', 'No subject')}"
                                                elif hasattr(item, 'display_name'):
                                                    info = f"Name: {getattr(item, 'display_name', 'No name')}"
                                                elif hasattr(item, 'title'):
                                                    info = f"Title: {getattr(item, 'title', 'No title')}"
                                                print(f"{indent}      {item_type}[{i}]: {info}")
                                            else:
                                                print(f"{indent}      {item_type}[{i}]: None")
                                        except Exception as e:
                                            print(f"{indent}      {item_type}[{i}]: Error - {e}")
                                            
                            except Exception as e:
                                print(f"{indent}    {count_method}: Error - {e}")
            
            # Check subfolders (only important ones or if at root level)
            important_folders = ['Inbox', 'Sent Items', 'Drafts', 'Contacts', 'Calendar', 'Tasks']
            if level == 0 or name in important_folders:
                try:
                    subfolder_count = 0
                    if hasattr(folder, 'number_of_sub_folders'):
                        subfolder_count = folder.number_of_sub_folders
                    elif hasattr(folder, 'get_number_of_sub_folders'):
                        subfolder_count = folder.get_number_of_sub_folders()
                    
                    for i in range(min(subfolder_count, 10)):
                        try:
                            subfolder = folder.get_sub_folder(i)
                            if subfolder:
                                check_folder_items(subfolder, level=level+1)
                        except Exception as e:
                            print(f"{indent}  Subfolder {i}: Error - {e}")
                            
                except Exception as e:
                    print(f"{indent}Error checking subfolders: {e}")
        
        check_folder_items(root_folder, "Root")
        
        pff_file.close()
        
    except ImportError:
        print("pypff not available")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python check_pst_items.py <pst_file>")
        sys.exit(1)
    
    pst_file = sys.argv[1]
    check_all_items(pst_file)