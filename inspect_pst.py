#!/usr/bin/env python3
"""
Simple PST inspection script
"""

import sys
from pathlib import Path

def inspect_with_pypff(pst_file):
    """Inspect PST file directly with pypff"""
    try:
        import pypff
        
        print(f"Opening PST file: {pst_file}")
        pff_file = pypff.file()
        pff_file.open(str(pst_file))
        
        root_folder = pff_file.get_root_folder()
        print(f"Root folder: {root_folder}")
        print(f"Root folder name: {getattr(root_folder, 'name', 'No name')}")
        
        # Print all available attributes
        print("\nAvailable attributes on root folder:")
        for attr in dir(root_folder):
            if not attr.startswith('_'):
                try:
                    value = getattr(root_folder, attr)
                    print(f"  {attr}: {type(value)} = {value}")
                except:
                    print(f"  {attr}: <error accessing>")
        
        # Try to get subfolders
        print("\nTrying to access subfolders...")
        for method in ['number_of_sub_folders', 'get_number_of_sub_folders']:
            if hasattr(root_folder, method):
                try:
                    if callable(getattr(root_folder, method)):
                        count = getattr(root_folder, method)()
                    else:
                        count = getattr(root_folder, method)
                    print(f"  {method}: {count}")
                except Exception as e:
                    print(f"  {method}: Error - {e}")
        
        # Try to get messages
        print("\nTrying to access messages...")
        for method in ['number_of_messages', 'get_number_of_messages']:
            if hasattr(root_folder, method):
                try:
                    if callable(getattr(root_folder, method)):
                        count = getattr(root_folder, method)()
                    else:
                        count = getattr(root_folder, method)
                    print(f"  {method}: {count}")
                except Exception as e:
                    print(f"  {method}: Error - {e}")
        
        pff_file.close()
        
    except ImportError:
        print("pypff not available")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python inspect_pst.py <pst_file>")
        sys.exit(1)
    
    pst_file = sys.argv[1]
    inspect_with_pypff(pst_file)