#!/usr/bin/env python3
import os
import sys
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET


def extract_settings_xml(docx_path, temp_dir):
    """Extract settings.xml from a docx file to a temporary directory."""
    if not os.path.exists(docx_path):
        print(f"Error: File '{docx_path}' does not exist.")
        return None

    # Extract the docx file
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        try:
            # Extract only settings.xml
            zip_ref.extract('word/settings.xml', temp_dir)
            return os.path.join(temp_dir, 'word', 'settings.xml')
        except KeyError:
            print(f"Error: settings.xml not found in {docx_path}")
            return None


def get_protection_status(settings_path):
    """Extract document protection information from settings.xml"""
    tree = ET.parse(settings_path)
    root = tree.getroot()
    
    # Define namespace mapping
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    protection_info = {
        'has_protection': False,
        'edit_mode': None,
        'enforcement': None,
        'protection_details': {}
    }
    
    # Check for document protection
    doc_protection = root.find('.//w:documentProtection', namespaces=ns)
    if doc_protection is not None:
        protection_info['has_protection'] = True
        
        # Get all attributes of document protection
        for key, value in doc_protection.attrib.items():
            # Convert namespace prefixed attributes to readable form
            if '}' in key:
                key = key.split('}')[1]  # Remove namespace prefix
            protection_info['protection_details'][key] = value
            
            # Store specific important attributes
            if key == 'enforcement':
                protection_info['enforcement'] = value
            elif key == 'edit':
                protection_info['edit_mode'] = value
    
    # Check for track revisions
    track_revisions = root.find('.//w:trackRevisions', namespaces=ns)
    protection_info['track_revisions'] = track_revisions is not None
    
    # Get rsids (revision IDs)
    rsids_element = root.find('.//w:rsids', namespaces=ns)
    if rsids_element is not None:
        rsid_elements = rsids_element.findall('./w:rsid', namespaces=ns)
        protection_info['rsids'] = [rsid.get('{%s}val' % ns['w']) for rsid in rsid_elements]
    else:
        protection_info['rsids'] = []
    
    return protection_info


def compare_settings(file1, file2):
    """Compare settings.xml from two docx files and display user-friendly differences."""
    temp_dir1 = tempfile.mkdtemp()
    temp_dir2 = tempfile.mkdtemp()
    
    try:
        # Extract settings.xml from both files
        settings1_path = extract_settings_xml(file1, temp_dir1)
        settings2_path = extract_settings_xml(file2, temp_dir2)
        
        if not settings1_path or not settings2_path:
            return False
        
        # Get protection info for both files
        protection1 = get_protection_status(settings1_path)
        protection2 = get_protection_status(settings2_path)
        
        # Display comparison in a user-friendly format
        print(f"\n=== COMPARISON OF SETTINGS.XML BETWEEN FILES ===")
        print(f"File 1: {os.path.basename(file1)}")
        print(f"File 2: {os.path.basename(file2)}")
        print("\n=== DOCUMENT PROTECTION ===")
        
        if protection1['has_protection'] != protection2['has_protection']:
            print(f"Document Protection: {'ENABLED' if protection1['has_protection'] else 'DISABLED'} in File 1, "
                  f"{'ENABLED' if protection2['has_protection'] else 'DISABLED'} in File 2")
        else:
            status = "ENABLED" if protection1['has_protection'] else "DISABLED"
            print(f"Document Protection: {status} in both files")
        
        if protection1['has_protection'] and protection2['has_protection']:
            # Compare enforcement
            if protection1['enforcement'] != protection2['enforcement']:
                print(f"Protection Enforcement: {protection1['enforcement']} in File 1, {protection2['enforcement']} in File 2")
            
            # Compare edit mode
            if protection1['edit_mode'] != protection2['edit_mode']:
                print(f"Edit Mode: {protection1['edit_mode']} in File 1, {protection2['edit_mode']} in File 2")
            
            # Compare other protection details
            all_keys = set(protection1['protection_details'].keys()) | set(protection2['protection_details'].keys())
            for key in all_keys:
                val1 = protection1['protection_details'].get(key, 'Not set')
                val2 = protection2['protection_details'].get(key, 'Not set')
                if val1 != val2:
                    print(f"Protection attribute '{key}': {val1} in File 1, {val2} in File 2")
        
        # Compare track revisions
        if protection1['track_revisions'] != protection2['track_revisions']:
            print(f"\nTrack Revisions: {'ENABLED' if protection1['track_revisions'] else 'DISABLED'} in File 1, "
                  f"{'ENABLED' if protection2['track_revisions'] else 'DISABLED'} in File 2")
        
        # Compare revision IDs
        print("\n=== REVISION IDs ===")
        rsids1 = set(protection1['rsids'])
        rsids2 = set(protection2['rsids'])
        
        if rsids1 == rsids2:
            print(f"Both files have the same revision IDs: {len(rsids1)} IDs")
        else:
            print(f"File 1 has {len(rsids1)} revision IDs")
            print(f"File 2 has {len(rsids2)} revision IDs")
            
            if len(rsids1 - rsids2) > 0:
                print(f"Revision IDs unique to File 1: {', '.join(list(rsids1 - rsids2)[:5])}" + 
                      (f" and {len(rsids1 - rsids2) - 5} more..." if len(rsids1 - rsids2) > 5 else ""))
            
            if len(rsids2 - rsids1) > 0:
                print(f"Revision IDs unique to File 2: {', '.join(list(rsids2 - rsids1)[:5])}" + 
                      (f" and {len(rsids2 - rsids1) - 5} more..." if len(rsids2 - rsids1) > 5 else ""))
        
        return True
        
    finally:
        # Clean up
        shutil.rmtree(temp_dir1)
        shutil.rmtree(temp_dir2)


def main():
    if len(sys.argv) != 3:
        print("Usage: python compare.py <docx_file1> <docx_file2>")
        sys.exit(1)
    
    docx_file1 = sys.argv[1]
    docx_file2 = sys.argv[2]
    
    if not compare_settings(docx_file1, docx_file2):
        sys.exit(1)


if __name__ == "__main__":
    main()
