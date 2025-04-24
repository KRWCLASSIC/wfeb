import os
import sys
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET

def modify_settings_xml(file_path):
    # Parse the XML file
    tree = ET.parse(file_path)
    root = tree.getroot()
    
    # Define namespace for Word documents
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # Set w:edit to "edit" and enforcement to 0 in w:documentProtection
    document_protection = root.find('.//w:documentProtection', namespaces=ns)
    if document_protection is not None:
        document_protection.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}edit', 'edit')
        document_protection.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}enforcement', '0')
        print("Document protection disabled successfully.")
    else:
        print("No document protection found in settings.xml.")
    
    # Remove track revisions if it exists
    track_revisions = root.find('.//w:trackRevisions', namespaces=ns)
    if track_revisions is not None:
        parent = track_revisions.getparent() if hasattr(track_revisions, 'getparent') else root
        parent.remove(track_revisions)
        print("Track revisions disabled successfully.")

    # Write the modified XML back to the file with proper XML declaration
    tree.write(file_path, encoding='UTF-8', xml_declaration=True)

def process_docx(docx_path):
    if not os.path.exists(docx_path):
        print(f"Error: File '{docx_path}' does not exist.")
        return False

    # Create a temporary directory
    temp_dir = tempfile.mkdtemp()
    try:
        # Extract the docx file (which is a zip file)
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Path to settings.xml
        settings_path = os.path.join(temp_dir, 'word', 'settings.xml')
        
        if not os.path.exists(settings_path):
            print("Error: settings.xml not found in the document.")
            return False

        # Modify the settings.xml file
        modify_settings_xml(settings_path)

        # Create a new docx file
        backup_path = docx_path + '.bak'
        shutil.copy2(docx_path, backup_path)
        print(f"Backup created at: {backup_path}")

        # Create a new zip file - Fix the compression to match original Office format
        new_docx_path = docx_path + '.new'
        with zipfile.ZipFile(new_docx_path, 'w', compression=zipfile.ZIP_DEFLATED) as new_docx:
            # First add the [Content_Types].xml file
            content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
            if os.path.exists(content_types_path):
                new_docx.write(content_types_path, '[Content_Types].xml')
            
            # Then add the _rels folder
            rels_dir = os.path.join(temp_dir, '_rels')
            if os.path.exists(rels_dir):
                for foldername, _, filenames in os.walk(rels_dir):
                    for filename in filenames:
                        file_path = os.path.join(foldername, filename)
                        arcname = os.path.relpath(file_path, temp_dir)
                        new_docx.write(file_path, arcname)
            
            # Then proceed with the rest of the files
            for foldername, _, filenames in os.walk(temp_dir):
                if os.path.basename(foldername) == '_rels' or foldername == rels_dir:
                    continue  # Skip _rels as it's already processed
                
                for filename in filenames:
                    if filename == '[Content_Types].xml':
                        continue  # Skip content types as it's already processed
                    
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    new_docx.write(file_path, arcname)

        # Replace the original with the new file
        os.remove(docx_path)
        os.rename(new_docx_path, docx_path)

        print(f"Successfully processed: {docx_path}")
        return True

    finally:
        # Clean up the temporary directory
        shutil.rmtree(temp_dir)

def main():
    """Main entry point for the command-line interface."""
    if len(sys.argv) < 2:
        print("Usage: wfeb <path_to_docx_file>")
        return 1

    docx_path = sys.argv[1]
    if not process_docx(docx_path):
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
