# (Microsoft) Word File Edit Bypass (WFEB) `1.0`

A simple utility to remove edit protection from Microsoft Word documents (.docx).

## What This Tool Does

- **Removes document edit protection** by modifying the `settings.xml` file inside the Word document
- **Disables tracking revisions** to prevent Word from tracking edits (Removes manual accepts)
- **Creates backups** of original files before modification
- **Works with .docx files** (modern Word format)

## What This Tool Does NOT Do

- **Does NOT remove password encryption** - if entire file is encrypted with a password, you still need the password to open it
- **Does NOT change document content** - all text, formatting, and other content remains intact

## How It Works

The tool:

1. Creates a temporary directory
2. Extracts the .docx file (which is a ZIP archive)
3. Locates and modifies the settings.xml file to disable document protection and track changes
4. Carefully repackages the document to preserve Word format compatibility
5. Preserves a backup of the original file with .bak extension

## Usage

```bash
python wfeb.py document.docx
```

Or if installed as a package:

```bash
wfeb document.docx
```

## Installation

```bash
pip install git+https://github.com/KRWCLASSIC/wfeb.git
```

Or for development:

```bash
git clone https://github.com/KRWCLASSIC/wfeb.git
cd wfeb
pip install -e .
```

## Technical Details

When a Word document is protected for editing, it has specific settings in the `settings.xml` file:

- The `documentProtection` element with `enforcement="1"` indicates protection is enabled
- The script changes this to `enforcement="0"` and `edit="edit"` to disable protection
- The script also removes the `trackRevisions` element to prevent Word from telling user to manually accept changes

## Requirements

- Python 3.6 or higher
