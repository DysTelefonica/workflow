---
name: vba-cleaner
description: Clean and fix encoding issues in exported VBA code from Microsoft Access. Use when the user uploads VBA files (.bas, .cls, .frm) with encoding problems, BOM (Byte Order Mark) characters, or requests to "clean", "fix encoding", or "remove BOM" from VBA/Access code files. CRITICAL - Also trigger automatically whenever Claude modifies, edits, or generates any VBA code to ensure the output is clean UTF-8 without BOM before presenting to the user. Automatically detects encoding, removes BOM, and converts to clean UTF-8.
---

# VBA Code Cleaner

Fix encoding issues and BOM (Byte Order Mark) problems in VBA files exported from Microsoft Access.

## ⚠️ CRITICAL AUTOMATIC RULE

**ALWAYS run this cleaning script on ANY VBA code before presenting it to the user**, whether:
- User uploaded a VBA file that Claude edited
- Claude generated new VBA code from scratch
- Claude modified existing VBA code
- Claude refactored, fixed bugs, or made any changes to VBA code

**Workflow mandate:**
1. Make your code changes in a temporary file
2. ALWAYS run `python scripts/clean_vba.py` on the file
3. ONLY THEN present the cleaned file to the user

This ensures the user can copy-paste the code directly into their VBA IDE without encoding issues.

## Overview

When exporting VBA modules from Microsoft Access, files often contain:
- Incorrect encoding (non-UTF-8)
- BOM (Byte Order Mark) characters that appear as "BOOM" or other artifacts
- Mixed line endings
- Character encoding errors

This skill provides a Python script that automatically:
1. Detects the original file encoding
2. Removes all types of BOM markers
3. Converts to clean UTF-8 without BOM
4. Normalizes line endings
5. Preserves the VBA code structure and functionality

## Usage

### Single File

Clean a single VBA file in place:
```bash
python scripts/clean_vba.py Module1.bas
```

Create a cleaned copy with a new name:
```bash
python scripts/clean_vba.py Module1.bas Module1_clean.bas
```

### Multiple Files (Directory)

Clean all .bas files in a directory:
```bash
python scripts/clean_vba.py /path/to/vba_modules --dir
```

Clean all .cls files (class modules):
```bash
python scripts/clean_vba.py /path/to/vba_modules --dir --pattern *.cls
```

Clean all VBA files recursively in subdirectories:
```bash
python scripts/clean_vba.py /path/to/vba_modules --dir --pattern *.bas --recursive
```

## Workflow

### Automatic Usage (Required for all VBA code modifications)

When Claude modifies ANY VBA code:
```bash
# 1. Claude makes changes to the VBA file
# 2. MANDATORY: Clean the file before presenting to user
python scripts/clean_vba.py /home/claude/modified_module.bas /mnt/user-data/outputs/Module1.bas

# 3. Present the cleaned file to user
```

### Manual Usage (User-initiated cleaning)

1. User uploads VBA file(s) to `/mnt/user-data/uploads`
2. Run the cleaning script on the file(s)
3. Move cleaned files to `/mnt/user-data/outputs`
4. Present cleaned files to user

Example:
```bash
# Clean a single uploaded file
python scripts/clean_vba.py /mnt/user-data/uploads/Module1.bas /mnt/user-data/outputs/Module1.bas

# Or clean all VBA files from uploads
for file in /mnt/user-data/uploads/*.bas; do
    filename=$(basename "$file")
    python scripts/clean_vba.py "$file" "/mnt/user-data/outputs/$filename"
done
```

## Supported File Types

- `.bas` - Standard modules
- `.cls` - Class modules  
- `.frm` - Form modules
- Any text file with VBA code

## Dependencies

The script requires the `chardet` library for encoding detection:
```bash
pip install chardet --break-system-packages
```

If `chardet` is not available, the script falls back to trying common encodings (UTF-8, CP1252, Latin1).
