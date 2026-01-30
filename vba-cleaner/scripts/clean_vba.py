#!/usr/bin/env python3
"""
VBA Code Cleaner
Fixes encoding issues and removes BOM from exported Access VBA files.
"""

import sys
import os
from pathlib import Path
import chardet


def detect_encoding(file_path):
    """Detect the encoding of a file."""
    with open(file_path, 'rb') as f:
        raw_data = f.read()
    result = chardet.detect(raw_data)
    return result['encoding'], raw_data


def clean_vba_file(input_path, output_path=None):
    """
    Clean a VBA file by:
    1. Detecting the current encoding
    2. Removing BOM (Byte Order Mark)
    3. Converting to UTF-8 without BOM
    
    Args:
        input_path: Path to the input VBA file
        output_path: Path for the cleaned file (if None, overwrites input)
    
    Returns:
        True if successful, False otherwise
    """
    try:
        input_path = Path(input_path)
        
        if not input_path.exists():
            print(f"Error: File not found: {input_path}")
            return False
        
        # Detect encoding and read raw bytes
        detected_encoding, raw_data = detect_encoding(input_path)
        
        # Remove BOM if present
        # UTF-8 BOM
        if raw_data.startswith(b'\xef\xbb\xbf'):
            raw_data = raw_data[3:]
        # UTF-16 LE BOM
        elif raw_data.startswith(b'\xff\xfe'):
            raw_data = raw_data[2:]
        # UTF-16 BE BOM
        elif raw_data.startswith(b'\xfe\xff'):
            raw_data = raw_data[2:]
        # UTF-32 LE BOM
        elif raw_data.startswith(b'\xff\xfe\x00\x00'):
            raw_data = raw_data[4:]
        # UTF-32 BE BOM
        elif raw_data.startswith(b'\x00\x00\xfe\xff'):
            raw_data = raw_data[4:]
        
        # Decode using detected encoding (or fallback encodings)
        text = None
        encodings_to_try = [detected_encoding, 'utf-8', 'cp1252', 'latin1', 'iso-8859-1']
        
        for encoding in encodings_to_try:
            if encoding is None:
                continue
            try:
                text = raw_data.decode(encoding)
                break
            except (UnicodeDecodeError, LookupError):
                continue
        
        if text is None:
            # Last resort: decode with errors='replace'
            text = raw_data.decode('utf-8', errors='replace')
        
        # Normalize line endings to \n
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        
        # Determine output path
        if output_path is None:
            output_path = input_path
        else:
            output_path = Path(output_path)
        
        # Write as UTF-8 without BOM
        with open(output_path, 'w', encoding='utf-8', newline='\n') as f:
            f.write(text)
        
        print(f"✓ Cleaned: {input_path.name}")
        if detected_encoding:
            print(f"  Original encoding: {detected_encoding}")
        print(f"  Output: {output_path}")
        
        return True
        
    except Exception as e:
        print(f"Error processing {input_path}: {str(e)}")
        return False


def clean_directory(directory_path, pattern="*.bas", recursive=False):
    """
    Clean all VBA files in a directory.
    
    Args:
        directory_path: Path to the directory
        pattern: File pattern to match (default: *.bas)
        recursive: Whether to search recursively
    
    Returns:
        Number of files successfully cleaned
    """
    directory = Path(directory_path)
    
    if not directory.is_dir():
        print(f"Error: Not a directory: {directory_path}")
        return 0
    
    if recursive:
        files = list(directory.rglob(pattern))
    else:
        files = list(directory.glob(pattern))
    
    if not files:
        print(f"No files matching '{pattern}' found in {directory_path}")
        return 0
    
    success_count = 0
    for file_path in files:
        if clean_vba_file(file_path):
            success_count += 1
    
    print(f"\nProcessed {success_count}/{len(files)} files successfully")
    return success_count


def main():
    """Main entry point for command-line usage."""
    if len(sys.argv) < 2:
        print("Usage:")
        print("  Single file:  python clean_vba.py <file_path> [output_path]")
        print("  Directory:    python clean_vba.py <directory_path> --dir [--pattern *.bas] [--recursive]")
        print("\nExamples:")
        print("  python clean_vba.py Module1.bas")
        print("  python clean_vba.py Module1.bas Module1_clean.bas")
        print("  python clean_vba.py ./vba_modules --dir")
        print("  python clean_vba.py ./vba_modules --dir --pattern *.cls --recursive")
        sys.exit(1)
    
    path = sys.argv[1]
    
    # Check if directory mode
    if '--dir' in sys.argv:
        pattern = "*.bas"
        recursive = False
        
        # Check for pattern flag
        if '--pattern' in sys.argv:
            pattern_idx = sys.argv.index('--pattern')
            if pattern_idx + 1 < len(sys.argv):
                pattern = sys.argv[pattern_idx + 1]
        
        # Check for recursive flag
        if '--recursive' in sys.argv:
            recursive = True
        
        clean_directory(path, pattern, recursive)
    else:
        # Single file mode
        output_path = sys.argv[2] if len(sys.argv) > 2 else None
        clean_vba_file(path, output_path)


if __name__ == "__main__":
    main()
