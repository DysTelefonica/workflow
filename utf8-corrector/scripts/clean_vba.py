import os
import sys
import codecs
import argparse

def detect_encoding(file_path):
    """
    Detects encoding based on BOM or attempts to decode with common encodings.
    Prioritizes: UTF-8-SIG (BOM), UTF-16, UTF-8, Windows-1252.
    """
    with open(file_path, 'rb') as f:
        raw = f.read(4)  # Read enough for BOMs
    
    if raw.startswith(codecs.BOM_UTF8):
        return 'utf-8-sig'
    if raw.startswith(codecs.BOM_UTF16_LE) or raw.startswith(codecs.BOM_UTF16_BE):
        return 'utf-16'
    
    # If no BOM, try reading the whole file with different encodings
    encodings = ['utf-8', 'cp1252', 'latin1', 'mbcs']
    
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                f.read()
            return enc
        except UnicodeDecodeError:
            continue
            
    return None

def clean_file(file_path, dry_run=False):
    """
    Reads a file, converts to UTF-8 without BOM, and saves it.
    """
    encoding = detect_encoding(file_path)
    if not encoding:
        print(f"❌ Error: Could not detect encoding for {file_path}")
        return False

    try:
        with open(file_path, 'r', encoding=encoding) as f:
            content = f.read()
        
        # Basic cleanup if needed (e.g. normalize line endings? VBA usually likes CRLF)
        # For now, we just ensure it's valid string content.
        
        if dry_run:
            print(f"✅ Would convert {file_path} from {encoding} to utf-8")
            return True

        # Write back as utf-8 (no BOM)
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
            
        print(f"✅ Converted {file_path}: {encoding} -> utf-8")
        return True
        
    except Exception as e:
        print(f"❌ Error processing {file_path}: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Clean VBA files: Fix encoding to UTF-8 and remove BOM.")
    parser.add_argument("files", nargs='+', help="Files to clean")
    parser.add_argument("--dry-run", action="store_true", help="Check without modifying files")
    
    args = parser.parse_args()
    
    for file_path in args.files:
        if os.path.isdir(file_path):
            # Recursively find files? Or just skip? User said "files or multiple files".
            # Let's assume glob expansion by shell, but if dir passed, walk it.
            for root, dirs, files in os.walk(file_path):
                for file in files:
                    if file.lower().endswith(('.bas', '.cls', '.frm', '.txt')):
                        clean_file(os.path.join(root, file), args.dry_run)
        else:
            if os.path.exists(file_path):
                clean_file(file_path, args.dry_run)
            else:
                print(f"⚠️ File not found: {file_path}")

if __name__ == "__main__":
    main()
