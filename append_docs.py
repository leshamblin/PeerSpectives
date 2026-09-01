import os
import argparse
from docxcompose.composer import Composer
from docx import Document

def append_document_to_files(target_directory, source_document_path):
    """
    Appends the source document to every .docx file in the target directory.
    """
    # Verify paths
    if not os.path.isdir(target_directory):
        print(f"Error: Target directory not found: {target_directory}")
        return

    if not os.path.isfile(source_document_path):
        print(f"Error: Source document not found: {source_document_path}")
        return

    # Load the source document once (we'll reload it for each composition to be safe, 
    # or handle it within the loop if docxcompose modifies it)
    
    files_processed = 0
    errors = 0

    print(f"Scanning directory: {target_directory}")
    
    for filename in os.listdir(target_directory):
        if filename.endswith(".docx") and not filename.startswith("~$"):
            target_file_path = os.path.join(target_directory, filename)
            
            # Skip the source file if it's in the same directory
            if os.path.abspath(target_file_path) == os.path.abspath(source_document_path):
                continue

            print(f"Processing: {filename}...")

            try:
                # Open the target document
                target_doc = Document(target_file_path)
                
                # We need to open the source document fresh for each composition 
                # to avoid any state carrying over or modification issues
                source_doc = Document(source_document_path)

                # Compose
                composer = Composer(target_doc)
                composer.append(source_doc)

                # Save
                composer.save(target_file_path)
                print(f"  -> Successfully appended to {filename}")
                files_processed += 1

            except Exception as e:
                print(f"  -> Failed to process {filename}: {e}")
                errors += 1

    print("\nSummary:")
    print(f"Files processed: {files_processed}")
    print(f"Errors: {errors}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Append a Word document to all .docx files in a directory.")
    parser.add_argument("target_dir", help="Path to the directory containing the team files")
    parser.add_argument("source_doc", help="Path to the source .docx file to append")

    args = parser.parse_args()

    append_document_to_files(args.target_dir, args.source_doc)
