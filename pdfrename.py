import os
import sys

def rename_pdfs(folder_path="."):
    # Get all PDF files in the folder, sorted for consistent ordering
    pdf_files = sorted([
        f for f in os.listdir(folder_path)
        if f.lower().endswith(".pdf")
    ])

    if not pdf_files:
        print("No PDF files found in the specified folder.")
        return

    print(f"Found {len(pdf_files)} PDF file(s). Renaming...")

    for index, filename in enumerate(pdf_files, start=1):
        new_name = f"Scan_{index:03d}.pdf"
        old_path = os.path.join(folder_path, filename)
        new_path = os.path.join(folder_path, new_name)

        # Avoid overwriting if the new name already exists and it's a different file
        if os.path.exists(new_path) and old_path != new_path:
            print(f"  Skipping: {new_name} already exists.")
            continue

        os.rename(old_path, new_path)
        print(f"  {filename}  →  {new_name}")

    print("Done.")

if __name__ == "__main__":
    # Optionally pass a folder path as a command-line argument
    folder = sys.argv[1] if len(sys.argv) > 1 else "."
    
    if not os.path.isdir(folder):
        print(f"Error: '{folder}' is not a valid directory.")
        sys.exit(1)

    rename_pdfs(folder)
