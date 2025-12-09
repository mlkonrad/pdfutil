import os
import json
import shutil
from glob import glob
from pypdf import PdfReader, PdfWriter
from PIL import Image
import time 

# --- CONFIGURATION FILE ---
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.json')

def load_config():
    """Loads the home folder path from the config file or uses the current directory."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            try:
                config = json.load(f)
                saved_path = config.get('home_folder', os.getcwd())
                # Ensure saved path is still valid
                return saved_path if os.path.isdir(saved_path) else os.getcwd()
            except json.JSONDecodeError:
                return os.getcwd()
    return os.getcwd()

def save_config(home_folder):
    """Saves the home folder path to the config file."""
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump({'home_folder': home_folder}, f)
        print("✅ Configuration saved successfully!")
    except Exception as e:
        print(f"❌ ERROR: Could not save configuration. {e}")

# --- HELPER FUNCTION: Convert Images ---

def convert_images_to_pdf(target_folder):
    """
    Finds JPG and BMP files, converts them to single-page PDFs, and saves 
    them temporarily to a *non-synced* path before merging.
    Returns the path to the temporary folder.
    """
    # 🔥 FIX: Create temp folder in the system's TEMP directory to avoid OneDrive locks.
    temp_dir_base = os.path.join(os.environ.get('TEMP', os.getcwd()), "pdf_utility_temps")
    # Use process ID (pid) to create a unique folder for safety
    temp_folder = os.path.join(temp_dir_base, f"temp_pdfs_{os.getpid()}") 
    
    os.makedirs(temp_folder, exist_ok=True) 

    # Find common image files (JPG/JPEG, BMP)
    image_extensions = ('*.jpg', '*.jpeg', '*.bmp')
    all_image_paths = []
    
    for ext in image_extensions:
        all_image_paths.extend(glob(os.path.join(target_folder, ext)))
        
    if not all_image_paths:
        print("ℹ️ No image files found for conversion.")
        return None

    print(f"Found {len(all_image_paths)} images. Converting to temporary PDFs...")
    
    # Process each image
    for i, img_path in enumerate(all_image_paths):
        img = None 
        try:
            img = Image.open(img_path).convert('RGB')
            base_name = os.path.splitext(os.path.basename(img_path))[0]
            temp_pdf_path = os.path.join(temp_folder, f"{base_name}.pdf")
            
            img.save(temp_pdf_path, "PDF", resolution=100.0) 
            print(f"  Converted {os.path.basename(img_path)} -> {os.path.basename(temp_pdf_path)}")
            
        except Exception as e:
            print(f"  ⚠️ WARNING: Could not convert image {os.path.basename(img_path)}. Skipping. Error: {e}")
        
        finally:
            if img:
                img.close() 

    return temp_folder

# --- PDF MERGING LOGIC ---

def merge_pdf_in_folder(target_folder, output_filename):
    """
    Converts images to PDF, then merges all PDF files (original and converted) 
    in the target folder with maximum compression applied.
    """
    print(f"\n--- Starting Merge in {target_folder} ---")
    temp_folder = None
    all_pdf_files = []
    
    try:
        # 1. Convert Images to temporary PDFs (in a local temp path)
        temp_folder = convert_images_to_pdf(target_folder)

        # 2. Collect ALL PDFs (Originals from target_folder + Temporaries from temp_folder)
        original_pdf_files = sorted(glob(os.path.join(target_folder, '*.pdf')))
        
        if temp_folder:
            temp_pdf_files = sorted(glob(os.path.join(temp_folder, '*.pdf')))
            all_pdf_files.extend(original_pdf_files)
            all_pdf_files.extend(temp_pdf_files)
        else:
            all_pdf_files.extend(original_pdf_files)
        
        if not all_pdf_files:
            print("❌ No PDF files or convertible images found to merge.")
            return

        # 3. Merging Process
        writer = PdfWriter()
        print(f"\nFound {len(all_pdf_files)} total documents to merge. Processing...")

        for i, file_path in enumerate(all_pdf_files):
            reader = None 
            try:
                reader = PdfReader(file_path)
                print(f"  [{i+1}/{len(all_pdf_files)}] Adding: {os.path.basename(file_path)}")
                
                for page in reader.pages:
                    writer.add_page(page)
                    
            except Exception as e:
                print(f"  ⚠️ WARNING: Could not read {os.path.basename(file_path)}. Skipping. Error: {e}")
            
            finally:
                if reader:
                    reader.close() 

        # 4. Finalize and Write Output with Compression
        output_path = os.path.join(target_folder, output_filename)
        
        print("\nApplying compression...")
        
        # Compress all content streams (works with all pypdf versions)
        for page_num, page in enumerate(writer.pages):
            try:
                # Compress text and vector graphics
                page.compress_content_streams()
                            
            except Exception as e:
                print(f"  ⚠️ Could not compress page {page_num + 1}: {e}")
        
        print("Compression applied successfully!")
        
        output_file = None 
        try:
            output_file = open(output_path, "wb") 
            writer.write(output_file) 
            
            print(f"\n✅ SUCCESS! Documents merged into: {output_path}")

        except Exception as e:
            print(f"\n❌ FATAL ERROR writing file: {e}")
            
        finally:
            if output_file:
                output_file.close()
            
    finally:
        # 5. Clean up the temporary PDF folder with robust retry logic
        if temp_folder and os.path.exists(temp_folder):
            
            MAX_RETRIES = 5
            DELAY_SECONDS = 1 
            
            print("\n🧹 Attempting robust cleanup of temporary files...")
            
            for i in range(MAX_RETRIES):
                try:
                    shutil.rmtree(temp_folder)
                    print("🧹 Cleaned up temporary PDF files successfully.")
                    break 
                except Exception as e:
                    if i == MAX_RETRIES - 1:
                        print(f"❌ Could not clean up temp folder after {MAX_RETRIES} attempts! Error: {e}")
                        print(f"Please delete this manually: {temp_folder}.")
                    else:
                        print(f"   Cleanup failed (Retry {i + 1}/{MAX_RETRIES}). Waiting {DELAY_SECONDS}s...")
                        time.sleep(DELAY_SECONDS)


# --- PASSWORD PROTECTION LOGIC ---

def add_password_to_pdf(input_pdf, output_pdf, password):
    """Add password protection to a PDF file."""
    reader = None
    output_file = None
    try:
        reader = PdfReader(input_pdf)
        writer = PdfWriter()
        
        # Copy all pages
        for page in reader.pages:
            writer.add_page(page)
        
        # Add password protection
        writer.encrypt(password)
        
        # Save the encrypted PDF
        output_file = open(output_pdf, 'wb')
        writer.write(output_file)
        
        print(f"  ✅ PDF encrypted: {os.path.basename(output_pdf)}")
        return True
        
    except Exception as e:
        print(f"  ❌ ERROR encrypting {os.path.basename(input_pdf)}: {e}")
        return False
        
    finally:
        if reader:
            reader.close()
        if output_file:
            output_file.close()


def add_password_to_excel(input_file, output_file, password):
    """Add password protection to Excel file using msoffcrypto-tool (full encryption)."""
    try:
        import msoffcrypto
        
        input_stream = None
        output_stream = None
        
        try:
            input_stream = open(input_file, 'rb')
            office_file = msoffcrypto.OfficeFile(input_stream)
            office_file.load_key(password=password)
            
            output_stream = open(output_file, 'wb')
            office_file.encrypt(password, output_stream)
            
            print(f"  ✅ Excel file encrypted: {os.path.basename(output_file)}")
            return True
            
        finally:
            if input_stream:
                input_stream.close()
            if output_stream:
                output_stream.close()
        
    except ImportError:
        print("  ❌ ERROR: msoffcrypto-tool library not installed. Run: pip install msoffcrypto-tool")
        return False
    except Exception as e:
        print(f"  ❌ ERROR encrypting {os.path.basename(input_file)}: {e}")
        return False


def protect_files_in_folder(target_folder, password):
    """Add password protection to all PDF and Excel files in the folder."""
    print(f"\n--- Starting Password Protection in {target_folder} ---")
    
    # Create protected subfolder
    protected_folder = os.path.join(target_folder, "protected")
    os.makedirs(protected_folder, exist_ok=True)
    
    # Find PDF files
    pdf_files = glob(os.path.join(target_folder, '*.pdf'))
    # Find Excel files
    excel_files = glob(os.path.join(target_folder, '*.xlsx')) + glob(os.path.join(target_folder, '*.xls'))
    
    total_files = len(pdf_files) + len(excel_files)
    
    if total_files == 0:
        print("❌ No PDF or Excel files found in this folder.")
        return
    
    print(f"\nFound {len(pdf_files)} PDF file(s) and {len(excel_files)} Excel file(s).")
    print(f"Protected files will be saved to: {protected_folder}\n")
    
    success_count = 0
    
    # Process PDF files
    if pdf_files:
        print("Processing PDF files...")
        for pdf_file in pdf_files:
            basename = os.path.basename(pdf_file)
            output_path = os.path.join(protected_folder, basename)
            
            if add_password_to_pdf(pdf_file, output_path, password):
                success_count += 1
    
    # Process Excel files
    if excel_files:
        print("\nProcessing Excel files...")
        for excel_file in excel_files:
            basename = os.path.basename(excel_file)
            output_path = os.path.join(protected_folder, basename)
            
            if add_password_to_excel(excel_file, output_path, password):
                success_count += 1
    
    print(f"\n✅ SUCCESS! {success_count}/{total_files} files protected.")
    print(f"Protected files saved in: {protected_folder}")


def run_password_protection(start_path):
    """Handles the user interaction for selecting a folder and setting password."""
    print("Launching folder navigator...")
    selected_folder = text_file_navigator(start_path)
    
    if selected_folder:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("=" * 50)
        print(f"Selected Folder: {selected_folder}")
        print("=" * 50)
        print("⚠️  WARNING: All PDF and Excel files in this folder will be")
        print("   password-protected and saved to a 'protected' subfolder.")
        print("-" * 50)
        
        confirm = input("Do you want to continue? (Y/N): ").strip().upper()
        
        if confirm == 'Y':
            while True:
                password = input("\nEnter password for protection: ").strip()
                if password:
                    password_confirm = input("Confirm password: ").strip()
                    if password == password_confirm:
                        break
                    else:
                        print("❌ Passwords don't match. Try again.")
                else:
                    print("❌ Password cannot be empty.")
            
            protect_files_in_folder(selected_folder, password)
        else:
            print("Password protection cancelled.")
    else:
        print("Password protection cancelled.")
    
    input("\nPress Enter to return to the Main Menu...")


# --- TEXT-BASED FILE NAVIGATOR ---

def text_file_navigator(start_path):
    """Allows the user to navigate directories via text input."""
    current_path = start_path
    
    while True:
        if not os.path.isdir(current_path):
            print(f"Path does not exist: {current_path}. Resetting to current directory.")
            current_path = os.getcwd()

        os.system('cls' if os.name == 'nt' else 'clear') 
        print("=" * 50)
        print("📁 FOLDER NAVIGATOR (Start at HOME FOLDER)")
        print("=" * 50)
        print(f"CURRENT PATH: {current_path}")
        print("-" * 50)
        
        try:
            items = [d for d in os.listdir(current_path) if not d.startswith('.')]
        except PermissionError:
            print("❌ Permission denied to access this folder.")
            input("Press Enter to go up one level...")
            current_path = os.path.dirname(current_path)
            continue
            
        folders = sorted([item for item in items if os.path.isdir(os.path.join(current_path, item))])
        
        print("AVAILABLE FOLDERS (Type folder number to enter):")
        if not folders:
            print("  (No subfolders found)")

        for i, folder in enumerate(folders):
            print(f"  [{i+1}] {folder}")

        print("\nOPTIONS:")
        print("  [..] Go Up one level")
        print("  [S]  Select this folder")
        print("  [Q]  Quit to Main Menu")
        print("-" * 50)
        
        choice = input("Enter option, folder number, or folder name: ").strip()

        if choice.upper() == 'Q':
            return None
        
        elif choice == '..':
            parent_path = os.path.dirname(current_path)
            if parent_path != current_path: 
                current_path = parent_path
            
        elif choice.upper() == 'S':
            return current_path
            
        else:
            new_path = None
            
            try:
                index = int(choice) - 1
                if 0 <= index < len(folders):
                    target_folder = folders[index]
                    new_path = os.path.join(current_path, target_folder)
            except ValueError:
                new_path = os.path.join(current_path, choice)
            
            if new_path and os.path.isdir(new_path):
                current_path = new_path
            else:
                input("Invalid selection. Press Enter to continue...")


def run_merge_process(start_path):
    """Handles the user interaction for selecting a folder and naming the output file."""
    print("Launching folder navigator...")
    selected_folder = text_file_navigator(start_path)
    
    if selected_folder:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("=" * 50)
        print(f"Selected Folder for Merge: {selected_folder}")
        print("=" * 50)

        while True:
            filename = input("Please, type the filename for the merged PDF (e.g., final_report.pdf): ").strip()
            if filename.lower().endswith('.pdf'):
                break
            elif filename:
                filename = filename + ".pdf"
                break
            else:
                print("Filename cannot be empty.")
        
        merge_pdf_in_folder(selected_folder, filename)
    else:
        print("Merge process cancelled.")
    
    input("\nPress Enter to return to the Main Menu...")


def set_home_folder_cli(current_home):
    """Prompts the user for a new home folder path in the terminal."""
    os.system('cls' if os.name == 'nt' else 'clear')
    print("=" * 50)
    print("🏠 SET HOME FOLDER")
    print("=" * 50)
    print(f"Current Home Folder: {current_home}")
    print("\nTo set the new Home Folder, please paste the full directory path.")
    print("-" * 50)
    
    new_path = input("Enter new Home Folder path (or press Enter to cancel): ").strip()
    
    if new_path:
        new_path = os.path.abspath(new_path)
        if os.path.isdir(new_path):
            save_config(new_path)
            return new_path
        else:
            print("❌ Invalid path entered. Folder does not exist.")
            input("Press Enter to continue...")
    return current_home

# --- MAIN CLI LOOP ---

def main_menu():
    """The main command-line interface menu loop."""
    home_folder = load_config()
    
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("=" * 50)
        print("📄 PDF UTILITY MAIN MENU")
        print("=" * 50)
        print(f"🏠 Home Folder: {home_folder}")
        print("-" * 50)
        print("1. Set Home Folder (Updates starting directory)")
        print("2. Merge PDF (Start folder navigation)")
        print("3. Password Protect Files (PDF & Excel)")
        print("Q. Quit")
        print("-" * 50)
        
        choice = input("Enter your choice: ").strip()

        if choice == '1':
            new_home = set_home_folder_cli(home_folder)
            home_folder = new_home
            
        elif choice == '2':
            run_merge_process(home_folder)
            
        elif choice == '3':
            run_password_protection(home_folder)
            
        elif choice.upper() == 'Q':
            print("Exiting PDF Utility. Goodbye!")
            break
            
        else:
            input("Invalid choice. Press Enter to try again...")

if __name__ == '__main__':
    try:
        main_menu()
    except KeyboardInterrupt:
        print("\nProgram interrupted. Exiting...")