from pathlib import Path
import os
import sys
import platform

convertedFiles = []

current_os=platform.system()

if (current_os=="Windows"):
    import win32com.client
    from msvcrt import getch
    def convertFile(powerpoint, file: Path):
        output_path = file.with_suffix(".pdf")
        try:
            presentation = powerpoint.Presentations.Open(str(file))
        except Exception as e:
            print(f'\nError opening "{file}": {e}\n')
            return

        try:
            presentation.SaveAs(str(output_path), 32)  # 32 = PDF
            convertedFiles.append(file)
            print(f'Saved: "{output_path}"')
        except Exception as e:
            print(f'\nError converting "{file}": {e}\n')
        finally:
            presentation.Close()


    def ppt_to_pdf(paths):
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        print("Converting to pdf...")

        try:
            for p in paths:
                path = Path(p)
                if path.is_dir():
                    for file in path.iterdir():
                        if file.is_file() and file.suffix.lower() in (".ppt", ".pptx"):
                            convertFile(powerpoint, file)

                elif path.is_file() and path.suffix.lower() in (".ppt", ".pptx"):
                    convertFile(powerpoint, path)

        finally:
            try:
                powerpoint.Quit()
            except Exception:
                pass

        if len(convertedFiles) > 0:
            print("Delete PPT/PPTX file(s) for converted PDF files? (y/n)")
            confirmDelete = getch().decode().lower()
            if confirmDelete == "y":
                print("Cleaning up...")
                for file in convertedFiles:
                    try:
                        os.remove(str(file))
                        print(f"Deleted: {file}")
                    except Exception as e:
                        print(f'Error deleting "{file}": {e}')
            else:
                print("Skipping clean-up...")
            print("Done.")

elif (current_os=="Linux"):
    import subprocess
    def ppt_to_pdf_l(paths):
        for p in paths:
            path = Path(p)
            if not path.exists():
                print(f"Path does not exist: {path}")
                continue
                
            if path.is_dir():
                print(f"Processing directory: {path}")
                for file in path.iterdir():
                    if file.is_file() and file.suffix.lower() in (".ppt", ".pptx"):
                        convert_single_file_l(file)
            elif path.is_file() and path.suffix.lower() in (".ppt", ".pptx"):
                convert_single_file_l(path)
            else:
                print(f"Skipping: {path} (not a PowerPoint file or directory)")

    def convert_single_file_l(file_path):
        print(f"Converting: {file_path}")
        try:
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(file_path.parent), str(file_path)],
                check=True,
                capture_output=True,
                text=True
            )
            print(f"Successfully converted: {file_path}")
        except Exception as e:
            print(f"An unexpected error occurred while converting {file_path}: {e}")

if __name__ == "__main__":
    
    inputPath = sys.argv[1:]

    if not inputPath:
        print("Enter path for a file. No input path detected.")
        inputPath = [input().strip().strip("\"")]
    
    if current_os == "Linux":
        ppt_to_pdf_l(inputPath)
        print("Done.")
    elif current_os == "Windows":
        if not inputPath:
            print("Enter the path to a file or directory (all Powerpoint files in the directory will be converted to pdf):")
            inputPath = [input().replace('"', "")]
        ppt_to_pdf(inputPath)
        print("\nPress any key to exit...")
        getch()