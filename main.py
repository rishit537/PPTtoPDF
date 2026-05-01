from pathlib import Path
import sys
import platform
import subprocess

current_os = platform.system()
convertedFiles = []


def convertFile(powerpoint, file: Path, current_os=current_os):
    # For Windows, use Powerpoint to convert to PDF
    if current_os == "Windows":
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
            try:
                presentation.Close()
            except:
                pass

    # For Linux, use LibreOffice in headless mode to convert to PDF
    elif current_os == "Linux":
        output_path = file.with_suffix(".pdf")
        try:
            subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    str(file),
                    "--outdir",
                    str(output_path.parent),
                ],
                check=True,
            )
            if output_path.exists():
                convertedFiles.append(file)
                print(f'Saved: "{output_path}"')
            else:
                print(f'\nError: PDF not created for "{file}".\n')
        except Exception as e:
            print(f'\nError converting "{file}": {e}\n')


def ppt_to_pdf(paths, current_os=current_os):
    print("Converting to pdf...")

    # Open powerpoint if the os is Windows
    if current_os == "Windows":
        import win32com.client

        global powerpoint
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")

    # Convert each file to pdf
    try:
        for p in paths:
            path = Path(p)
            if path.is_dir():
                for (
                    file
                ) in path.iterdir():  # Iterate through the contents of the directory
                    if file.is_file() and file.suffix.lower() in (".ppt", ".pptx"):
                        convertFile(powerpoint=powerpoint, file=file)

            elif path.is_file() and path.suffix.lower() in (".ppt", ".pptx"):
                convertFile(powerpoint if current_os == "Windows" else None, path)

    # Attempt to close PPT
    finally:
        try:
            if current_os == "Windows":
                powerpoint.Quit()
        except Exception:
            pass

    # Provide the option to delete the original PPT file
    if len(convertedFiles) > 0:
        print("Delete PPT/PPTX file(s) for converted PDF files? (y/n)")
        confirmDelete = input().lower()
        if confirmDelete == "y":
            print("Cleaning up...")
            for file in convertedFiles:
                try:
                    file.unlink()
                    print(f'Deleted: "{file}"')
                except Exception as e:
                    print(f'Error deleting "{file}": {e}')
        else:
            print("Skipping clean-up...")
        print("Done.")


if __name__ == "__main__":
    inputPath = sys.argv[1:]

    if not inputPath:
        print(
            "Enter the path to a file or directory (all Powerpoint files in the directory will be converted to pdf):"
        )
        inputPath = [input().strip().strip('"').strip("'")]

    ppt_to_pdf(inputPath)
    print("\nPress enter to exit...")
    input()
