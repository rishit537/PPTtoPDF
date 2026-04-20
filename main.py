from pathlib import Path
import os
import sys
import win32com.client
from msvcrt import getch


convertedFiles = []


def convertFile(powerpoint, file: Path):
    output_path = file.with_suffix(".pdf")
    presentation = powerpoint.Presentations.Open(str(file))
    try:
        presentation.SaveAs(str(output_path), 32)
        convertedFiles.append(file)
    except Exception as e:
        print(f"Error: {e}")
    presentation.Close()
    print(f'Saved: "{output_path}"')


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
    except Exception as e:
        print(f"Error: {e}")
    finally:
        try:
            powerpoint.Quit()
        except Exception as e:
            pass
    print("Delete PPT/PPTX file(s) for converted PDF files? (y/n)")
    confirmDelete = getch().decode().lower()
    if confirmDelete.lower() == "y":
        print("Cleaning up...")
        for file in convertedFiles:
            os.remove(str(file))
            print(f"Deleted: {file}")
    else:
        print("Skipping clean-up...")
    print("Done.")


if __name__ == "__main__":
    inputPath = sys.argv[1:]

    if not inputPath:
        print(
            "Enter the path to a file or directory (all Powerpoint files in the directory will be converted to pdf):"
        )
        inputPath = [Path(input().replace('"', ""))]

    ppt_to_pdf(inputPath)
    print("\nPress any key to exit...")
    getch()
