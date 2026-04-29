from pathlib import Path
import os
import sys
import platform
import subprocess

current_os = platform.system()
convertedFiles = []


def convertFile(powerpoint, file: Path, current_os=current_os):
    if (current_os=="Windows"):
        import win32com.client
        from msvcrt import getch
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
    elif (current_os=="Linux"):
        output_path = file.with_suffix(".pdf")
        try:
            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", str(file), "--outdir", str(output_path.parent)], check=True)
            if output_path.exists():
                convertedFiles.append(file)
                print(f'Saved: "{output_path}"')
            else:
                print(f'\nError: PDF not created for "{file}".\n')
        except Exception as e:
            print(f'\nError converting "{file}": {e}\n')


def ppt_to_pdf(paths, current_os=current_os):
    if (current_os=="Windows"):
        import win32com.client
        from msvcrt import getch
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
        print("Converting to pdf...")
        for p in paths:
            path = Path(p)
            if path.is_dir():
                for file in path.iterdir():
                    if file.is_file() and file.suffix.lower() in (".ppt", ".pptx"):
                        convertFile(None, file, current_os)

            elif path.is_file() and path.suffix.lower() in (".ppt", ".pptx"):
                convertFile(None, path, current_os)
        print("Done.")
    else:
        print(f"Unsupported OS: {current_os}. Only Windows and Linux are supported.")


if __name__ == "__main__":
    inputPath = sys.argv[1:]

    if not inputPath:
        print(
            "Enter the path to a file or directory (all Powerpoint files in the directory will be converted to pdf):"
        )
        inputPath = [input().replace('"', "")]

    ppt_to_pdf(inputPath)
    print("\nPress enter to exit...")
    input()
