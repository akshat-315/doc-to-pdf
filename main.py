from fastapi import FastAPI, File, UploadFile
import os
import subprocess
import platform
from pathlib import Path
import shutil

app = FastAPI()

UPLOAD_FOLDER = "uploaded_docs"

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


def install_libreoffice():
    system = platform.system()
    
    if system == "Windows":
        # For Windows, download the installer and execute it
        installer_url = "https://download.documentfoundation.org/libreoffice/stable/7.6.2/win/x86_64/LibreOffice_7.6.2_Win_x64.msi"
        installer_path = Path("LibreOffice_Installer.msi")
        
        # Download the installer
        subprocess.run(["powershell", "-Command", f"Invoke-WebRequest -Uri {installer_url} -OutFile {installer_path}"], check=True)
        
        # Install LibreOffice silently
        subprocess.run(["msiexec", "/i", str(installer_path), "/quiet", "/norestart"], check=True)
        
        # Clean up
        installer_path.unlink(missing_ok=True)
    
    elif system == "Linux":
        # Install LibreOffice using package manager
        subprocess.run(["sudo", "apt-get", "update"], check=True)
        subprocess.run(["sudo", "apt-get", "install", "-y", "libreoffice"], check=True)
    
    elif system == "Darwin":  # macOS
        # For macOS, download the installer DMG and install
        installer_url = "https://download.documentfoundation.org/libreoffice/stable/7.6.2/mac/x86_64/LibreOffice_7.6.2_MacOS_x86-64.dmg"
        installer_path = Path("LibreOffice_Installer.dmg")
        
        # Download the installer
        subprocess.run(["curl", "-o", str(installer_path), installer_url], check=True)
        
        # Mount the DMG
        subprocess.run(["hdiutil", "attach", str(installer_path)], check=True)
        
        # Install LibreOffice
        subprocess.run(["sudo", "cp", "-R", "/Volumes/LibreOffice/LibreOffice.app", "/Applications"], check=True)
        
        # Unmount the DMG
        subprocess.run(["hdiutil", "detach", "/Volumes/LibreOffice"], check=True)
        
        # Clean up
        installer_path.unlink(missing_ok=True)
    
    else:
        raise OSError(f"Unsupported OS: {system}")


def get_libreoffice_path() -> str:
    if platform.system() == "Windows":
        potential_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        ]
        for path in potential_paths:
            if Path(path).exists():
                return path
        install_libreoffice()
        for path in potential_paths:
            if Path(path).exists():
                return path

    else:
        libreoffice_cmd = shutil.which("libreoffice") or shutil.which("soffice")
        if libreoffice_cmd:
            return libreoffice_cmd
        
        install_libreoffice()
        libreoffice_cmd = shutil.which("libreoffice") or shutil.which("soffice")
        if libreoffice_cmd:
            return libreoffice_cmd
    
    raise FileNotFoundError("LibreOffice installation failed or not found.")


def convert_to_pdf(input_file: str, output_folder: str) -> str:
    libreoffice_path = get_libreoffice_path()
    output_file = os.path.join(output_folder, f"{Path(input_file).stem}.pdf")

    command = [
        libreoffice_path,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_folder,
        input_file
    ]

    try:
        subprocess.run(command, check=True)
        if not Path(output_file).exists():
            raise FileNotFoundError(f"PDF not generated at {output_file}.")
        return output_file
    except subprocess.CalledProcessError as e:
        raise Exception(f"Error during conversion: {e}")


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    file_location = os.path.join(UPLOAD_FOLDER, file.filename)

    # Save the uploaded file
    with open(file_location, "wb") as f:
        f.write(await file.read())

    # Convert the file to PDF
    try:
        pdf_file = convert_to_pdf(file_location, UPLOAD_FOLDER)
        return {"info": f"File '{file.filename}' converted to PDF at '{pdf_file}'"}
    except Exception as e:
        return {"error": str(e)}
