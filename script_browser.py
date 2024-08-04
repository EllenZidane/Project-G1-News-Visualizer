from robocorp.tasks import task
import os
import subprocess
import requests
import zipfile
from pathlib import Path
import logging
import re
import winreg
import shutil

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def download_file(url, local_filename):
    try:
        with requests.get(url, stream=True) as response:
            response.raise_for_status()
            with open(local_filename, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
        logging.info(f"Downloaded {local_filename} successfully.")
    except requests.RequestException as e:
        logging.error(f"Error downloading {local_filename}: {e}")
        raise


def install_edge():
    logging.info("Downloading and installing Microsoft Edge...")
    edge_download_page = "https://www.microsoft.com/edge"

    # Fetch the page to extract the installer URL
    try:
        response = requests.get(edge_download_page)
        response.raise_for_status()
        installer_url = re.search(
            r"https://msedge\.sf\.dl\.delivery\.mp\.microsoft\.com/.*?/MicrosoftEdgeSetup\.exe",
            response.text,
        )
        if installer_url:
            installer_url = installer_url.group(0)
            installer_path = "MicrosoftEdgeSetup.exe"
            download_file(installer_url, installer_path)
            subprocess.run([installer_path, "/silent"], check=True)
            os.remove(installer_path)
            logging.info("Microsoft Edge installed successfully!")
        else:
            logging.error("Could not find the installer URL on the Edge download page.")
            raise Exception("Installer URL not found")
    except requests.RequestException as e:
        logging.error(f"Error accessing Edge download page: {e}")
        raise


def get_edge_version(path):
    try:
        try:
            key_path = r"SOFTWARE\Microsoft\Edge\BLBeacon"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                version, _ = winreg.QueryValueEx(key, "version")
                return version
        except FileNotFoundError:
            try:
                key_path = r"SOFTWARE\Microsoft\Edge\BLBeacon"
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    version, _ = winreg.QueryValueEx(key, "version")
                    return version
            except FileNotFoundError:
                return "Microsoft Edge not found in the registry."
    except Exception as e:
        return f"An unexpected error occurred: {e}"


def install_webdriver(edge_version):
    logging.info("Downloading and installing Microsoft Edge WebDriver...")
    webdriver_url = (
        f"https://msedgedriver.azureedge.net/{edge_version}/edgedriver_win64.zip"
    )
    zip_path = "edgedriver_win64.zip"
    extract_dir = "edgedriver"
    path = "C:/ProgramData"

    download_file(webdriver_url, zip_path)
    try:
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(extract_dir)
        logging.info("WebDriver extracted successfully.")

        # Move WebDriver to a directory in PATH
        webdriver_path = Path(extract_dir) / "msedgedriver.exe"
        destination_path = Path(path)
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(str(webdriver_path), str(destination_path))
        logging.info("WebDriver installed successfully!")
    except zipfile.BadZipFile as e:
        logging.error(f"Error extracting WebDriver: {e}")
        raise
    finally:
        os.remove(zip_path)


def get_webdriver_version():
    try:
        result = subprocess.run(
            ["msedgedriver", "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=True,
            text=True,
        )
        return result.stdout.strip()
    except FileNotFoundError:
        return "WebDriver not found."
    except subprocess.CalledProcessError as e:
        return f"Error getting WebDriver version: {e}"


def main():
    edge_paths = [
        "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
        "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
    ]

    webdriver_paths = [
        "C:/Program Files (x86)/Microsoft/Edge/Application/msedgedriver.exe",
        "C:/Program Files/WebDriver/msedgedriver.exe",
    ]

    # Check if Microsoft Edge is installed
    edge_paths_status = {}
    for file in edge_paths:
        if os.path.exists(file):
            edge_paths_status[file] = True
            path = file
        else:
            edge_paths_status[file] = False
    

    # Check if files do not exist and take action
    if not True in edge_paths_status.values():
        logging.info("Microsoft Edge not found, installing...")
        install_edge()

    if not path:
        path = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"

    edge_version = get_edge_version(path)
    logging.info(f"Microsoft Edge version: {edge_version}")

    # Check if WebDriver is installed
    webdriver_paths_status = {}
    for file in webdriver_paths:
        if not os.path.exists(file):
            webdriver_paths_status[file] = False
        else:
            webdriver_paths_status[file] = True

    # Check if files do not exist and take action
    if not True in webdriver_paths_status.values():
        webdriver_version = get_webdriver_version()
        logging.info(f"WebDriver version: {webdriver_version}")
        if "not found" in edge_version.lower():
            install_edge()
        install_webdriver(edge_version)


if __name__ == "__main__":
    main()
