import os
import subprocess
import sys
import requests
from packaging import version

VERSION_URL = "https://raw.githubusercontent.com/alonsuissa12/upload-Patients/master/dist/local_version_macabi.txt"
LATEST_API = "https://api.github.com/repos/alonsuissa12/upload-Patients/releases/latest"

def get_local_version(version_file):
    if not os.path.exists(version_file):
        return "0.0.0"
    with open(version_file, "r") as f:
        return f.read().strip()


def write_local_version(version_file, new_version):
    with open(version_file, "w") as f:
        f.write(new_version)


def get_latest_url():
    data = requests.get(LATEST_API).json()
    for asset in data["assets"]:
        if asset["name"] == "clalit.exe":
            return asset["browser_download_url"]
    return None

def main():
    base = os.path.dirname(sys.argv[0])
    version_file = os.path.join(base, "local_version_macabi.txt")
    local_version = get_local_version(version_file)

    print("looking for updates...")
    remote_ver = requests.get(VERSION_URL).text.strip()

    base = os.path.dirname(sys.argv[0])
    updater = os.path.join(base, "updater.exe")
    main_app = os.path.join(base, "clalit.exe")

    if version.parse(remote_ver) > version.parse(local_version):
        print("🔔 Update found:", remote_ver)

        url = get_latest_url()

        subprocess.run([updater, main_app, url], check=True)

        write_local_version(version_file, remote_ver)


    subprocess.Popen([main_app])

if __name__ == "__main__":
    main()
