import requests
import subprocess
import sys
import os
from packaging import version

# local version
LOCAL_VERSION = "1.0.0"

# link to the version on github
VERSION_URL = "https://raw.githubusercontent.com/alonsuissa12/upload-Patients/master/version.txt"

# the GitHub releases API (does NOT contain the exe itself)
LATEST_API_URL = "https://api.github.com/repos/alonsuissa12/upload-Patients/releases/latest"


def get_latest_exe_url():
    # get the link for download
    response = requests.get(LATEST_API_URL, timeout=5)
    data = response.json()

    for asset in data["assets"]:
        if asset["name"].endswith(".exe"):
            return asset["browser_download_url"]

    return None


def check_for_update():
    try:
        # check the version
        remote_ver = requests.get(VERSION_URL, timeout=5).text.strip()

        if version.parse(remote_ver) > version.parse(LOCAL_VERSION):
            print(f"🔔 there is a new version: {remote_ver},downloading...")

            download_url = get_latest_exe_url()
            if not download_url:
                print("could not find EXE file")
                return

            # update the exe
            run_updater(remote_ver, download_url)
            sys.exit(0)
        else:
            print("the program is up to date.")

    except Exception as e:
        print("could not get the updates: ", e)


def run_updater(new_version, download_url):

    subprocess.Popen([
        sys.executable,
        "updater.py",
        LOCAL_VERSION,
        new_version,
        download_url,
        os.path.abspath(sys.argv[0])
    ])
    print("🔄 running updater...")
    sys.exit(0)


# --- start ---
check_for_update()

