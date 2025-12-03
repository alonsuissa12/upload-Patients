import sys
import time
import requests
import os

TARGET_EXE = sys.argv[1]
DOWNLOAD_URL = sys.argv[2]

def wait_for_close(path):
    for _ in range(30):
        try:
            os.remove(path)
            return True
        except PermissionError:
            time.sleep(0.2)
    return False

def download_new_file(url, path):
    r = requests.get(url, stream=True)
    r.raise_for_status()
    with open(path, "wb") as f:
        for chunk in r.iter_content(8192):
            f.write(chunk)

if not wait_for_close(TARGET_EXE):
    print("❌ file locked")
    sys.exit(1)

download_new_file(DOWNLOAD_URL, TARGET_EXE)

