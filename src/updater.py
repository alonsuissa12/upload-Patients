import sys
import time
import requests
import os

"""
Args:
1 - old version (unused)
2 - new version
3 - download URL
4 - main app path
"""

_, old_ver, new_ver, download_url, main_app_path = sys.argv


def download_new_version(url, output_path):
    r = requests.get(url, stream=True)
    r.raise_for_status()  # check for web errors

    with open(output_path, "wb") as f:
        for chunk in r.iter_content(1024):
            if chunk:
                f.write(chunk)


def wait_for_file_release(path, retries=10, delay=0.2):
    for _ in range(retries):
        try:
            if os.path.exists(path):
                os.remove(path)
            return True
        except PermissionError:
            time.sleep(delay)
    return False


def update_main_app():
    temp_path = main_app_path + ".new"

    # 1. הורדה
    download_new_version(download_url, temp_path)

    # 2. מחיקה בטוחה — אם הקובץ עדיין נעול, נחכה שוב
    if not wait_for_file_release(main_app_path):
        print("❌ הקובץ הראשי נעול — לא ניתן לעדכן.")
        return

    # 3. החלפה
    os.rename(temp_path, main_app_path)
    print(f"✔ עדכון הושלם → גרסה {new_ver}")

    # 4. הפעלה מחדש
    try:
        os.startfile(main_app_path)
    except Exception as e:
        print("⚠ לא ניתן להפעיל מחדש:", e)


update_main_app()
