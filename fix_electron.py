import urllib.request
import zipfile
import os

url = "https://github.com/electron/electron/releases/download/v31.7.7/electron-v31.7.7-win32-x64.zip"
zip_path = "node_modules/electron/electron.zip"
extract_to = "node_modules/electron/dist"

print("Downloading Electron...")
urllib.request.urlretrieve(url, zip_path)
print("Extracting...")
with zipfile.ZipFile(zip_path, 'r') as z:
    z.extractall(extract_to)
os.remove(zip_path)
print("Done. Run: npx electron --version")