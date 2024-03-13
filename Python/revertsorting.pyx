# cython: language_level=3

import os
from os import listdir, getenv, mkdir, walk
from os.path import isfile, join, exists, isdir, basename, islink, getsize
import platform
import shutil
import mimetypes
from time import sleep
from pathlib import Path
from datetime import datetime
from subprocess import check_output

# to move a file from downloads to its organized directory
def moveFile(file, parent_path):
    curPath = Path(file)
    file_stem = curPath.stem

    file_ext = curPath.suffix

    attempt_insertion = join(parent_path, file_stem  + file_ext)

    if not exists(attempt_insertion):
        shutil.move(file, parent_path)

    else:
        modified = f'{join(parent_path, file_stem)} - {datetime.now().strftime("%Y-%m-%d %H%M%S")}{file_ext}'
        shutil.move(source, modified)

# to move a folder from downloads to its organized directory
def moveFolder(folder, parent_path):
    parent_folder_sort = f'{downloads_path}\_FOLDERS'
    attempt_insertion = join(parent_folder_sort, basename(folder))

    if not exists(attempt_insertion):
        shutil.move(folder, parent_folder_sort)

    else:
        timestamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
        modified_folder_name = f'{attempt_insertion} - {timestamp}'
    
        shutil.move(folder, modified_folder_name)
        
def getDownloadsDirectory():
    if platform.system() == 'Windows':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
            os.environ["DOWNLOADS_PATH"] = location
        return location
    # macos and linux
    else:
        return os.path.join(os.path.expanduser('~'), 'Downloads')
    
def getPid(name):
    return check_output(["pidof", name])
    
# if we have already created the environment variable in the user's system, simply set downloads_path to that directory, otherwise find the directory
if "DOWNLOADS_PATH" in os.environ:
    downloads_path = str(os.environ.get("DOWNLOADS_PATH"))
else:
    downloads_path = getDownloadsDirectory()
    
# init the default paths for sorting and craete the folder if it does not exist
audio_path = join(downloads_path, 'Audio')
video_path = join(downloads_path, 'Video')
image_path = join(downloads_path, 'Images')
zip_path = join(downloads_path, 'Zipped')
text_path = join(downloads_path, 'Text')
app_path = join(downloads_path, 'Apps')
pdf_path = join(downloads_path, 'Pdf')
misc_path = join(downloads_path, 'Misc')
work_path = join(downloads_path, 'Work')    

folders_path = join(downloads_path, '_FOLDERS')

    
# paths for each organized folder
paths = [audio_path, video_path, image_path, zip_path, text_path, app_path, misc_path, pdf_path, work_path, folders_path]

# kill the task
os.system('taskkill /f /im DownloadsSimplified.exe')
os.system('taskkill /f /im DownloadsSimplified-nostartup.exe')

# get startup path to remove the program from startup
program_name = 'DownloadsSimplified.exe'
appdata_path = os.getenv('APPDATA')
default_startup = join(appdata_path, 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
startup_with_program = join(default_startup, program_name)

if exists(startup_with_program):
    os.remove(startup_with_program)

# go thru each path
for path in paths:
    files = [f for f in listdir(path) if isfile(join(path, f))]
    folders = [f for f in listdir(path) if isdir(join(path, f))]
    
    # move each file out of it's organized folder, into downloads
    for file in files:
        source = join(path, file)
        moveFile(source, downloads_path)
    
    # move each folder out of it's organized folder, into downloads
    for folder in folders:
        source = join(path, folder)
        moveFile(source, downloads_path)

# after moving all of the files out, we can 
for path in paths:
    if exists(path):
        os.rmdir(path)