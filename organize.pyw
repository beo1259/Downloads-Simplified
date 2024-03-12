import os
from os import listdir, getenv, mkdir, walk
from os.path import isfile, join, exists, isdir, basename, islink, getsize
import shutil
import mimetypes
from time import sleep
from pathlib import Path
from datetime import datetime
import string
import winreg

# to autoscan each organized folder and move any files that aren't where they should be 
# (this would occur if the user places something in the wrong folder manually)
def autoScan(folder, mimetype):
    files = [f for f in listdir(folder) if isfile(join(folder, f))]
    
    # need this to be a tuple so that it checks every element in the mimetypes lists
    mimetype = tuple(mimetype)
    
    for file in files:
        source = join(folder, file) # set the source as the current file in the given organized directory

        file_type = mimetypes.guess_type(file)[0] # get the mime type of the current file 
        
        # finds files that are not in the directory that their folder is supposed to hold
        if file_type and not file_type.startswith(mimetype):
            for mtypekey, mtypevalue in paths_and_their_mimetypes.items():
                # go if not a list of mimetypes (zips, office)
                if not isinstance(mtypevalue, list) and file_type.startswith(mtypevalue):
                    moveFile(source, join(mtypekey, file))
                    break
                    # for all the folders that are asisgned multiple mimetypes (are lists)
                elif isinstance(mtypevalue, list) and file_type.startswith(tuple(mtypevalue)):
                    if mimetype in mtypevalue:
                        moveFile(source, join(mtypekey, file))
                    break
        continue
    
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


# this is the deafult windows downloads direct
def defaultDownloads():
    return os.path.join(os.path.expanduser('~'), 'Downloads')

# on windows we can get the downloads folder easily
def getDownloadsFolderWindows():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders')
    downloads_path, _ = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')
    downloads_path = os.path.expandvars(downloads_path)  # expand any environment variables
    return downloads_path

if os.name == 'nt':  # if running on Windows
    downloads_path = getDownloadsFolderWindows()
    os.environ["DOWNLOADS_PATH"] = downloads_path


# if it isnt in the deafult directory and the user isn't on windows
def findDownloads(drives):
    for drive in drives:
        for root, dirs, files in os.walk(drive, topdown=True): 
            if 'Downloads' in dirs:
                downloads_path = os.path.join(root, 'Downloads')
                os.environ["DOWNLOADS_PATH"] = downloads_path
                return downloads_path  
    return None

# sort drives from largest to smallest (size)
def sortDrives(drive):
    valid_drives = []
    for each_drive in drive:
        if os.path.exists(each_drive+':\\'):
            valid_drives.append(each_drive+":\\")
    
    size_dict = {}

    for drive in valid_drives:
        du_tuple = shutil.disk_usage(drive)
        size_dict[drive] = du_tuple.used

    sorted_sizes = dict(sorted(size_dict.items(), key=lambda item: item[1], reverse=True))
    sorted_drive_arr = []

    for key in sorted_sizes.keys():
        sorted_drive_arr.append(key)

    return sorted_drive_arr

# main loop
while True:

    #**************************************************************************************************************************
    # *********UNCOMMENT NEXT 7 LINES (UP TO AND INCLUDING: 'shutil.copy2('./Downloads Simplified.exe', startup_path)') FOR CREATING EXE
    # nesting itself in startup directory
    home_path = str(Path.home())
    default_startup = 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\organize.pyw'
    startup_path = join(home_path, default_startup)
    
    # nest script in startup
    if not exists(startup_path):
        shutil.copy2('./organize.pyw', startup_path)

    # if we have already created the environment variable in the user's system, simply set downloads_path to that directory, otherwise find the directory
    if "DOWNLOADS_PATH" in os.environ:
        downloads_path = str(os.environ.get("DOWNLOADS_PATH"))
    else:
        downloads_path = str(getDownloadsFolderWindows())
        if os.path.exists(downloads_path):
            os.environ["DOWNLOADSPATH"] = downloads_path
        else:
            # get drives and sort them from largest to smallest (typically the larger drive will contain downloads)
            drive = string.ascii_uppercase
            sorted_drives = sortDrives(drive)
            findDownloads(sorted_drives)

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

    # all the types of zipped files to be stored in zipped
    zip_mimetypes = ['application/vnd.rar', 'application/x-rar-compressed', 'application/octet-stream',
                'application/zip', 'application/octet-stream', 'application/x-zip-compressed', 'multipart/x-zip',
                'application/vnd.cncf.helm.chart.content.v1.tar+gzip']
    
    # all the mimetypes of office folders
    office_mimetypes = ['application/msword',
                        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
                        'application/vnd.ms-word.document.macroEnabled.12',
                        'application/vnd.ms-word.template.macroEnabled.12',
                        'application/vnd.ms-excel',
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
                        'application/vnd.ms-excel.sheet.macroEnabled.12',
                        'application/vnd.ms-excel.template.macroEnabled.12',
                        'application/vnd.ms-excel.addin.macroEnabled.12',
                        'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
                        'application/vnd.ms-powerpoint',
                        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                        'application/vnd.openxmlformats-officedocument.presentationml.template',
                        'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
                        'application/vnd.ms-powerpoint.addin.macroEnabled.12',
                        'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
                        'application/vnd.ms-powerpoint.template.macroEnabled.12',
                        'application/vnd.ms-powerpoint.slideshow.macroEnabled.12'
                        ]

    # paths for each organized folder
    paths = [audio_path, video_path, image_path, zip_path, text_path, app_path, misc_path, pdf_path, work_path, folders_path]
    
    # each folder's mimetype
    paths_and_their_mimetypes = {
        audio_path:'audio',
        video_path:'video',
        image_path:'image',
        zip_path:zip_mimetypes,
        text_path:'text',
        pdf_path:['application/pdf', 'application/x-pdf'],
        app_path:'application',
        work_path:office_mimetypes,
    }

    # create directories if they don't exist
    for path in paths:
        if not exists(path):
            mkdir(path)

    # first go through each folder and check that they are organized
    for pathkey, pathval in paths_and_their_mimetypes.items():
        if not isinstance(pathval, list):
            typeArr = []
            typeArr.append(pathval)
            autoScan(pathkey, typeArr)
        else:
            autoScan(pathkey, pathval)
        

    # get each file in downloads
    files = [f for f in listdir(downloads_path) if isfile(join(downloads_path, f))]

    # add each file to the directory according to its mime type
    for file in files:

        source = join(downloads_path, file) # set the source as the current file in downloads

        file_type = mimetypes.guess_type(file)[0] # get the mime type of the current file

        if file_type and file_type in office_mimetypes:
            autoScan(work_path, office_mimetypes)
            moveFile(source, work_path)
            continue
        
        elif file_type and (file_type == 'application/pdf' or file_type == 'application/x-pdf'):
            moveFile(source, pdf_path)
            continue

        elif file_type and file_type.startswith('application') and file_type not in zip_mimetypes:            
            moveFile(source, app_path)
            continue

        elif file_type and file_type.startswith('audio'):
            moveFile(source, audio_path)
            autoScan(audio_path, ['audio'])
            continue

        elif file_type and file_type.startswith('image'):
            moveFile(source, image_path)
            continue

        elif file_type and file_type.startswith('text'):
            moveFile(source, text_path)
            continue

        if file_type and file_type.startswith('video'):
            moveFile(source, video_path)
            continue

        elif file_type in zip_mimetypes:
            moveFile(source, zip_path)
            continue
        
        else:
            moveFile(source, misc_path)

    # handling folders, moving them all into a directory
    folders = [f for f in listdir(downloads_path) if isdir(join(downloads_path, f))]

    for folder in folders:

        source_folder = join(downloads_path, folder)

        # make sure we dont try to remove the sort folders
        if source_folder not in paths:
            moveFolder(source_folder, folders_path)

    sleep(21600) # check to sort again 6 hours from now
