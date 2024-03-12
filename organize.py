from os import listdir, getenv, mkdir, walk
from os.path import isfile, join, exists, isdir, basename, islink, getsize
import shutil
import mimetypes
from time import sleep
from pathlib import Path
from datetime import datetime


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


def moveFolder(folder, parent_path):
    parent_folder_sort = f'{downloads_path}\_FOLDERS'
    attempt_insertion = join(parent_folder_sort, basename(folder))

    if not exists(attempt_insertion):
        shutil.move(folder, parent_folder_sort)

    else:
        timestamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
        modified_folder_name = f'{attempt_insertion} - {timestamp}'
    
        shutil.move(folder, modified_folder_name)

def get_size(start_path):
    total_size = 0
    for dirpath, dirnames, filenames in walk(start_path):
        for f in filenames:
            fp = join(dirpath, f)
            # skip if it is symbolic link
            if not islink(fp):
                total_size += getsize(fp)

    return total_size

while True:

    #**************************************************************************************************************************
    # *********UNCOMMENT NEXT 6 LINES (UP TO AND INCLUDING: 'shutil.copy2('./Downloads Simplified.exe', startup_path)') FOR CREATING EXE
    # home_path = str(Path.home())
    # startup_path = join(home_path, 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Downloads Simplified.exe')
    
    # # nest script in startup
    # if not exists(startup_path):
    #     shutil.copy2('./Downloads Simplified.exe', startup_path)

    # default file path for downloads
    downloads_path = str(Path.home() / "Downloads")

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

    paths = [audio_path, video_path, image_path, zip_path, text_path, app_path, misc_path, pdf_path, work_path, folders_path]

    for path in paths:
        if not exists(path):
            mkdir(path)

    # all the types of zipped files to be stored in zipped
    zip_mimetypes = ['application/vnd.rar', 'application/x-rar-compressed', 'application/octet-stream',
                'application/zip', 'application/octet-stream', 'application/x-zip-compressed', 'multipart/x-zip',
                'application/vnd.cncf.helm.chart.content.v1.tar+gzip']
    
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

    # get each file in downloads
    files = [f for f in listdir(downloads_path) if isfile(join(downloads_path, f))]

    # add each file to the directory according to its mime type
    for file in files:

        source = join(downloads_path, file) # set the source as the current file in downloads

        file_type = mimetypes.guess_type(file)[0] # get the mime type of the current file

        if file_type and file_type in office_mimetypes:
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

    sleep(21600)
