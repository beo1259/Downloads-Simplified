from os import listdir, getenv, mkdir
from os.path import isfile, join, exists, isdir, basename
import shutil
import mimetypes
from time import sleep
from pathlib import Path
from datetime import datetime

downloads_path = str(Path.home() / "Downloads")

app_path = join(downloads_path, 'Apps')
work_path = join(downloads_path, 'Work')

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


files = [f for f in listdir(app_path) if isfile(join(app_path, f))]

for file in files:
    source = join(app_path, file) # set the source as the current file in downloads

    file_type = mimetypes.guess_type(file)[0] # get the mime type of the current file
    print(file_type)
    if file_type in office_mimetypes:
        print('true')
        shutil.move(source, work_path)
        continue
