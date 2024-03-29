import os
from os import listdir, getenv, mkdir, walk
from os.path import isfile, join, exists, isdir, basename, islink, getsize
import platform
import shutil
import mimetypes
from time import sleep
from pathlib import Path
from datetime import datetime
import subprocess
import tkinter as tk
from tkinter import *
from tkinter import messagebox, Checkbutton, OptionMenu, ttk
import re

customtkinter.set_appearance_mode("System")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

top = Tk()

app = customtkinter.CTk()  # create CTk window like you do with the Tk window
app.geometry("400x240")

frame = ttk.Frame(top)
frame.pack(expand=True, fill=tk.BOTH)

launchOnStartupVar = IntVar()

def setLaunchOnStartup():
    global launchOnStartup
    launchOnStartup = launchOnStartupVar.get()
    print("Launch on startup:", bool(launchOnStartup))
    
c1 = customtkinter.CTkButton(top, text='Launch On Startup?', variable=launchOnStartupVar, onvalue=1, offvalue=0, command=setLaunchOnStartup)
c1.pack()

repeatOptions = ['1 hour', '2 hours', '3 hours', '4 hours', '5 hours', '6 hours', '7 hours', '8 hours']
selectedRepeat = StringVar(top)
selectedRepeat.set("Repeat Every")

def on_selection_change(*args):
    hours = int(re.search(r'\d+', selectedRepeat.get()).group())
    seconds = hours * 3600
    return seconds

question_menu = ttk.OptionMenu(top, selectedRepeat, *repeatOptions)
question_menu.pack()



def getDownloadsDirectory():
    """Returns the default downloads path for Linux, macOS, and Windows."""
    if platform.system() == 'Windows':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
            os.environ["DOWNLOADS_PATH"] = location
        return location
    elif platform.system() == 'Darwin':
        return os.path.join(os.path.expanduser('~'), 'Downloads')
    else:
        return os.path.join(os.path.expanduser('~'), 'Downloads')
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
def moveFolder(folder, downloads_path):
    parent_folder_sort = rf'{downloads_path}\_FOLDERS'
    attempt_insertion = join(parent_folder_sort, basename(folder))

    if not exists(attempt_insertion):
        shutil.move(folder, parent_folder_sort)

    else:
        timestamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
        modified_folder_name = f'{attempt_insertion} - {timestamp}'
    
        shutil.move(folder, modified_folder_name)

def revertSorting():
    msg_box = tk.messagebox.askquestion('Revert Sorting', 'Are you sure you want to disorganize your downloads folder? (You can always re-sort it)',
                                        icon='warning')
    if msg_box == 'yes':
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

        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        try:
            subprocess.Popen(["taskkill", "/im", "DownloadsSimplified.exe", '/f'], startupinfo=si)
        finally:
                pass
        try:
            subprocess.Popen(["taskkill", "/im", "DownloadsSimplified-nostartup.exe", '/f'], startupinfo=si)
        finally:
                pass

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
    else:
        tk.messagebox.showinfo('Return', 'You will now return to the application screen')

button1 = ttk.Button(top, text='Exit Application', command=revertSorting)
button1.pack()

def main():
    global launchOnStartup
    # to autoscan each organized folder and move any files that aren't where they should be 
    # (this would occur if the user places something in the wrong folder manually)
    def autoScan(folder, mimetype):
        files = [f for f in listdir(folder) if isfile(join(folder, f))]
        
        # need this to be a tuple so that it checks every element in the mimetypes lists
        mimetype = tuple(mimetype)
        
        for file in files:
            global source 
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


    def getDownloadsDirectory():
        """Returns the default downloads path for Linux, macOS, and Windows."""
        if platform.system() == 'Windows':
            import winreg
            sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
            downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
                location = winreg.QueryValueEx(key, downloads_guid)[0]
                os.environ["DOWNLOADS_PATH"] = location
            return location
        elif platform.system() == 'Darwin':
            return os.path.join(os.path.expanduser('~'), 'Downloads')
        else:
            return os.path.join(os.path.expanduser('~'), 'Downloads')

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

    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    
    # user cannot run both at same time
    # try:
    #     subprocess.Popen(["taskkill", "/im", "DownloadsSimplified-nostartup.exe", '/f'], startupinfo=si)
    # finally:
    #     pass
    
    # get startup path to remove the program from startup
    program_name = 'DownloadsSimplified.exe'
    appdata_path = os.getenv('APPDATA')
    default_startup = join(appdata_path, 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    startup_with_program = join(default_startup, program_name)
    # nest script in startup
    if bool(launchOnStartupVar.get()) and not exists(startup_with_program):
        shutil.copy2('./DownloadsSimplified.exe', startup_with_program)
    elif not bool(launchOnStartupVar.get()) and exists(startup_with_program):
        os.remove(join('./DownloadsSimplified.exe', startup_with_program))

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
            moveFolder(source_folder, downloads_path)

button2 = ttk.Button(top, text='Sort Downloads', command=main)
button2.pack()


button1.pack()

top.mainloop()


