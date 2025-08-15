#!/usr/bin/env python3

import os
from pathlib import Path
from win32com.shell import shell, shellcon
import shutil

#define file formats
docs = (".txt", ".pdf", ".doc", ".docx", ".odt", ".xls", ".xlsx", ".ppt", ".pptx", ".rtf", ".ipynb", ".tex", ".csv")
archives = (".zip", ".rar")
vid = (".mp4", ".avi", ".mkv", ".mov", ".flv", ".wmv", ".webm")
pic = (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".svg", ".webp")
mus = (".mp3", ".wav", ".flac", ".aac", ".ogg", ".m4a")

#get folder locations
downloads = shell.SHGetKnownFolderPath(shellcon.FOLDERID_Downloads)
documents = shell.SHGetKnownFolderPath(shellcon.FOLDERID_Documents)
videos = shell.SHGetKnownFolderPath(shellcon.FOLDERID_Videos)
pictures = shell.SHGetKnownFolderPath(shellcon.FOLDERID_Pictures)
music = shell.SHGetKnownFolderPath(shellcon.FOLDERID_Music)

#create folders for the files from downloads
Path(documents + "/downloaded_docs").mkdir(exist_ok=True)
Path(videos + "/downloaded_videos").mkdir(exist_ok=True)
Path(pictures + "/downloaded_pictures").mkdir(exist_ok=True)
Path(music + "/downloaded_music").mkdir(exist_ok=True)
Path(documents + "downloaded_archives").mkdir(exist_ok=True)

#move files
os.chdir(downloads)
for file in os.listdir():
    name, ext = os.path.splitext(file)
    if ext in docs:
        shutil.move(file, documents + "/downloaded_docs")
    if ext in archives:
        shutil.move(file, documents + "/downloaded_archives")
    if ext in vid:
        shutil.move(file, videos + "/downloaded_videos")
    if ext in pic:
        shutil.move(file, pictures + "/downloaded_pictures")
    if ext in mus:
        shutil.move(file, music + "/downloaded_music")