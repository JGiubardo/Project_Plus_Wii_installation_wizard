import os
import sys
from tkinter import messagebox, Tk, Button
from psutil import disk_usage, disk_partitions
from py7zr import SevenZipFile
from win32api import GetLogicalDriveStrings

MAX_DRIVE_SIZE = 32 * 1024 * 1024 * 1024
if getattr(sys, 'frozen', False):
    P_PLUS_ZIP = file = os.path.join(sys._MEIPASS, 'files\\PPlus2.3.2.7z')
else:
    P_PLUS_ZIP = 'PPlus2.3.2.7z'
# P_PLUS_ZIP = 'test.7z'
REQUIRED_FREE_SPACE = 1766703104  # size in bytes of the extracted zip
ALLOWED_FILE_SYSTEMS = {"FAT32", "FAT", "FAT16"}


class BadLocation(Exception):
    pass


def select_drive(ignore_problems=False):
    drive_selector_gui(ignore_problems)
    try:
        path = drive
        print(path)
    except NameError:
        messagebox.showerror("Error", "Drive not selected")
        sys.exit()
    if not ignore_problems:
        try:
            check_for_problems(path)
        except BadLocation as e:
            messagebox.showerror("Error", e.args[0])
            sys.exit()
    return path


def drive_selected(gui, d):
    gui.destroy()
    global drive
    drive = d


def drive_selector_gui(ignore_problems):
    gui = Tk()
    gui.title("Select Drive")
    drives = get_drives(ignore_problems)
    font = ('Consolas', 20, 'bold')
    for i, drv in enumerate(drives):
        Button(text=drv, font=font, width=5, command=lambda d=drv: drive_selected(gui, d)).grid(row=i // 5,
                                                                                                column=i % 5,
                                                                                                padx=8, pady=5)
    gui.mainloop()


def get_drives(ignore_problems):
    if ignore_problems:
        drives = [x[:3] for x in GetLogicalDriveStrings().split('\x00')[:-1]]
    else:
        drives = get_eligible_drives()
        if len(drives) == 0:
            messagebox.showwarning("Warning", "No eligible drives found, now showing all drives")
            drives = get_drives(True)
    return drives


def drive_too_big(path):
    return MAX_DRIVE_SIZE < disk_usage(path).total


def wont_fit(path):
    return REQUIRED_FREE_SPACE > disk_usage(path).free


def wrong_filesystem(filesystem):
    return filesystem not in ALLOWED_FILE_SYSTEMS


def get_eligible_drives():
    drives = []
    for part in disk_partitions():
        path = part.mountpoint
        if not wrong_filesystem(part.fstype) and not drive_too_big(path) and not wont_fit(path):
            drives.append(path)
    return drives


def check_for_problems(path):
    if drive_too_big(path):
        raise BadLocation("Drive is too big. SD card should be 32GB or smaller!")
    if wont_fit(path):
        raise BadLocation("The mod needs more space, remove items from your SD or use a different one")
    check_file_system(path)


def check_file_system(path):
    for part in disk_partitions():
        if part.device.startswith(path):
            if wrong_filesystem(part.fstype):
                raise BadLocation("Wrong file system, format the drive as FAT32 or use a different one")
            else:
                break


def extract_to_drive(drive):
    print("Extracting...")
    with SevenZipFile(P_PLUS_ZIP, 'r') as zip:
        zip.extractall(drive)


if __name__ == '__main__':
    drive = select_drive()
    extract_to_drive(drive)
    messagebox.showinfo("Complete",
                        "Mod extracted, place SD in console and boot through stage builder or homebrew channel. An "
                        "NTSC Brawl disc or backup is required to play")
