from tkinter import messagebox, Tk, Button
from win32api import GetLogicalDriveStrings
from py7zr import SevenZipFile
from psutil import disk_usage, disk_partitions
import sys, os

MAX_DRIVE_SIZE = 32 * 1024 * 1024 * 1024
if getattr(sys, 'frozen', False):
    P_PLUS_ZIP = file=os.path.join(sys._MEIPASS, 'files\\PPlus2.3.2.7z')
else:
    P_PLUS_ZIP = 'PPlus2.3.2.7z'
# P_PLUS_ZIP = 'test.7z'
REQUIRED_FREE_SPACE = 1766703104  # size in bytes of the extracted zip
ALLOWED_FILE_SYSTEMS = {"FAT32", "FAT", "FAT16"}


class BadLocation(Exception):
    pass


def select_drive(ignore_problems=False):
    drive_selector_gui()
    try:
        path = drive + "\\"
        print(path)
    except NameError as e:
        messagebox.showerror("Error", "Drive not selected")
        exit()
    if not ignore_problems:
        try:
            check_for_problems(path)
        except BadLocation as e:
            messagebox.showerror("Error", e.args[0])
            exit()
    return path


def drive_selected(gui, d):
    gui.destroy()
    global drive
    drive = d


def drive_selector_gui():
    gui = Tk()
    gui.title("Select Drive")
    drives = [x[:2] for x in GetLogicalDriveStrings().split('\x00')[:-1]]
    font = ('Consolas', 18, 'bold')
    for i, drv in enumerate(drives):
        Button(text=drv, font=font, width=5, command=lambda d=drv: drive_selected(gui, d)).grid(row=i // 5,
                                                                                                column=i % 5,
                                                                                                padx=5, pady=3)
    gui.mainloop()


def check_for_problems(path):
    if MAX_DRIVE_SIZE < disk_usage(path).total:
        raise BadLocation("Drive is too big. SD card should be 32GB or smaller!")
    if REQUIRED_FREE_SPACE > disk_usage(path).free:
        raise BadLocation("The mod needs more space, remove items from your SD or use a different one")
    check_file_system(path)


def check_file_system(path):
    for part in disk_partitions():
        if part.device.startswith(path):
            if part.fstype not in ALLOWED_FILE_SYSTEMS:
                raise BadLocation("Wrong file system, format the drive or use a different one")
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
                        "Mod extracted, place SD in console and boot through stage builder or homebrew channel. A "
                        "Brawl disc or backup is required to play")

