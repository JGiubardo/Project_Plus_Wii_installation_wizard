import os
import sys
from tkinter import messagebox, Tk, OptionMenu, Button, font, StringVar, Label, BOTTOM
from psutil import disk_usage, disk_partitions
from py7zr import SevenZipFile
from win32api import GetLogicalDriveStrings
from win32file import GetDriveType

VERSION_NUMBER = "v0.3.0"
MAX_DRIVE_SIZE = 32 * 1024 * 1024 * 1024

if getattr(sys, 'frozen', False):
    P_PLUS_ZIP = os.path.join(sys._MEIPASS, 'files\\PPlus2.3.2.7z')
    PLUS_ICON = os.path.join(sys._MEIPASS, 'files\\pplus.ico')
else:
    P_PLUS_ZIP = 'PPlus2.3.2.7z'
    PLUS_ICON = 'pplus.ico'
"""
P_PLUS_ZIP = 'test.7z'   # lets you test the application with a test file to eliminate time to extract to the SD
PLUS_ICON = 'pplus.ico'
"""
REQUIRED_FREE_SPACE = 1766703104  # size in bytes of the extracted zip
ALLOWED_FILE_SYSTEMS = {"FAT32", "FAT", "FAT16"}
REMOVABLE_DRIVE_TYPE = 2   # Get_Drive_Type returns 2 if the drive is removable


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


def drive_selector_gui_old(ignore_problems):
    gui = Tk()
    gui.title("Select Drive")
    drives = get_drives(ignore_problems)
    font = ('Consolas', 20, 'bold')
    for i, drv in enumerate(drives):
        Button(text=drv, font=font, width=5, command=lambda d=drv: drive_selected(gui, d)).grid(row=i // 5,
                                                                                                column=i % 5,
                                                                                                padx=8, pady=5)
    gui.mainloop()


def drive_selector_gui(ignore_problems):
    gui = Tk()
    gui.title("Select Drive")
    gui.iconbitmap(PLUS_ICON)
    drives = get_drives(ignore_problems)
    gui.geometry('300x300')
    value_inside = StringVar(gui)
    value_inside.set("Select a drive")
    segoe_font = font.Font(family='Segoe UI', size=18)
    menu = OptionMenu(gui, value_inside, *drives)
    widget = gui.nametowidget(menu.menuname)
    menu.config(font=segoe_font, width=40)
    widget.config(font=segoe_font)
    menu.pack()
    gui.focus_force()
    drive_info_text = StringVar(gui)
    drive_info_text.set("No drive selected")
    textbox = Label(gui, textvariable=drive_info_text, font=segoe_font, width=20, height=4)
    textbox.pack()
    select_button = Button(gui, text="Select", font=segoe_font, width=5,
                           command=lambda: drive_selected(gui, value_inside.get()))
    select_button.pack(side=BOTTOM, pady=5)
    value_inside.trace('w', lambda *args: display_drive_info(drive_info_text, value_inside.get()))
    gui.mainloop()


def gigabyte_string(size) -> str:
    gbs = size / (1024*1024*1024)
    return f"{gbs:.2f}"


def display_drive_info(drive_info_text: StringVar, drive_selected):
    size, free_space, filesystem, drive_type = drive_info(drive_selected)
    size_in_gb = gigabyte_string(size)
    space_in_gb = gigabyte_string(free_space)
    if drive_type == REMOVABLE_DRIVE_TYPE:
        type_text = "Drive is removable"
    else:
        type_text = "Drive is not removable"
    drive_info_text.set(f"Total Size: {size_in_gb} GB\n"
                        f"Free Space: {space_in_gb} GB\n"
                        f"Filesystem: {filesystem}\n"
                        f"{type_text}")


def drive_info(path):
    for part in disk_partitions():
        if part.device.startswith(path):
            filesystem = part.fstype
            break
    return disk_usage(path).total, disk_usage(path).free, filesystem, GetDriveType(path)


def get_drives(ignore_problems):
    if ignore_problems:
        drives = [x[:3] for x in GetLogicalDriveStrings().split('\x00')[:-1]]
    else:
        drives = get_eligible_drives()
        if len(drives) == 0:
            messagebox.showwarning("Warning", "No eligible drives found, now showing all drives")
            drives = get_drives(True)
    return drives


def drive_too_big(path) -> bool:
    return MAX_DRIVE_SIZE < disk_usage(path).total


def wont_fit(path) -> bool:
    return REQUIRED_FREE_SPACE > disk_usage(path).free


def wrong_filesystem(filesystem) -> bool:
    return filesystem not in ALLOWED_FILE_SYSTEMS


def drive_not_removable(path) -> bool:
    return GetDriveType(path) != REMOVABLE_DRIVE_TYPE


def get_eligible_drives() -> list:
    drives = []
    for part in disk_partitions():
        path = part.mountpoint
        if not wrong_filesystem(part.fstype) and not drive_too_big(path) and \
                not wont_fit(path) and not drive_not_removable(path):
            drives.append(path)
    return drives


def check_for_problems(path):
    if drive_too_big(path):
        raise BadLocation("Drive is too big. SD card should be 32GB or smaller!")
    if wont_fit(path):
        raise BadLocation("The mod needs more space, remove items from your SD or use a different one.")
    if drive_not_removable(path):
        raise BadLocation("Drive isn't a removable device. Should be installed on an SD card.")
    check_file_system(path)


def check_file_system(path):
    for part in disk_partitions():
        if part.device.startswith(path):
            if wrong_filesystem(part.fstype):
                raise BadLocation("Wrong file system, format the drive as FAT32 or use a different one.")
            else:
                break


def extract_to_drive(drive):
    print("Extracting...")
    with SevenZipFile(P_PLUS_ZIP, 'r') as zip:
        zip.extractall(drive)


def welcome():
    print(f"Project+ Wii Installation Wizard {VERSION_NUMBER}")
    root = Tk()
    root.overrideredirect(True)
    root.withdraw()
    messagebox.showinfo("Project+ Wii Installation Wizard", "Please make sure your SD card is unlocked "
                                                            "and inserted into your computer.")
    root.destroy()


if __name__ == '__main__':
    welcome()
    drive = select_drive()
    extract_to_drive(drive)
    messagebox.showinfo("Complete",
                        "Mod extracted; place SD in console and boot through stage builder or homebrew channel. An "
                        "NTSC Brawl disc or backup is required to play.")
    sys.exit()
