import os
import sys
import requests
import webbrowser
from xml.dom import minidom
from packaging import version
from tkinter import messagebox, Tk, OptionMenu, Button, font, StringVar, Label, BOTTOM
from psutil import disk_usage, disk_partitions
from py7zr import SevenZipFile
from win32api import GetLogicalDriveStrings
from win32file import GetDriveType
from shutil import rmtree

VERSION_NUMBER = "v0.4.0"
P_PLUS_VERSION_NUMBER = "2.4.0"
RELEASES_PAGE = "https://github.com/JGiubardo/Project_Plus_Wii_installation_wizard/releases/"
RELEASES_PAGE_API = "https://api.github.com/repos/JGiubardo/Project_Plus_Wii_installation_wizard/releases"
MAX_DRIVE_SIZE = 32 * 1024 * 1024 * 1024    # 32 GB in bytes
REQUIRED_FREE_SPACE = 1766703104    # size in bytes of the extracted zip
ALLOWED_FILE_SYSTEMS = {"FAT32", "FAT", "FAT16"}
REMOVABLE_DRIVE_TYPE = 2    # GetDriveType returns 2 if the drive is removable

if getattr(sys, 'frozen', False):       # The program is being run as a pyinstaller executable
    P_PLUS_ZIP = os.path.join(sys._MEIPASS, 'files", "PPlus2.4.0.7z')
    PLUS_ICON = os.path.join(sys._MEIPASS, 'files", "pplus.ico')
else:                                   # The program is being run as a standalone python file
    P_PLUS_ZIP = 'PPlus2.4.0.7z'
    PLUS_ICON = 'pplus.ico'
"""
P_PLUS_ZIP = 'test.7z'   # lets you test the application with a test file to eliminate time to extract to the SD
PLUS_ICON = 'pplus.ico'
"""


class BadLocation(Exception):
    pass


def check_installer_updates():
    try:
        response = requests.get(RELEASES_PAGE_API)
    except requests.exceptions.RequestException:  # if there's a problem, skip checking for an update
        return
    latest = response.json()[0]["tag_name"]
    if version.parse(latest) > version.parse(VERSION_NUMBER):
        root = Tk()
        root.overrideredirect(True)
        root.withdraw()
        result = messagebox.askyesno("Update available", "An update for the installer is available. "
                                                         "Would you like to go to the download page?")
        root.destroy()
        if result:
            webbrowser.open(RELEASES_PAGE)
            sys.exit()


def check_p_plus_updates(drive):
    path = os.path.join(drive, "apps", "projplus", "meta.xml")
    if p_plus_installed(drive):
        if os.path.exists(path):
            file = minidom.parse(path)
            installed_version = file.getElementsByTagName("version")[0].firstChild.wholeText
            if version.parse(installed_version) >= version.parse(P_PLUS_VERSION_NUMBER):
                ask_to_delete_or_skip(drive, "Project+ is up to date.")
            else:
                ask_to_delete_or_skip(drive, "Outdated P+ installation detected.")
        else:           # P+ is installed but can't determine version
            ask_to_delete_or_skip(drive, "Unknown P+ version detected.")


def p_plus_installed(drive) -> bool:
    folder_path = os.path.join(drive, "Project+", "")
    return os.path.exists(folder_path)


def delete_files(drive):
    rmtree(os.path.join(drive, "Project+", ""), True)
    rmtree(os.path.join(drive, "private", "wii", "app", "RSBE", ""), True)
    rmtree(os.path.join(drive, "apps", "projplus", ""), True)
    elf_path = os.path.join(drive, "boot.elf")
    if os.path.isfile(elf_path):
        os.remove(elf_path)


def ask_to_delete_or_skip(drive, current_status):
    """If P+ is already installed, gives the option to delete files or exit."""
    root = Tk()
    root.overrideredirect(True)
    root.withdraw()
    answer = messagebox.askokcancel("P+ installation found", f"{current_status} Delete all files and reinstall P+?")
    root.destroy()
    if answer:
        delete_files(drive)
    else:
        sys.exit()


def select_drive(ignore_problems=False) -> str:
    drive_selector_gui(ignore_problems)
    try:
        path = drive    # drive_selected has assigned the global variable
        print(path)
    except NameError:   # window was closed before a drive was selected
        root = Tk()
        root.overrideredirect(True)
        root.withdraw()
        messagebox.showerror("Error", "Drive not selected")
        root.destroy()
        sys.exit()
    if not ignore_problems:
        try:
            check_for_problems(path)
        except BadLocation as e:    # Incompatible drive
            root = Tk()
            root.overrideredirect(True)
            root.withdraw()
            messagebox.showerror("Error", e.args[0])
            root.destroy()
            sys.exit()
    return path


def drive_selected(gui, d):
    """After a drive is selected, make it a global drive and destroy the gui"""
    gui.destroy()
    global drive
    drive = d


def drive_selector_gui(ignore_problems):  # TODO add indicator in UI for problems
    gui = Tk()
    gui.title("Select Drive")
    gui.iconbitmap(PLUS_ICON)
    gui.geometry('300x300')
    drives = get_drives(ignore_problems)
    value_inside = StringVar(gui)
    value_inside.set("Select a drive")
    menu = OptionMenu(gui, value_inside, *drives)
    widget = gui.nametowidget(menu.menuname)

    # Formatting stuff
    segoe_font = font.Font(family='Segoe UI', size=18)
    menu.config(font=segoe_font, width=12)
    widget.config(font=segoe_font)
    menu.pack()
    gui.focus_force()

    drive_info_text = StringVar(gui)
    drive_info_text.set("No drive selected")
    textbox = Label(gui, textvariable=drive_info_text, font=segoe_font, width=20, height=5)
    textbox.pack()

    # When the button is pressed, send the selected drive
    select_button = Button(gui, text="Select", font=segoe_font, width=7,
                           command=lambda: drive_selected(gui, value_inside.get()))
    select_button.pack(side=BOTTOM, pady=5)
    value_inside.trace('w', lambda *args: display_drive_info(drive_info_text, value_inside.get()))
    gui.mainloop()


def gigabyte_string(size: int) -> str:
    """Takes the size of a drive in bytes and returns a String in the format of X.XX"""
    gbs = size / (1024 * 1024 * 1024)
    return f"{gbs:.2f}"


def display_drive_info(drive_info_text: StringVar, drive_selected):
    """Edits drive_info_text to show basic information about drive_selected"""
    size, free_space, filesystem, drive_type = drive_info(drive_selected)
    size_in_gb = gigabyte_string(size)
    space_in_gb = gigabyte_string(free_space)
    if drive_type == REMOVABLE_DRIVE_TYPE:
        type_text = "Drive is removable"
    else:
        type_text = "Drive is not removable"
    drive_info_text.set(f"Total Size: {size_in_gb} GB\n"
                        f"Free Space: {space_in_gb} GB\n"
                        f"Format: {filesystem}\n"
                        f"{type_text}\n"
                        )  # emoji test "\N{check mark}\N{heavy check mark}\N{cross mark}\N{prohibited sign}"


def drive_info(path) -> (int, int, str, any):
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
        if not drives:
            messagebox.showwarning("Warning", "No eligible drives found, now showing all drives")
            drives = get_drives(True)
    return drives


def drive_too_big(path) -> bool:
    return MAX_DRIVE_SIZE < disk_usage(path).total


def too_big_for_hackless(path) -> bool:
    return disk_usage(path).total > (2 * 1024 * 1024 * 1024)


def too_big_for_hackless_message(path) -> str:
    if too_big_for_hackless(path):
        return ""
    else:
        return "stage builder or "


def wont_fit(path) -> bool:
    return REQUIRED_FREE_SPACE > disk_usage(path).free


def wont_fit_ever(path) -> bool:
    return REQUIRED_FREE_SPACE > disk_usage(path).total


def wrong_filesystem(filesystem) -> bool:
    return filesystem not in ALLOWED_FILE_SYSTEMS


def drive_not_removable(path) -> bool:
    if os.name == "nt":         # Windows
        return drive_not_removable_windows(path)
    elif os.name == "posix":    # Linux
        return drive_not_removable_linux(path)
    else:
        return False


def drive_not_removable_windows(path) -> bool:
    return GetDriveType(path) != REMOVABLE_DRIVE_TYPE


def drive_not_removable_linux(path) -> bool:        # TODO pass the block device to this instead of the mount point
    # Check if the path is a block device.
    try:
        with open("/sys/class/block/{}/removable".format(path), "r") as f:
            removable = f.read().strip()
            f.close()
    except FileNotFoundError:
        return False

    # Return True if the path is a block device and the device is removable.
    return removable == "1"


def get_eligible_drives() -> list:
    drives = []
    for part in disk_partitions():
        path = part.mountpoint
        if not wrong_filesystem(part.fstype) and not drive_too_big(path) and \
                not wont_fit_ever(path) and not drive_not_removable(path):
            if wont_fit(path):
                if p_plus_installed(path):      # drive would be compatible if P+ was deleted
                    drives.append(path)
            else:
                drives.append(path)
    return drives


def check_for_problems(path):
    """Raises an Exception if there's an issue with the drive's compatibility with P+"""
    if drive_too_big(path):
        raise BadLocation("Drive is too big. SD card should be 32GB or smaller! Use a smaller SD.")
    if wont_fit_ever(path):
        raise BadLocation("The mod needs more space. Use a larger SD.")
    if wont_fit(path) and not p_plus_installed(path):
        raise BadLocation("The mod needs more space. Remove items from your SD or use a larger SD.")
    if drive_not_removable(path):
        raise BadLocation("Drive isn't a removable device. The mod should be installed on an SD card.")
    check_file_system(path)


def check_file_system(path):
    for part in disk_partitions():
        if part.device.startswith(path):    # finds the drive the path points to
            if wrong_filesystem(part.fstype):
                raise BadLocation("Wrong filesystem. Format the drive as FAT32 or use a different SD.")
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
    check_installer_updates()
    welcome()
    drive = select_drive()
    check_p_plus_updates(drive)
    extract_to_drive(drive)
    root = Tk()
    root.overrideredirect(True)
    root.withdraw()
    messagebox.showinfo("Complete",
                        f"Mod extracted; place SD in console and boot through {too_big_for_hackless_message(drive)}"
                        "homebrew channel. An NTSC Brawl disc or backup is required to play.")
    root.destroy()
    sys.exit()
