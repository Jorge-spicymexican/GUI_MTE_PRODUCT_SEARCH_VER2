# from win32com.client import Dispatch
from plyer import notification

"""
File Imported display
"""


def FileInsertedNotfi(filename):
    notification.notify(
        title="File Imported",
        message=filename,
        app_icon='./assets/tci_logo_Csx_icon.ico',
        timeout=6,
        toast=True
    )


"""
File Could not be imparted 
"""


def FileInsertErrorNotfi():
    notification.notify(
        title="File Error",
        message="Selected File could not be imported please try again.",
        app_icon='./assets/tci_logo_Csx_icon.ico',
        timeout=6,
        toast=True
    )


"""
File could not be opened
"""


def filenotopened():
    notification.notify(
        title="File Error",
        message="File Could not be opened",
        app_icon='./assets/tci_logo_Csx_icon.ico',
        timeout=6,
        toast=True
    )


"""
Excel Import error
"""


def ImportError():
    notification.notify(
        title="Excel  Error",
        message="File could not be imported",
        app_icon='./assets/tci_logo_Csx_icon.ico',
        timeout=6,
        toast=True
    )


"""
Excel Read error
"""


def ExcelReadError():
    notification.notify(
        title="Excel Error",
        message="File Could not be Read Correctly",
        app_icon='./assets/tci_logo_Csx_icon.ico',
        timeout=6,
        toast=True
    )


"""
Excel not found 
"""


def NotFoundError():
    notification.notify(
        title="Excel Error",
        message="File Could not be Found",
        app_icon='./assets/tci_logo_Csx_icon.ico',
        timeout=6,
        toast=True
    )


"""
Please insert an excel file
"""


def InsertExcel():
    notification.notify(
        title="Excel Not Inserted",
        message="Please Insert an Excel File",
        app_icon='./assets/tci_logo_Csx_icon.ico',
        timeout=6,
        toast=True
    )


"""
Excel read only mode
"""


def ExcelReadonlyError():
    notification.notify(
        title="Excel Error",
        message="File Could not be Read due to Read Only Mode",
        app_icon='./assets/tci_logo_Csx_icon.ico',
        timeout=6,
        toast=True
    )


"""
This function checks the chrome version path that the user has and also checks if the a chrome driver already exist 
in his application. 
"""


def Chrome_path_check():
    paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
             r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
    version = list(filter(None, [get_version_via_com(p) for p in paths]))[0]
    return version


"""
This functions returns and gets the file version for chrome.
"""


def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        return None
    return version


if __name__ == "__main__":
    print("FILE SHOULD BE RUN AS MAIN.\n")
