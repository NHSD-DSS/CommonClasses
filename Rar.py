import os
import rarfile  # extracts rar files
from colorama import *


class RarArchive:
    def __init__(self, name: str):
        """
RarArchive class - holds data for a RAR archive
        :param name: Name of RAR Archive
        :type name: str
        """
        if not name.endswith(".rar"):
            name += ".rar"

        self.file_name = name
        self.file_path = ""
        self.file = None

    def SetPath(self, folder: str) -> bool:
        """
Sets the path variable for an instance of RarArchive.
        :param folder: The folder that the archive is in. This is paired with the name to build the path.
        :type folder: str
        :return: Optionally returns a bool when path is valid/invalid
        :rtype: bool
        """
        full_path = os.path.join(folder, self.file_name)
        print(f"{Fore.LIGHTBLUE_EX}Looking for: {self.file_name} \nin \n{folder}\n")
        if not os.path.isfile(full_path):
            print(f"{Fore.RED}Cannot find file! Exiting...")
            return False
        print(f"{Fore.GREEN}File found -- path successfully updated to:\n{full_path}\n")
        self.file_path = full_path
        self.LoadInstance()
        return True

    def LoadInstance(self):
        """
Calls the rarfile module to load an instance of a rar file into the class instance
        """
        rarfile.UNRAR_TOOL = "unrar"
        self.file = rarfile.RarFile(self.file_path, mode="r")
