from typing import Any
from extract_msg import Message
from colorama import *

import extract_msg
import colorama
import os


class Mail:
    name: str
    s_name: str
    path: str
    msg: Message
    content: Any

    def __init__(self, filename, **kwargs):
        """
A .msg file taken from Outlook
        :param filename: Filename of the mail being loaded
        :type filename: str
        :param kwargs: Optional arguments
        """
        self.name = filename
        self.s_name = kwargs.get("short_name", "")
        self.category = kwargs.get("category")
        self.path = ""
        self.ext = ".msg"
        self.content = None

    def SetPath(self, mail_folder_path: str, **kwargs):
        """
Set the file path for the .msg file to allow the mail to be loaded and parsed. Updates class' file path.
        :param mail_folder_path: Path to the folder that the mail is located in
        :type mail_folder_path: str
        :param kwargs: "mail_name" - allows the name of the mail to be changed/respecified
        :return: Returns absolute path to the mail (optional)
        :rtype: str
        """
        print(f"\nSetting {self.s_name} mail path...")
        mail_name = str(kwargs.get("mail_name", self.name))
        if mail_name != self.name:
            self.name = mail_name

        if mail_name.endswith(".msg"):
            mail_name = mail_name[:-4]
        try:
            if not os.path.exists(mail_folder_path):
                os.mkdir(mail_folder_path)
        except Exception as e:
            print(f"Error encountered: {e}")

        self.path = os.path.join(mail_folder_path, mail_name + self.ext)
        self.LoadInstance()
        return self.path

    def LoadInstance(self):
        """
Calls the extract_msg module to load an instance into the class
        """
        print("Loading message instance into class...")
        self.msg = extract_msg.openMsg(self.path)

    def LoadContentFromMail(self):
        print(f"\nLoading content for {Fore.GREEN}{self.s_name}{Fore.WHITE}:")
        if self.category == "Something":
            

        elif self.category == "Something else":
           
        else:
            print("No content retrieval method specified!")
