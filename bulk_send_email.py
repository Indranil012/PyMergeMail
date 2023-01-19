"""

this module can send bulk personalized mail with excel data

Author: Indranil012
Email: indranilchowdhury0@gmail.com
version: 6.4

"""

import os
import json
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid
from jinja2 import (Environment,
                    FileSystemLoader,
                    select_autoescape,
                    meta)
import pandas as pd


def check_input(message):
    """
        loop until input key match
    """
    input_keys = ["y", "n"]
    input_key = None
    while input_key not in input_keys:
        input_key = input(message).lower()
        if input_key not in input_keys:
            print("invalid input key")
    return input_key

def full_path(path):
    """
        provide full path from current dir
    """
    return f"{os.getcwd()}/{path}"

class PyMergeEmail:
    """
    main class
    """
    def __init__(self, cred_file_path,
                       data_file_path,
                       template_directory_path,
                       subject_file_name,
                       body_file_name,
                       cid_fields=None,
        ):

        self.cred_file_path = cred_file_path
        self.df = pd.read_excel(data_file_path)

        self.df.columns = self.df.columns.str.lower()
        print(f"\nPrinting source data...\n{self.df}")
        env = Environment(loader=FileSystemLoader(template_directory_path),
                          autoescape=select_autoescape()
              )
        self.subject = env.get_template(subject_file_name)
        self.body = env.get_template(body_file_name)
        self.cid_fields = cid_fields

        body_source = env.loader.get_source(env, body_file_name)
        subject_source = env.loader.get_source(env, subject_file_name)
        parsed_content = env.parse(body_source + subject_source)
        self.variables = meta.find_undeclared_variables(parsed_content)

        with open(self.cred_file_path, encoding="UTF-8") as cred_file:
            self.cred_dict = json.load(cred_file)

    def show_cred(self):
        """
            print cred on terminal
        """
        print(f"your current credentials: {self.cred_dict}")

    def change_cred(self):
        """
            to change cred on the go
        """
        if check_input("Do you want to change SMTP Keys? (y/n): ") == "y":
            input_email = input("Enter email: ")
            input_pass = input("Enter app password: ")
            with open(self.cred_file_path, "r+", encoding="UTF-8") as cred_file:
                self.cred_dict.update({"email":input_email, "pass":input_pass})
                cred_file.seek(0)
                json.dump(self.cred_dict, cred_file, indent=4)
                cred_file.truncate()
            print(f"New keys updated...\nupdated keys: {self.cred_dict}")

    def merge_context(self, row):
        """
            obtain required info from excel
        """
        context = {}
        for variable in self.variables:
            if variable in row:
                context[variable] = row[variable]

        img_path_cid = {}
        for cid_field in self.cid_fields:
            path = full_path(row[cid_field])
            img_cid = make_msgid(cid_field)
            img_path_cid[path] = img_cid
            context[cid_field] = img_cid[1:-1]

        return context, img_path_cid

    def email_details(self, row):
        """
            set email details
        """
        context, img_path_cid = self.merge_context(row)
        msg = EmailMessage()
        msg["From"] = f"{self.cred_dict['alias']}<{self.cred_dict['email']}>"
        msg["Subject"] = self.subject.render(context)
        msg["To"] = f"{row['name']}<{row['email']}>"
        msg.set_content(
            "This is a HTML mail please use supported client to render properly"
        )
        body = self.body.render(context)
        msg.add_alternative(body, "html")

        for path, cid in img_path_cid.items():
            with open(path, "rb") as img:
                msg.get_payload()[1].add_related(img.read(),
                                                 "image", "png",
                                                 cid=cid
                                     )

        # for attachment
        # with suppress(Exception):
        attach_paths = row['attachment'].split(', ')
        for path in attach_paths:
            file = full_path(path)
            with open(file, "rb") as attach_file:
                attachment = attach_file.read()
            msg.add_attachment( attachment,
                                maintype="application",
                                subtype="octet-stream",
                                filename=os.path.basename(file)
                )
        return msg

    def send_mail(self):
        """
            main function
        """
        self.show_cred()
        # self.change_cred()

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            #smtp.set_debuglevel(1)
            smtp.login(self.cred_dict["email"],
                       self.cred_dict["pass"])
            print("loged in")

            for _, row in self.df.iterrows():
                msg = self.email_details(row)
                print(f"sending mail to {row['email']}")
                smtp.send_message(msg,
                                  from_addr=self.cred_dict["email"],
                                  to_addrs=row["email"]
                     )
                print (f"mail sent to {row['email']}")

        print("all done!")

if __name__ == "__main__":
    CRED_FILE_PATH = full_path("key.json")
    DATA_FILE_PATH = full_path("example_data.xlsx")
    TEMPLATE_DIRECTORY_PATH = os.getcwd()  # Directory path
    SUBJECT_FILE_NAME = "subject.txt"
    BODY_FILE_NAME = "test.html"
    CID_FIELDS = ["img_path", "sig_path"]

    user1 = PyMergeEmail(CRED_FILE_PATH,
                         DATA_FILE_PATH,
                         TEMPLATE_DIRECTORY_PATH,
                         SUBJECT_FILE_NAME,
                         BODY_FILE_NAME,
                         CID_FIELDS,
            )
    user1.send_mail()
