"""
this module can send bulk personalized mail with excel data

Author: Indranil012
version: 7.0
"""

import os
import asyncio
import json
from email.message import EmailMessage
from email.utils import make_msgid
import aiosmtplib
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
                       cid_fields=None):
        self.cred_file_path = cred_file_path
        self.cid_fields = cid_fields

        self.df = pd.read_excel(data_file_path)
        self.df.columns = self.df.columns.str.lower()
        print(f"\nPrinting source data...\n{self.df}")

        env = Environment(loader=FileSystemLoader(template_directory_path),
                          autoescape=select_autoescape())
        self.subject = env.get_template(subject_file_name)
        self.body = env.get_template(body_file_name)

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

    async def change_cred(self):
        """
            to change cred on the go
        """
        if check_input("Do you want to change self.smtp Keys? (y/n): ") == "y":
            input_email = input("Enter email: ")
            input_pass = input("Enter app password: ")
            with open(self.cred_file_path, "r+", encoding="UTF-8") as cred_file:
                self.cred_dict.update({"email":input_email, "pass":input_pass})
                cred_file.seek(0)
                json.dump(self.cred_dict, cred_file, indent=4)
                cred_file.truncate()
            print(f"New keys updated...\nupdated keys: {self.cred_dict}")

    async def get_context(self, row):
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

    async def setup_msg(self, row):
        """
            set email details
        """
        context, img_path_cid = await self.get_context(row)

        msg = EmailMessage()
        msg["From"] = f"{self.cred_dict['alias']}<{self.cred_dict['email']}>"
        msg["Subject"] = self.subject.render(context)
        msg["To"] = f"{row['name']}<{row['email']}>"
        msg["Cc"] = row['cc']
        msg["Bcc"] = row['bcc']

        msg.set_content("""
            This is a HTML mail please use supported client to render properly
                        """)

        body = self.body.render(context)
        msg.add_alternative(body, "html")

        for path, cid in img_path_cid.items():
            with open(path, "rb") as img:
                msg.get_payload()[1].add_related(img.read(),
                                                 "image", "png",
                                                 cid=cid)

        # for attachment
        attach_paths = row['attachment'].split(', ')
        for path in attach_paths:
            file = full_path(path)
            with open(file, "rb") as attach_file:
                attachment = attach_file.read()

            msg.add_attachment(attachment,
                               maintype="application",
                               subtype="octet-stream",
                               filename=os.path.basename(file))
        # print(hash(msg))

        return msg

    async def get_msg(self):
        """
           get address and msg dict as key, value
        """
        emails_msg = {}
        for _, row in self.df.iterrows():
            msg = await self.setup_msg(row)
            emails_msg[row['email']] = msg
            print(f"sending mail to {row['email']}")

        return emails_msg

    async def login(self):
        """
            todo
        """
        smtp = aiosmtplib.SMTP(hostname="smtp.gmail.com",
                               username=self.cred_dict['email'],
                               password=self.cred_dict['pass'],
                               )
        await smtp.connect()
        print("loged in")

        return smtp

    async def send_mail(self, smtp, msg, email):
        """
            todo
        """
        try:
            await smtp.send_message(msg,
                                    self.cred_dict['email'],
                                    email,)
            print(f"Mail sent to {email}")

        except Exception as identifier:
            print(f"Failed to send mail to {email}\n{identifier}")

    async def main(self):
        """
            main function
        """
        self.show_cred()
        # await self.change_cred()
        smtp = await self.login()
        emails_msg = await self.get_msg()
        # print(f"printing emails_msg dict\n{emails_msg}")
        for email, msg in emails_msg.items():
            await self.send_mail(smtp, msg, email)

        print("all done!")
        await smtp.quit()


if __name__ == "__main__":
    CRED_FILE_PATH = full_path("key.json")
    DATA_FILE_PATH = full_path("source_data.xlsx")
    TEMPLATE_DIRECTORY_PATH = os.getcwd() # body, subject directory path
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
    # loop = asyncio.get_event_loop()
    # loop.run_until_complete(user1.main())
    asyncio.run(main=user1.main(), debug=True)
