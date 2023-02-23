"""
this module can send bulk personalized mail with excel data

Author: Indranil012
version: 0.1.0
"""

import os
import sys
import asyncio
import json
from email.message import EmailMessage
from email.utils import make_msgid
import smtplib
from jinja2 import (Environment,
                    FileSystemLoader,
                    select_autoescape,
                    meta)
import pandas as pd

def color_print(text, color):
    """
    can print color text
    """
    return f"\x1b[38;5;{color}m{text}\x1b[m"

def check_blank(check_input, assign_with):
    """
    todo
    """
    if check_input=="":
        check_input = assign_with

async def cred(cred_file_path, change=False):
    """
    print cred on terminal
    """
    with open(cred_file_path, encoding="UTF-8") as cred_file:
        cred_dict = json.load(cred_file)
    print(f"your current credentials: {cred_dict}")
    if change is True:
        input_email = input("Enter email: ")
        check_blank(input_email, cred_dict['email'])
        input_pass = input("Enter app password: ")
        check_blank(input_pass, cred_dict['pass'])

        with open(cred_file_path, "r+", encoding="UTF-8") as cred_file:
            cred_dict.update({"email":input_email, "pass":input_pass})
            cred_file.seek(0)
            json.dump(cred_dict, cred_file, indent=4)
            cred_file.truncate()
        print(f"New keys updated...\nupdated keys: {cred_dict}")

    return cred_dict

async def get_template(file_path, var_only=False):
    """
    todo
    """
    template_dir = os.path.dirname(file_path)
    env = Environment(
        loader=FileSystemLoader(
            template_dir),
                      autoescape=select_autoescape())
    file_name = os.path.basename(file_path)

    if var_only is True:
        source = env.loader.get_source(env, file_name)
        parced_content = env.parse(source)
        return meta.find_undeclared_variables(parced_content)

    return env.get_template(file_name)

async def get_var(subject_file_path,
                  body_file_path):
    """
    todo
    """
    body_var = await get_template(body_file_path, var_only=True)
    subject_var = await get_template(subject_file_path, var_only=True)

    return body_var | subject_var

async def get_context(row,
                      variables,
                      cid_fields):
    """
        obtain required info from excel
    """
    context = {}
    for variable in variables:
        if variable in row:
            context[variable] = row[variable]

    img_path_cid = {}
    for cid_field in cid_fields:
        path = row.get(cid_field)
        img_cid = make_msgid(cid_field)
        img_path_cid[path] = img_cid
        context[cid_field] = img_cid[1:-1]

    return context, img_path_cid

async def setup_msg(row,
                    subject_file_path,
                    body_file_path,
                    cred_dict,
                    context,
                    img_path_cid,
                    attach_field):
    """
        set email details
    """
    subject = await get_template(subject_file_path)
    body = await get_template(body_file_path)
    body = body.render(context)

    msg = EmailMessage()
    msg["From"] = f"{cred_dict['alias']}<{cred_dict['email']}>"
    msg["Subject"] = subject.render(context)
    msg["To"] = f"{row['name']}<{row['email']}>"
    msg["cc"] = row['cc']
    msg["Bcc"] = row['bcc']

    msg.set_content("""
        This is a HTML mail please use supported client to render properly
                    """)
    msg.add_alternative(body, "html")

    for path, cid in img_path_cid.items():
        if path is not None:
            with open(path, "rb") as img:
                msg.get_payload()[1].add_related(img.read(),
                                                 "image", "png",
                                                 cid=cid)

    # for attachment
    if row.get(attach_field) is not None:
        attach_paths = row[attach_field].split(', ')
        for path in attach_paths:
            try:
                with open(path, "rb") as attach_file:
                    attachment = attach_file.read()

                msg.add_attachment(attachment,
                                   maintype="application",
                                   subtype="octet-stream",
                                   filename=os.path.basename(path))
            except Exception:
                print(color_print("wrong attachment path", 160))
                input_str = input("Do you want to send mail without attachment?\n"
                                  "enter 'y' to continue: ")
                if input_str != "y":
                    sys.exit()

    return msg

async def result(count_successful, count_unsuccessful):
    """
    todo
    """
    total_email = count_successful + count_unsuccessful
    if count_unsuccessful == 0:
        print(color_print(f"{count_successful}/{total_email} Mail successfully sent", 157))
        # 157 for green
    else:
        print(color_print(f"{count_successful}/{total_email} Mail successfully sent and", 157))
        # 157 for green
        print(color_print(f"{count_unsuccessful}/{total_email} Mail couldn't be sent", 160))
        # 160 for red

async def main(cred_file_path,
               data_file_path,
               subject_file_path,
               body_file_path,
               cid_fields,
               attach_field):
    """
        main function
    """
    cred_dict = await cred(cred_file_path)

    df = pd.read_excel(data_file_path)
    df.columns = df.columns.str.lower()
    print(f"\nPrinting source data...\n{df}")

    variables = await get_var(subject_file_path,
                              body_file_path)

    email_msg = {}
    for _, row in df.iterrows():
        context, img_path_cid = await get_context(row,
                                                  variables,
                                                  cid_fields)
        msg = await setup_msg(row,
                              subject_file_path,
                              body_file_path,
                              cred_dict,
                              context,
                              img_path_cid,
                              attach_field)
        email_msg[row['email']] = msg

    count_successful = 0
    count_unsuccessful = 0
    with smtplib.SMTP_SSL("smtp.gmail.com") as smtp:
        # smtp.set_debuglevel(1)
        print("Trying to log in")
        smtp.login(user=cred_dict['email'],
                   password=cred_dict['pass'])
        print(color_print("Log in successful", 157))

        for email, msg in email_msg.items():
            print(f"Sending mail to {email}")
            try:
                smtp.send_message(msg)
                print(color_print(f"Mail sent to {email}", 157))
                count_successful += 1

            except Exception as error:
                print(color_print(f"Failed to send mail to {email}\n{type(error)}: {error}", 160))
                count_unsuccessful += 1

    await result(count_successful, count_unsuccessful)


if __name__ == "__main__":
    CRED_FILE_PATH = "source_data/key.json"
    DATA_FILE_PATH = "source_data/source_data.xlsx"
    SUBJECT_FILE_PATH = "source_data/subject.txt"
    BODY_FILE_PATH = "source_data/test.html"
    CID_FIELDS = ["img_path", "sig_path"]
    ATTACH_FIELD = "attachment"

    asyncio.run(main=main(CRED_FILE_PATH,
                          DATA_FILE_PATH,
                          SUBJECT_FILE_PATH,
                          BODY_FILE_PATH,
                          CID_FIELDS,
                          ATTACH_FIELD))
