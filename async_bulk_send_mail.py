"""
this module can send bulk personalized mail with excel data

Author: Indranil012
version: 0.0.2
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

async def cred(cred_file_path, change=False):
    """
        print cred on terminal
    """
    with open(cred_file_path, encoding="UTF-8") as cred_file:
        cred_dict = json.load(cred_file)
    print(f"your current credentials: {cred_dict}")
    if change is True:
        input_email = input("Enter email: ")
        input_pass = input("Enter app password: ")
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

    final_obj = None
    if var_only is True:
        source = env.loader.get_source(env, file_name)
        parced_content = env.parse(source)
        final_obj = meta.find_undeclared_variables(parced_content)
    else:
        final_obj = env.get_template(file_name)

    return final_obj

async def get_context(row,
                      subject_file_path,
                      body_file_path,
                      cid_fields):
    """
        obtain required info from excel
    """
    body_var = await get_template(body_file_path, var_only=True)
    subject_var = await get_template(subject_file_path, var_only=True)
    variables = body_var | subject_var

    context = {}
    for variable in variables:
        if variable in row:
            context[variable] = row[variable]

    img_path_cid = {}
    for cid_field in cid_fields:
        path = row[cid_field]
        img_cid = make_msgid(cid_field)
        img_path_cid[path] = img_cid
        context[cid_field] = img_cid[1:-1]

    return context, img_path_cid

async def setup_msg(row,
                    subject_file_path,
                    body_file_path,
                    cred_dict,
                    cid_fields):
    """
        set email details
    """
    context, img_path_cid = await get_context(row,
                                              subject_file_path,
                                              body_file_path,
                                              cid_fields)
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
        with open(path, "rb") as img:
            msg.get_payload()[1].add_related(img.read(),
                                             "image", "png",
                                             cid=cid)

    # for attachment
    attach_paths = row['attachment'].split(', ')
    for path in attach_paths:
        with open(path, "rb") as attach_file:
            attachment = attach_file.read()

        msg.add_attachment(attachment,
                           maintype="application",
                           subtype="octet-stream",
                           filename=os.path.basename(path))

    return msg

async def get_msg(data_file_path,
                  subject_file_path,
                  body_file_path,
                  cred_dict,
                  cid_fields):
    """
       get address(str) : msg(obj) dict.
    """
    df = pd.read_excel(data_file_path)
    df.columns = df.columns.str.lower()
    print(f"\nPrinting source data...\n{df}")
    emails_msg = {}
    for _, row in df.iterrows():
        msg = await setup_msg(row,
                              subject_file_path,
                              body_file_path,
                              cred_dict,
                              cid_fields)

        emails_msg[row['email']] = msg
        print(f"sending mail to {row['email']}")

    return emails_msg

async def login(cred_dict):
    """
        todo
    """
    smtp = aiosmtplib.SMTP(hostname="smtp.gmail.com",
                           username=cred_dict['email'],
                           password=cred_dict['pass'])
    await smtp.connect()
    print("loged in")

    return smtp

async def result(count_successful, count_unsuccessful):
    """
    todo
    """
    total_email = count_successful + count_unsuccessful
    if count_unsuccessful == 0:
        print(f"{count_successful}/{total_email} Mail successfully sent ")
    else:
        print(f"{count_successful}/{total_email} Mail successfully sent and")
        print(f"{count_unsuccessful}/{total_email} Mail couldn't be sent ")

async def main(cred_file_path,
               data_file_path,
               subject_file_path,
               body_file_path,
               cid_fields):
    """
        main function
    """
    cred_dict = await cred(cred_file_path) # pass arg change=True to change cred.
    smtp = await login(cred_dict)

    emails_msg = await get_msg(data_file_path,
                               subject_file_path,
                               body_file_path,
                               cred_dict,
                               cid_fields)
    count_successful = 0
    count_unsuccessful = 0
    for email, msg in emails_msg.items():
        try:
            await smtp.send_message(msg)
            print(f"Mail sent to {email}")
            count_successful += 1

        except Exception as error:
            print(f"Failed to send mail to {email}\n"
                  f"{type(error)}: {error}")
            count_unsuccessful += 1

    await result(count_successful, count_unsuccessful)
    await smtp.quit()


if __name__ == "__main__":
    CRED_FILE_PATH = "key.json"
    DATA_FILE_PATH = "source_data.xlsx"
    SUBJECT_FILE_PATH = "subject.txt"
    BODY_FILE_PATH = "test.html"
    CID_FIELDS = ["img_path", "sig_path"]

    asyncio.run(main=main(CRED_FILE_PATH,
                          DATA_FILE_PATH,
                          SUBJECT_FILE_PATH,
                          BODY_FILE_PATH,
                          CID_FIELDS,))
