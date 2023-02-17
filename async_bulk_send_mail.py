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

# def check_input(message):
#     """
#         loop until input key match
#     """
#     input_keys = ["y", "n"]
#     input_key = None
#     while input_key not in input_keys:
#         input_key = input(message).lower()
#         if input_key not in input_keys:
#             print("invalid input key")
#
#     return input_key

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

async def get_template_dir(path):
    """
    todo
    """
    return os.path.dirname(path)

async def get_file_name(path):
    """
    todo
    """
    return os.path.basename(path)
    # return os.path.splitext(fname_with_ext)[0]

async def get_env(template_directory):
    """
    todo
    """
    env = Environment(
        loader=FileSystemLoader(
            template_directory),
                      autoescape=select_autoescape())

    return env

async def get_context(row, 
                      subject_file_path,
                      body_file_path,
                      cid_fields,
                      ):
    """
        obtain required info from excel
    """
    template_directory = await get_template_dir(body_file_path)
    env = await get_env(template_directory)

    body_file_name = await get_file_name(body_file_path)
    body_source = env.loader.get_source(env, body_file_name)

    subject_file_name = await get_file_name(subject_file_path)
    subject_source = env.loader.get_source(env, subject_file_name)

    parsed_content = env.parse(body_source + subject_source)
    variables = meta.find_undeclared_variables(parsed_content)

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
                    cid_fields,
                    ):
    """
        set email details
    """
    context, img_path_cid = await get_context(row,
                                              subject_file_path,
                                              body_file_path,
                                              cid_fields)
    template_directory = await get_template_dir(body_file_path)
    env = await get_env(template_directory)

    body_file_name = await get_file_name(body_file_path)
    subject_file_name = await get_file_name(subject_file_path)

    subject = env.get_template(subject_file_name)
    body = env.get_template(body_file_name)

    msg = EmailMessage()
    msg["From"] = f"{cred_dict['alias']}<{cred_dict['email']}>"
    msg["Subject"] = subject.render(context)
    msg["To"] = f"{row['name']}<{row['email']}>"
    msg["cc"] = row['cc']
    msg["Bcc"] = row['bcc']

    msg.set_content("""
        This is a HTML mail please use supported client to render properly
                    """)

    body = body.render(context)
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
                  cid_fields,
                  ):
    """
       get address and msg dict as key, value
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
                              cid_fields,
                              )

        emails_msg[row['email']] = msg
        print(f"sending mail to {row['email']}")

    return emails_msg

async def login(cred_dict):
    """
        todo
    """
    smtp = aiosmtplib.SMTP(hostname="smtp.gmail.com",
                           username=cred_dict['email'],
                           password=cred_dict['pass'],
                      )
    await smtp.connect()
    print("loged in")

    return smtp

async def result(count_successful, count_unsuccessful):
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
               cid_fields,
               ):
    """
        main function
    """
    cred_dict = await cred(cred_file_path) # pass arg change=True to change cred.
    print(cred_dict)

    smtp = await login(cred_dict)

    emails_msg = await get_msg(data_file_path,
                               subject_file_path,
                               body_file_path,
                               cred_dict,
                               cid_fields,)
    # print(f"printing emails_msg dict\n{emails_msg}")
    count_successful = 0
    count_unsuccessful = 0
    for email, msg in emails_msg.items():
        # await send_mail(smtp, msg, email)
        try:
            await smtp.send_message(msg)
            print(f"Mail sent to {email}")
            count_successful += 1

        except Exception as identifier:
            print(f"Failed to send mail to {email}\n{identifier}")
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
