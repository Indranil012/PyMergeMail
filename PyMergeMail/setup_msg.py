import os
import sys
from email.message import EmailMessage
from PyMergeMail.color_print import color_print

async def setup_msg(cred: dict,
                    recv_data: dict,
                    subject: "Template",
                    body: "Template",
                    img_path_cid: dict = None,
                    attach_paths: list = None):
    """
        set email details
    """
    msg = EmailMessage()
    msg["From"] = f"{cred.get('alias')}<{cred.get('email')}>"
    msg["Subject"] = subject
    msg["To"] = f"{recv_data.get('name')}<{recv_data.get('email')}>"
    msg["cc"] = recv_data.get('cc')
    msg["Bcc"] = recv_data.get('bcc')
    msg.set_content("""
        This is a HTML mail please use supported client to render properly
                    """)
    msg.add_alternative(body, "html")

    # to embed img in mail
    for path, cid in img_path_cid.items():
        if path is not None:
            with open(path, "rb") as img:
                msg.get_payload()[1].add_related(img.read(),
                                                 "image", "png",
                                                 cid=cid)

    # for attachment
    if attach_paths is not None:
        for path in attach_paths:
            try:
                with open(path, "rb") as attach_file:
                    attachment = attach_file.read()

                msg.add_attachment(attachment,
                                   maintype="application",
                                   subtype="octet-stream",
                                   filename=os.path.basename(path))
            except Exception:
                color_print("wrong attachment path", 160)
                input_str = input("Do you want to send mail without attachment?\n"
                                  "enter 'y' to continue: ")
                if input_str != "y":
                    sys.exit()

    return msg
