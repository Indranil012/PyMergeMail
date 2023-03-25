import os
import sys
from email.message import EmailMessage
from PyMergeMail.get_template import template
from PyMergeMail.color_print import color_print

async def setup_msg(data_dict,
                    cred_dict,
                    subject_file_path=None,
                    body_file_path=None,
                    context=None,
                    img_path_cid=None,
                    attach_field=None):
    """
        set email details
    """
    subject = await template(subject_file_path).get_template()
    body = await template(body_file_path).get_template()
    body = body.render(context)

    msg = EmailMessage()
    msg["From"] = f"{cred_dict['alias']}<{cred_dict['email']}>"
    msg["Subject"] = subject.render(context)
    msg["To"] = f"{data_dict['name']}<{data_dict['email']}>"
    msg["cc"] = data_dict['cc']
    msg["Bcc"] = data_dict['bcc']

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
    if data_dict.get(attach_field) is not None:
        attach_paths = data_dict[attach_field].split(', ')
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

