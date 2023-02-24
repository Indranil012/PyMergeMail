import smtplib
import pandas as pd
from PyMergeMail.cred import cred
from PyMergeMail.get_var import get_var
from PyMergeMail.get_context import get_context
from PyMergeMail.setup_msg import setup_msg
from PyMergeMail.color_print import color_print
from PyMergeMail.result import result

async def send(cred_file_path,
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
        color_print("Log in successful", 157)

        for email, msg in email_msg.items():
            print(f"Sending mail to {email}")
            try:
                smtp.send_message(msg)
                color_print(f"Mail sent to {email}", 157)
                count_successful += 1

            except Exception as error:
                color_print(f"Failed to send mail to {email}\n{type(error)}: {error}", 160)
                count_unsuccessful += 1

    await result(count_successful, count_unsuccessful)
