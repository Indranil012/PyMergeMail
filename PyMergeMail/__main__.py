import smtplib
import pandas as pd
from PyMergeMail.cred import cred
from PyMergeMail.get_template import get_variables
from PyMergeMail.get_context import get_context
from PyMergeMail.setup_msg import setup_msg
from PyMergeMail.color_print import color_print, color
from PyMergeMail.result import result
from tqdm import tqdm
# from PyMergeMail.split import split

class Send:
    def __init__(self,
                 cred_file_path: str,
                 data_file_path: str,
                 subject_file_path: str = None,
                 body_file_path: str = None,
                 cid_fields: list = None,
                 attach_field: str = None):
        self.cred_dict = cred(cred_file_path)

        self.df = pd.read_excel(data_file_path)
        self.df.columns = self.df.columns.str.lower()
        print(f"\nPrinting source data...\n{self.df}")

        self.subject_file_path = subject_file_path
        self.body_file_path = body_file_path
        self.cid_fields = cid_fields
        self.attach_field = attach_field

        self.count_successful = 0
        self.count_unsuccessful = 0

    async def get_email_msg(self):
        variables = await get_variables(self.subject_file_path, self.body_file_path)
        email_msg = {}
        for _, row in self.df.iterrows():
            context, img_path_cid = await get_context(row, variables, self.cid_fields)
            msg = await setup_msg(row,
                                  self.cred_dict,
                                  self.subject_file_path,
                                  self.body_file_path,
                                  context,
                                  img_path_cid,
                                  self.attach_field)
            email_msg[row['email']] = msg
        return email_msg

    async def mail(self,):
        """
            main function
        """
        email_msg = await self.get_email_msg()

        with smtplib.SMTP_SSL("smtp.gmail.com") as smtp:
            # smtp.set_debuglevel(1)
            print("Trying to log in")
            smtp.login(user=self.cred_dict['email'], password=self.cred_dict['pass'])
            color_print("Log in successful", 157)

            for email, msg in tqdm(email_msg.items()):
                tqdm.write(f"Sending mail to {email}")
                try:
                    smtp.send_message(msg)
                    tqdm.write(color(f"Mail sent to {email}", 157))
                    self.count_successful += 1

                except Exception as error:
                    tqdm.write(color(f"Failed to send mail to {email}\n{type(error)}: {error}", 160))
                    self.count_unsuccessful += 1

        await result(self.count_successful, self.count_unsuccessful)
