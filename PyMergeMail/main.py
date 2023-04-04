import smtplib
from PyMergeMail.get_template import ToTemplate, get_variables
from PyMergeMail.get_context import Context
from PyMergeMail.setup_msg import setup_msg
from PyMergeMail.color_print import color_print, color
from PyMergeMail.result import result
from PyMergeMail.cred import get_cred
from tqdm import tqdm
from pandas import read_excel
# from PyMergeMail.split import split

async def mail(cred_file_path, 
               data_file_path = None, 
               subject_file_path = None, 
               body_file_path = None, 
               cid_fields = None, 
               attach_field = None):
    """main function
    """
    cred = get_cred(cred_file_path)
    subject_template = ToTemplate(subject_file_path).get_template() 
    body_template = ToTemplate(body_file_path).get_template() 
    variables = get_variables(subject_file_path, body_file_path)

    data_frame = read_excel(data_file_path)
    data_frame.columns = data_frame.columns.str.lower()
    print(f"\nPrinting source data...\n{data_frame}")

    email_msg = {}
    for _, recv_data in data_frame.iterrows():
        Context_ = Context(recv_data, variables, cid_fields)
        context = await Context_.get_context()
        img_cid =  await Context_.get_img_cid()
        subject = subject_template.render(context)
        body = body_template.render(context)
        
        if (attach_paths := recv_data.get(attach_field)) is not None:
            attach_paths = attach_paths.split(', ')

        msg = await setup_msg(cred, recv_data, subject, body, img_cid, attach_paths)
        email_msg[recv_data['email']] = msg

    count_successful = 0
    count_unsuccessful = 0

    with smtplib.SMTP_SSL("smtp.gmail.com") as smtp:
        # smtp.set_debuglevel(1)
        print("Trying to log in")
        smtp.login(user=cred['email'], password=cred['pass'])
        color_print("Log in successful", 157)

        for email, msg in tqdm(email_msg.items()):
            tqdm.write(f"Sending mail to {email}")
            try:
                smtp.send_message(msg)
                tqdm.write(color(f"Mail sent to {email}", 157))
                count_successful += 1

            except Exception as error:
                tqdm.write(color(f"Failed to send mail to {email}\n{type(error)}: {error}", 160))
                count_unsuccessful += 1

    await result(count_successful, count_unsuccessful)
