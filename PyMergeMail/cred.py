import json
from PyMergeMail.check_blank import check_blank
def cred(cred_file_path, change=False):
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
